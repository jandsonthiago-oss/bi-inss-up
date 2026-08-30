from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import tempfile
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

import pandas as pd
import pyarrow.parquet as pq

from main import BASE_FOLDER, Graph


APP_VERSION = "1.0.0"
POWERBI_FOLDER = os.getenv("POWERBI_FOLDER", f"{BASE_FOLDER}/powerbi")
POWERBI_CONTROL_FOLDER = os.getenv(
    "POWERBI_CONTROL_FOLDER",
    f"{BASE_FOLDER}/_powerbi_controle",
)
ROWS_PER_CSV = int(os.getenv("POWERBI_ROWS_PER_CSV", "150000"))
DOWNLOAD_CHUNK = 8 * 1024 * 1024
MIN_FREE_DISK = 1536 * 1024 * 1024

CATEGORIES = (
    "concedidos",
    "emitidos",
    "indeferidos",
    "mantidos",
    "cat",
    "solicitados",
    "pendentes",
    "unidades",
    "folha_emitidos",
)

HEADER_KEYWORDS = {
    "competencia",
    "mes",
    "ano",
    "uf",
    "municipio",
    "beneficio",
    "especie",
    "sexo",
    "idade",
    "cid",
    "cnae",
    "cbo",
    "despacho",
    "clientela",
    "filiacao",
    "ramo",
    "atividade",
    "quantidade",
    "valor",
    "data",
    "codigo",
    "nome",
    "unidade",
    "requerimento",
    "protocolo",
    "situacao",
    "status",
    "acidente",
    "agencia",
    "residencia",
}


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def norm(value: Any) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", text.strip().lower())


def column_name(value: Any, fallback: str) -> str:
    text = norm(value).replace("%", " percentual ")
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text or fallback


def unique_columns(values: list[Any]) -> list[str]:
    seen: dict[str, int] = {}
    output: list[str] = []

    for index, value in enumerate(values, start=1):
        base = column_name(value, f"coluna_{index}")
        seen[base] = seen.get(base, 0) + 1
        output.append(base if seen[base] == 1 else f"{base}_{seen[base]}")

    return output


def safe_name(value: Any, limit: int = 100) -> str:
    text = column_name(value, "arquivo").replace("_", "-")
    return text[:limit]


def check_disk(path: Path, required_extra: int = 0) -> None:
    free = shutil.disk_usage(path).free
    if free - required_extra < MIN_FREE_DISK:
        raise RuntimeError(
            f"Espaco em disco insuficiente. Livre={free / 1024**3:.2f} GiB"
        )


def source_control_path(category: str) -> str:
    return f"{BASE_FOLDER}/_controle/{category}.json"


def powerbi_control_path(category: str) -> str:
    return f"{POWERBI_CONTROL_FOLDER}/{category}.json"


def load_json(graph: Graph, path: str) -> dict[str, Any]:
    payload = graph.read_json(path)
    if not payload:
        return {}
    if not isinstance(payload, dict):
        raise RuntimeError(f"JSON invalido: {path}")
    return payload


def save_powerbi_control(
    graph: Graph,
    category: str,
    control: dict[str, Any],
) -> None:
    control["app_version"] = APP_VERSION
    control["updated_at_utc"] = now_iso()
    graph.upload_bytes(
        json.dumps(control, ensure_ascii=False, indent=2, sort_keys=True).encode("utf-8"),
        powerbi_control_path(category),
        "application/json; charset=utf-8",
    )


def delete_remote(graph: Graph, remote_path: str) -> None:
    response = graph.get_item_by_path(remote_path)
    if response.status_code == 404:
        return
    if response.status_code != 200:
        raise RuntimeError(
            f"Falha ao consultar {remote_path}: HTTP {response.status_code}"
        )

    item_id = response.json()["id"]
    graph.request(
        "DELETE",
        f"/drives/{graph.drive_id}/items/{item_id}",
        expected=(204,),
    )


def download_remote(graph: Graph, remote_path: str, target: Path) -> None:
    encoded = quote(remote_path.strip("/"), safe="/")
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{graph.drive_id}"
        f"/root:/{encoded}:/content"
    )

    for auth_attempt in range(2):
        with graph.session.get(
            url,
            headers={"Authorization": f"Bearer {graph.token}"},
            stream=True,
            allow_redirects=True,
            timeout=(30, 600),
        ) as response:
            if response.status_code == 401 and auth_attempt == 0:
                graph.refresh_token()
                continue
            if response.status_code != 200:
                raise RuntimeError(
                    f"Download SharePoint HTTP {response.status_code}: "
                    f"{remote_path} | {response.text[:500]}"
                )

            with target.open("wb") as handle:
                for chunk in response.iter_content(DOWNLOAD_CHUNK):
                    if chunk:
                        handle.write(chunk)
                        check_disk(target.parent)
            break
    else:
        raise RuntimeError(f"Falha de autenticacao ao baixar {remote_path}")

    if not target.exists() or target.stat().st_size == 0:
        raise RuntimeError(f"Arquivo SharePoint vazio: {remote_path}")


def schema_is_bad(names: list[str]) -> bool:
    normalized = [norm(name) for name in names]
    generic = sum(
        bool(re.fullmatch(r"(?:coluna|column)_?\d+", name))
        for name in normalized
    )
    first = normalized[0] if normalized else ""
    title_like = any(
        token in first
        for token in (
            "dados abertos",
            "beneficios concedidos",
            "beneficios mantidos",
            "beneficios emitidos",
            "beneficios indeferidos",
            "requerimentos",
            "comunicacoes de acidente",
            "perfil das unidades",
            "folha de pagamento",
        )
    )
    return generic >= max(2, len(names) // 4) or (title_like and generic >= 1)


def looks_numeric(value: str) -> bool:
    text = value.strip()
    if not text:
        return False
    text = text.replace(".", "").replace(",", ".")
    return bool(re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text))


def header_row_score(values: list[Any]) -> tuple[int, int, int, int]:
    cells = ["" if value is None else str(value).strip() for value in values]
    nonempty = [cell for cell in cells if cell]
    if not nonempty:
        return (-1, -1, -1, -1)

    normalized = [norm(cell) for cell in nonempty]
    keyword_hits = sum(
        1
        for cell in normalized
        if any(keyword in cell for keyword in HEADER_KEYWORDS)
    )
    text_like = sum(
        1
        for cell in nonempty
        if re.search(r"[A-Za-zÀ-ÿ]", cell)
        and not looks_numeric(cell)
        and len(cell) <= 120
    )
    numeric_like = sum(1 for cell in nonempty if looks_numeric(cell))
    unique_count = len(set(normalized))

    return (
        keyword_hits,
        text_like - numeric_like,
        unique_count,
        len(nonempty),
    )


def detect_headers(parquet_file: pq.ParquetFile) -> tuple[list[str], int]:
    original = list(parquet_file.schema_arrow.names)
    if not schema_is_bad(original):
        return unique_columns(original), 0

    first_batch = next(
        parquet_file.iter_batches(batch_size=40),
        None,
    )
    if first_batch is None or first_batch.num_rows == 0:
        return unique_columns(original), 0

    preview = first_batch.to_pandas()
    required_nonempty = max(2, min(len(original), max(3, len(original) // 3)))

    best_index = 0
    best_score = (-1, -1, -1, -1)

    for index in range(min(20, len(preview))):
        values = preview.iloc[index].tolist()
        nonempty_count = sum(
            value is not None and str(value).strip() != ""
            for value in values
        )
        if nonempty_count < required_nonempty:
            continue

        score = header_row_score(values)
        if score > best_score:
            best_score = score
            best_index = index

    row = preview.iloc[best_index].tolist()
    promoted: list[Any] = []
    for index, value in enumerate(row):
        if value is None or str(value).strip() == "":
            promoted.append(original[index] if index < len(original) else f"coluna_{index + 1}")
        else:
            promoted.append(value)

    headers = unique_columns(promoted)
    return headers, best_index + 1


def resource_year(resource: dict[str, Any], input_path: str) -> str:
    match = re.search(r"/silver/[^/]+/(20\d{2})/", input_path)
    if match:
        return match.group(1)

    years = resource.get("years") or []
    numeric_years = [int(year) for year in years if str(year).isdigit()]
    if numeric_years:
        return str(max(numeric_years))

    return "sem_ano"


def input_signature(resource: dict[str, Any]) -> str:
    payload = {
        "version": resource.get("version"),
        "source_sha256": resource.get("source_sha256"),
        "outputs": resource.get("outputs"),
    }
    return hashlib.sha256(
        json.dumps(payload, sort_keys=True, ensure_ascii=False).encode("utf-8")
    ).hexdigest()


def clean_dataframe(
    dataframe: pd.DataFrame,
    headers: list[str],
    *,
    category: str,
    year: str,
    resource_id: str,
    resource_name: str,
    input_path: str,
) -> pd.DataFrame:
    dataframe = dataframe.copy()
    dataframe.columns = headers

    dataframe = dataframe.dropna(how="all")
    if dataframe.empty:
        return dataframe

    for column in dataframe.columns:
        dataframe[column] = dataframe[column].map(
            lambda value: "" if pd.isna(value) else str(value).strip()
        )

    # Remove linhas completamente vazias depois da conversao para texto.
    mask_nonempty = dataframe.apply(
        lambda row: any(str(value).strip() for value in row),
        axis=1,
    )
    dataframe = dataframe.loc[mask_nonempty]
    if dataframe.empty:
        return dataframe

    # Remove cabecalho repetido dentro dos dados, quando existir.
    comparison_headers = [norm(header) for header in headers]

    def is_repeated_header(row: pd.Series) -> bool:
        values = [norm(value) for value in row.tolist()]
        comparable = min(len(values), len(comparison_headers))
        if comparable == 0:
            return False
        matches = sum(
            1
            for index in range(comparable)
            if values[index] and values[index] == comparison_headers[index]
        )
        return matches >= max(3, comparable // 2)

    if len(dataframe) <= 250_000:
        repeated = dataframe.apply(is_repeated_header, axis=1)
        dataframe = dataframe.loc[~repeated]

    dataframe["fonte_categoria"] = category
    dataframe["fonte_ano"] = year
    dataframe["fonte_recurso_id"] = resource_id
    dataframe["fonte_recurso"] = resource_name
    dataframe["fonte_arquivo_silver"] = input_path

    return dataframe


def convert_one_parquet(
    graph: Graph,
    *,
    category: str,
    resource_id: str,
    resource: dict[str, Any],
    input_path: str,
    output_index: int,
    local_parquet: Path,
    workdir: Path,
) -> tuple[list[str], list[str], int]:
    parquet_file = pq.ParquetFile(local_parquet)
    headers, rows_to_skip = detect_headers(parquet_file)
    year = resource_year(resource, input_path)
    resource_name = str(resource.get("resource_name") or resource_id)

    print(
        f"CABECALHO | {input_path} | colunas={len(headers)} | "
        f"linhas_iniciais_removidas={rows_to_skip}"
    )

    remote_folder = (
        f"{POWERBI_FOLDER}/{category}/{year}/"
        f"{safe_name(resource_id, 70)}/{output_index:02d}"
    )

    uploaded_paths: list[str] = []
    total_rows = 0
    global_position = 0
    part_number = 1

    for batch in parquet_file.iter_batches(batch_size=ROWS_PER_CSV):
        dataframe = batch.to_pandas()
        batch_start = global_position
        global_position += len(dataframe)

        if rows_to_skip > batch_start:
            skip_here = min(rows_to_skip - batch_start, len(dataframe))
            dataframe = dataframe.iloc[skip_here:]

        if dataframe.empty:
            continue

        if len(dataframe.columns) != len(headers):
            raise RuntimeError(
                f"Quantidade de colunas mudou em {input_path}: "
                f"esperado={len(headers)} recebido={len(dataframe.columns)}"
            )

        dataframe = clean_dataframe(
            dataframe,
            headers,
            category=category,
            year=year,
            resource_id=resource_id,
            resource_name=resource_name,
            input_path=input_path,
        )
        if dataframe.empty:
            continue

        local_csv = workdir / f"part-{part_number:05d}.csv"
        dataframe.to_csv(
            local_csv,
            index=False,
            encoding="utf-8-sig",
            sep=",",
            quoting=csv.QUOTE_MINIMAL,
            lineterminator="\n",
        )

        remote_csv = f"{remote_folder}/{local_csv.name}"
        graph.upload_file(local_csv, remote_csv)
        uploaded_paths.append(remote_csv)
        total_rows += len(dataframe)

        print(
            f"CSV {part_number:05d} OK | {len(dataframe):,} linhas | "
            f"{local_csv.stat().st_size / 1024**2:,.1f} MB"
        )

        local_csv.unlink(missing_ok=True)
        part_number += 1
        check_disk(workdir)

    if not uploaded_paths:
        raise RuntimeError(f"Nenhuma linha de dados produzida para {input_path}")

    return headers, uploaded_paths, total_rows


def process_resource(
    graph: Graph,
    *,
    category: str,
    resource_id: str,
    resource: dict[str, Any],
) -> dict[str, Any]:
    outputs = resource.get("outputs") or []
    if not isinstance(outputs, list) or not outputs:
        raise RuntimeError(f"Recurso sem outputs silver: {resource_id}")

    year = resource_year(resource, str(outputs[0].get("path") or ""))
    resource_folder = (
        f"{POWERBI_FOLDER}/{category}/{year}/{safe_name(resource_id, 70)}"
    )

    # Substituicao atomica por recurso: remove a versao Power BI anterior
    # somente quando o mesmo recurso precisa ser reconstruido.
    delete_remote(graph, resource_folder)

    all_columns: list[str] = []
    all_remote_outputs: list[str] = []
    total_rows = 0

    with tempfile.TemporaryDirectory(prefix=f"powerbi_{category}_") as tmp:
        workdir = Path(tmp)

        for output_index, output in enumerate(outputs, start=1):
            input_path = str(output.get("path") or "")
            if not input_path.lower().endswith(".parquet"):
                continue

            print("\n" + "=" * 88)
            print(f"{category.upper()} | {resource.get('resource_name') or resource_id}")
            print(f"SILVER: {input_path}")
            print("=" * 88)

            local_parquet = workdir / f"input-{output_index:02d}.parquet"
            download_remote(graph, input_path, local_parquet)
            print(
                f"DOWNLOAD OK | {local_parquet.stat().st_size / 1024**2:,.1f} MB"
            )

            headers, uploaded, rows = convert_one_parquet(
                graph,
                category=category,
                resource_id=resource_id,
                resource=resource,
                input_path=input_path,
                output_index=output_index,
                local_parquet=local_parquet,
                workdir=workdir,
            )

            all_columns.extend(headers)
            all_remote_outputs.extend(uploaded)
            total_rows += rows
            local_parquet.unlink(missing_ok=True)

    return {
        "status": "success",
        "source_version": resource.get("version"),
        "source_sha256": resource.get("source_sha256"),
        "signature": input_signature(resource),
        "resource_name": resource.get("resource_name"),
        "processed_at_utc": now_iso(),
        "rows": total_rows,
        "columns": sorted(set(all_columns)),
        "outputs": all_remote_outputs,
    }


def build_category(
    graph: Graph,
    category: str,
    max_inputs: int,
    force: bool,
) -> tuple[int, int, int]:
    source_control = load_json(graph, source_control_path(category))
    resources = source_control.get("resources") or {}
    if not isinstance(resources, dict):
        raise RuntimeError(f"Controle silver invalido para {category}")

    target_control = load_json(graph, powerbi_control_path(category))
    target_resources = target_control.setdefault("resources", {})
    if not isinstance(target_resources, dict):
        target_resources = {}
        target_control["resources"] = target_resources

    candidates: list[tuple[str, dict[str, Any]]] = []

    for resource_id, resource in resources.items():
        if not isinstance(resource, dict):
            continue
        outputs = resource.get("outputs") or []
        if not outputs:
            continue

        current_signature = input_signature(resource)
        previous = target_resources.get(resource_id)
        completed = (
            isinstance(previous, dict)
            and previous.get("status") == "success"
            and previous.get("signature") == current_signature
        )

        if force or not completed:
            candidates.append((str(resource_id), resource))

    candidates.sort(
        key=lambda pair: str(pair[1].get("processed_at_utc") or ""),
        reverse=True,
    )

    if max_inputs > 0:
        candidates = candidates[:max_inputs]

    print(
        f"[{category}] silver={len(resources)} | "
        f"pendentes_powerbi={len(candidates)}"
    )

    success = 0
    errors = 0
    skipped = max(0, len(resources) - len(candidates))

    for resource_id, resource in candidates:
        try:
            target_resources[resource_id] = process_resource(
                graph,
                category=category,
                resource_id=resource_id,
                resource=resource,
            )
            save_powerbi_control(graph, category, target_control)
            success += 1
        except Exception as exc:
            errors += 1
            print("\n" + "!" * 88)
            print(f"ERRO POWERBI | {category} | {resource_id}")
            print(f"{type(exc).__name__}: {exc}")
            print("!" * 88)

    return success, errors, skipped


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Gera camada CSV limpa para Power BI a partir do silver Parquet."
    )
    parser.add_argument(
        "--dataset",
        choices=("all", *CATEGORIES),
        default=os.getenv("POWERBI_DATASET", "all"),
    )
    parser.add_argument(
        "--max-inputs",
        type=int,
        default=int(os.getenv("POWERBI_MAX_INPUTS", "0")),
        help="0 = todos os recursos silver pendentes da categoria.",
    )
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()

    if args.max_inputs < 0:
        raise SystemExit("--max-inputs nao pode ser negativo")

    categories = CATEGORIES if args.dataset == "all" else (args.dataset,)

    print("=" * 88)
    print("UNIVERSO PREVIDENCIARIO - CAMADA POWER BI")
    print(f"APP_VERSION: {APP_VERSION}")
    print(f"Destino: {POWERBI_FOLDER}")
    print("Formato: CSV UTF-8 padronizado, particionado e auditavel")
    print("=" * 88)

    graph = Graph()
    graph.ensure_folder(POWERBI_FOLDER)
    graph.ensure_folder(POWERBI_CONTROL_FOLDER)
    print("SharePoint conectado.")

    total_success = 0
    total_errors = 0
    total_skipped = 0

    for category in categories:
        success, errors, skipped = build_category(
            graph,
            category,
            args.max_inputs,
            args.force,
        )
        total_success += success
        total_errors += errors
        total_skipped += skipped

    print("\n" + "=" * 88)
    print(f"SUCESSO: {total_success}")
    print(f"ERROS: {total_errors}")
    print(f"JA CONCLUIDOS: {total_skipped}")
    print("=" * 88)

    return 1 if total_errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
