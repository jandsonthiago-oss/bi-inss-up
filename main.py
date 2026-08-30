from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import time
import unicodedata
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlparse

import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq
import requests
from openpyxl import load_workbook
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ============================================================
# BI INSS - UNIVERSO PREVIDENCIARIO
# INSS 2025+ -> Parquet -> SharePoint -> Power BI
# ============================================================

CKAN_BASE = "https://dadosabertos.inss.gov.br"
CKAN_ORG = "instituto-nacional-de-seguro-social-inss"
ANO_INICIAL = int(os.getenv("ANO_INICIAL", "2025"))

SITE_ID = os.getenv(
    "SHAREPOINT_SITE_ID",
    "universoprevidenciario.sharepoint.com,"
    "b4ca13e8-7517-4e58-bb82-4cb8f8e06866,"
    "c2bf221a-26c3-4ff9-8dd3-648cb5cc1552",
)
BASE_FOLDER = os.getenv("SHAREPOINT_BASE_FOLDER", "INSS")

CSV_CHUNK_ROWS = 100_000
DOWNLOAD_CHUNK = 8 * 1024 * 1024
UPLOAD_CHUNK = 10 * 1024 * 1024   # multiplo de 320 KiB
SIMPLE_UPLOAD_LIMIT = 240 * 1024 * 1024

RULES = {
    "concedidos": "beneficios concedidos",
    "emitidos": "beneficios emitidos",
    "indeferidos": "beneficios indeferidos",
    "mantidos": "beneficios mantidos",
    "cat": "comunicacoes de acidente de trabalho",
    "solicitados": "dados quantitativos de requerimentos administrativos solicitados",
    "pendentes": "dados quantitativos de requerimentos administrativos pendentes",
    "unidades": "perfil das unidades",
    "folha_emitidos": "dados agregados da folha de pagamento em relacao aos beneficios emitidos",
}


def norm(value: Any) -> str:
    s = "" if value is None else str(value)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s.lower().strip())


def safe_name(value: Any, limit: int = 100) -> str:
    s = re.sub(r"[^a-z0-9._-]+", "-", norm(value))
    return (re.sub(r"-+", "-", s).strip("-._") or "arquivo")[:limit]


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def years_of(resource: dict[str, Any]) -> list[int]:
    # O ano deve vir do proprio recurso, nunca do titulo/faixa do dataset.
    # Ordem: nome do recurso -> nome do arquivo na URL -> descricao.
    candidates = [
        str(resource.get("name") or ""),
        Path(urlparse(str(resource.get("url") or "")).path).name,
        str(resource.get("description") or ""),
    ]
    for text in candidates:
        years = sorted({int(x) for x in re.findall(r"\b(20\d{2})\b", text)})
        if years:
            return years
    return []


def resource_is_valid(resource: dict[str, Any]) -> bool:
    years = years_of(resource)
    return bool(years) and max(years) >= ANO_INICIAL


def resource_version(resource: dict[str, Any]) -> str:
    return str(
        resource.get("last_modified")
        or resource.get("created")
        or resource.get("url")
        or resource.get("id")
        or ""
    )


def build_session() -> requests.Session:
    retry = Retry(
        total=5,
        backoff_factor=1,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET", "HEAD", "POST", "PUT", "DELETE"}),
        respect_retry_after_header=True,
    )
    s = requests.Session()
    s.mount("https://", HTTPAdapter(max_retries=retry))
    s.headers.update({"User-Agent": "UniversoPrevidenciario-BI-INSS/3.0"})
    return s


HTTP = build_session()


def ckan(action: str, params: dict[str, Any]) -> Any:
    r = HTTP.get(
        f"{CKAN_BASE}/api/3/action/{action}",
        params=params,
        timeout=(30, 180),
    )
    r.raise_for_status()
    payload = r.json()
    if not payload.get("success"):
        raise RuntimeError(f"CKAN success=false: {action}")
    return payload["result"]


def load_catalog() -> list[dict[str, Any]]:
    result = ckan(
        "package_search",
        {"q": f"organization:{CKAN_ORG}", "rows": 1000, "start": 0},
    )
    return list(result.get("results", []))


def dataset_matches(category: str, title: str) -> bool:
    t = norm(title)
    rule = RULES[category]

    if category in {"concedidos", "emitidos", "indeferidos", "mantidos", "cat"}:
        return t.startswith(rule)

    return rule in t


def choose_dataset(
    packages: list[dict[str, Any]],
    category: str,
) -> dict[str, Any]:
    candidates = [
        p for p in packages
        if dataset_matches(category, str(p.get("title") or ""))
    ]
    if not candidates:
        raise RuntimeError(f"Dataset nao encontrado: {category}")

    def score(p: dict[str, Any]) -> tuple[int, int, str]:
        resources = p.get("resources") or []
        valid_count = sum(resource_is_valid(r) for r in resources)
        current_plan = int("plano de dados abertos" in norm(p.get("title")))
        return valid_count, current_plan, str(p.get("metadata_modified") or "")

    return max(candidates, key=score)


def list_resources(dataset: dict[str, Any]) -> list[dict[str, Any]]:
    resources = [r for r in (dataset.get("resources") or []) if resource_is_valid(r)]
    resources.sort(
        key=lambda r: (
            max(years_of(r) or [0]),
            str(r.get("last_modified") or r.get("created") or ""),
            str(r.get("name") or ""),
        ),
        reverse=True,
    )
    return resources


class Graph:
    def __init__(self) -> None:
        self.session = build_session()
        self.token = ""
        self.drive_id = ""
        self.root_id = ""
        self.refresh_token()
        self.load_drive()

    def refresh_token(self) -> None:
        result = subprocess.run(
            [
                "az", "account", "get-access-token",
                "--resource-type", "ms-graph",
                "--query", "accessToken",
                "--output", "tsv",
            ],
            check=True,
            capture_output=True,
            text=True,
        )
        self.token = result.stdout.strip()
        if not self.token:
            raise RuntimeError("Token Microsoft Graph vazio.")

    def request(
        self,
        method: str,
        endpoint: str,
        *,
        json_body: Any | None = None,
        data: Any | None = None,
        headers: dict[str, str] | None = None,
        expected: tuple[int, ...] = (200,),
        retry_auth: bool = True,
    ) -> requests.Response:
        url = endpoint if endpoint.startswith("http") else f"https://graph.microsoft.com/v1.0{endpoint}"
        h = {"Authorization": f"Bearer {self.token}"}
        if headers:
            h.update(headers)

        r = self.session.request(
            method,
            url,
            json=json_body,
            data=data,
            headers=h,
            timeout=(30, 300),
        )

        if r.status_code == 401 and retry_auth:
            self.refresh_token()
            return self.request(
                method,
                endpoint,
                json_body=json_body,
                data=data,
                headers=headers,
                expected=expected,
                retry_auth=False,
            )

        if r.status_code not in expected:
            raise RuntimeError(f"Graph HTTP {r.status_code}: {r.text[:1200]}")
        return r

    def load_drive(self) -> None:
        self.drive_id = self.request(
            "GET",
            f"/sites/{SITE_ID}/drive?$select=id",
        ).json()["id"]
        self.root_id = self.request(
            "GET",
            f"/drives/{self.drive_id}/root?$select=id",
        ).json()["id"]

    def get_item_by_path(self, path: str) -> requests.Response:
        enc = quote(path.strip("/"), safe="/")
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{enc}"
        r = self.session.get(
            url,
            headers={"Authorization": f"Bearer {self.token}"},
            timeout=(30, 120),
        )
        if r.status_code == 401:
            self.refresh_token()
            r = self.session.get(
                url,
                headers={"Authorization": f"Bearer {self.token}"},
                timeout=(30, 120),
            )
        return r

    def ensure_folder(self, path: str) -> None:
        parts = [p for p in path.strip("/").split("/") if p]
        parent_id = self.root_id
        current: list[str] = []

        for part in parts:
            current.append(part)
            current_path = "/".join(current)
            r = self.get_item_by_path(current_path)

            if r.status_code == 200:
                parent_id = r.json()["id"]
                continue
            if r.status_code != 404:
                raise RuntimeError(f"Falha pasta {current_path}: HTTP {r.status_code}")

            try:
                created = self.request(
                    "POST",
                    f"/drives/{self.drive_id}/items/{parent_id}/children",
                    json_body={
                        "name": part,
                        "folder": {},
                        "@microsoft.graph.conflictBehavior": "fail",
                    },
                    expected=(201,),
                )
                parent_id = created.json()["id"]
            except RuntimeError:
                r2 = self.get_item_by_path(current_path)
                if r2.status_code != 200:
                    raise
                parent_id = r2.json()["id"]

    def read_json(self, path: str) -> dict[str, Any]:
        enc = quote(path.strip("/"), safe="/")
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{enc}:/content"
        r = self.session.get(
            url,
            headers={"Authorization": f"Bearer {self.token}"},
            timeout=(30, 120),
            allow_redirects=True,
        )
        if r.status_code == 404:
            return {}
        if r.status_code == 401:
            self.refresh_token()
            return self.read_json(path)
        if r.status_code != 200:
            raise RuntimeError(f"Falha ao ler {path}: HTTP {r.status_code}")
        return r.json()

    def upload_bytes(self, content: bytes, remote: str, content_type: str) -> None:
        folder = str(Path(remote).parent).replace("\\", "/")
        if folder != ".":
            self.ensure_folder(folder)

        enc = quote(remote.strip("/"), safe="/")
        self.request(
            "PUT",
            f"/drives/{self.drive_id}/root:/{enc}:/content",
            data=content,
            headers={"Content-Type": content_type},
            expected=(200, 201),
        )

    def upload_file(self, local: Path, remote: str) -> None:
        folder = str(Path(remote).parent).replace("\\", "/")
        if folder != ".":
            self.ensure_folder(folder)

        if local.stat().st_size <= SIMPLE_UPLOAD_LIMIT:
            enc = quote(remote.strip("/"), safe="/")
            url = (
                f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}"
                f"/root:/{enc}:/content"
            )

            for attempt in range(1, 6):
                with local.open("rb") as f:
                    r = requests.put(
                        url,
                        data=f,
                        headers={
                            "Authorization": f"Bearer {self.token}",
                            "Content-Type": "application/octet-stream",
                        },
                        timeout=(30, 300),
                    )

                if r.status_code in (200, 201):
                    return
                if r.status_code == 401:
                    self.refresh_token()
                    continue
                if r.status_code in (429, 500, 502, 503, 504):
                    time.sleep(min(30, attempt * 3))
                    continue
                raise RuntimeError(
                    f"Upload simples HTTP {r.status_code}: {r.text[:800]}"
                )

            raise RuntimeError("Upload simples excedeu tentativas.")

        enc = quote(remote.strip("/"), safe="/")
        session_info = self.request(
            "POST",
            f"/drives/{self.drive_id}/root:/{enc}:/createUploadSession",
            json_body={
                "item": {
                    "@microsoft.graph.conflictBehavior": "replace",
                    "name": Path(remote).name,
                }
            },
        ).json()

        upload_url = session_info["uploadUrl"]
        total = local.stat().st_size

        with local.open("rb") as f:
            start = 0
            while start < total:
                chunk = f.read(UPLOAD_CHUNK)
                if not chunk:
                    break

                end = start + len(chunk) - 1
                headers = {
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {start}-{end}/{total}",
                }

                for attempt in range(1, 6):
                    r = requests.put(
                        upload_url,
                        data=chunk,
                        headers=headers,
                        timeout=(30, 300),
                    )
                    if r.status_code in (200, 201, 202):
                        break
                    if r.status_code in (429, 500, 502, 503, 504):
                        time.sleep(min(30, attempt * 3))
                        continue
                    raise RuntimeError(f"Upload grande HTTP {r.status_code}: {r.text[:800]}")
                else:
                    raise RuntimeError("Upload grande excedeu tentativas.")

                start = end + 1


def detect_csv(path: Path) -> tuple[str, str]:
    with path.open("rb") as f:
        sample = f.read(250_000)

    encoding = "utf-8"
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            sample.decode(enc)
            encoding = enc
            break
        except UnicodeDecodeError:
            pass

    text = sample.decode(encoding)
    try:
        delimiter = csv.Sniffer().sniff(text[:100_000], delimiters=",;\t|").delimiter
    except csv.Error:
        delimiter = ";" if text.count(";") > text.count(",") else ","

    return encoding, delimiter


def unique_columns(columns: list[Any]) -> list[str]:
    used: dict[str, int] = {}
    result: list[str] = []

    for i, c in enumerate(columns, start=1):
        base = str(c).strip() if c is not None else ""
        base = base or f"coluna_{i}"
        used[base] = used.get(base, 0) + 1
        result.append(base if used[base] == 1 else f"{base}__{used[base]}")
    return result


def chunk_to_table(df: pd.DataFrame) -> pa.Table:
    columns = unique_columns(list(df.columns))
    df.columns = columns

    arrays = {}
    for c in columns:
        arrays[c] = pa.array(
            [None if pd.isna(v) else str(v) for v in df[c].tolist()],
            type=pa.string(),
        )
    return pa.table(arrays)


def csv_to_parquet(source: Path, target: Path) -> int:
    encoding, delimiter = detect_csv(source)
    writer: pq.ParquetWriter | None = None
    rows = 0

    try:
        reader = pd.read_csv(
            source,
            sep=delimiter,
            encoding=encoding,
            dtype=str,
            chunksize=CSV_CHUNK_ROWS,
            keep_default_na=False,
            na_filter=False,
            on_bad_lines="error",
        )

        for chunk in reader:
            table = chunk_to_table(chunk)
            if writer is None:
                writer = pq.ParquetWriter(
                    target,
                    table.schema,
                    compression="zstd",
                    use_dictionary=True,
                )
            writer.write_table(table)
            rows += len(chunk)
    finally:
        if writer:
            writer.close()

    if writer is None:
        raise RuntimeError(f"CSV sem dados: {source.name}")
    return rows


def xlsx_to_parquet(source: Path, output_dir: Path, prefix: str) -> list[tuple[Path, int]]:
    wb = load_workbook(source, read_only=True, data_only=True)
    outputs: list[tuple[Path, int]] = []

    try:
        for ws in wb.worksheets:
            it = ws.iter_rows(values_only=True)
            header = None

            for _ in range(30):
                try:
                    row = next(it)
                except StopIteration:
                    break
                if any(v is not None and str(v).strip() for v in row):
                    header = unique_columns(list(row))
                    break

            if not header:
                continue

            suffix = f"__{safe_name(ws.title)}" if len(wb.worksheets) > 1 else ""
            target = output_dir / f"{prefix}{suffix}.parquet"
            schema = pa.schema([(c, pa.string()) for c in header])
            writer = pq.ParquetWriter(target, schema, compression="zstd")
            rows = 0
            batch: list[list[Any]] = []

            try:
                for row in it:
                    batch.append(list(row))
                    if len(batch) >= 50_000:
                        data = {
                            c: pa.array(
                                [
                                    None if i >= len(r) or r[i] is None else str(r[i])
                                    for r in batch
                                ],
                                type=pa.string(),
                            )
                            for i, c in enumerate(header)
                        }
                        writer.write_table(pa.table(data))
                        rows += len(batch)
                        batch.clear()

                if batch:
                    data = {
                        c: pa.array(
                            [
                                None if i >= len(r) or r[i] is None else str(r[i])
                                for r in batch
                            ],
                            type=pa.string(),
                        )
                        for i, c in enumerate(header)
                    }
                    writer.write_table(pa.table(data))
                    rows += len(batch)
            finally:
                writer.close()

            outputs.append((target, rows))
    finally:
        wb.close()

    if not outputs:
        raise RuntimeError(f"XLSX sem dados tabulares: {source.name}")
    return outputs



def xls_to_parquet(source: Path, output_dir: Path, prefix: str) -> list[tuple[Path, int]]:
    sheets = pd.read_excel(source, sheet_name=None, dtype=str, engine="xlrd")
    outputs: list[tuple[Path, int]] = []

    for sheet_name, df in sheets.items():
        if df.empty and len(df.columns) == 0:
            continue
        df.columns = unique_columns(list(df.columns))
        suffix = f"__{safe_name(sheet_name)}" if len(sheets) > 1 else ""
        target = output_dir / f"{prefix}{suffix}.parquet"
        table = chunk_to_table(df)
        pq.write_table(table, target, compression="zstd", use_dictionary=True)
        outputs.append((target, len(df)))

    if not outputs:
        raise RuntimeError(f"XLS sem dados tabulares: {source.name}")
    return outputs

def extension_of(resource: dict[str, Any]) -> str:
    fmt = norm(resource.get("format")).replace(".", "")
    if fmt in {"csv", "xlsx", "xls", "zip", "txt"}:
        return f".{fmt}"

    ext = Path(urlparse(str(resource.get("url") or "")).path).suffix.lower()
    return ext if ext in {".csv", ".xlsx", ".xls", ".zip", ".txt"} else ".bin"


def download_resource(resource: dict[str, Any], folder: Path) -> tuple[Path, str]:
    url = str(resource.get("url") or "")
    if not url:
        raise RuntimeError("Recurso sem URL.")

    target = folder / f"origem{extension_of(resource)}"
    digest = hashlib.sha256()

    with HTTP.get(url, stream=True, timeout=(30, 300)) as r:
        r.raise_for_status()
        with target.open("wb") as f:
            for chunk in r.iter_content(DOWNLOAD_CHUNK):
                if chunk:
                    f.write(chunk)
                    digest.update(chunk)

    return target, digest.hexdigest()


def convert_resource(
    source: Path,
    resource: dict[str, Any],
    folder: Path,
) -> list[tuple[Path, int]]:
    out = folder / "out"
    out.mkdir(exist_ok=True)

    prefix = f"{safe_name(resource.get('id'), 50)}__{safe_name(resource.get('name'), 80)}"

    if source.suffix.lower() in {".csv", ".txt"}:
        target = out / f"{prefix}.parquet"
        return [(target, csv_to_parquet(source, target))]

    if source.suffix.lower() == ".xlsx":
        return xlsx_to_parquet(source, out, prefix)

    if source.suffix.lower() == ".xls":
        return xls_to_parquet(source, out, prefix)

    if source.suffix.lower() != ".zip":
        raise RuntimeError(f"Formato nao suportado: {source.suffix}")

    with zipfile.ZipFile(source) as z:
        members = [n for n in z.namelist() if not n.endswith("/") and "__MACOSX" not in n]

        csvs = [n for n in members if Path(n).suffix.lower() in {".csv", ".txt"}]
        excels = [n for n in members if Path(n).suffix.lower() in {".xlsx", ".xls"}]

        # Quando o ZIP traz CSV + JSON espelho, usamos CSV para evitar duplicidade.
        chosen = csvs or excels
        if not chosen:
            raise RuntimeError("ZIP sem CSV/TXT/XLSX tabular.")

        outputs: list[tuple[Path, int]] = []

        for i, member in enumerate(chosen, start=1):
            ext = Path(member).suffix.lower()
            extracted = folder / f"zip_{i}{ext}"

            with z.open(member) as src, extracted.open("wb") as dst:
                shutil.copyfileobj(src, dst, length=DOWNLOAD_CHUNK)

            inner_prefix = f"{prefix}__{safe_name(Path(member).stem)}"

            if ext in {".csv", ".txt"}:
                target = out / f"{inner_prefix}.parquet"
                outputs.append((target, csv_to_parquet(extracted, target)))
            elif ext == ".xlsx":
                outputs.extend(xlsx_to_parquet(extracted, out, inner_prefix))
            else:
                outputs.extend(xls_to_parquet(extracted, out, inner_prefix))

            extracted.unlink(missing_ok=True)

        return outputs


def control_path(category: str) -> str:
    return f"{BASE_FOLDER}/_controle/{category}.json"


def load_control(graph: Graph, category: str) -> dict[str, Any]:
    try:
        control = graph.read_json(control_path(category))
    except Exception:
        control = {}

    if not isinstance(control, dict):
        control = {}
    control.setdefault("resources", {})
    return control


def save_control(graph: Graph, category: str, control: dict[str, Any]) -> None:
    control["updated_at_utc"] = now_iso()
    graph.upload_bytes(
        json.dumps(control, ensure_ascii=False, indent=2).encode("utf-8"),
        control_path(category),
        "application/json; charset=utf-8",
    )


def needs_processing(
    control: dict[str, Any],
    resource: dict[str, Any],
    force: bool,
) -> bool:
    if force:
        return True

    rid = str(resource.get("id") or "")
    previous = control["resources"].get(rid)
    return not previous or previous.get("version") != resource_version(resource)


def process_one(
    graph: Graph,
    control: dict[str, Any],
    category: str,
    dataset: dict[str, Any],
    resource: dict[str, Any],
) -> None:
    rid = str(resource.get("id") or "")
    name = str(resource.get("name") or rid)
    year = str(max(years_of(resource)))

    print("\n" + "=" * 80)
    print(f"{category.upper()} | {name}")
    print("=" * 80)

    with tempfile.TemporaryDirectory(prefix="bi_inss_") as tmp:
        folder = Path(tmp)

        print("1/4 Download oficial...")
        source, sha = download_resource(resource, folder)
        print(f"OK | {source.stat().st_size / 1024 / 1024:,.1f} MB")

        print("2/4 Conversao Parquet em lotes...")
        outputs = convert_resource(source, resource, folder)
        print(f"OK | {sum(r for _, r in outputs):,} linhas")

        print("3/4 Upload SharePoint...")
        remote_outputs = []
        for local, rows in outputs:
            remote = f"{BASE_FOLDER}/silver/{category}/{year}/{local.name}"
            graph.upload_file(local, remote)
            remote_outputs.append(
                {"path": remote, "rows": rows, "bytes": local.stat().st_size}
            )
            print(f"OK | {remote}")

        print("4/4 Checkpoint...")
        control["resources"][rid] = {
            "category": category,
            "dataset_id": dataset.get("id"),
            "dataset_slug": dataset.get("name"),
            "dataset_title": dataset.get("title"),
            "resource_name": name,
            "resource_url": resource.get("url"),
            "resource_format": resource.get("format"),
            "version": resource_version(resource),
            "source_sha256": sha,
            "years": years_of(resource),
            "processed_at_utc": now_iso(),
            "outputs": remote_outputs,
        }
        save_control(graph, category, control)
        print("CHECKPOINT OK")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--dataset",
        choices=["all", *RULES.keys()],
        default=os.getenv("INSS_DATASET", "all"),
    )
    parser.add_argument(
        "--max-resources",
        type=int,
        default=int(os.getenv("MAX_RESOURCES", "3")),
        help="0 = sem limite",
    )
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--catalog-only", action="store_true")
    args = parser.parse_args()

    categories = list(RULES) if args.dataset == "all" else [args.dataset]

    print("=" * 80)
    print("UNIVERSO PREVIDENCIARIO - BI INSS CLOUD")
    print(f"Periodo: {ANO_INICIAL} em diante")
    print("Processamento: um recurso por vez")
    print("=" * 80)

    packages = load_catalog()
    print(f"Datasets encontrados no portal: {len(packages)}")

    selected: dict[str, dict[str, Any]] = {}
    for category in categories:
        ds = choose_dataset(packages, category)
        selected[category] = ds
        print(
            f"[{category}] {ds.get('title')} | "
            f"{len(list_resources(ds))} recursos {ANO_INICIAL}+"
        )

    if args.catalog_only:
        print("CATALOGO OK - nenhum arquivo baixado.")
        return 0

    graph = Graph()
    graph.ensure_folder(f"{BASE_FOLDER}/_controle")
    print("SharePoint conectado.")

    ok = 0
    errors = 0

    for category, dataset in selected.items():
        control = load_control(graph, category)
        pending = [
            r for r in list_resources(dataset)
            if needs_processing(control, r, args.force)
        ]

        if args.max_resources > 0:
            pending = pending[: args.max_resources]

        print(f"\n[{category}] pendentes nesta execucao: {len(pending)}")

        for resource in pending:
            try:
                process_one(graph, control, category, dataset, resource)
                ok += 1
            except Exception as e:
                errors += 1
                print(f"ERRO | {category} | {resource.get('name')}")
                print(str(e))
                print("Nao foi marcado como concluido; sera tentado novamente.")

    print("\n" + "=" * 80)
    print(f"SUCESSO: {ok}")
    print(f"ERROS: {errors}")
    print("=" * 80)

    # Erros parciais nao apagam nem invalidam o que ja foi concluido.
    # A proxima execucao retenta somente os recursos ainda nao checkpointados.
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
