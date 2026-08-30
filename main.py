from __future__ import annotations

import json
import re
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen


BASE_URL = "https://dadosabertos.inss.gov.br"
ORGANIZACAO = "instituto-nacional-de-seguro-social-inss"
ANO_INICIAL = 2025

# Lista controlada das bases que interessam ao BI.
# Não usamos filtro genérico como "emitidos" ou "requerimentos",
# para evitar capturar datasets errados.
BASES_ALVO = (
    "beneficios concedidos",
    "beneficios indeferidos",
    "dados de requerimentos administrativos solicitados",
    "dados quantitativos de requerimentos administrativos solicitados",
    "dados de requerimentos administrativos pendentes",
    "dados quantitativos de requerimentos administrativos pendentes",
    "beneficios emitidos",
    "beneficios mantidos",
    "comunicacoes de acidente de trabalho",
    "perfil das unidades",
    "dados agregados da folha de pagamento em relacao aos beneficios emitidos",
)

ARQUIVO_SAIDA = Path("catalogo_inss_2025.json")


def normalizar(texto: object) -> str:
    texto = "" if texto is None else str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = texto.encode("ascii", "ignore").decode("ascii")
    texto = texto.lower().strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def api_get(acao: str, parametros: dict) -> dict:
    query = urlencode(parametros)
    url = f"{BASE_URL}/api/3/action/{acao}?{query}"

    req = Request(
        url,
        headers={
            "User-Agent": "UniversoPrevidenciario-BI-INSS/1.0",
            "Accept": "application/json",
        },
    )

    with urlopen(req, timeout=120) as resposta:
        payload = json.loads(resposta.read().decode("utf-8"))

    if not payload.get("success"):
        raise RuntimeError(f"CKAN retornou success=false em {acao}")

    return payload["result"]


def eh_base_alvo(titulo: str) -> bool:
    titulo_n = normalizar(titulo)
    return any(base in titulo_n for base in BASES_ALVO)


def anos_do_recurso(nome: str) -> list[int]:
    # IMPORTANTE:
    # O ano é lido do NOME DO RECURSO, e não do título do dataset.
    # Isso evita interpretar "PDA Jun/2023 a Jun/2027"
    # como se o arquivo fosse de 2027.
    return [int(a) for a in re.findall(r"\b(20\d{2})\b", nome or "")]


def recurso_desde_2025(nome: str) -> bool:
    anos = anos_do_recurso(nome)

    # Recursos sem ano explícito podem ser dicionários,
    # cadastros ou arquivos de referência e são preservados.
    if not anos:
        return True

    return max(anos) >= ANO_INICIAL


def main() -> None:
    print("=" * 70)
    print("UNIVERSO PREVIDENCIARIO - API INSS")
    print(f"Janela oficial do projeto: {ANO_INICIAL} em diante")
    print("Modo: SOMENTE CATALOGO - nenhum arquivo pesado sera baixado")
    print("=" * 70)

    resultado = api_get(
        "package_search",
        {
            "fq": f"organization:{ORGANIZACAO}",
            "rows": 1000,
            "start": 0,
        },
    )

    datasets_api = resultado.get("results", [])
    selecionados = []

    for dataset in datasets_api:
        titulo = dataset.get("title") or dataset.get("name") or ""

        if not eh_base_alvo(titulo):
            continue

        recursos_validos = []

        for recurso in dataset.get("resources") or []:
            nome = recurso.get("name") or recurso.get("description") or ""

            if not recurso_desde_2025(nome):
                continue

            recursos_validos.append(
                {
                    "resource_id": recurso.get("id"),
                    "nome": nome,
                    "formato": recurso.get("format"),
                    "url": recurso.get("url"),
                    "created": recurso.get("created"),
                    "last_modified": recurso.get("last_modified"),
                    "mimetype": recurso.get("mimetype"),
                    "datastore_active": bool(
                        recurso.get("datastore_active", False)
                    ),
                }
            )

        if recursos_validos:
            selecionados.append(
                {
                    "dataset_id": dataset.get("id"),
                    "dataset_slug": dataset.get("name"),
                    "dataset_titulo": titulo,
                    "metadata_modified": dataset.get("metadata_modified"),
                    "quantidade_recursos": len(recursos_validos),
                    "recursos": recursos_validos,
                }
            )

    catalogo = {
        "gerado_em_utc": datetime.now(timezone.utc).isoformat(),
        "fonte": BASE_URL,
        "ano_inicial": ANO_INICIAL,
        "datasets_encontrados_no_portal": len(datasets_api),
        "datasets_selecionados": len(selecionados),
        "datasets": selecionados,
    }

    ARQUIVO_SAIDA.write_text(
        json.dumps(catalogo, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print()
    print(f"Datasets encontrados no portal: {len(datasets_api)}")
    print(f"Datasets selecionados para o BI: {len(selecionados)}")
    print()

    total_recursos = 0

    for dataset in selecionados:
        qtd = dataset["quantidade_recursos"]
        total_recursos += qtd
        print(f"[OK] {dataset['dataset_titulo']}")
        print(f"     Recursos de 2025 em diante/referencia: {qtd}")

    print()
    print(f"TOTAL DE RECURSOS CONTROLADOS: {total_recursos}")
    print(f"Catalogo salvo em: {ARQUIVO_SAIDA}")
    print()
    print("TESTE CONCLUIDO.")
    print("Nenhum CSV, XLSX ou ZIP do INSS foi baixado nesta etapa.")


if __name__ == "__main__":
    main()
