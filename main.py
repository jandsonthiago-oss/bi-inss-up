from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen


BASE_URL = "https://dadosabertos.inss.gov.br"
ORGANIZACAO = "instituto-nacional-de-seguro-social-inss"
ANO_INICIAL = 2025

# SOMENTE os datasets oficiais atuais que interessam ao BI.
# Filtro exato pelo slug CKAN: elimina series legadas e coincidencias por nome.
DATASETS_ALVO = {
    "beneficios-concedidos-plano-de-dados-abertos-jun-2023-a-jun-2025",
    "beneficios-indeferidos-plano-de-dados-abertos-jun-2023-a-jun-2025",
    "beneficios-emitidos-plano-de-dados-abertos-jun-2023-a-jun-2025",
    "beneficios-mantidos-plano-de-dados-abertos-jun-2023-a-jun-2025",
    "comunicacoes-de-acidente-de-trabalho-cat-plano-de-dados-abertos-jun-2023-a-jun-2025",
    "perfil-das-unidades-plano-de-dados-abertos-jun-2023-a-jun-2025",
    "dados-de-requerimentos-administrativos-solicitados-plano-de-dados-abertos-jun-2023-a-jun-2025",
    "dados-de-requerimentos-administrativos-pendentes-plano-de-dados-abertos-jun-2023-a-jun-2025",
    "dados-agregados-da-folha-de-pagamento-beneficios-emitidos-plano-de-dados-abertos-jun-2023-a-jun-2027",
}

ARQUIVO_SAIDA = Path("catalogo_inss_2025.json")


def api_get(acao: str, parametros: dict) -> dict:
    query = urlencode(parametros)
    url = f"{BASE_URL}/api/3/action/{acao}?{query}"

    req = Request(
        url,
        headers={
            "User-Agent": "UniversoPrevidenciario-BI-INSS/2.0",
            "Accept": "application/json",
        },
    )

    with urlopen(req, timeout=120) as resposta:
        payload = json.loads(resposta.read().decode("utf-8"))

    if not payload.get("success"):
        raise RuntimeError(f"CKAN retornou success=false em {acao}")

    return payload["result"]


def anos_do_recurso(nome: str) -> list[int]:
    return [int(a) for a in re.findall(r"\b(20\d{2})\b", nome or "")]


def recurso_desde_2025(nome: str) -> bool:
    anos = anos_do_recurso(nome)

    # Sem ano explicito = dicionario/referencia.
    # Preservamos porque pode ser necessario para interpretar os dados.
    if not anos:
        return True

    return max(anos) >= ANO_INICIAL


def main() -> None:
    print("=" * 72)
    print("UNIVERSO PREVIDENCIARIO - CATALOGO OFICIAL INSS")
    print(f"Periodo de producao: {ANO_INICIAL} em diante")
    print("Modo seguro: nenhum CSV/XLSX/ZIP pesado sera baixado")
    print("=" * 72)

    resultado = api_get(
        "package_search",
        {
            "fq": f"organization:{ORGANIZACAO}",
            "rows": 1000,
            "start": 0,
        },
    )

    datasets_api = resultado.get("results", [])

    encontrados_slugs = set()
    selecionados = []
    resource_ids_vistos = set()

    for dataset in datasets_api:
        slug = (dataset.get("name") or "").strip()

        if slug not in DATASETS_ALVO:
            continue

        encontrados_slugs.add(slug)

        titulo = dataset.get("title") or slug
        recursos_validos = []

        for recurso in dataset.get("resources") or []:
            nome = recurso.get("name") or recurso.get("description") or ""

            if not recurso_desde_2025(nome):
                continue

            resource_id = recurso.get("id")

            # Protecao contra o mesmo resource_id aparecer duas vezes.
            if resource_id in resource_ids_vistos:
                raise RuntimeError(
                    f"RESOURCE_ID DUPLICADO DETECTADO: {resource_id}"
                )

            resource_ids_vistos.add(resource_id)

            recursos_validos.append(
                {
                    "resource_id": resource_id,
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

        selecionados.append(
            {
                "dataset_id": dataset.get("id"),
                "dataset_slug": slug,
                "dataset_titulo": titulo,
                "metadata_modified": dataset.get("metadata_modified"),
                "quantidade_recursos": len(recursos_validos),
                "recursos": recursos_validos,
            }
        )

    # Se o INSS mudar/remover um dataset, o processo PARA.
    # Nao seguimos silenciosamente com uma base incompleta.
    faltantes = DATASETS_ALVO - encontrados_slugs

    if faltantes:
        print("\nERRO: datasets oficiais esperados nao encontrados:")
        for slug in sorted(faltantes):
            print(f"  - {slug}")

        raise RuntimeError(
            "Catalogo incompleto. Nenhum processamento deve continuar."
        )

    catalogo = {
        "gerado_em_utc": datetime.now(timezone.utc).isoformat(),
        "fonte": BASE_URL,
        "ano_inicial": ANO_INICIAL,
        "datasets_encontrados_no_portal": len(datasets_api),
        "datasets_esperados": len(DATASETS_ALVO),
        "datasets_selecionados": len(selecionados),
        "datasets": selecionados,
    }

    ARQUIVO_SAIDA.write_text(
        json.dumps(catalogo, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print()
    print(f"Datasets encontrados no portal: {len(datasets_api)}")
    print(f"Datasets oficiais esperados: {len(DATASETS_ALVO)}")
    print(f"Datasets selecionados: {len(selecionados)}")
    print()

    total_recursos = 0

    for dataset in sorted(
        selecionados,
        key=lambda d: d["dataset_titulo"].lower()
    ):
        qtd = dataset["quantidade_recursos"]
        total_recursos += qtd

        print(f"[OK] {dataset['dataset_titulo']}")
        print(f"     Slug: {dataset['dataset_slug']}")
        print(f"     Dataset ID CKAN: {dataset['dataset_id']}")
        print(f"     Recursos 2025+/referencia: {qtd}")

    print()
    print("=" * 72)
    print(f"DATASETS VALIDOS: {len(selecionados)}/9")
    print(f"TOTAL DE RECURSOS CONTROLADOS: {total_recursos}")
    print(f"Catalogo salvo em: {ARQUIVO_SAIDA}")
    print("VALIDACAO CONCLUIDA COM SUCESSO.")
    print("Nenhum arquivo pesado do INSS foi baixado.")
    print("=" * 72)


if __name__ == "__main__":
    main()
