"""
Carga única de um ano fechado (ex: 2025) para uma das 4 bases que
suportam isso - fora do fluxo diário/mensal normal (main.py). "Carteira e
Parceiros" não precisa deste script (já baixa o ano inteiro sozinha).

Uso:
    python backfill_ano_fechado.py --base meta_financiamento_seguro --ano 2025
    python backfill_ano_fechado.py --all --ano 2025    # roda as 4 bases
    python backfill_ano_fechado.py --base numero_contratos --ano 2025 --debug

Todos os 4 filtros usados aqui (Safra Mês/Referencia Month -> "is in the
year"; Dt Relatorio Date -> intervalo customizado) foram validados ao vivo
contra o portal em 25/08/2026 antes deste script existir - ver
looker_automation.download_ano_fechado_report.

Roda uma vez só, por ano. Depois de completar, o ano fica congelado: o
fluxo diário normal (main.py) nunca mais escreve nele (ver
data_processor._eh_ano_corrente) - rodar este script de novo para o mesmo
ano/base atualiza a planilha (mesma regra de cor: verde para chave nova,
amarelo para chave existente com dado alterado), então é seguro executar
de novo se uma base falhar no meio da carga.
"""

import argparse
import logging

from playwright.sync_api import sync_playwright

import config
import data_processor
import looker_automation as la

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.FileHandler(config.LOG_DIR / "rpa.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("backfill_ano_fechado")


def backfill_base(context, page, base_id: str, ano: int):
    logger.info("=== Carga do ano fechado %d: '%s' ===", ano, base_id)
    downloaded_path = la.download_ano_fechado_report(context, page, base_id, ano)
    origem_path = data_processor.processar_ano_fechado(downloaded_path, base_id, ano)
    logger.info("=== '%s' (%d) concluída com sucesso: %s ===", base_id, ano, origem_path)


def main():
    parser = argparse.ArgumentParser(description="Carga única de um ano fechado (backfill)")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--base", choices=la.BASES_COM_ANO_FECHADO, help="id da base a carregar")
    group.add_argument("--all", action="store_true", help="roda as 4 bases que suportam ano fechado")
    parser.add_argument("--ano", type=int, required=True, help="ano fechado a carregar (ex: 2025)")
    parser.add_argument("--debug", action="store_true", help="abre o navegador visível em vez de headless")
    args = parser.parse_args()

    config.validar_ambiente()
    base_ids = [args.base] if args.base else list(la.BASES_COM_ANO_FECHADO)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.debug)
        context = browser.new_context(accept_downloads=True, viewport={"width": 1600, "height": 900})
        page = context.new_page()
        try:
            la.login(page)
            for base_id in base_ids:
                try:
                    backfill_base(context, page, base_id, args.ano)
                except Exception:
                    logger.exception("Falha na carga do ano fechado %d para '%s'", args.ano, base_id)
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
