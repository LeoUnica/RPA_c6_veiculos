"""
Orquestrador da RPA - Bases C6 Veículos.

Uso:
    python main.py --base numero_contratos       # roda uma base específica
    python main.py --all                          # roda todas as bases
    python main.py --frequencia diaria             # roda só as bases diárias
                                                     (útil para agendar no
                                                     Task Scheduler / cron)
"""

import argparse
import logging
from pathlib import Path

import config
import looker_automation
import data_processor
import sharepoint_sync

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.FileHandler(config.LOG_DIR / "rpa.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("main")


def run_base(base: dict):
    logger.info("=== Iniciando base: %s ===", base["nome"])
    try:
        # 1. Baixa a base original atual do SharePoint (para o merge ficar certo)
        original_local = config.STAGING_DIR / f"{base['id']}_original.xlsx"
        try:
            sharepoint_sync.download_original_base(base, original_local)
        except Exception:
            logger.warning("Não foi possível baixar a base original (pode ser a primeira execução).")

        # 2. Baixa o relatório novo do Looker
        downloaded_path = looker_automation.download_base(base)

        # 3. Trata e mescla os dados
        final_path = data_processor.process_base(downloaded_path, base)

        # 4. Sobe o resultado final para o SharePoint
        sharepoint_sync.upload_processed_base(final_path, base)

        logger.info("=== Base '%s' concluída com sucesso ===", base["nome"])
    except Exception:
        logger.exception("Falha ao processar a base '%s'", base["nome"])
        raise


def main():
    parser = argparse.ArgumentParser(description="RPA - Bases C6 Veículos")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--base", help="id da base a rodar (ver config.py)")
    group.add_argument("--all", action="store_true", help="roda todas as bases")
    group.add_argument("--frequencia", choices=["diaria", "semanal", "semanal_segunda"],
                        help="roda todas as bases dessa frequência")
    args = parser.parse_args()

    if args.base:
        bases_to_run = [config.get_base_by_id(args.base)]
    elif args.all:
        bases_to_run = config.BASES
    else:
        bases_to_run = [b for b in config.BASES if b["frequencia"] == args.frequencia]

    for base in bases_to_run:
        run_base(base)


if __name__ == "__main__":
    main()
