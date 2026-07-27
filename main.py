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


def _usa_sharepoint(base: dict) -> bool:
    # Bases com planilha de origem local (Número de Contratos, Dias sem
    # Produção, Meta Financiamento e Seguro, Carteira e Parceiros) não usam
    # SharePoint: os dados vão direto para o arquivo local (ver
    # data_processor._process_*).
    modo = base["regras"].get("modo") or ""
    return not modo.startswith("planilha_origem_local")


def run_bases(bases: list[dict], headless: bool = True):
    """
    Executa o pipeline completo para uma ou mais bases. O login no portal
    C6 é feito **uma única vez** para todas as bases desta chamada (ver
    looker_automation.download_bases) - não há necessidade de sair da conta
    e entrar de novo a cada procedimento, já que cada relatório é acessado
    a partir da mesma aba logada do WebAutorizador.
    """
    # 1. Baixa a base original atual do SharePoint de cada base (para o merge
    # ficar certo) - independente do portal C6/Looker.
    for base in bases:
        if _usa_sharepoint(base):
            original_local = config.STAGING_DIR / f"{base['id']}_original.xlsx"
            try:
                sharepoint_sync.download_original_base(base, original_local)
            except Exception:
                logger.warning("Não foi possível baixar a base original de '%s' (pode ser a primeira execução).", base["nome"])

    # 2. Login único no portal C6 e download de todas as bases do Looker
    logger.info("=== Iniciando download no portal C6 (login único para %d base(s)) ===", len(bases))
    downloaded_paths = looker_automation.download_bases(bases, headless=headless)

    # 3. Trata, mescla e sobe cada base que baixou com sucesso
    for base in bases:
        downloaded_path = downloaded_paths.get(base["id"])
        if downloaded_path is None:
            logger.error("Base '%s' pulada: download do Looker falhou (ver erro acima).", base["nome"])
            continue
        try:
            final_path = data_processor.process_base(downloaded_path, base)
            if _usa_sharepoint(base):
                sharepoint_sync.upload_processed_base(final_path, base)
            logger.info("=== Base '%s' concluída com sucesso ===", base["nome"])
        except Exception:
            logger.exception("Falha ao processar a base '%s'", base["nome"])


def main():
    parser = argparse.ArgumentParser(description="RPA - Bases C6 Veículos")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--base", help="id da base a rodar (ver config.py)")
    group.add_argument("--all", action="store_true", help="roda todas as bases")
    group.add_argument("--frequencia", choices=["diaria", "semanal", "semanal_segunda", "mensal"],
                        help="roda todas as bases dessa frequência")
    parser.add_argument("--debug", action="store_true", help="abre o navegador visível em vez de headless")
    args = parser.parse_args()

    if args.base:
        bases_to_run = [config.get_base_by_id(args.base)]
    elif args.all:
        bases_to_run = config.BASES
    else:
        bases_to_run = [b for b in config.BASES if b["frequencia"] == args.frequencia]

    run_bases(bases_to_run, headless=not args.debug)


if __name__ == "__main__":
    main()
