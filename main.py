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
import os
import time
from contextlib import contextmanager
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

LOCK_PATH = config.LOG_DIR / "rpa.lock"
# Mesmo limite do ExecutionTimeLimit do Task Scheduler (2h): uma trava mais
# velha que isso é considerada abandonada (processo anterior crashou sem limpar).
LOCK_MAX_IDADE_SEGUNDOS = 2 * 60 * 60


class ExecucaoConcorrenteError(RuntimeError):
    """Já existe uma execução em andamento (ver `_trava_execucao_unica`)."""


@contextmanager
def _trava_execucao_unica():
    """
    Garante que só uma execução do RPA rode por vez via arquivo de trava
    (`logs/rpa.lock`). Necessário porque há mais de uma tarefa agendada
    (diária/semanal/mensal) e todas usam "Start When Available": se uma
    execução anterior ainda estiver rodando quando outra disparar - inclusive
    duas tarefas que caiam no mesmo horário exato (ex: diária e semanal numa
    segunda-feira) -, duas instâncias escrevendo nos mesmos arquivos do
    OneDrive ao mesmo tempo poderiam corromper dados, ou os dois processos
    logarem ao mesmo tempo no portal C6 e derrubar a sessão um do outro (o
    portal só permite uma sessão ativa por usuário - ver `login`).

    A criação do arquivo usa `os.O_CREAT | os.O_EXCL`, atômica no nível do
    SO: se dois processos tentarem criar a trava ao mesmo tempo, o sistema
    operacional garante que só um consegue - o outro recebe `FileExistsError`
    na hora, sem a janela de corrida que existiria entre "verificar se existe"
    e "escrever o arquivo" como duas operações separadas.
    """
    while True:
        try:
            fd = os.open(LOCK_PATH, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            idade_segundos = time.time() - LOCK_PATH.stat().st_mtime
            if idade_segundos < LOCK_MAX_IDADE_SEGUNDOS:
                raise ExecucaoConcorrenteError(
                    f"Já existe uma execução em andamento (trava criada há "
                    f"{int(idade_segundos)}s em '{LOCK_PATH}') - abortando para não "
                    "rodar duas instâncias ao mesmo tempo sobre os mesmos arquivos. "
                    "Se tiver certeza de que não há nenhuma execução ativa agora, "
                    f"apague o arquivo '{LOCK_PATH}' manualmente e rode de novo."
                )
            logger.warning(
                "Trava de execução encontrada, mas já tem %ds (> %ds do limite) - "
                "a execução anterior provavelmente travou/crashou sem limpar a "
                "trava. Removendo a trava velha e tentando de novo.",
                int(idade_segundos), LOCK_MAX_IDADE_SEGUNDOS,
            )
            try:
                LOCK_PATH.unlink()
            except FileNotFoundError:
                pass
            continue
        else:
            break

    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(str(os.getpid()))
    try:
        yield
    finally:
        try:
            LOCK_PATH.unlink()
        except FileNotFoundError:
            pass


def _usa_sharepoint(base: dict) -> bool:
    """Bases com planilha de origem local não usam SharePoint (dados vão direto
    para o arquivo local, ver data_processor._process_*)."""
    modo = base["regras"].get("modo") or ""
    return not modo.startswith("planilha_origem_local")


def run_bases(bases: list[dict], headless: bool = True):
    """
    Executa o pipeline completo para uma ou mais bases. O login no portal
    C6 é feito uma única vez para todas as bases desta chamada (ver
    looker_automation.download_bases).
    """
    # 1. Baixa a base original atual do SharePoint (para o merge ficar certo).
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
            # linhas_baixadas=None (em vez de 0) distingue "falha técnica" de "sem dados no período".
            data_processor.registrar_historico(base["nome"], None, observacao="Falha no download (ver logs/rpa.log)")
            continue
        try:
            final_path = data_processor.process_base(downloaded_path, base)
            if _usa_sharepoint(base):
                sharepoint_sync.upload_processed_base(final_path, base)
            logger.info("=== Base '%s' concluída com sucesso ===", base["nome"])
        except Exception:
            logger.exception("Falha ao processar a base '%s'", base["nome"])
            data_processor.registrar_historico(base["nome"], None, observacao="Falha ao processar (ver logs/rpa.log)")


def main():
    parser = argparse.ArgumentParser(description="RPA - Bases C6 Veículos")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--base", help="id da base a rodar (ver config.py)")
    group.add_argument("--all", action="store_true", help="roda todas as bases")
    group.add_argument("--frequencia", choices=["diaria", "semanal", "semanal_segunda", "mensal"],
                        help="roda todas as bases dessa frequência")
    parser.add_argument("--debug", action="store_true", help="abre o navegador visível em vez de headless")
    args = parser.parse_args()

    try:
        config.validar_ambiente()
    except RuntimeError:
        logger.exception("Configuração inválida - abortando antes de abrir qualquer navegador.")
        raise

    if args.base:
        bases_to_run = [config.get_base_by_id(args.base)]
    elif args.all:
        bases_to_run = config.BASES
    else:
        bases_to_run = [b for b in config.BASES if b["frequencia"] == args.frequencia]

    try:
        with _trava_execucao_unica():
            run_bases(bases_to_run, headless=not args.debug)
    except ExecucaoConcorrenteError:
        logger.error("Execução abortada: já existe outra em andamento (ver mensagem acima).")
        raise


if __name__ == "__main__":
    main()
