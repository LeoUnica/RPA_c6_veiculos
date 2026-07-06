"""
Automação do download dos relatórios no Looker via Playwright.

IMPORTANTE: os seletores (get_by_text, get_by_role, etc.) abaixo são um
ESQUELETO baseado no fluxo descrito nos slides. Você vai precisar abrir o
Looker de vocês, inspecionar os elementos reais (botões de menu, filtro de
data, botão "Download") e ajustar os seletores marcados com "# TODO".

Rodar `python looker_automation.py --base numero_contratos --debug` abre o
navegador visível (headless=False) para você comparar com o site real.
"""

import argparse
import logging
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, Page

import config

logger = logging.getLogger("looker_automation")


def login(page: Page):
    page.goto(config.LOOKER_URL)
    # TODO: ajustar para o fluxo real de login (SSO, usuário/senha, etc.)
    page.get_by_label("Usuário").fill(config.LOOKER_USER)
    page.get_by_label("Senha").fill(config.LOOKER_PASSWORD)
    page.get_by_role("button", name="Entrar").click()
    page.wait_for_load_state("networkidle")


def navigate_menu(page: Page, looker_path: list[str]):
    """Clica sequencialmente nos itens de menu até chegar no relatório."""
    for item in looker_path:
        # TODO: confirmar se o menu do Looker usa <a>, <button> ou <div role="menuitem">
        page.get_by_text(item, exact=True).click()
        page.wait_for_timeout(800)  # pequena espera para o submenu carregar


def apply_value_filter(page: Page, filtro_valor: str):
    """Aplica o filtro de Valor: este mês / este ano."""
    filtro_map = {
        "este_mes": "Is in this month",
        "este_ano": "In this Year",
    }
    texto_filtro = filtro_map[filtro_valor]

    # TODO: ajustar seletor do campo de filtro "Valor"
    page.get_by_label("Valor").click()
    page.get_by_text(texto_filtro, exact=True).click()
    page.wait_for_load_state("networkidle")


def open_bloco_if_needed(page: Page, bloco: str | None):
    if bloco:
        # TODO: ajustar seletor do bloco (ex: "Bloco de Metas - Por Filial")
        page.get_by_text(bloco, exact=True).click()
        page.wait_for_timeout(800)


def download_report(page: Page, base_id: str) -> Path:
    """Clica em Download > Excel > All results e salva o arquivo."""
    with page.expect_download() as download_info:
        # TODO: ajustar para o menu real de export do Looker
        page.get_by_role("button", name="Download").click()
        page.get_by_text("Excel", exact=True).click()
        # marca "All results" conforme especificado no material da equipe
        page.get_by_label("All results").check()
        page.get_by_role("button", name="Download", exact=True).click()

    download = download_info.value
    dest_path = config.DOWNLOAD_DIR / f"{base_id}_{int(time.time())}.xlsx"
    download.save_as(dest_path)
    logger.info("Arquivo baixado: %s", dest_path)
    return dest_path


def download_base(base: dict, headless: bool = True) -> Path:
    """Executa o fluxo completo de download para uma base configurada."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            login(page)
            navigate_menu(page, base["looker_path"])
            open_bloco_if_needed(page, base.get("bloco"))
            apply_value_filter(page, base["filtro_valor"])
            path = download_report(page, base["id"])
        finally:
            context.close()
            browser.close()

    return path


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    parser = argparse.ArgumentParser(description="Download de relatório do Looker")
    parser.add_argument("--base", required=True, help="id da base (ver config.py)")
    parser.add_argument("--debug", action="store_true", help="abre o navegador visível")
    args = parser.parse_args()

    base_cfg = config.get_base_by_id(args.base)
    download_base(base_cfg, headless=not args.debug)
