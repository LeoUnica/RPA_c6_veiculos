"""
Automação do download dos relatórios no Looker via Playwright.

IMPORTANTE: os seletores (get_by_text, get_by_role, etc.) abaixo são um
ESQUELETO baseado no fluxo descrito nos slides. Você vai precisar abrir o
Looker de vocês, inspecionar os elementos reais (botões de menu, filtro de
data, botão "Download") e ajustar os seletores marcados com "# TODO".

A base "numero_contratos" ("Acompanhamento Veículos" > "Analítico") já tem
um fluxo detalhado (fornecido pela equipe) implementado em
`download_numero_contratos_report` - os seletores de ícone (svg) e texto
foram copiados do HTML real informado pela equipe, mas alguns pontos ainda
têm "# TODO" onde a estrutura exata da página não foi confirmada.

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

# --------------------------------------------------------------------------
# Seletores (svg path) copiados do HTML real fornecido pela equipe para o
# fluxo de "Acompanhamento Veículos" > "Analítico" (base numero_contratos)
# --------------------------------------------------------------------------
ICON_FILTER_PANEL_PATH = "M10 18h4v-2h-4v2zM3 6v2h18V6H3zm3 7h12v-2H6v2z"
ICON_MORE_VERT_PATH = (
    "M12 8c1.1 0 2-.9 2-2s-.9-2-2-2-2 .9-2 2 .9 2 2 2zm0 2c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2z"
    "m0 6c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2z"
)


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


# --------------------------------------------------------------------------
# Fluxo dedicado - base "numero_contratos" (Acompanhamento Veículos > Analítico)
# --------------------------------------------------------------------------

def _click_icon_button(page: Page, svg_path_d: str):
    """Clica no botão cujo ícone svg tem o atributo 'd' informado."""
    button = page.locator(f'button:has(svg path[d="{svg_path_d}"])').first
    button.click()


def open_acompanhamento_veiculos_analitico(page: Page, base: dict):
    """
    Navega até o relatório Analítico de Acompanhamento Veículos:
    Relatórios > Relatórios Gerenciais > Auto (painel) > card "Acompanhamento"
    > "Acompanhamento Veículos" > botão "Analítico".
    """
    navigate_menu(page, base["looker_path"])  # Relatórios > Relatórios Gerenciais > Auto

    # No card "Acompanhamento", seleciona "Acompanhamento Veículos"
    page.get_by_text(base["card_acompanhamento"], exact=True).click()
    page.wait_for_load_state("networkidle")

    # Clica no botão "Analítico"
    page.get_by_role("button", name=base["aba_relatorio"]).click()
    page.wait_for_load_state("networkidle")


def apply_analitico_filters(page: Page, filtros: dict):
    """
    Abre o painel de filtros do lado direito e configura:
      - "Tipo Exibição" -> mantém somente a opção informada (ex: "Valor")
      - "Dt Relatorio Date" -> período relativo (ex: "Last 30 Days")
    """
    # Abre o painel de filtros (ícone de filtro no lado direito)
    _click_icon_button(page, ICON_FILTER_PANEL_PATH)
    page.wait_for_timeout(500)

    # --- Tipo Exibição ---
    # TODO: confirmar se é necessário expandir o card do filtro antes de
    # marcar a opção (ex: clicar no título "Tipo Exibição" para abrir a lista)
    page.get_by_text("Tipo Exibição", exact=True).click()
    page.get_by_text(filtros["tipo_exibicao"], exact=True).click()

    # --- Dt Relatorio Date -> "Last 30 Days" ---
    page.get_by_text("Dt Relatorio Date", exact=True).click()
    page.get_by_text(filtros["periodo_dt_relatorio"], exact=True).click()

    page.wait_for_load_state("networkidle")


def update_report_data(page: Page):
    """Clica no botão 'Update' para atualizar os dados do relatório."""
    page.locator('button[aria-labelledby="page-freshness-indicator"]').click()
    page.wait_for_load_state("networkidle")
    # pequena espera extra para garantir que o refresh dos dados terminou
    page.wait_for_timeout(1500)


def download_analitico_spreadsheet(page: Page, base_id: str) -> Path:
    """
    Rola até a planilha "Analítico", abre o menu de 3 pontinhos, clica em
    "Download data", seleciona o formato Excel, expande "Advanced data
    options" e marca as opções de exportação completa antes de baixar.
    """
    analitico_title = page.get_by_text("Analítico", exact=True)
    analitico_title.scroll_into_view_if_needed()

    # O botão de 3 pontinhos fica praticamente invisível até passar o mouse
    # por cima da linha/planilha (revelado via hover no CSS) - por isso
    # damos hover no container antes de clicar, e usamos force=True como
    # segurança caso o Playwright ainda considere o elemento "oculto".
    container = analitico_title.locator(
        f'xpath=ancestor::*[.//button[.//svg/path[@d="{ICON_MORE_VERT_PATH}"]]][1]'
    ).first
    container.hover()
    more_button = container.locator(f'button:has(svg path[d="{ICON_MORE_VERT_PATH}"])').first
    more_button.click(force=True)

    page.get_by_text("Download data", exact=True).click()

    # Formato do download
    page.get_by_text("Excel Spreadsheet (Excel 2007 or later)", exact=True).click()

    # Expande "Advanced data options"
    page.get_by_text("Advanced data options", exact=True).click()

    # Results -> "With visualizations options applied"
    page.get_by_text("With visualizations options applied", exact=True).click()

    # Data Values -> "Formatted"
    page.get_by_text("Formatted", exact=True).click()

    # Number of rows to include -> "All results"
    page.get_by_text("All results", exact=True).click()

    with page.expect_download() as download_info:
        page.get_by_role("button", name="Download", exact=True).click()

    download = download_info.value
    dest_path = config.DOWNLOAD_DIR / f"{base_id}_{int(time.time())}.xlsx"
    download.save_as(dest_path)
    logger.info("Arquivo baixado: %s", dest_path)
    return dest_path


def download_numero_contratos_report(page: Page, base: dict) -> Path:
    """Fluxo completo específico da base 'numero_contratos'."""
    open_acompanhamento_veiculos_analitico(page, base)
    apply_analitico_filters(page, base["filtros"])
    update_report_data(page)
    return download_analitico_spreadsheet(page, base["id"])


def download_base(base: dict, headless: bool = True) -> Path:
    """Executa o fluxo completo de download para uma base configurada."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            login(page)

            if base["id"] == "numero_contratos":
                path = download_numero_contratos_report(page, base)
            else:
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
