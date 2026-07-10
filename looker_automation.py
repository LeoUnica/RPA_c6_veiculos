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
import re
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, BrowserContext, Locator, Page

import config

logger = logging.getLogger("looker_automation")

# --------------------------------------------------------------------------
# Seletor (svg path) do ícone de "3 pontinhos" (Tile actions) usado no
# fluxo de download da planilha "Analítico" (base numero_contratos)
# --------------------------------------------------------------------------
ICON_MORE_VERT_PATH = (
    "M12 8c1.1 0 2-.9 2-2s-.9-2-2-2-2 .9-2 2 .9 2 2 2zm0 2c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2z"
    "m0 6c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2z"
)


def login(page: Page):
    """
    Login no portal C6 Consig (WebAutorizador - página ASP.NET clássica,
    sem <label>). Seletores confirmados inspecionando o HTML real da página:
      - Usuário: input#EUsuario_CAMPO
      - Senha:   input#ESenha_CAMPO
      - Entrar:  <a id="lnkEntrar"> (link com postback, não é um <button>)

    O portal costuma mostrar um confirm() JS ("Usuário já autenticado em
    outra estação. Deseja desconectar-se...") quando já existe uma sessão
    ativa - aceitamos automaticamente para forçar a nova sessão.
    """
    page.on("dialog", lambda dialog: dialog.accept())

    page.goto(config.LOOKER_URL)
    page.locator("#EUsuario_CAMPO").fill(config.LOOKER_USER)
    page.locator("#ESenha_CAMPO").fill(config.LOOKER_PASSWORD)
    page.locator("#lnkEntrar").click()
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
#
# Este fluxo foi validado rodando de verdade contra o portal (não é mais um
# esqueleto/chute): o relatório é hospedado no Google Looker de verdade,
# embutido dentro do WebAutorizador via duas janelas pop-up sucessivas.
# --------------------------------------------------------------------------

def open_acompanhamento_veiculos_analitico(context: BrowserContext, page: Page) -> Page:
    """
    Navega até o dashboard "Acompanhamento Veículos" e retorna a Page do
    Looker onde ele foi aberto (é uma nova aba/pop-up, não a mesma página).

    Fluxo real confirmado:
      1. O menu "Relatórios" só revela "Relatórios Gerenciais" com hover
         (não com click).
      2. Clicar em "Relatórios Gerenciais" abre uma pop-up nova com o
         catálogo de relatórios do Looker (dashboards/371).
      3. Dentro dessa pop-up, o card "🚗 Auto" (o texto vem com emoji e
         espaço, por isso o match é parcial) leva ao dashboard "One Page -
         Auto" (dashboards/513), na mesma aba.
      4. O card "Acompanhamento Veículos" tem DOIS elementos com o mesmo
         texto: o título do card (não clicável) e, mais abaixo, o link de
         fato (por isso usamos `.nth(1)`, não `.first`).
      5. Clicar nesse link abre OUTRA pop-up com o dashboard final
         (corp_consignado_embed::00050_producao).
    """
    page.get_by_text("Relatórios", exact=True).first.hover()
    page.wait_for_timeout(500)

    with context.expect_page(timeout=15000) as popup_info:
        page.get_by_text("Relatórios Gerenciais", exact=True).first.click()
    catalogo = popup_info.value
    catalogo.wait_for_load_state("networkidle", timeout=20000)
    catalogo.wait_for_timeout(5000)

    catalogo.get_by_text("Auto", exact=False).first.click()
    catalogo.wait_for_timeout(3000)
    catalogo.wait_for_load_state("networkidle", timeout=15000)
    catalogo.wait_for_timeout(2000)

    link_acompanhamento = catalogo.get_by_text("Acompanhamento Veículos", exact=True).nth(1)
    with context.expect_page(timeout=10000) as popup_info2:
        link_acompanhamento.click(force=True)
    final_page = popup_info2.value

    # O dashboard final faz polling contínuo em segundo plano, então
    # "networkidle" nunca conclui aqui - usamos espera fixa.
    final_page.wait_for_load_state("domcontentloaded", timeout=20000)
    final_page.wait_for_timeout(8000)
    return final_page


def apply_analitico_filters(final_page: Page, filtros: dict):
    """
    Clica na aba "Analítico" e configura, no painel de filtros:
      - "Tipo Exibição" -> mantém somente a opção informada (ex: "Valor")
      - "Dt Relatorio Date" -> período relativo (ex: "Last 30 Days")

    A aba é um link cujo texto acessível inclui um emoji (ex: "❗ Analítico"),
    por isso usamos correspondência parcial via role em vez de texto exato.
    """
    final_page.get_by_role("link", name="Analítico", exact=False).first.click()
    final_page.wait_for_timeout(5000)

    # Abre o painel de filtros (botão "NN filters", o número varia por aba)
    final_page.get_by_text("filters", exact=False).first.click()
    final_page.wait_for_timeout(1500)

    # --- Tipo Exibição ---
    # O chip de valor do filtro mostra o texto atual (ex: "is Qtde" ou
    # "is Valor"). "Tipo Exibição" é sempre o primeiro filtro do painel,
    # então o primeiro chip que começa com "is " é o dele.
    final_page.get_by_text(re.compile(r"^is "), exact=False).first.click()
    final_page.wait_for_timeout(500)
    final_page.get_by_text(filtros["tipo_exibicao"], exact=True).click()
    final_page.wait_for_timeout(500)

    # --- Dt Relatorio Date ---
    # Na prática já vem com o valor certo por padrão salvo no dashboard
    # (confirmado via querystring "Dt+Relatorio+Date=30+day"). Só avisamos
    # no log se algum dia vier diferente - o clique para trocar esse filtro
    # específico ainda não foi mapeado/validado.
    if final_page.get_by_text(filtros["periodo_dt_relatorio"], exact=True).count() == 0:
        logger.warning(
            "Filtro 'Dt Relatorio Date' não está em '%s' - ajuste manual pode "
            "ser necessário (fluxo de troca ainda não mapeado).",
            filtros["periodo_dt_relatorio"],
        )


def update_report_data(final_page: Page):
    """Clica no botão 'Update' para atualizar os dados do relatório."""
    final_page.locator('button[aria-labelledby="page-freshness-indicator"]').click()
    # espera fixa: networkidle não é confiável nesse dashboard (polling contínuo)
    final_page.wait_for_timeout(5000)


def _find_tile_actions_button(final_page: Page, near: Locator) -> Locator:
    """
    Encontra o botão "Tile actions" (3 pontinhos) mais próximo verticalmente
    do elemento `near`. O mesmo ícone svg é reaproveitado ~40x na página
    (cada coluna do crosstab tem um mini-ícone igual no cabeçalho, e há um
    menu global de dashboard também com o mesmo ícone) - por isso filtramos
    por altura do botão (os de tile ficam com 24px, diferente dos 36px do
    menu global e dos ~21px dos ícones de coluna) e pegamos o mais próximo
    em Y do elemento de referência.
    """
    near_box = near.bounding_box()
    candidatos = final_page.locator(f'button:has(svg path[d="{ICON_MORE_VERT_PATH}"])')
    melhor = None
    menor_distancia = None
    for i in range(candidatos.count()):
        el = candidatos.nth(i)
        box = el.bounding_box()
        if not box or abs(box["height"] - 24) > 2:
            continue
        distancia = abs(box["y"] - near_box["y"])
        if menor_distancia is None or distancia < menor_distancia:
            menor_distancia = distancia
            melhor = el
    return melhor


def _complete_download_dialog(final_page: Page, base_id: str) -> Path:
    """
    A partir do menu "Tile actions" já aberto, clica em "Download data",
    seleciona o formato Excel, expande "Advanced data options" e marca as
    opções de exportação completa antes de baixar. Compartilhado por todas
    as bases que usam este mesmo fluxo de download do Looker.
    """
    final_page.get_by_text("Download data", exact=True).click()
    final_page.wait_for_timeout(1500)

    # O combobox de formato parece usar Shadow DOM fechado: com ele fechado
    # não dá pra ler nem clicar no valor atual ("CSV") por texto. Precisa
    # abrir a lista primeiro - só aí as opções viram texto normal acessível.
    final_page.get_by_role("combobox").first.click()
    final_page.wait_for_timeout(500)
    final_page.get_by_text("Excel Spreadsheet (Excel 2007 or later)", exact=True).click()
    final_page.wait_for_timeout(300)
    final_page.keyboard.press("Escape")  # garante que a lista feche antes de continuar
    final_page.wait_for_timeout(1000)

    final_page.get_by_text("Advanced data options", exact=True).click(force=True)
    final_page.wait_for_timeout(1000)

    # Results -> "With visualizations options applied"
    final_page.get_by_text("With visualizations options applied", exact=True).click()

    # Data Values -> "Formatted" (já vem marcado por padrão; clicar de novo é inofensivo)
    final_page.get_by_text("Formatted", exact=True).click()

    # Number of rows to include -> "All results"
    final_page.get_by_text("All results", exact=True).click()

    with final_page.expect_download(timeout=60000) as download_info:
        final_page.get_by_role("button", name="Download", exact=True).click()

    download = download_info.value
    dest_path = config.DOWNLOAD_DIR / f"{base_id}_{int(time.time())}.xlsx"
    download.save_as(dest_path)
    logger.info("Arquivo baixado: %s", dest_path)
    return dest_path


def download_analitico_spreadsheet(final_page: Page, base_id: str) -> Path:
    """
    Rola até a planilha "Analítico", abre o menu de 3 pontinhos daquela
    planilha (fica quase invisível até o hover) e completa o download.
    """
    analitico_section = final_page.get_by_text("Analítico", exact=True).last
    analitico_section.scroll_into_view_if_needed()
    final_page.wait_for_timeout(1000)

    tile_button = _find_tile_actions_button(final_page, analitico_section)
    tile_button.hover()
    final_page.wait_for_timeout(300)
    tile_button.click(force=True)
    final_page.wait_for_timeout(1000)

    return _complete_download_dialog(final_page, base_id)


def download_numero_contratos_report(context: BrowserContext, page: Page, base: dict) -> Path:
    """Fluxo completo específico da base 'numero_contratos'."""
    final_page = open_acompanhamento_veiculos_analitico(context, page)
    apply_analitico_filters(final_page, base["filtros"])
    update_report_data(final_page)
    return download_analitico_spreadsheet(final_page, base["id"])


# --------------------------------------------------------------------------
# Fluxo dedicado - base "dias_sem_producao" (SLA - Última Atuação Comercial
# - Analítico), validado rodando de verdade contra o portal.
# --------------------------------------------------------------------------

def open_sla_analitico(context: BrowserContext, page: Page, base: dict) -> Page:
    """
    Navega até o dashboard "SLA Última Atuação Comercial - Analítico":
    Relatórios (hover) > Relatórios Gerenciais (abre pop-up com o catálogo)
    > card "Auto" > link "SLA - Última atuação comercial - Analítico"
    (dentro do card "SLA - Última atuação da loja") - abre outra pop-up já
    na aba certa ("SLA Analítico"), com o filtro "Referencia Month" já em
    "is this month" por padrão.
    """
    page.get_by_text("Relatórios", exact=True).first.hover()
    page.wait_for_timeout(500)

    with context.expect_page(timeout=15000) as popup_info:
        page.get_by_text("Relatórios Gerenciais", exact=True).first.click()
    catalogo = popup_info.value
    catalogo.wait_for_load_state("networkidle", timeout=20000)
    catalogo.wait_for_timeout(5000)

    catalogo.get_by_text("Auto", exact=False).first.click()
    catalogo.wait_for_timeout(3000)
    catalogo.wait_for_load_state("networkidle", timeout=15000)
    catalogo.wait_for_timeout(2000)

    link = catalogo.get_by_text(base["link_relatorio"], exact=True).first
    with context.expect_page(timeout=10000) as popup_info2:
        link.click(force=True)
    final_page = popup_info2.value

    final_page.wait_for_load_state("domcontentloaded", timeout=20000)
    final_page.wait_for_timeout(8000)
    return final_page


def verify_referencia_month_filter(final_page: Page):
    """
    Abre o painel de filtros e confere que "Referencia Month" já está em
    "is this month" (confirmado via querystring "Referencia+Month=this+
    month"). Só avisa no log se algum dia vier diferente - o clique para
    trocar esse filtro específico (um seletor de data relativa composto,
    tipo "is this" + "month") ainda não foi mapeado/validado.
    """
    final_page.get_by_text("filters", exact=False).first.click()
    final_page.wait_for_timeout(1500)

    if final_page.get_by_text("is this month", exact=True).count() == 0:
        logger.warning(
            "Filtro 'Referencia Month' não está em 'is this month' - ajuste "
            "manual pode ser necessário (fluxo de troca ainda não mapeado)."
        )
    # Não precisa fechar o painel de filtros - o botão "Update" continua
    # clicável normalmente com o painel aberto.


def download_sla_analitico_spreadsheet(final_page: Page, base_id: str) -> Path:
    """
    Localiza o botão "Tile actions" da tabela SLA Analítico usando como
    referência o cabeçalho de coluna "Cnpj Da Loja" - esse relatório não
    tem uma faixa de título separada acima da tabela (como "Analítico" em
    Número de Contratos), então usar o título da página como referência
    pega o botão errado (uma tabela de navegação interna escondida). A
    própria coluna da tabela funciona como ponto de referência correto.
    """
    referencia = final_page.get_by_text("Cnpj Da Loja", exact=True).first
    referencia.scroll_into_view_if_needed()
    final_page.wait_for_timeout(1000)

    tile_button = _find_tile_actions_button(final_page, referencia)
    tile_button.hover()
    final_page.wait_for_timeout(300)
    tile_button.click(force=True)
    final_page.wait_for_timeout(1000)

    return _complete_download_dialog(final_page, base_id)


def download_dias_sem_producao_report(context: BrowserContext, page: Page, base: dict) -> Path:
    """Fluxo completo específico da base 'dias_sem_producao'."""
    final_page = open_sla_analitico(context, page, base)
    verify_referencia_month_filter(final_page)
    update_report_data(final_page)
    return download_sla_analitico_spreadsheet(final_page, base["id"])


def download_base(base: dict, headless: bool = True) -> Path:
    """Executa o fluxo completo de download para uma base configurada."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            login(page)

            if base["id"] == "numero_contratos":
                path = download_numero_contratos_report(context, page, base)
            elif base["id"] == "dias_sem_producao":
                path = download_dias_sem_producao_report(context, page, base)
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
