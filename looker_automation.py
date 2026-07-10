"""
Automação do download dos relatórios no Looker via Playwright.

As 4 bases (numero_contratos, dias_sem_producao, meta_financiamento_seguro,
carteira_parceiros) têm cada uma seu próprio fluxo dedicado de navegação e
download, validado rodando de verdade contra o portal - não são mais um
esqueleto. O relatório é hospedado no Google Looker de verdade, embutido
dentro do WebAutorizador via janelas pop-up sucessivas.

Rodar `python looker_automation.py --base <id> --debug` abre o navegador
visível (headless=False) para acompanhar o fluxo no site real.
"""

import argparse
import logging
import re
import time
from datetime import date, timedelta
from pathlib import Path

import holidays
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


def _complete_download_dialog(final_page: Page, base_id: str, download_timeout_ms: int = 60000) -> Path:
    """
    A partir do menu "Tile actions" já aberto, clica em "Download data",
    seleciona o formato Excel, expande "Advanced data options" e marca as
    opções de exportação completa antes de baixar. Compartilhado por todas
    as bases que usam este mesmo fluxo de download do Looker.

    `download_timeout_ms` pode ser aumentado para bases com volumes maiores
    de dados (ex: Carteira e Parceiros, que baixa o ano inteiro) - o Looker
    demora mais para gerar o arquivo antes do download começar.
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

    with final_page.expect_download(timeout=download_timeout_ms) as download_info:
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


# --------------------------------------------------------------------------
# Fluxo dedicado - base "meta_financiamento_seguro" (Apuração Parceiro -
# Resumo > Bloco de Metas - Por Filial), validado rodando de verdade contra
# o portal.
# --------------------------------------------------------------------------

def _dia_util_mg(dia: date) -> bool:
    """Considera dia útil: seg-sex e não feriado nacional/estadual de MG."""
    if dia.weekday() >= 5:  # 5=sábado, 6=domingo
        return False
    feriados_mg = holidays.Brazil(state="MG", years=dia.year)
    return dia not in feriados_mg


def deve_usar_janela_curta_safra_mes(hoje: date | None = None) -> bool:
    """
    Regra da virada de mês de "Meta Financiamento e Seguro": se hoje é dia 1
    ou 2 do mês E o último dia do mês anterior não foi dia útil em MG
    (fim de semana ou feriado, calculado com a biblioteca `holidays`), a
    apuração de fim do mês anterior pode ainda não ter sido processada -
    nesse caso usamos "is in the last 3 days" em vez de "is this month" no
    filtro "Safra Mês", para não perder esses dados.
    """
    hoje = hoje or date.today()
    if hoje.day not in (1, 2):
        return False
    ultimo_dia_mes_anterior = date(hoje.year, hoje.month, 1) - timedelta(days=1)
    return not _dia_util_mg(ultimo_dia_mes_anterior)


def open_resumo_parceiro(context: BrowserContext, page: Page, base: dict) -> Page:
    """
    Navega até o dashboard "Apuração Parceiro - Resumo": Relatórios (hover)
    > Relatórios Gerenciais (abre pop-up com o catálogo) > card "Auto" >
    link "Resumo Apuração Parceiro 2.0" (dentro do card "Apuração Parceiro
    2.0") - abre outra pop-up já na aba "Resumo".
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


def apply_safra_mes_filter(final_page: Page):
    """
    Abre o painel de filtros e configura "Safra Mês" (sempre o primeiro
    filtro do painel - mesmo truque de regex "^is " usado em Número de
    Contratos). O valor padrão salvo no dashboard é "is in the last 6
    months", então SEMPRE precisamos trocar (diferente das outras bases,
    onde o padrão já vinha certo):
      - Caso normal: muda para "is this" + "month".
      - Caso especial (`deve_usar_janela_curta_safra_mes()`): muda para
        "is in the last" + "3" + "days", para não perder a apuração de fim
        do mês anterior quando não houve dia útil antes do dia 01.
    """
    final_page.get_by_text("filters", exact=False).first.click()
    final_page.wait_for_timeout(1500)

    final_page.get_by_text(re.compile(r"^is "), exact=False).first.click()
    final_page.wait_for_timeout(800)

    if deve_usar_janela_curta_safra_mes():
        # já abre em "is in the last" por padrão - só ajusta número e unidade
        final_page.locator('input[type="number"]').first.fill("3")
        final_page.wait_for_timeout(300)
        unidade = final_page.locator('input[type="text"][role="combobox"]').nth(1)
        unidade.click(force=True)
        final_page.wait_for_timeout(500)
        final_page.get_by_text("days", exact=True).first.click(force=True)
        logger.info("Safra Mês: usando janela curta 'is in the last 3 days' (virada de mês sem dia útil antes).")
    else:
        final_page.get_by_text("is in the last", exact=True).first.click(force=True)
        final_page.wait_for_timeout(800)
        final_page.get_by_text("is this", exact=True).first.click(force=True)
        final_page.wait_for_timeout(800)

    final_page.keyboard.press("Escape")
    final_page.wait_for_timeout(500)


def download_bloco_metas_spreadsheet(final_page: Page, base_id: str, secao_tabela: str) -> Path:
    """
    Rola até a seção "Bloco de Metas - Por Filial" (faixa de título cinza
    acima da tabela, igual ao padrão de "Analítico" em Número de
    Contratos) e completa o download.
    """
    secao = final_page.get_by_text(secao_tabela, exact=True).last
    secao.scroll_into_view_if_needed()
    final_page.wait_for_timeout(1000)

    tile_button = _find_tile_actions_button(final_page, secao)
    tile_button.hover()
    final_page.wait_for_timeout(300)
    tile_button.click(force=True)
    final_page.wait_for_timeout(1000)

    return _complete_download_dialog(final_page, base_id)


def download_meta_financiamento_seguro_report(context: BrowserContext, page: Page, base: dict) -> Path:
    """Fluxo completo específico da base 'meta_financiamento_seguro'."""
    final_page = open_resumo_parceiro(context, page, base)
    apply_safra_mes_filter(final_page)
    update_report_data(final_page)
    return download_bloco_metas_spreadsheet(final_page, base["id"], base["secao_tabela"])


# --------------------------------------------------------------------------
# Fluxo dedicado - base "carteira_parceiros" (Painel Carteira), validado
# rodando de verdade contra o portal.
# --------------------------------------------------------------------------

def open_painel_carteira(context: BrowserContext, page: Page, base: dict) -> Page:
    """
    Navega até o dashboard "Painel Carteira": Relatórios (hover) >
    Relatórios Gerenciais (abre pop-up com o catálogo) > card "Auto" > link
    "Carteira" (dentro do card "Carteira") - abre outra pop-up já na tabela
    certa (não tem abas, diferente das outras bases).
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

    # O card "Carteira" tem DOIS elementos com o mesmo texto: o título do
    # card (H2, não clicável) e o link de fato (SPAN) - usamos `.nth(1)`.
    link = catalogo.get_by_text(base["link_relatorio"], exact=True).nth(1)
    with context.expect_page(timeout=10000) as popup_info2:
        link.click(force=True)
    final_page = popup_info2.value

    final_page.wait_for_load_state("domcontentloaded", timeout=20000)
    final_page.wait_for_timeout(8000)
    return final_page


def apply_referencia_year_filter(final_page: Page):
    """
    Abre o painel de filtros e configura "Referência" (sempre o primeiro
    filtro do painel). O valor padrão salvo no dashboard é "is this month",
    então trocamos a segunda parte do seletor composto de "month" para
    "year" (mantendo o tipo "is this").
    """
    final_page.get_by_text("filters", exact=False).first.click()
    final_page.wait_for_timeout(1500)

    final_page.get_by_text(re.compile(r"^is "), exact=False).first.click()
    final_page.wait_for_timeout(800)

    unidade = final_page.locator('input[type="text"][role="combobox"]').nth(1)
    unidade.click(force=True)
    final_page.wait_for_timeout(500)
    final_page.get_by_text("year", exact=True).first.click(force=True)
    final_page.wait_for_timeout(800)

    final_page.keyboard.press("Escape")
    final_page.wait_for_timeout(500)


def download_carteira_spreadsheet(final_page: Page, base_id: str) -> Path:
    """
    Localiza o botão "Tile actions" da tabela usando o cabeçalho de coluna
    "Cnpj Da Loja" como referência (esse relatório não tem uma faixa de
    título separada acima da tabela, mesma situação de Dias sem Produção).
    O timeout de download é maior (120s) porque baixa o ano inteiro.
    """
    referencia = final_page.get_by_text("Cnpj Da Loja", exact=True).first
    referencia.scroll_into_view_if_needed()
    final_page.wait_for_timeout(1000)

    tile_button = _find_tile_actions_button(final_page, referencia)
    tile_button.hover()
    final_page.wait_for_timeout(300)
    tile_button.click(force=True)
    final_page.wait_for_timeout(1000)

    return _complete_download_dialog(final_page, base_id, download_timeout_ms=120000)


def download_carteira_parceiros_report(context: BrowserContext, page: Page, base: dict) -> Path:
    """Fluxo completo específico da base 'carteira_parceiros'."""
    final_page = open_painel_carteira(context, page, base)
    apply_referencia_year_filter(final_page)
    update_report_data(final_page)
    return download_carteira_spreadsheet(final_page, base["id"])


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
            elif base["id"] == "meta_financiamento_seguro":
                path = download_meta_financiamento_seguro_report(context, page, base)
            elif base["id"] == "carteira_parceiros":
                path = download_carteira_parceiros_report(context, page, base)
            else:
                raise ValueError(f"Base '{base['id']}' não tem fluxo de download implementado")
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
