"""
Automação do download dos relatórios no Looker via Playwright.

As 5 bases têm cada uma seu próprio fluxo dedicado de navegação e
download. O relatório é hospedado no Google Looker, embutido dentro do
WebAutorizador via janelas pop-up sucessivas.

Rodar `python looker_automation.py --base <id> --debug` abre o navegador
visível (headless=False) para acompanhar o fluxo no site real.

--------------------------------------------------------------------------
Hierarquia de seletores (ver `_click_com_prioridade`)
--------------------------------------------------------------------------
Todo clique/interação neste módulo segue, quando possível, esta ordem de
prioridade (do mais estável para o menos estável):

  1. id fixo do portal ASP.NET ou do Looker (ex: `#WFP2010_MPCNSRELGER`,
     `#qr-export-modal-download`, `#listbox-input-qr-export-modal-format`)
     - controle de servidor/chrome do próprio Looker, não editável pelo
       time que mantém os relatórios.
  2. atributo semântico (href com o id do dashboard, aria-label, role,
     type) - também não é conteúdo editável do relatório.
  3. texto visível - único recurso para elementos que o Looker renderiza
     como `<span>`/`<div>` de styled-components sem id/data-testid (comum
     nos menus/opções do visualizador Looker Studio) - texto fixo da
     INTERFACE do Looker (Google), então razoavelmente estável.

Quando existe um seletor de nível mais alto confirmado, ele é tentado
primeiro; o texto que já funcionava antes fica como fallback automático
(`_click_com_prioridade`), então uma mudança no id/atributo interno do
Looker não quebra a execução - só loga um aviso pedindo para revalidar.
"""

import argparse
import logging
import re
import time
from datetime import date, timedelta
from pathlib import Path
from typing import Callable

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

# --------------------------------------------------------------------------
# ids/atributos fixos do DOM real (não são texto de relatório, então não
# mudam se o time renomear um botão/aba) - centralizados aqui.
# --------------------------------------------------------------------------
ID_LINK_RELATORIOS_GERENCIAIS = "WFP2010_MPCNSRELGER"  # <a id="..."> no menu ASP.NET do portal, usado pelas 5 bases
ID_COMBOBOX_FORMATO_EXPORT = "listbox-input-qr-export-modal-format"  # modal de export do Looker (chrome, todas as bases)
ID_BOTAO_DOWNLOAD_EXPORT = "qr-export-modal-download"  # idem


def _click_com_prioridade(
    tentativas: list[tuple[str, Callable[[], Locator]]], *,
    force: bool = False, timeout: int | None = None, no_wait_after: bool = False,
) -> str:
    """
    Clica seguindo uma HIERARQUIA de seletores, do mais estável para o
    menos estável, em vez de depender só de texto visível (que pode ser
    renomeado no Looker/portal sem avisar e quebrar a automação
    silenciosamente) - ver a explicação completa no docstring do módulo.

    `tentativas`: lista de (descrição, função que retorna o Locator) em
    ordem de prioridade. Tenta cada uma, usando a primeira que encontrar
    >=1 elemento. Loga um AVISO (não erro) se precisou cair para um nível
    mais baixo que o primeiro - a execução não quebra, mas é sinal de que
    o seletor prioritário pode ter parado de bater. Levanta erro só se
    NENHUM nível funcionar. Retorna a descrição da tentativa que funcionou.

    `timeout=None` (padrão) usa o timeout padrão do Playwright (30s).

    `no_wait_after=True` pula a espera padrão do Playwright por uma
    "navegação agendada" depois do clique - útil para botões que disparam
    um download em vez de navegar (o Playwright fica parado em "waiting
    for scheduled navigations to finish" até estourar o timeout, mesmo com
    o download já capturado por `expect_download`).
    """
    ultimo_erro = None
    for i, (descricao, factory) in enumerate(tentativas):
        try:
            locator = factory()
            if locator.count() == 0:
                continue
            kwargs = {"timeout": timeout} if timeout is not None else {}
            if no_wait_after:
                kwargs["no_wait_after"] = True
            locator.first.click(force=force, **kwargs)
            if i > 0:
                logger.warning(
                    "Seletor prioritário de '%s' não encontrado - clicou via fallback "
                    "'%s' (nível %d). Vale revalidar o seletor de nível mais alto.",
                    tentativas[0][0], descricao, i + 1,
                )
            return descricao
        except Exception as ex:
            ultimo_erro = ex
            continue
    raise RuntimeError(
        f"Nenhum seletor da hierarquia funcionou (tentativas: {[d for d, _ in tentativas]})"
    ) from ultimo_erro


def _abrir_painel_filtros(final_page: Page):
    """
    Abre o painel de filtros (botão "NN filters", o número varia por
    aba) - texto é o único seletor disponível: o botão é um
    styled-components sem id/data-testid/aria-label (confirmado por
    inspeção ao vivo), mas "filters" é rótulo fixo do próprio Looker
    Studio, não editável pelo time. Compartilhado por todas as bases que
    têm painel de filtros (Número de Contratos, Dias sem Produção, Meta
    Financiamento e Seguro, Carteira e Parceiros).
    """
    final_page.get_by_text("filters", exact=False).first.click()
    final_page.wait_for_timeout(1500)


def _confirmar_dados_cadastrais_se_necessario(page: Page):
    """
    Depois do login, o portal às vezes exige passar por uma tela
    "Atualizar meus Dados Cadastrais" antes de liberar o acesso normal -
    sem isso, o menu "Relatórios" nunca aparece e a navegação trava logo no
    primeiro passo. Se a tela aparecer, clica em "Confirmar" mantendo os
    dados como já estão preenchidos; se não aparecer (caso normal), não
    faz nada. Tela condicional sem id/atributo estável conhecido - texto é
    o único seletor disponível.
    """
    page.wait_for_timeout(1500)
    confirmar = page.get_by_text("Confirmar", exact=True)
    if confirmar.count() == 0:
        return
    logger.info(
        "Tela 'Atualizar meus Dados Cadastrais' apareceu após o login - "
        "confirmando com os dados já preenchidos."
    )
    confirmar.first.click()
    page.wait_for_load_state("networkidle")


def login(page: Page):
    """
    Login no portal C6 Consig (WebAutorizador - página ASP.NET clássica,
    sem <label>):
      - Usuário: input#EUsuario_CAMPO
      - Senha:   input#ESenha_CAMPO
      - Entrar:  <a id="lnkEntrar"> (link com postback, não é um <button>)
    Todos ids fixos de controle ASP.NET - topo da hierarquia de seletores.

    O portal costuma mostrar um confirm() JS ("Usuário já autenticado em
    outra estação...") quando já existe uma sessão ativa - aceitamos
    automaticamente para forçar a nova sessão. Também pode exigir
    confirmar os dados cadastrais (ver `_confirmar_dados_cadastrais_se_necessario`).
    """
    page.on("dialog", lambda dialog: dialog.accept())

    page.goto(config.LOOKER_URL)
    page.locator("#EUsuario_CAMPO").fill(config.LOOKER_USER)
    page.locator("#ESenha_CAMPO").fill(config.LOOKER_PASSWORD)
    page.locator("#lnkEntrar").click()
    page.wait_for_load_state("networkidle")
    _confirmar_dados_cadastrais_se_necessario(page)


# --------------------------------------------------------------------------
# Navegação compartilhada pelas 5 bases: Relatórios (hover) > Relatórios
# Gerenciais (abre pop-up com o catálogo) > card "Auto" - idêntica nas 5,
# só o link clicado DENTRO do catálogo muda por base (ver cada open_*).
# --------------------------------------------------------------------------

def _abrir_catalogo_auto(context: BrowserContext, page: Page) -> Page:
    """
    Abre o menu "Relatórios" (hover), clica em "Relatórios Gerenciais"
    (pop-up com o catálogo do Looker) e entra no card "Auto". Retorna a
    Page do catálogo já dentro da seção "Auto", pronta para o caller
    clicar no link do relatório específico daquela base.

    "Relatórios Gerenciais" tem id fixo (`WFP2010_MPCNSRELGER`, controle
    ASP.NET do menu do portal) - usado como topo da hierarquia, com o texto
    como fallback automático. O card "Auto" é um `<h3>` sem id/data-testid
    - texto é o único seletor possível.
    """
    # "Relatórios" só precisa de hover (não passa por _click_com_prioridade) -
    # o `role="button"` do link não tem name acessível único.
    page.get_by_text("Relatórios", exact=True).first.hover()
    page.wait_for_timeout(500)  # tempo do dropdown CSS abrir antes do próximo elemento ficar clicável

    with context.expect_page(timeout=15000) as popup_info:
        _click_com_prioridade([
            (f"id #{ID_LINK_RELATORIOS_GERENCIAIS}", lambda: page.locator(f"#{ID_LINK_RELATORIOS_GERENCIAIS}")),
            ("texto 'Relatórios Gerenciais'", lambda: page.get_by_text("Relatórios Gerenciais", exact=True)),
        ])
    catalogo = popup_info.value
    catalogo.wait_for_load_state("networkidle", timeout=20000)
    catalogo.wait_for_timeout(5000)  # conteúdo do catálogo ainda renderiza após "networkidle" (SPA)

    _click_com_prioridade([
        ("texto 'Auto' (card)", lambda: catalogo.get_by_text("Auto", exact=False)),
    ])
    catalogo.wait_for_timeout(3000)
    catalogo.wait_for_load_state("networkidle", timeout=15000)
    catalogo.wait_for_timeout(2000)  # idem - conteúdo da seção "Auto" ainda renderiza após "networkidle"

    return catalogo


def _clicar_link_relatorio_e_abrir_popup(
    context: BrowserContext, catalogo: Page, texto_link: str, *, indice: int = 0,
) -> Page:
    """
    Clica no link de um relatório dentro do catálogo (card específico da
    base) e retorna a Page da pop-up que abre com o dashboard final.

    Alguns cards do catálogo têm DOIS elementos com o mesmo texto: o
    título do card (não clicável, geralmente um heading) e, mais abaixo,
    o link de fato. Quando `indice=1` (ver callers), tentamos primeiro
    isolar o elemento com role="link" (mais confiável que contar posição
    às cegas, já que só o link de verdade tem esse role) e só caímos para
    ".nth(1)" por posição se o role não resolver sozinho.
    """
    if indice == 0:
        tentativas = [
            (f"texto '{texto_link}'", lambda: catalogo.get_by_text(texto_link, exact=True)),
        ]
    else:
        tentativas = [
            (f"role link '{texto_link}'", lambda: catalogo.get_by_role("link", name=texto_link, exact=True)),
            (f"texto '{texto_link}' (posição {indice})", lambda: catalogo.get_by_text(texto_link, exact=True).nth(indice)),
        ]

    with context.expect_page(timeout=10000) as popup_info:
        _click_com_prioridade(tentativas, force=True)
    final_page = popup_info.value
    catalogo.close()

    # O dashboard final faz polling contínuo em segundo plano, então
    # "networkidle" nunca conclui aqui - usamos espera fixa.
    final_page.wait_for_load_state("domcontentloaded", timeout=20000)
    final_page.wait_for_timeout(8000)
    return final_page


# --------------------------------------------------------------------------
# Fluxo dedicado - base "numero_contratos" (Acompanhamento Veículos > Analítico)
# --------------------------------------------------------------------------

def open_acompanhamento_veiculos_analitico(context: BrowserContext, page: Page) -> Page:
    """
    Navega até o dashboard "Acompanhamento Veículos" e retorna a Page do
    Looker onde ele foi aberto (nova aba/pop-up).

    `_abrir_catalogo_auto` cuida de "Relatórios" > "Relatórios Gerenciais" >
    card "Auto" (compartilhado com as outras bases). O card "Acompanhamento
    Veículos" tem DOIS elementos com o mesmo texto (título + link de fato)
    - `_clicar_link_relatorio_e_abrir_popup` isola pelo role="link"
    primeiro, com posição (`.nth(1)`) como fallback.
    """
    catalogo = _abrir_catalogo_auto(context, page)
    return _clicar_link_relatorio_e_abrir_popup(context, catalogo, "Acompanhamento Veículos", indice=1)


HREF_DASHBOARD_ANALITICO_NUMERO_CONTRATOS = "corp_consignado_embed::00087"  # id fixo do dashboard "Analítico" (numero_contratos)


def apply_analitico_filters(final_page: Page, filtros: dict):
    """
    Clica na aba "Analítico" e configura, no painel de filtros:
      - "Tipo Exibição" -> mantém somente a opção informada (ex: "Valor")
      - "Dt Relatorio Date" -> período relativo (ex: "Last 90 Days")

    O dashboard abre inicialmente na aba "Produção" - só serve de ponto de
    partida até clicarmos em "Analítico".

    A barra de abas é navegação nativa do Looker Studio, renderizada como
    `<a href="/embed/dashboards/{HREF_DASHBOARD_ANALITICO_NUMERO_CONTRATOS}?...">`
    - href é o seletor mais estável (não depende de texto/ícone). Cada aba
    tem um ícone antes do rótulo (ex: "❗ Analítico"), então o texto NÃO é
    exatamente "Analítico" - o fallback por texto usa regex que exige
    terminar em "Analítico" (exclui "Analítico Aprov no dia").
    """
    logger.info("Aguardando a aba 'Analítico' aparecer no dashboard...")
    aba_analitico = final_page.locator(f'a[href*="{HREF_DASHBOARD_ANALITICO_NUMERO_CONTRATOS}"]')

    try:
        aba_analitico.first.wait_for(state="visible", timeout=15000)
    except Exception:
        logger.warning(
            "Aba 'Analítico' não encontrada pelo href do dashboard - tentando pelo texto (fallback)."
        )
        aba_analitico = final_page.get_by_text(re.compile(r"Analítico$", re.IGNORECASE))
        try:
            aba_analitico.first.wait_for(state="visible", timeout=45000)
        except Exception:
            _logar_diagnostico_aba_analitico_nao_encontrada(final_page)
            raise

    aba_analitico.first.click()
    logger.info("Aba 'Analítico' selecionada.")
    final_page.wait_for_timeout(5000)

    _abrir_painel_filtros(final_page)

    # --- Tipo Exibição ---
    # O chip de valor do filtro mostra o texto atual (ex: "is Qtde" ou
    # "is Valor") - tem um `id`, mas é um UUID gerado por instância de
    # render (confirmado por inspeção ao vivo: muda a cada carregamento),
    # não um id fixo reaproveitável. "Tipo Exibição" é sempre o primeiro
    # filtro do painel, então o primeiro chip que começa com "is " é o
    # dele - continua sendo o seletor mais confiável disponível aqui.
    final_page.get_by_text(re.compile(r"^is "), exact=False).first.click()
    final_page.wait_for_timeout(500)
    final_page.get_by_text(filtros["tipo_exibicao"], exact=True).click()
    final_page.wait_for_timeout(500)

    # --- Dt Relatorio Date --- (ver `_alterar_periodo_dt_relatorio`)
    _alterar_periodo_dt_relatorio(final_page, filtros["periodo_dt_relatorio"])
    logger.info("Filtros da aba 'Analítico' configurados com sucesso.")


def apply_analitico_filters_ano_fechado(final_page: Page, tipo_exibicao: str, ano: int):
    """
    Variante de `apply_analitico_filters` usada só na carga única de um
    ano fechado (ex: 2025), fora do fluxo diário normal: mesmos passos de
    selecionar a aba "Analítico" e "Tipo Exibição", mas troca "Dt
    Relatorio Date" para um intervalo customizado (1º de janeiro a 31 de
    dezembro do ano informado) em vez de um preset relativo - ver
    `_selecionar_intervalo_customizado_dt_relatorio`.
    """
    logger.info("Aguardando a aba 'Analítico' aparecer no dashboard...")
    aba_analitico = final_page.locator(f'a[href*="{HREF_DASHBOARD_ANALITICO_NUMERO_CONTRATOS}"]')
    try:
        aba_analitico.first.wait_for(state="visible", timeout=15000)
    except Exception:
        logger.warning(
            "Aba 'Analítico' não encontrada pelo href do dashboard - tentando pelo texto (fallback)."
        )
        aba_analitico = final_page.get_by_text(re.compile(r"Analítico$", re.IGNORECASE))
        try:
            aba_analitico.first.wait_for(state="visible", timeout=45000)
        except Exception:
            _logar_diagnostico_aba_analitico_nao_encontrada(final_page)
            raise

    aba_analitico.first.click()
    logger.info("Aba 'Analítico' selecionada.")
    final_page.wait_for_timeout(5000)

    _abrir_painel_filtros(final_page)

    final_page.get_by_text(re.compile(r"^is "), exact=False).first.click()
    final_page.wait_for_timeout(500)
    final_page.get_by_text(tipo_exibicao, exact=True).click()
    final_page.wait_for_timeout(500)

    _selecionar_intervalo_customizado_dt_relatorio(final_page, f"{ano}/01/01", f"{ano}/12/31")
    logger.info("Filtros da aba 'Analítico' configurados para o ano fechado %d.", ano)


def _logar_diagnostico_aba_analitico_nao_encontrada(final_page: Page):
    """
    Diagnóstico de apoio quando nem o href nem o texto encontram a aba
    "Analítico" - loga URL/título atuais, todos os elementos cujo texto
    contém "Analítico" (para conferir como o Looker renderizou o rótulo
    daquela vez) e salva um screenshot em `downloads/` antes de propagar o
    erro original para quem chamou.
    """
    logger.error("Aba 'Analítico' não foi encontrada (nem por href, nem por texto).")
    logger.error("URL atual do dashboard: %s", final_page.url)
    try:
        logger.error("Título atual da página: %s", final_page.title())
    except Exception:
        pass

    elementos_analitico = final_page.get_by_text(re.compile(r"Analítico", re.IGNORECASE))
    logger.error("Quantidade de elementos contendo 'Analítico': %d", elementos_analitico.count())
    for i in range(elementos_analitico.count()):
        try:
            logger.error("Analítico[%d] = %r", i, elementos_analitico.nth(i).inner_text())
        except Exception:
            pass

    try:
        screenshot_path = config.DOWNLOAD_DIR / "debug_analitico_falha.png"
        final_page.screenshot(path=str(screenshot_path), full_page=False)
        logger.error("Screenshot de diagnóstico salvo em: %s", screenshot_path)
    except Exception as ex:
        logger.warning("Não foi possível salvar screenshot de diagnóstico: %s", ex)


def _alterar_periodo_dt_relatorio(final_page: Page, valor_alvo: str):
    """
    Troca o filtro "Dt Relatorio Date" para o preset informado (ex:
    "Last 90 Days", "Year To Date"). É um dropdown com abas
    "Presets"/"Custom" e uma lista fixa de opções clicáveis (Today,
    Yesterday, Last 7 Days, ..., Year To Date, More...) - clicar no chip do
    valor atual abre o dropdown, então basta clicar no preset alvo pelo
    texto exato.
    """
    logger.info("Alterando 'Dt Relatorio Date' para '%s'...", valor_alvo)

    # Localiza o chip atual do filtro (mostra o preset em vigor, ex: "Last 30 Days").
    chip_periodo = final_page.get_by_text(re.compile(r"^Last \d+ Days?$", re.IGNORECASE))
    chip_periodo.first.wait_for(state="visible", timeout=30000)
    chip_periodo.first.click(force=True)
    final_page.wait_for_timeout(800)

    # Dropdown abre com a lista de presets - clica na opção alvo pelo texto exato.
    opcao_preset = final_page.get_by_text(valor_alvo, exact=True)
    opcao_preset.first.wait_for(state="visible", timeout=15000)
    opcao_preset.first.click()
    final_page.wait_for_timeout(500)

    # Fecha o dropdown, caso ainda esteja aberto.
    final_page.keyboard.press("Escape")
    final_page.wait_for_timeout(500)

    if final_page.get_by_text(valor_alvo, exact=True).count() == 0:
        logger.warning(
            "Após tentar trocar 'Dt Relatorio Date' para '%s', "
            "o chip não confirmou esse valor. "
            "Revalidar visualmente com --debug.",
            valor_alvo,
        )
    else:
        logger.info("'Dt Relatorio Date' configurado para '%s'.", valor_alvo)


def _selecionar_intervalo_customizado_dt_relatorio(final_page: Page, data_inicio: str, data_fim: str):
    """
    Troca o filtro "Dt Relatorio Date" para um intervalo de datas livre
    (aba "Custom" do dropdown, ao lado de "Presets") - usado só para a
    carga única de um ano fechado (ex: 2025), fora do fluxo diário normal
    (que usa `_alterar_periodo_dt_relatorio` com um preset).

    `data_inicio`/`data_fim` no formato "YYYY/MM/DD" (ex: "2025/01/01").
    Confirmado ao vivo contra o portal em 25/08/2026: por baixo do texto
    exibido existem `input[data-testid="date-from-text-input"]` e
    `input[data-testid="date-to-text-input"]` editáveis diretamente
    (`.fill()`), sem precisar abrir o calendário visual.
    """
    logger.info("Alterando 'Dt Relatorio Date' para intervalo customizado %s - %s...", data_inicio, data_fim)

    chip_periodo = final_page.get_by_text(re.compile(r"^Last \d+ Days?$", re.IGNORECASE))
    chip_periodo.first.wait_for(state="visible", timeout=30000)
    chip_periodo.first.click(force=True)
    final_page.wait_for_timeout(800)

    final_page.get_by_text("Custom", exact=True).click()
    final_page.wait_for_timeout(800)

    final_page.locator('input[data-testid="date-from-text-input"]').fill(data_inicio)
    final_page.wait_for_timeout(300)
    final_page.locator('input[data-testid="date-to-text-input"]').fill(data_fim)
    final_page.wait_for_timeout(300)
    final_page.keyboard.press("Enter")
    final_page.wait_for_timeout(800)
    final_page.keyboard.press("Escape")
    final_page.wait_for_timeout(500)

    valor_esperado = f"{data_inicio} - {data_fim}"
    if final_page.get_by_text(valor_esperado, exact=False).count() == 0:
        logger.warning(
            "Após tentar trocar 'Dt Relatorio Date' para '%s', "
            "o chip não confirmou esse valor. Revalidar visualmente com --debug.",
            valor_esperado,
        )
    else:
        logger.info("'Dt Relatorio Date' configurado para '%s'.", valor_esperado)


def update_report_data(final_page: Page):
    """
    Clica no botão 'Update' para atualizar os dados do relatório. Usa
    `aria-labelledby="page-freshness-indicator"` - atributo semântico
    fixo, topo da hierarquia de seletores.
    """
    final_page.locator('button[aria-labelledby="page-freshness-indicator"]').click()
    # espera fixa: networkidle não é confiável nesse dashboard (polling
    # contínuo em segundo plano) - não há uma condição específica de DOM
    # que sinalize "dados atualizados", então o sleep aqui é necessário.
    final_page.wait_for_timeout(5000)


def _find_tile_actions_button(final_page: Page, near: Locator, titulo_tile: str | None = None) -> Locator:
    """
    Encontra o botão "Tile actions" (3 pontinhos) de um tile específico.

    Hierarquia de tentativas:
      1. Quando `titulo_tile` é informado (a referência usada para rolar
         até o tile É plausivelmente o próprio título dele - caso de
         "Analítico"/`secao_tabela`, que ficam numa faixa de título acima
         da tabela): tenta por `role="button"` + `aria-label` contendo
         `titulo_tile` (o botão real tem `aria-label="{Título do tile} -
         Tile actions"`) - atributo semântico, não depende de posição.
      2. Heurística geométrica: o ícone svg (`ICON_MORE_VERT_PATH`) é
         reaproveitado ~40x na página (colunas do crosstab + menu global) -
         filtramos por altura do botão (tile = 24px, diferente do menu
         global e dos ícones de coluna) e pegamos o mais próximo em Y do
         elemento `near`. Usada quando não há `titulo_tile` confiável (ex:
         SLA/Carteira) ou como fallback se a tentativa 1 não achar nada.
    """
    if titulo_tile:
        candidato = final_page.get_by_role(
            "button", name=re.compile(re.escape(titulo_tile), re.IGNORECASE),
        )
        if candidato.count() > 0:
            return candidato.first
        logger.warning(
            "Botão 'Tile actions' não encontrado por aria-label contendo '%s' - "
            "usando heurística geométrica (posição) como fallback.", titulo_tile,
        )

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
    as bases que usam este fluxo de download do Looker.

    O modal de export é chrome fixo do Looker (mesma estrutura nas 5
    bases) - combobox de formato e botão "Download" têm id fixo
    (`#listbox-input-qr-export-modal-format`, `#qr-export-modal-download`),
    com role como fallback. As demais opções do modal são rótulos fixos da
    interface do Looker sem id - texto é o único seletor possível.

    `download_timeout_ms` pode ser aumentado para bases com volumes maiores
    (ex: Carteira e Parceiros, que baixa o ano inteiro).
    """
    final_page.get_by_text("Download data", exact=True).click()
    final_page.wait_for_timeout(1500)

    _click_com_prioridade([
        (f"id #{ID_COMBOBOX_FORMATO_EXPORT}", lambda: final_page.locator(f"#{ID_COMBOBOX_FORMATO_EXPORT}")),
        ("role combobox (formato)", lambda: final_page.get_by_role("combobox")),
    ])
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
        _click_com_prioridade([
            (f"id #{ID_BOTAO_DOWNLOAD_EXPORT}", lambda: final_page.locator(f"#{ID_BOTAO_DOWNLOAD_EXPORT}")),
            ("role button 'Download'", lambda: final_page.get_by_role("button", name="Download", exact=True)),
        ], no_wait_after=True)

    download = download_info.value
    dest_path = config.DOWNLOAD_DIR / f"{base_id}_{int(time.time())}.xlsx"
    download.save_as(dest_path)
    logger.info("Arquivo baixado: %s", dest_path)
    return dest_path


def download_analitico_spreadsheet(final_page: Page, base_id: str) -> Path:
    """
    Rola até a planilha "Analítico", abre o menu de 3 pontinhos daquela
    planilha (fica quase invisível até o hover) e completa o download.
    Passa "Analítico" como dica de título para `_find_tile_actions_button`
    tentar primeiro por aria-label (o tile real chama-se "Digitação
    Analítico", então "Analítico" bate como substring) antes de cair para
    a heurística geométrica.
    """
    analitico_section = final_page.get_by_text("Analítico", exact=True).last
    analitico_section.scroll_into_view_if_needed()
    final_page.wait_for_timeout(1000)

    tile_button = _find_tile_actions_button(final_page, analitico_section, titulo_tile="Analítico")
    tile_button.hover()
    final_page.wait_for_timeout(300)
    tile_button.click(force=True)
    final_page.wait_for_timeout(1000)

    # timeout maior que o padrão (60s): filtro "Dt Relatorio Date" = "Year To
    # Date" exporta o ano inteiro, e o Looker demora mais para gerar o arquivo.
    return _complete_download_dialog(final_page, base_id, download_timeout_ms=240000)


def download_numero_contratos_report(context: BrowserContext, page: Page, base: dict) -> Path:
    """Fluxo completo específico da base 'numero_contratos'."""
    final_page = open_acompanhamento_veiculos_analitico(context, page)
    apply_analitico_filters(final_page, base["filtros"])
    update_report_data(final_page)
    path = download_analitico_spreadsheet(final_page, base["id"])
    final_page.close()
    return path


# --------------------------------------------------------------------------
# Fluxo dedicado - base "dias_sem_producao" (SLA - Última Atuação Comercial
# - Analítico)
# --------------------------------------------------------------------------

def open_sla_analitico(context: BrowserContext, page: Page, base: dict) -> Page:
    """
    Navega até o dashboard "SLA Última Atuação Comercial - Analítico":
    `_abrir_catalogo_auto` (Relatórios > Relatórios Gerenciais > card
    "Auto") + link "SLA - Última atuação comercial - Analítico" (dentro do
    card "SLA - Última atuação da loja") - abre outra pop-up já na aba
    certa ("SLA Analítico"), com o filtro "Referencia Month" já em "is
    this month" por padrão.
    """
    catalogo = _abrir_catalogo_auto(context, page)
    return _clicar_link_relatorio_e_abrir_popup(context, catalogo, base["link_relatorio"])


def apply_referencia_month_ano(final_page: Page, ano: int):
    """
    Troca "Referencia Month" (mesmo widget de "Safra Mês") para "is in the
    year {ano}" - usado tanto no fluxo diário normal da base
    "dias_sem_producao" (com `ano = date.today().year`, para buscar o ano
    corrente inteiro a cada execução, não só o mês atual) quanto na carga
    única de um ano fechado de histórico (ex: 2025, via
    `download_ano_fechado_report`). Confirmado ao vivo contra o portal em
    25/08/2026 - ver `_selecionar_ano`.

    O padrão salvo no dashboard é "is this month" - por isso sempre
    precisamos trocar, mesmo no fluxo diário.
    """
    _abrir_painel_filtros(final_page)
    final_page.get_by_text("is this month", exact=True).click(force=True)
    final_page.wait_for_timeout(800)
    _selecionar_ano(final_page, ano)
    final_page.keyboard.press("Escape")
    final_page.wait_for_timeout(500)


def download_sla_analitico_spreadsheet(final_page: Page, base_id: str, download_timeout_ms: int = 60000) -> Path:
    """
    Localiza o botão "Tile actions" da tabela SLA Analítico usando como
    referência o cabeçalho de coluna "Cnpj Da Loja" - esse relatório não
    tem uma faixa de título separada acima da tabela (como "Analítico" em
    Número de Contratos), então usar o título da página como referência
    pega o botão errado (uma tabela de navegação interna escondida). A
    própria coluna da tabela funciona como ponto de referência correto.

    Sem `titulo_tile`: "Cnpj Da Loja" é um cabeçalho de COLUNA, não o
    título do tile - não faz sentido tentar aria-label com esse texto
    (sempre falharia e só geraria log de aviso falso a cada execução) -
    mantido na heurística geométrica, já baseada em atributo (svg path do
    ícone) + posição, que é a estratégia correta quando não há faixa de
    título disponível para usar como sinal semântico.

    `download_timeout_ms` pode ser aumentado para downloads maiores (ex:
    carga de ano fechado inteiro, ver `download_ano_fechado_report`).
    """
    referencia = final_page.get_by_text("Cnpj Da Loja", exact=True).first
    referencia.scroll_into_view_if_needed(timeout=90000)
    final_page.wait_for_timeout(1000)

    tile_button = _find_tile_actions_button(final_page, referencia)
    tile_button.hover()
    final_page.wait_for_timeout(300)
    tile_button.click(force=True)
    final_page.wait_for_timeout(1000)

    return _complete_download_dialog(final_page, base_id, download_timeout_ms=download_timeout_ms)


def download_dias_sem_producao_report(context: BrowserContext, page: Page, base: dict) -> Path:
    """
    Fluxo completo específico da base 'dias_sem_producao'. Busca o ANO
    CORRENTE inteiro a cada execução (filtro "Referencia Month" -> "is in
    the year", ver `apply_referencia_month_ano`), não só o mês atual -
    igual ao padrão já usado por Número de Contratos (Year To Date) e
    Carteira e Parceiros (Referência = Este Ano). O acúmulo e a marcação
    de cor continuam cuidando de não duplicar nem perder linha (ver
    `data_processor._process_dias_sem_producao`).
    """
    final_page = open_sla_analitico(context, page, base)
    apply_referencia_month_ano(final_page, date.today().year)
    update_report_data(final_page)
    path = download_sla_analitico_spreadsheet(final_page, base["id"])
    final_page.close()
    return path


# --------------------------------------------------------------------------
# Fluxo dedicado - base "meta_financiamento_seguro" (Apuração Parceiro -
# Resumo > Bloco de Metas - Por Filial)
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
    Navega até o dashboard "Apuração Parceiro - Resumo": `_abrir_catalogo_auto`
    (Relatórios > Relatórios Gerenciais > card "Auto") + link "Resumo
    Apuração Parceiro 2.0" (dentro do card "Apuração Parceiro 2.0") - abre
    outra pop-up já na aba "Resumo".
    """
    catalogo = _abrir_catalogo_auto(context, page)
    return _clicar_link_relatorio_e_abrir_popup(context, catalogo, base["link_relatorio"])


def _selecionar_ano(final_page: Page, ano: int):
    """
    Troca o filtro de período (Safra Mês ou Referencia Month - mesmo
    widget) para "is in the year {ano}" - usado tanto no fluxo diário/
    mensal de Meta Financiamento e Seguro e Dias sem Produção (ano
    corrente, buscando o ano inteiro a cada execução) quanto na carga
    única de um ano fechado de histórico (ex: 2025).

    Confirmado ao vivo contra o portal em 25/08/2026: ao abrir a lista de
    operadores existe a opção nativa "is in the year", que abre um campo
    numérico pré-preenchido com o ano corrente - basta sobrescrever com o
    ano desejado. O combobox de operador é sempre o primeiro
    `input[role="combobox"]` do mini-widget que abre ao clicar no chip do
    filtro - o texto atual dele varia por base ("is in the last" no padrão
    de Safra Mês/Meta Financiamento e Seguro, "is this" no padrão de
    Referencia Month/Dias sem Produção), por isso é localizado pela
    posição/role em vez de um texto fixo.
    """
    final_page.locator('input[type="text"][role="combobox"]').first.click(force=True)
    final_page.wait_for_timeout(800)
    final_page.get_by_role("option", name="is in the year", exact=True).click(force=True)
    final_page.wait_for_timeout(800)
    final_page.locator('input[type="number"]').first.fill(str(ano))
    final_page.wait_for_timeout(500)


def apply_safra_mes_filter(final_page: Page, periodo: str = "mes_atual", ano: int | None = None):
    """
    Abre o painel de filtros e configura "Safra Mês" (sempre o primeiro
    filtro do painel - mesmo truque de regex "^is " usado em Número de
    Contratos). O valor padrão salvo no dashboard é "is in the last 6
    months", então SEMPRE precisamos trocar (diferente das outras bases,
    onde o padrão já vinha certo).

    `periodo`:
      - "mes_atual" (padrão): mês corrente - usado pela base
        "comissao_a_vista".
          - Caso normal: muda para "is this" + "month".
          - Caso especial (`deve_usar_janela_curta_safra_mes()`): muda para
            "is in the last" + "3" + "days", para não perder a apuração de
            fim do mês anterior quando não houve dia útil antes do dia 01.
      - "ano": ano específico (`ano`, obrigatório nesse caso) - "is in the
        year {ano}" (ver `_selecionar_ano`). Usado pelo fluxo diário/
        mensal de "meta_financiamento_seguro" (ano corrente, buscando o
        ano inteiro a cada execução) e pela carga única de histórico
        fechado (ver `download_ano_fechado_report`).
    """
    _abrir_painel_filtros(final_page)

    final_page.get_by_text(re.compile(r"^is "), exact=False).first.click()
    final_page.wait_for_timeout(800)

    if periodo == "ano":
        if ano is None:
            raise ValueError("periodo='ano' exige o parâmetro 'ano'.")
        _selecionar_ano(final_page, ano)
        final_page.keyboard.press("Escape")
        final_page.wait_for_timeout(500)
        return

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


def download_bloco_metas_spreadsheet(
    final_page: Page, base_id: str, secao_tabela: str, download_timeout_ms: int = 60000,
) -> Path:
    """
    Rola até a seção "Bloco de Metas - Por Filial" (faixa de título cinza
    acima da tabela, igual ao padrão de "Analítico" em Número de
    Contratos) e completa o download. Passa `secao_tabela` como dica de
    título para `_find_tile_actions_button` tentar primeiro por
    aria-label (mesmo padrão de `download_analitico_spreadsheet`).

    `download_timeout_ms` pode ser aumentado para downloads maiores (ex:
    carga de ano fechado inteiro, ver `download_ano_fechado_report`).
    """
    secao = final_page.get_by_text(secao_tabela, exact=True).last
    secao.scroll_into_view_if_needed()
    final_page.wait_for_timeout(1000)

    tile_button = _find_tile_actions_button(final_page, secao, titulo_tile=secao_tabela)
    tile_button.hover()
    final_page.wait_for_timeout(300)
    tile_button.click(force=True)
    final_page.wait_for_timeout(1000)

    return _complete_download_dialog(final_page, base_id, download_timeout_ms=download_timeout_ms)


def download_meta_financiamento_seguro_report(context: BrowserContext, page: Page, base: dict) -> Path:
    """
    Fluxo completo da base 'meta_financiamento_seguro'. Busca o ANO
    CORRENTE inteiro a cada execução (filtro "Safra Mês" -> "is in the
    year", ver `apply_safra_mes_filter`), não só o mês atual - mesmo
    padrão já usado por Dias sem Produção e Número de Contratos (Year To
    Date). O acúmulo e a marcação de cor continuam cuidando de não
    duplicar nem perder linha (ver
    `data_processor._process_meta_financiamento_seguro`).
    """
    final_page = open_resumo_parceiro(context, page, base)
    apply_safra_mes_filter(final_page, periodo="ano", ano=date.today().year)
    update_report_data(final_page)
    # timeout maior que o padrão (60s): buscar o ano inteiro demora mais que
    # o mês atual - mesmo motivo do timeout maior de Comissão à Vista.
    path = download_bloco_metas_spreadsheet(
        final_page, base["id"], base["secao_tabela"], download_timeout_ms=240000,
    )
    final_page.close()
    return path


# --------------------------------------------------------------------------
# Fluxo dedicado - base "carteira_parceiros" (Painel Carteira)
# --------------------------------------------------------------------------

def open_painel_carteira(context: BrowserContext, page: Page, base: dict) -> Page:
    """
    Navega até o dashboard "Painel Carteira": `_abrir_catalogo_auto`
    (Relatórios > Relatórios Gerenciais > card "Auto") + link "Carteira"
    (dentro do card "Carteira") - abre outra pop-up já na tabela certa
    (não tem abas, diferente das outras bases). O card "Carteira" tem
    DOIS elementos com o mesmo texto (título + link de fato), mesmo caso
    de "Acompanhamento Veículos" - ver `_clicar_link_relatorio_e_abrir_popup`.
    """
    catalogo = _abrir_catalogo_auto(context, page)
    return _clicar_link_relatorio_e_abrir_popup(context, catalogo, base["link_relatorio"], indice=1)


def apply_referencia_year_filter(final_page: Page):
    """
    Abre o painel de filtros e configura "Referência" (sempre o primeiro
    filtro do painel). O valor padrão salvo no dashboard é "is this month",
    então trocamos a segunda parte do seletor composto de "month" para
    "year" (mantendo o tipo "is this").
    """
    _abrir_painel_filtros(final_page)

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
    Localiza o botão "Tile actions" usando o cabeçalho de coluna "Cnpj Da
    Loja" como referência (mesma situação de Dias sem Produção, sem
    `titulo_tile` - ver `download_sla_analitico_spreadsheet`). Timeout de
    download maior (240s) porque baixa o ano inteiro.
    """
    referencia = final_page.get_by_text("Cnpj Da Loja", exact=True).first
    referencia.scroll_into_view_if_needed(timeout=90000)
    final_page.wait_for_timeout(1000)

    tile_button = _find_tile_actions_button(final_page, referencia)
    tile_button.hover()
    final_page.wait_for_timeout(300)
    tile_button.click(force=True)
    final_page.wait_for_timeout(1000)

    return _complete_download_dialog(final_page, base_id, download_timeout_ms=240000)


def download_carteira_parceiros_report(context: BrowserContext, page: Page, base: dict) -> Path:
    """Fluxo completo específico da base 'carteira_parceiros'."""
    final_page = open_painel_carteira(context, page, base)
    apply_referencia_year_filter(final_page)
    update_report_data(final_page)
    path = download_carteira_spreadsheet(final_page, base["id"])
    final_page.close()
    return path


# --------------------------------------------------------------------------
# Fluxo dedicado - base "comissao_a_vista" (Apuração Comissão à Vista,
# dentro do mesmo card "Apuração Parceiro 2.0" de Meta Financiamento e
# Seguro, porém um link diferente dentro dele).
#
#   - Link do catálogo: "Apuração Comissão À Vista" ("À" maiúsculo -
#     `exact=True` é sensível a isso).
#   - Filtro de período: "Safra Mês" (mesmo widget de Meta Financiamento e
#     Seguro) - reaproveita `apply_safra_mes_filter`, inclusive a exceção
#     de janela curta na virada de mês.
#   - `secao_tabela` (config.py) é "Analítico", mesmo padrão de Número de
#     Contratos.
# Ver a chave de comparação em `regras["chave_comparacao"]` (config.py) e o
# filtro da linha de totais em `data_processor._process_comissao_a_vista`.
# --------------------------------------------------------------------------

def open_apuracao_comissao_a_vista(context: BrowserContext, page: Page, base: dict) -> Page:
    """
    Navega até o dashboard "Apuração Comissão à Vista": `_abrir_catalogo_auto`
    (Relatórios > Relatórios Gerenciais > card "Auto") + link "Apuração
    Comissão à Vista" (dentro do card "Apuração Parceiro 2.0") - abre
    outra pop-up com o dashboard final. Mesma estrutura de navegação de
    `open_resumo_parceiro` (mesmo card no catálogo), só muda o link
    clicado dentro dele.
    """
    catalogo = _abrir_catalogo_auto(context, page)
    return _clicar_link_relatorio_e_abrir_popup(context, catalogo, base["link_relatorio"])


def download_comissao_a_vista_spreadsheet(
    final_page: Page, base_id: str, secao_tabela: str, download_timeout_ms: int = 60000,
) -> Path:
    """
    Rola até a seção/tabela do relatório (referência configurada em
    `secao_tabela`, config.py) e completa o download - mesmo padrão de
    `download_bloco_metas_spreadsheet`, incluindo a dica de título para
    `_find_tile_actions_button`.

    `download_timeout_ms` pode ser aumentado para downloads maiores (ex:
    carga de ano fechado inteiro, ver `download_ano_fechado_report`) - o
    Looker demora mais para gerar o arquivo antes do download começar.
    """
    secao = final_page.get_by_text(secao_tabela, exact=True).last
    secao.scroll_into_view_if_needed()
    final_page.wait_for_timeout(1000)

    tile_button = _find_tile_actions_button(final_page, secao, titulo_tile=secao_tabela)
    tile_button.hover()
    final_page.wait_for_timeout(300)
    tile_button.click(force=True)
    final_page.wait_for_timeout(1000)

    return _complete_download_dialog(final_page, base_id, download_timeout_ms=download_timeout_ms)


def download_comissao_a_vista_report(context: BrowserContext, page: Page, base: dict) -> Path:
    """
    Fluxo completo específico da base 'comissao_a_vista'. O filtro "Safra
    Mês" precisa refletir o mesmo mês/ano usado pelo relatório "Analítico"
    (base numero_contratos) - como ambos derivam do mês/ano corrente do
    sistema (`config.periodo_referencia_atual`), isso já vale
    automaticamente contanto que as duas bases rodem dentro do mesmo mês
    civil (não precisa passar período explícito entre elas).
    """
    ano_ref, mes_ref = config.periodo_referencia_atual()
    logger.info(
        "Base 'Comissão à Vista': usando referência %02d/%d (mesmo mês/ano do relatório Analítico).",
        mes_ref, ano_ref,
    )
    final_page = open_apuracao_comissao_a_vista(context, page, base)
    apply_safra_mes_filter(final_page)
    update_report_data(final_page)
    path = download_comissao_a_vista_spreadsheet(final_page, base["id"], base["secao_tabela"])
    final_page.close()
    return path


def _download_single_base(context: BrowserContext, page: Page, base: dict) -> Path:
    """Despacha para o fluxo dedicado de download de uma base, assumindo que
    `page` já está logada no portal."""
    if base["id"] == "numero_contratos":
        return download_numero_contratos_report(context, page, base)
    elif base["id"] == "dias_sem_producao":
        return download_dias_sem_producao_report(context, page, base)
    elif base["id"] == "meta_financiamento_seguro":
        return download_meta_financiamento_seguro_report(context, page, base)
    elif base["id"] == "carteira_parceiros":
        return download_carteira_parceiros_report(context, page, base)
    elif base["id"] == "comissao_a_vista":
        return download_comissao_a_vista_report(context, page, base)
    else:
        raise ValueError(f"Base '{base['id']}' não tem fluxo de download implementado")


BASES_COM_ANO_FECHADO = ("meta_financiamento_seguro", "comissao_a_vista", "dias_sem_producao", "numero_contratos")


def download_ano_fechado_report(context: BrowserContext, page: Page, base_id: str, ano: int) -> Path:
    """
    Baixa o relatório de um ano fechado específico (ex: 2025) para uma das
    4 bases em `BASES_COM_ANO_FECHADO` - NÃO faz parte do fluxo diário/
    mensal normal (`download_bases`/`_download_single_base`), só do script
    de carga única de histórico. "Carteira e Parceiros" não precisa disso
    (já baixa o ano inteiro via `apply_referencia_year_filter`).

    Cada base usa seu próprio widget de período (ver
    `apply_safra_mes_filter`, `apply_referencia_month_ano`,
    `apply_analitico_filters_ano_fechado`), todos validados ao vivo contra
    o portal em 25/08/2026 antes de existir aqui.
    """
    if base_id not in BASES_COM_ANO_FECHADO:
        raise ValueError(f"Base '{base_id}' não suporta carga de ano fechado (ver BASES_COM_ANO_FECHADO)")

    base = config.get_base_by_id(base_id)
    base_id_arquivo = f"{base_id}_{ano}"

    if base_id == "meta_financiamento_seguro":
        final_page = open_resumo_parceiro(context, page, base)
        apply_safra_mes_filter(final_page, periodo="ano", ano=ano)
        update_report_data(final_page)
        path = download_bloco_metas_spreadsheet(
            final_page, base_id_arquivo, base["secao_tabela"], download_timeout_ms=240000,
        )
    elif base_id == "comissao_a_vista":
        final_page = open_apuracao_comissao_a_vista(context, page, base)
        apply_safra_mes_filter(final_page, periodo="ano", ano=ano)
        update_report_data(final_page)
        path = download_comissao_a_vista_spreadsheet(
            final_page, base_id_arquivo, base["secao_tabela"], download_timeout_ms=240000,
        )
    elif base_id == "dias_sem_producao":
        final_page = open_sla_analitico(context, page, base)
        apply_referencia_month_ano(final_page, ano)
        update_report_data(final_page)
        path = download_sla_analitico_spreadsheet(final_page, base_id_arquivo, download_timeout_ms=240000)
    elif base_id == "numero_contratos":
        final_page = open_acompanhamento_veiculos_analitico(context, page)
        apply_analitico_filters_ano_fechado(final_page, base["filtros"]["tipo_exibicao"], ano)
        update_report_data(final_page)
        path = download_analitico_spreadsheet(final_page, base_id_arquivo)

    final_page.close()
    return path


MAX_TENTATIVAS_POR_BASE = 2  # toda base tenta pelo menos 2x antes de ser considerada falha/pulada (ver download_bases)


def download_bases(bases: list[dict], headless: bool = True) -> dict[str, Path]:
    """
    Executa o fluxo completo de download para uma ou mais bases fazendo um
    único login no portal - cada fluxo de download parte da mesma página
    inicial (`page`, a aba do WebAutorizador logada) para abrir seu próprio
    caminho de pop-ups no menu Relatórios.

    Retorna um dict {base_id: caminho_do_arquivo_baixado} só com as bases
    que baixaram com sucesso - falha em uma base é logada e não impede as
    demais.

    Toda base tenta de novo em caso de falha técnica/de navegação, até
    `MAX_TENTATIVAS_POR_BASE` vezes, antes de ser considerada "pulada". A
    base "dias_sem_producao" (SLA) tem uma falha técnica intermitente
    conhecida (ver GUIA_TIME_DADOS.md seção 10), mas o retry vale para
    todas. Um download que funciona mas vem com a planilha vazia NÃO é uma
    falha (tratado à parte em `data_processor`).
    """
    resultados: dict[str, Path] = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        # viewport largo (não só a janela): no padrão do Playwright (1280x720)
        # a barra de abas do dashboard "Acompanhamento Veículos" some do DOM
        # (o Looker colapsa itens que não cabem), e o clique em "Analítico"
        # (`apply_analitico_filters`) nunca encontra o elemento.
        context = browser.new_context(accept_downloads=True, viewport={"width": 1600, "height": 900})
        page = context.new_page()

        try:
            login(page)
            for base in bases:
                if page.is_closed():
                    # A aba principal (WebAutorizador logado) foi fechada de forma
                    # inesperada (ex: crash do navegador) - sem ela, nenhuma base
                    # restante consegue nem começar a navegar. Para aqui em vez de
                    # deixar cada uma falhar com um erro confuso e repetitivo.
                    restantes = ", ".join(b["nome"] for b in bases[bases.index(base):])
                    logger.error(
                        "Sessão do portal perdida (aba principal fechada inesperadamente) - "
                        "pulando as bases restantes desta execução: %s",
                        restantes,
                    )
                    break

                max_tentativas = MAX_TENTATIVAS_POR_BASE
                for tentativa in range(1, max_tentativas + 1):
                    paginas_antes = set(context.pages)
                    try:
                        resultados[base["id"]] = _download_single_base(context, page, base)
                        break  # sucesso - não tenta de novo
                    except Exception:
                        ultima_tentativa = tentativa == max_tentativas
                        if not ultima_tentativa:
                            # Ainda tem tentativa sobrando - loga como aviso (não como
                            # falha final) e tenta de novo do zero, dando um tempo para
                            # o portal/dashboard terminar de carregar antes de reabrir.
                            logger.warning(
                                "Falha ao baixar a base '%s' do Looker (tentativa %d/%d) - "
                                "tentando de novo...",
                                base["nome"], tentativa, max_tentativas, exc_info=True,
                            )
                            page.wait_for_timeout(5000)
                        else:
                            logger.exception(
                                "Falha ao baixar a base '%s' do Looker (tentativa %d/%d, desistindo)",
                                base["nome"], tentativa, max_tentativas,
                            )
                            if base["id"] == "dias_sem_producao":
                                # Falha técnica conhecida de navegação/carregamento no
                                # portal (ver GUIA_TIME_DADOS.md seção 10), não ausência de dados.
                                logger.warning(
                                    "A base 'Dias sem Produção' (SLA) NÃO falhou por falta de "
                                    "dados - é um problema técnico já conhecido de navegação/"
                                    "carregamento no portal, mesmo após %d tentativas. A base "
                                    "foi pulada nesta execução; as demais bases não são afetadas.",
                                    max_tentativas,
                                )
                    finally:
                        # Se a base falhou antes do `final_page.close()` do seu próprio
                        # fluxo, a aba/pop-up fica aberta e "suja" a sessão para a
                        # próxima tentativa/base - fecha qualquer aba nova ainda aberta.
                        for pagina in context.pages:
                            if pagina not in paginas_antes and pagina is not page and not pagina.is_closed():
                                try:
                                    pagina.close()
                                except Exception:
                                    pass
        finally:
            context.close()
            browser.close()

    return resultados


def download_base(base: dict, headless: bool = True) -> Path:
    """Atalho para baixar uma única base (usado pelo CLI deste módulo)."""
    resultados = download_bases([base], headless=headless)
    if base["id"] not in resultados:
        raise RuntimeError(f"Falha ao baixar a base '{base['id']}' - ver log acima para o erro original")
    return resultados[base["id"]]


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    parser = argparse.ArgumentParser(description="Download de relatório do Looker")
    parser.add_argument("--base", required=True, help="id da base (ver config.py)")
    parser.add_argument("--debug", action="store_true", help="abre o navegador visível")
    args = parser.parse_args()

    base_cfg = config.get_base_by_id(args.base)
    download_base(base_cfg, headless=not args.debug)
