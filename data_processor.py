"""
Tratamento de dados das bases baixadas do Looker.

Cada uma das 5 bases (numero_contratos, dias_sem_producao,
meta_financiamento_seguro, carteira_parceiros, comissao_a_vista) tem seu
próprio fluxo de tratamento dedicado (`_process_*`), pois cada uma tem
regras próprias de seleção de colunas, identificação do "período atual" e
destino final - ver `process_base` para o dispatch entre elas.
"""

import logging
import time
from copy import copy
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

import config

logger = logging.getLogger("data_processor")

DATE_COLUMN_NUMERO_CONTRATOS = "Dt Relatório"

VERDE_LINHA_NOVA = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
AMARELO_LINHA_EDITADA = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
VERMELHO_SEM_DADOS = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

HISTORICO_COLUNAS = ["Data/Hora", "Base", "Linhas baixadas", "Linhas novas", "Linhas totais (destino)", "Observação"]

TENTATIVAS_ARQUIVO_BLOQUEADO = 5
ESPERA_ARQUIVO_BLOQUEADO_SEGUNDOS = 3


def _com_retry_arquivo_bloqueado(
    path: Path, salvar: Callable[[], None],
    *, tentativas: int = TENTATIVAS_ARQUIVO_BLOQUEADO, espera_segundos: int = ESPERA_ARQUIVO_BLOQUEADO_SEGUNDOS,
) -> None:
    """
    Executa `salvar()` (um `DataFrame.to_excel(...)` ou `Workbook.save(...)`)
    com retry em caso de arquivo bloqueado - situação comum em produção já
    que as planilhas ficam em pastas do OneDrive: alguém pode estar com o
    arquivo aberto no Excel, ou o próprio OneDrive pode estar no meio de
    uma sincronização quando a execução tenta gravar. Sem isso, qualquer
    um dos dois cenários derrubava a base inteira com um `PermissionError`
    (WinError 32) sem chance de recuperação, mesmo sendo uma condição
    tipicamente temporária.

    Tenta `tentativas` vezes com espera fixa entre elas; na última
    tentativa, deixa o erro se propagar com uma mensagem clara sobre a
    causa provável e qual arquivo fechar, em vez do `PermissionError` cru
    do openpyxl/pandas.
    """
    for tentativa in range(1, tentativas + 1):
        try:
            salvar()
            return
        except PermissionError:
            if tentativa == tentativas:
                raise PermissionError(
                    f"Não foi possível salvar '{path}' após {tentativas} tentativas - "
                    "o arquivo parece estar aberto no Excel ou o OneDrive ainda está "
                    "sincronizando. Feche a planilha (ou aguarde a sincronização) e "
                    "rode a base de novo."
                ) from None
            logger.warning(
                "Arquivo '%s' está bloqueado (aberto no Excel ou sincronizando no "
                "OneDrive) - tentativa %d/%d, tentando de novo em %ds...",
                path, tentativa, tentativas, espera_segundos,
            )
            time.sleep(espera_segundos)


def _historico_path() -> Path:
    return config.LOG_DIR / "historico_execucoes.xlsx"


def registrar_historico(
    base_nome: str,
    linhas_baixadas: int | None,
    linhas_novas: int | None = None,
    linhas_total: int | None = None,
    observacao: str = "",
):
    """
    Registra uma linha em `logs/historico_execucoes.xlsx` (separado do
    `rpa.log`, feito pra abrir/filtrar direto no Excel) com a quantidade de
    linhas baixadas do Looker nesta execução para uma base - permite
    conferir rapidamente, ao longo do tempo, se cada execução trouxe dados
    novos ou não, sem precisar vasculhar o log de texto.

    `linhas_baixadas=0` (ou `None`, quando nem chegou a baixar - ver
    `main.run_bases`) pinta a linha de vermelho, pra pular aos olhos ao
    abrir a planilha. Chamada uma vez por base a cada execução, pelo fluxo
    de sucesso (final de cada `_process_*`) e pelo caminho de falha no
    download (`main.run_bases`, quando a base é pulada sem nem gerar
    arquivo).
    """
    path = _historico_path()
    path.parent.mkdir(parents=True, exist_ok=True)

    if path.exists():
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Histórico"
        ws.append(HISTORICO_COLUNAS)

    ws.append([
        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        base_nome,
        linhas_baixadas,
        linhas_novas,
        linhas_total,
        observacao,
    ])

    if not linhas_baixadas:  # 0 ou None
        for cell in ws[ws.max_row]:
            cell.fill = VERMELHO_SEM_DADOS

    ws.auto_filter.ref = ws.dimensions
    _com_retry_arquivo_bloqueado(path, lambda: wb.save(path))


def _chave_como_serie(df: pd.DataFrame, chave) -> pd.Series:
    """Converte a(s) coluna(s) de `chave` numa série de valores comparáveis - uma
    tupla por linha se `chave` for uma lista de colunas, ou a linha inteira
    (como tupla) se `chave` for None.

    Dois ajustes evitam falsos positivos ao comparar com a execução anterior:
    - Colunas numéricas são arredondadas (6 casas decimais) - o Excel perde
      um pouco de precisão de ponto flutuante ao salvar/reabrir um valor
      (ex: 61653.33333333334 vira 61653.333333333336), o que faria uma linha
      sem nenhuma mudança real parecer "editada".
    - Células vazias (NaN) são trocadas por um marcador fixo, já que
      `NaN != NaN` em Python faria uma linha inalterada (mas com alguma
      célula vazia) ser sempre considerada "diferente" por engano.
    """
    if chave is None:
        colunas = df
    elif isinstance(chave, (list, tuple)):
        colunas = df[list(chave)]
    else:
        return df[chave].fillna("__NaN__")
    colunas = colunas.round(6).fillna("__NaN__")
    return colunas.apply(tuple, axis=1)


def _marcar_linhas_novas_e_editadas(path: Path, df: pd.DataFrame, chave, df_anterior: pd.DataFrame | None):
    """Pinta, na planilha "Prévia" salva em `path`, de **verde** as linhas cuja
    chave não existia na versão anterior da Prévia (linhas novas desta
    compilação do RPA) e de **amarelo** as linhas cuja chave já existia mas
    algum dado da linha mudou (linhas editadas) - para o usuário identificar
    visualmente o que foi adicionado/alterado na execução do dia."""
    if df.empty:
        return

    cols_chave = list(chave) if isinstance(chave, (list, tuple)) else ([chave] if chave is not None else [])
    tem_anterior = (
        df_anterior is not None
        and not df_anterior.empty
        and all(c in df_anterior.columns for c in cols_chave)
    )
    if tem_anterior:
        chave_anterior_serie = _chave_como_serie(df_anterior, chave)
        linha_anterior_serie = _chave_como_serie(df_anterior, None)
        linha_anterior_por_chave = dict(zip(chave_anterior_serie, linha_anterior_serie))
    else:
        linha_anterior_por_chave = {}

    chave_serie = _chave_como_serie(df, chave)
    linha_atual_serie = _chave_como_serie(df, None)

    wb = load_workbook(path)
    ws = wb.active
    mudou_algo = False
    for offset, (k, linha_atual) in enumerate(zip(chave_serie, linha_atual_serie), start=2):  # linha 1 = cabeçalho
        if k not in linha_anterior_por_chave:
            fill = VERDE_LINHA_NOVA
        elif linha_anterior_por_chave[k] != linha_atual:
            fill = AMARELO_LINHA_EDITADA
        else:
            fill = None
        if fill is not None:
            mudou_algo = True
            for cell in ws[offset]:
                cell.fill = fill
    if mudou_algo:
        _com_retry_arquivo_bloqueado(path, lambda: wb.save(path))


def _acumular_planilha(
    caminho: Path, df_novo: pd.DataFrame, chave,
) -> tuple[pd.DataFrame, pd.DataFrame | None]:
    """
    Lê a planilha salva na execução anterior (Prévia ou "Base Price" -
    ver `_acumular_e_colorir_origem`) e mescla com os dados novos desta
    execução (`df_novo`), em vez de substituir o arquivo por inteiro -
    garante que uma linha já vista **nunca** desapareça, mesmo se o Looker
    deixar de trazê-la numa execução seguinte (ex: instabilidade do
    relatório, virada de mês/ano) ou se o período já tiver "fechado" - só
    acumula, nunca descarta uma linha por período. Sempre mantém a versão
    mais recente de cada chave (`drop_duplicates(..., keep="last")`),
    então uma linha existente é atualizada se algum dado dela mudar.

    Se `df_novo` vier vazio (sem dados no período), a planilha anterior é
    devolvida sem nenhuma alteração - evita apagar todo o histórico por
    causa de uma execução sem dados no período.

    Retorna (df a gravar, versão anterior - usada em seguida por
    `_marcar_linhas_novas_e_editadas` para a marcação de cor).
    """
    df_anterior = pd.read_excel(caminho) if caminho.exists() else None

    if df_novo.empty:
        df_final = df_anterior if df_anterior is not None else df_novo
        return df_final, df_anterior

    if df_anterior is None or df_anterior.empty:
        return df_novo, df_anterior

    df_final = pd.concat([df_anterior, df_novo], ignore_index=True)
    df_final = df_final.drop_duplicates(subset=chave, keep="last")
    return df_final, df_anterior


def _acumular_e_colorir_origem(
    origem_path: Path, df_previa: pd.DataFrame, chave, aplicar_autofiltro: bool,
) -> tuple[pd.DataFrame, int]:
    """
    Mescla `df_previa` (a Prévia já acumulada) na planilha de origem
    oficial ("Base Price") em `origem_path`, usando o mesmo acúmulo
    "nunca descarta uma linha" da Prévia (`_acumular_planilha`) e a mesma
    marcação de cor (`_marcar_linhas_novas_e_editadas`): 🟩 verde para
    chave nova (acrescentada depois da última linha já preenchida), 🟨
    amarelo para chave já existente com algum dado alterado (atualizada
    onde já estava, nunca duplicada). Usada pelas bases "Dias sem
    Produção", "Meta Financiamento e Seguro" e "Carteira e Parceiros"
    (Número de Contratos e Comissão à Vista têm suas próprias regras de
    origem oficial - ver `_process_numero_contratos`/
    `_acumular_origem_comissao_a_vista`).

    Retorna (DataFrame final gravado em `origem_path`, quantidade de
    chaves genuinamente novas nesta execução - usada no histórico).
    """
    df_final, df_anterior = _acumular_planilha(origem_path, df_previa, chave)
    _com_retry_arquivo_bloqueado(origem_path, lambda: df_final.to_excel(origem_path, index=False))
    if aplicar_autofiltro:
        _apply_excel_autofilter(origem_path)
    _marcar_linhas_novas_e_editadas(origem_path, df_final, chave, df_anterior)
    return df_final, _contar_chaves_novas(df_final, df_anterior, chave)


def _contar_chaves_novas(df_final: pd.DataFrame, df_anterior: pd.DataFrame | None, chave) -> int:
    """Quantas chaves de `df_final` não existiam em `df_anterior` - usado só
    para o número reportado no histórico (`registrar_historico`), a
    marcação de cor em si é feita à parte por `_marcar_linhas_novas_e_editadas`."""
    if df_anterior is not None and not df_anterior.empty:
        chaves_anteriores = set(_chave_como_serie(df_anterior, chave))
        chaves_atuais = set(_chave_como_serie(df_final, chave))
        return len(chaves_atuais - chaves_anteriores)
    return len(df_final)


def _current_month_mask(df: pd.DataFrame, date_col: str) -> pd.Series:
    ano, mes = config.periodo_referencia_atual()
    dt = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
    return (dt.dt.month == mes) & (dt.dt.year == ano)


def _current_month_mask_com_virada(df: pd.DataFrame, date_col: str, dias_extra: int = 3) -> pd.Series:
    """
    Igual a `_current_month_mask`, mas no primeiro dia do mês (virada)
    também mantém os últimos `dias_extra` dias do mês anterior - alguns
    contratos de fim de mês só aparecem como "PROPOSTA PAGA" com um pequeno
    atraso. A partir do segundo dia do mês, volta a ser só o mês atual.
    """
    hoje = date.today()
    mask = _current_month_mask(df, date_col)

    if hoje.day == 1:
        dt = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
        primeiro_dia_mes_atual = pd.Timestamp(hoje.year, hoje.month, 1)
        limite_inferior = primeiro_dia_mes_atual - timedelta(days=dias_extra)
        mask = mask | ((dt >= limite_inferior) & (dt < primeiro_dia_mes_atual))

    return mask


def _select_columns(df_novo: pd.DataFrame, base: dict) -> pd.DataFrame:
    """Mantém somente as colunas listadas em `colunas_manter` (config.py) - as demais são excluídas."""
    colunas_manter = base["regras"]["colunas_manter"]
    faltando = [c for c in colunas_manter if c not in df_novo.columns]
    if faltando:
        logger.warning("Colunas esperadas não encontradas no arquivo baixado: %s", faltando)
    colunas_presentes = [c for c in colunas_manter if c in df_novo.columns]
    return df_novo[colunas_presentes]


def _apply_row_filters(df: pd.DataFrame, base: dict) -> pd.DataFrame:
    """Aplica filtros de linha (ex: Status Proposta) antes de qualquer corte de coluna."""
    status_filtro = base["regras"].get("filtro_status_proposta")
    if status_filtro:
        df = df[df["Status Proposta"] == status_filtro]
    return df


def _apply_excel_autofilter(path: Path):
    """Adiciona o filtro (AutoFilter) do Excel em todas as colunas da planilha final."""
    wb = load_workbook(path)
    ws = wb.active
    ws.auto_filter.ref = ws.dimensions
    _com_retry_arquivo_bloqueado(path, lambda: wb.save(path))


CHAVE_UNICA_NUMERO_CONTRATOS = "ID Proposta"
DIAS_JANELA_TRIMESTRE_NUMERO_CONTRATOS = 90  # janela móvel da planilha "Trimestre" - ver _process_numero_contratos


def _process_numero_contratos(downloaded_path: Path, base: dict) -> Path:
    """
    Fluxo específico da base "Número de Contratos":
      1. Filtra Status Proposta = PROPOSTA PAGA e seleciona as colunas certas.
      2. Descarta qualquer linha que não seja do mês atual - o relatório usa
         "Last 30 Days", então sempre traz um pedaço do mês anterior junto,
         que não deve entrar na Prévia nem ser considerado daqui pra frente.
         Exceção: no primeiro dia do mês (virada), também mantém os últimos
         3 dias do mês anterior, para não perder contratos de fim de mês
         que só aparecem como "PROPOSTA PAGA" com um pequeno atraso.
      3. Acumula o resultado na planilha "Prévia" (só o mês atual), sem
         duplicar contratos já vistos em downloads anteriores do mesmo mês -
         a deduplicação é por "ID Proposta", mantendo sempre a versão mais
         recente baixada. Contrato novo = 🟩 verde, contrato existente com
         dado alterado = 🟨 amarelo.
      4. Mescla essa mesma Prévia na planilha "Trimestre" (Digitação
         Analítico - {ano} - Trimestre.xlsx, mesmo ano) - contrato novo é
         acrescentado depois da última linha já preenchida (🟩 verde) e
         contrato existente é atualizado onde já estava se algum dado
         mudou (🟨 amarelo). Diferente das outras planilhas de origem
         oficial do projeto, esta NÃO acumula para sempre: a cada
         execução, qualquer contrato com data mais antiga que
         `DIAS_JANELA_TRIMESTRE_NUMERO_CONTRATOS` (90 dias) é removido -
         é uma janela móvel de ~1 trimestre, não um histórico permanente
         (pedido do time em 21/08/2026 - antes acumulava o ano inteiro).
      5. Mescla também a Prévia numa segunda planilha, "Digitação
         Analítico - {ano}_Anual" (mesma pasta da planilha do passo 4),
         usando `_acumular_e_colorir_origem` - mesma regra de cor do
         passo 4, mas SEM a janela de 90 dias e SEM reordenar por data:
         contrato novo é sempre acrescentado logo depois da última linha
         já preenchida, na ordem em que foi sendo processado ao longo do
         ano, funcionando como o único acumulado histórico permanente
         (nunca remove uma linha) - é esta planilha, não a do passo 4, que
         guarda o ano completo.

    Na planilha do passo 4, o resultado final fica ordenado por data
    crescente (do menor para o maior dia de cada mês, mês a mês) - não só
    o bloco novo, a tabela inteira é reordenada por data a cada execução.
    """
    chave = CHAVE_UNICA_NUMERO_CONTRATOS
    date_col = DATE_COLUMN_NUMERO_CONTRATOS

    def _ordenar_por_data(df: pd.DataFrame) -> pd.DataFrame:
        if date_col in df.columns:
            return df.sort_values(by=date_col, ascending=True, kind="stable").reset_index(drop=True)
        return df

    def _apenas_mes_atual(df: pd.DataFrame) -> pd.DataFrame:
        if date_col in df.columns and not df.empty:
            return df[_current_month_mask_com_virada(df, date_col)]
        return df

    def _apenas_janela_trimestre(df: pd.DataFrame) -> pd.DataFrame:
        if date_col not in df.columns or df.empty:
            return df
        dt = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
        limite_inferior = pd.Timestamp(date.today() - timedelta(days=DIAS_JANELA_TRIMESTRE_NUMERO_CONTRATOS))
        return df[dt >= limite_inferior]

    df_tratado = pd.read_excel(downloaded_path)
    linhas_baixadas = len(df_tratado)
    df_tratado = _apply_row_filters(df_tratado, base)
    df_tratado = _select_columns(df_tratado, base)
    df_tratado = _apenas_mes_atual(df_tratado)

    # --- 1. Acumula na "Prévia" (só o mês atual), sem duplicar por ID Proposta ---
    previa_path = config.caminho_previa_numero_contratos()
    previa_path.parent.mkdir(parents=True, exist_ok=True)

    df_previa_existente = pd.read_excel(previa_path) if previa_path.exists() else pd.DataFrame(columns=df_tratado.columns)
    df_previa_existente = _apenas_mes_atual(df_previa_existente)  # descarta sobra de mês anterior já acumulada

    df_previa = pd.concat([df_previa_existente, df_tratado], ignore_index=True)
    df_previa = df_previa.drop_duplicates(subset=chave, keep="last")
    df_previa = _ordenar_por_data(df_previa)

    _com_retry_arquivo_bloqueado(previa_path, lambda: df_previa.to_excel(previa_path, index=False))
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(previa_path)
    _marcar_linhas_novas_e_editadas(previa_path, df_previa, chave, df_previa_existente)
    logger.info("Prévia atualizada (sem duplicar '%s'): %s (%d linhas)", chave, previa_path, len(df_previa))

    # --- 2. Mescla na planilha "Trimestre" (janela móvel de 90 dias) ---
    ano = date.today().year
    origem_path = config.caminho_planilha_origem_numero_contratos(ano)
    origem_path.parent.mkdir(parents=True, exist_ok=True)
    autofiltro = bool(base["regras"].get("aplicar_autofiltro_excel"))

    df_final, df_origem_anterior = _acumular_planilha(origem_path, df_previa, chave)
    df_final = _apenas_janela_trimestre(df_final)  # descarta contratos com mais de 90 dias
    df_final = _ordenar_por_data(df_final)
    _com_retry_arquivo_bloqueado(origem_path, lambda: df_final.to_excel(origem_path, index=False))
    if autofiltro:
        _apply_excel_autofilter(origem_path)
    _marcar_linhas_novas_e_editadas(origem_path, df_final, chave, df_origem_anterior)
    linhas_novas = _contar_chaves_novas(df_final, df_origem_anterior, chave)

    logger.info(
        "Planilha 'Trimestre' atualizada: %s (+%d contratos novos, %d no total, janela de %d dias)",
        origem_path, linhas_novas, len(df_final), DIAS_JANELA_TRIMESTRE_NUMERO_CONTRATOS,
    )

    # --- 3. Mescla na planilha "_Anual" (mesmo ano, sem reordenar por data) ---
    origem_anual_path = config.caminho_planilha_origem_numero_contratos_anual(ano)
    df_final_anual, linhas_novas_anual = _acumular_e_colorir_origem(
        origem_anual_path, df_previa, chave, autofiltro,
    )
    logger.info(
        "Planilha '_Anual' atualizada: %s (+%d contratos novos, %d no total)",
        origem_anual_path, linhas_novas_anual, len(df_final_anual),
    )

    observacao = "Sem dados no período" if linhas_baixadas == 0 else ""
    registrar_historico(base["nome"], linhas_baixadas, linhas_novas, len(df_final), observacao)
    return origem_path


CHAVE_UNICA_DIAS_SEM_PRODUCAO = ["Cd Loja", "Safra Mes"]


def _process_dias_sem_producao(downloaded_path: Path, base: dict) -> Path:
    """
    Fluxo específico da base "Dias sem Produção":
      1. Seleciona as colunas certas (essa base não tem filtro de status).
      2. Acumula o resultado tratado na pasta "Prévia", mesclando com a
         Prévia da execução anterior sem duplicar por `Cd Loja` + `Safra
         Mes` (ver `_acumular_planilha`) - uma linha vista numa execução
         anterior NUNCA é descartada da Prévia, mesmo que o Looker não
         traga ela de novo numa execução seguinte. Chave nova = 🟩 verde,
         chave existente com dado alterado = 🟨 amarelo.
      3. Mescla essa mesma Prévia na planilha de origem oficial ("Base
         Price") - a chave nova é acrescentada depois da última linha já
         preenchida (🟩 verde) e a chave existente é atualizada onde já
         estava se algum dado mudou (🟨 amarelo); nunca remove uma linha
         (ver `_acumular_e_colorir_origem`).
    """
    chave = CHAVE_UNICA_DIAS_SEM_PRODUCAO

    df_tratado = pd.read_excel(downloaded_path)
    linhas_baixadas = len(df_tratado)
    df_tratado = _select_columns(df_tratado, base)

    previa_path = config.caminho_previa_dias_sem_producao()
    previa_path.parent.mkdir(parents=True, exist_ok=True)
    df_previa, df_previa_anterior = _acumular_planilha(previa_path, df_tratado, chave)

    _com_retry_arquivo_bloqueado(previa_path, lambda: df_previa.to_excel(previa_path, index=False))
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(previa_path)
    _marcar_linhas_novas_e_editadas(previa_path, df_previa, chave, df_previa_anterior)
    logger.info("Base '%s' tratada (prévia): %s (%d linhas)", base["nome"], previa_path, len(df_previa))

    origem_path = config.caminho_planilha_origem_dias_sem_producao()
    origem_path.parent.mkdir(parents=True, exist_ok=True)
    autofiltro = bool(base["regras"].get("aplicar_autofiltro_excel"))
    df_final, linhas_novas = _acumular_e_colorir_origem(origem_path, df_previa, chave, autofiltro)

    logger.info(
        "Planilha de origem atualizada: %s (+%d linhas novas, %d no total)",
        origem_path, linhas_novas, len(df_final),
    )
    observacao = "Sem dados no período" if linhas_baixadas == 0 else ""
    registrar_historico(base["nome"], linhas_baixadas, linhas_novas, len(df_final), observacao)
    return origem_path


CHAVE_UNICA_META_FINANCIAMENTO_SEGURO = ["Anomes Apuracao", "Filial"]


def _process_meta_financiamento_seguro(downloaded_path: Path, base: dict) -> Path:
    """
    Fluxo específico da base "Meta Financiamento e Seguro":
      1. Seleciona as colunas certas (essa base não tem filtro de status).
      2. Acumula o resultado tratado na pasta "Prévia", mesclando com a
         Prévia da execução anterior sem duplicar por `Anomes Apuracao` +
         `Filial` (ver `_acumular_planilha`) - uma linha vista numa
         execução anterior NUNCA é descartada da Prévia, mesmo que o
         Looker não traga ela de novo numa execução seguinte. Chave nova =
         🟩 verde, chave existente com dado alterado = 🟨 amarelo.
      3. Mescla essa mesma Prévia em cada planilha de origem oficial
         ("Base Price", uma por ano) - a chave nova é acrescentada depois
         da última linha já preenchida (🟩 verde) e a chave existente é
         atualizada onde já estava se algum dado mudou (🟨 amarelo); nunca
         remove uma linha (ver `_acumular_e_colorir_origem`). Roteado por
         ano a partir de TODOS os anos presentes na Prévia acumulada (não
         só no download desta execução), pra garantir que uma linha antiga
         preservada na Prévia também acabe entrando na planilha de origem
         correta.

    Cada ano tem seu próprio arquivo de origem (não subpasta, como em
    Número de Contratos): "Meta Financiamento Seguro - {ano}.xlsx". Se a
    Prévia acumulada abranger mais de um ano, cada ano é roteado para o
    arquivo correto.
    """
    chave = CHAVE_UNICA_META_FINANCIAMENTO_SEGURO

    df_tratado = pd.read_excel(downloaded_path)
    linhas_baixadas = len(df_tratado)
    df_tratado = _select_columns(df_tratado, base)

    previa_path = config.caminho_previa_meta_financiamento_seguro()
    previa_path.parent.mkdir(parents=True, exist_ok=True)
    df_previa, df_previa_anterior = _acumular_planilha(previa_path, df_tratado, chave)

    _com_retry_arquivo_bloqueado(previa_path, lambda: df_previa.to_excel(previa_path, index=False))
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(previa_path)
    _marcar_linhas_novas_e_editadas(previa_path, df_previa, chave, df_previa_anterior)
    logger.info("Base '%s' tratada (prévia): %s (%d linhas)", base["nome"], previa_path, len(df_previa))

    if df_previa.empty:
        # Sem "Anomes Apuracao" nenhum na Prévia (nem desta execução, nem
        # acumulado antes) - não dá pra saber qual planilha de origem
        # tocar. Não é um erro - loga e segue em frente sem mexer na
        # origem oficial, em vez de tentar `origem_paths[-1]` numa lista
        # vazia (o que travaria a execução).
        logger.warning(
            "Relatório '%s' baixado veio vazio e a Prévia também está vazia - "
            "nada para atualizar na planilha de origem oficial.",
            base["nome"],
        )
        registrar_historico(base["nome"], linhas_baixadas, 0, None, "Sem dados no período")
        return previa_path

    autofiltro = bool(base["regras"].get("aplicar_autofiltro_excel"))
    anos_presentes = sorted(df_previa["Anomes Apuracao"].astype(str).str[:4].unique())
    origem_paths = []
    linhas_total_todos_anos = 0
    linhas_novas_todos_anos = 0

    for ano_str in anos_presentes:
        ano = int(ano_str)
        df_ano_previa = df_previa[df_previa["Anomes Apuracao"].astype(str).str[:4] == ano_str]

        origem_path = config.caminho_planilha_origem_meta_financiamento_seguro(ano)
        origem_path.parent.mkdir(parents=True, exist_ok=True)
        df_final, linhas_novas = _acumular_e_colorir_origem(origem_path, df_ano_previa, chave, autofiltro)

        logger.info(
            "Planilha de origem atualizada: %s (+%d linhas novas, %d no total)",
            origem_path, linhas_novas, len(df_final),
        )
        origem_paths.append(origem_path)
        linhas_total_todos_anos += len(df_final)
        linhas_novas_todos_anos += linhas_novas

    observacao = "Sem dados no período" if linhas_baixadas == 0 else ""
    registrar_historico(base["nome"], linhas_baixadas, linhas_novas_todos_anos, linhas_total_todos_anos, observacao)
    return origem_paths[-1]


CHAVE_UNICA_CARTEIRA_PARCEIROS = ["Cnpj Da Loja", "Filial", "Anomes"]


def _process_carteira_parceiros(downloaded_path: Path, base: dict) -> Path:
    """
    Fluxo específico da base "Carteira e Parceiros":
      1. Não há filtro de colunas nem de status - o arquivo baixado é usado
         como está (todas as colunas).
      2. Acumula o resultado na "Prévia", mas só o mês atual (mesmo padrão
         de "Número de Contratos" - ver `_process_numero_contratos`):
         qualquer linha de um "Anomes" diferente do mês/ano corrente é
         descartada, tanto da Prévia já salva quanto do download desta
         execução, antes de mesclar sem duplicar por `Cnpj Da Loja` +
         `Filial` + `Anomes`. Chave nova = 🟩 verde, chave existente com
         dado alterado = 🟨 amarelo.
      3. A planilha de origem oficial ("Base Price", uma por ano) continua
         recebendo o download da execução INTEIRO (todos os meses do ano
         corrente, não só a Prévia recortada do passo 2) - a chave nova é
         acrescentada depois da última linha já preenchida (🟩 verde) e a
         chave existente é atualizada onde já estava se algum dado mudou
         (🟨 amarelo); nunca remove uma linha. Roteado por ano a partir de
         TODOS os anos presentes no download desta execução. Um ano já
         fechado (que o filtro "Referência = Este Ano" do Looker parou de
         trazer) continua preservado no próprio arquivo de origem daquele
         ano, já que `_acumular_planilha` nunca descarta uma linha
         existente na origem, mesmo sem receber dado novo dela.

    Em ambas as planilhas (Prévia e origem oficial), o resultado final fica
    ordenado por "Anomes" crescente (do mês mais antigo para o mais
    recente) - a ordenação é estável, então dentro de um mesmo Anomes os
    dados incluídos por último continuam por último, não embaralha nada.
    """
    def _ordenar_por_anomes(df: pd.DataFrame) -> pd.DataFrame:
        if "Anomes" in df.columns:
            return df.sort_values(by="Anomes", ascending=True, kind="stable").reset_index(drop=True)
        return df

    def _apenas_mes_atual(df: pd.DataFrame) -> pd.DataFrame:
        if "Anomes" in df.columns and not df.empty:
            ano, mes = config.periodo_referencia_atual()
            anomes_atual = ano * 100 + mes
            return df[df["Anomes"].astype(int) == anomes_atual]
        return df

    df_tratado = pd.read_excel(downloaded_path)
    linhas_baixadas = len(df_tratado)
    df_tratado = _ordenar_por_anomes(df_tratado)

    # --- 1. Acumula na "Prévia", mas só o mês atual (não o ano inteiro) ---
    previa_path = config.caminho_previa_carteira_parceiros()
    previa_path.parent.mkdir(parents=True, exist_ok=True)

    df_previa_existente = pd.read_excel(previa_path) if previa_path.exists() else pd.DataFrame(columns=df_tratado.columns)
    df_previa_existente = _apenas_mes_atual(df_previa_existente)  # descarta sobra de mês anterior já acumulada

    df_previa = pd.concat([df_previa_existente, _apenas_mes_atual(df_tratado)], ignore_index=True)
    df_previa = df_previa.drop_duplicates(subset=CHAVE_UNICA_CARTEIRA_PARCEIROS, keep="last")
    df_previa = _ordenar_por_anomes(df_previa)

    _com_retry_arquivo_bloqueado(previa_path, lambda: df_previa.to_excel(previa_path, index=False))
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(previa_path)
    _marcar_linhas_novas_e_editadas(previa_path, df_previa, CHAVE_UNICA_CARTEIRA_PARCEIROS, df_previa_existente)
    logger.info("Prévia atualizada (somente mês atual, sem duplicar %s): %s (%d linhas)", CHAVE_UNICA_CARTEIRA_PARCEIROS, previa_path, len(df_previa))

    if df_tratado.empty:
        logger.warning(
            "Relatório '%s' baixado veio vazio - nada para atualizar na planilha de origem oficial.",
            base["nome"],
        )
        registrar_historico(base["nome"], linhas_baixadas, 0, None, "Sem dados no período")
        return previa_path

    # --- 2. Mescla o download INTEIRO (todos os meses do ano) na origem oficial ---
    autofiltro = bool(base["regras"].get("aplicar_autofiltro_excel"))
    anos_presentes = sorted(df_tratado["Anomes"].astype(str).str[:4].unique())
    origem_paths = []
    linhas_total_todos_anos = 0
    linhas_novas_todos_anos = 0

    for ano_str in anos_presentes:
        ano = int(ano_str)
        df_ano_tratado = df_tratado[df_tratado["Anomes"].astype(str).str[:4] == ano_str]

        origem_path = config.caminho_planilha_origem_carteira_parceiros(ano)
        origem_path.parent.mkdir(parents=True, exist_ok=True)
        df_final, df_anterior = _acumular_planilha(origem_path, df_ano_tratado, CHAVE_UNICA_CARTEIRA_PARCEIROS)

        # `pd.concat` (dentro de `_acumular_planilha`) mantém a ordem de
        # colunas da planilha de origem já existente - como essa base não
        # tem uma lista fixa de `colunas_manter` (diferente das outras 3),
        # se o Looker já tiver reordenado alguma coluna desde a última vez
        # que a origem foi criada/salva (ex: "Loja Nova"), a origem ia
        # divergir da Prévia (que sempre reflete a ordem atual do Looker) e
        # ficar "presa" na ordem antiga para sempre. Reordena para bater
        # com a ordem atual (`df_tratado.columns`) antes de salvar -
        # qualquer coluna legada que não exista mais no download vai para
        # o final, sem perder nenhum dado.
        colunas_atuais = [c for c in df_tratado.columns if c in df_final.columns]
        colunas_legado = [c for c in df_final.columns if c not in df_tratado.columns]
        df_final = df_final[colunas_atuais + colunas_legado]
        df_final = _ordenar_por_anomes(df_final)

        _com_retry_arquivo_bloqueado(origem_path, lambda: df_final.to_excel(origem_path, index=False))
        if autofiltro:
            _apply_excel_autofilter(origem_path)
        _marcar_linhas_novas_e_editadas(origem_path, df_final, CHAVE_UNICA_CARTEIRA_PARCEIROS, df_anterior)
        linhas_novas = _contar_chaves_novas(df_final, df_anterior, CHAVE_UNICA_CARTEIRA_PARCEIROS)

        logger.info(
            "Planilha de origem atualizada: %s (+%d linhas novas, %d no total)",
            origem_path, linhas_novas, len(df_final),
        )
        origem_paths.append(origem_path)
        linhas_total_todos_anos += len(df_final)
        linhas_novas_todos_anos += linhas_novas

    observacao = "Sem dados no período" if linhas_baixadas == 0 else ""
    registrar_historico(base["nome"], linhas_baixadas, linhas_novas_todos_anos, linhas_total_todos_anos, observacao)
    return origem_paths[-1]


def _valor_para_excel(v):
    """Converte um valor de célula pandas/numpy para um tipo que o openpyxl
    aceita diretamente em `Worksheet.append` (usado por
    `_acumular_origem_comissao_a_vista`, que grava linha a linha via
    openpyxl - as demais bases usam `DataFrame.to_excel`, que já faz essa
    conversão sozinho)."""
    if pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if isinstance(v, np.generic):
        v = v.item()
    return v


def _valor_como_texto(v):
    """Converte `v` para string, sem deixar `.0` sobrando quando o valor é
    um float "inteiro" (ex: 202608.0 -> "202608", não "202608.0") - usado
    para colunas que a planilha já guarda como texto (ver
    `_colunas_como_texto`)."""
    if pd.isna(v):
        return None
    if isinstance(v, np.generic):
        v = v.item()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _colunas_como_texto(ws, linha_referencia: int, colunas) -> set:
    """
    Identifica quais colunas a planilha já guarda como **texto** (não
    número), olhando o tipo do valor em algumas linhas de referência - a
    planilha oficial de "Comissão à Vista" guarda "Cnpj Master", "Cnpj
    Corban" e "Anomes Apuracao" como texto, mas o relatório baixado do
    Looker traz esses mesmos campos como número. Sem converter, CNPJs
    (números de 14 dígitos) aparecem em notação científica no Excel (ex:
    "4,11635E+13") nas linhas novas, diferente do resto da planilha.

    Confere um pequeno intervalo de linhas (não só uma) e considera texto
    se achar QUALQUER valor string numa delas - uma única linha de
    referência com aquela célula vazia poderia mascarar uma coluna que na
    prática sempre vem como texto.
    """
    colunas_texto = set()
    ultima_linha = min(linha_referencia + 9, ws.max_row)
    for idx, nome in enumerate(colunas, start=1):
        for row in range(linha_referencia, ultima_linha + 1):
            if isinstance(ws.cell(row=row, column=idx).value, str):
                colunas_texto.add(nome)
                break
    return colunas_texto


def _formatar_celula_como(celula, referencia):
    """Copia número/fonte/preenchimento/borda/alinhamento da célula
    `referencia` para `celula` - usado ao anexar linhas na planilha de
    origem oficial de "Comissão à Vista" para seguir o modelo que a
    planilha já tinha antes desta base existir no RPA (`ws.append()`
    sozinho não herda nenhuma formatação, e a referência é sempre a 1ª
    linha de dado real, que segue o alinhamento padrão da planilha - sem
    forçar alinhamento à direita)."""
    celula.number_format = referencia.number_format
    celula.font = copy(referencia.font)
    celula.fill = copy(referencia.fill)
    celula.border = copy(referencia.border)
    celula.alignment = copy(referencia.alignment)


def _separar_linha_totais(df: pd.DataFrame) -> tuple[pd.DataFrame, int | None]:
    """
    A planilha de origem oficial de "Comissão à Vista" já tinha, antes
    desta base existir no RPA, uma linha de totais (identificação vazia,
    soma/média nas colunas numéricas - mesmo padrão da linha de rodapé que
    o próprio relatório do Looker traz, detectada aqui pela mesma coluna
    "Cd Contrato" vazia). Separa essa linha (se existir) do restante dos
    dados. Retorna (df_sem_totais, posição 0-based da linha de totais no
    DataFrame original, ou None se não havia).
    """
    if "Cd Contrato" not in df.columns or df.empty:
        return df, None
    mask_totais = df["Cd Contrato"].isna()
    if not mask_totais.any():
        return df, None
    posicao = int(df.index[mask_totais][0])
    return df[~mask_totais].reset_index(drop=True), posicao


def _calcular_linha_totais(df: pd.DataFrame) -> dict:
    """
    Recalcula a linha de totais da planilha de origem oficial de "Comissão
    à Vista" a partir de TODOS os dados atuais (antigos + novos desta
    execução): soma nas colunas que começam com "R$", média nas que
    começam com "%", e vazio nas colunas de identificação - mesmo padrão
    que a planilha já usava antes desta base existir no RPA.
    """
    linha = {}
    for col in df.columns:
        if col.startswith("R$"):
            linha[col] = df[col].sum()
        elif col.startswith("%"):
            linha[col] = df[col].mean()
        else:
            linha[col] = None
    return linha


def _acumular_origem_comissao_a_vista(
    path: Path, df_previa: pd.DataFrame, chave, subset, nome_planilha: str, aplicar_autofiltro: bool,
) -> tuple[int, int]:
    """
    Mescla `df_previa` (a Prévia já acumulada) na planilha de origem
    oficial de "Comissão à Vista" sem duplicar por `chave` - cria a
    planilha do zero se ainda não existir; senão, ACRESCENTA depois da
    última linha as chaves genuinamente novas (🟩 verde) e ATUALIZA no
    lugar as chaves já existentes cujo algum dado mudou (🟨 amarelo) -
    nunca remove uma linha. Mesmo padrão de cor das outras 4 bases
    (`_marcar_linhas_novas_e_editadas`), mas aplicado à mão célula a
    célula (em vez de `DataFrame.to_excel`), porque esta planilha precisa
    preservar a formatação (texto/alinhamento) herdada da 1ª linha de dado
    já existente e a linha de totais no final.

    A planilha já vinha, antes desta base existir no RPA, com uma linha de
    totais no final (identificação vazia, soma nas colunas "R$..." e
    média nas colunas "%..." - ver `_calcular_linha_totais`). Ela é
    SEMPRE removida de onde estiver, recalculada com todos os dados atuais
    (existentes + atualizados + novos) e recolocada como a última linha -
    assim nunca fica "presa" no meio conforme mais dados forem adicionados
    nas próximas execuções.

    Retorna (linhas novas + editadas nesta execução, linhas totais depois -
    sem contar a própria linha de totais).
    """
    if not path.exists():
        _com_retry_arquivo_bloqueado(path, lambda: df_previa.to_excel(path, index=False))
        wb = load_workbook(path)
        ws = wb.active
        if len(df_previa) >= 1:
            totais = _calcular_linha_totais(df_previa)
            ws.append([_valor_para_excel(totais.get(c)) for c in df_previa.columns])
            linha_referencia = ws.cell(row=2, column=1).row  # 1ª linha de dado real
            for col in range(1, ws.max_column + 1):
                _formatar_celula_como(
                    ws.cell(row=ws.max_row, column=col), ws.cell(row=linha_referencia, column=col),
                )
        _com_retry_arquivo_bloqueado(path, lambda: wb.save(path))
        if aplicar_autofiltro:
            _apply_excel_autofilter(path)
        _marcar_linhas_novas_e_editadas(path, df_previa, chave, None)
        logger.info("Planilha '%s' criada (primeira execução): %s (%d linhas)", nome_planilha, path, len(df_previa))
        return len(df_previa), len(df_previa)

    df_existente_completo = pd.read_excel(path)
    df_existente, posicao_totais = _separar_linha_totais(df_existente_completo)
    linha_totais_excel = posicao_totais + 2 if posicao_totais is not None else None  # +2: 0-based -> 1-based, +1 pelo cabeçalho

    faltando = [c for c in df_existente.columns if c not in df_previa.columns]
    extras = [c for c in df_previa.columns if c not in df_existente.columns]
    if faltando or extras:
        logger.warning(
            "Colunas da Prévia divergem da planilha '%s' existente "
            "(faltando=%s, extras=%s) - alinhando pelas colunas já existentes na planilha.",
            nome_planilha, faltando, extras,
        )
    df_previa_alinhada = df_previa.reindex(columns=df_existente.columns)
    df_previa_alinhada = df_previa_alinhada.drop_duplicates(subset=subset, keep="last").reset_index(drop=True)

    if not df_existente.empty:
        chave_existente_serie = _chave_como_serie(df_existente, chave)
        linha_existente_serie = _chave_como_serie(df_existente, None)
        posicao_por_chave = {k: i for i, k in enumerate(chave_existente_serie)}
        linha_existente_por_chave = dict(zip(chave_existente_serie, linha_existente_serie))
    else:
        posicao_por_chave = {}
        linha_existente_por_chave = {}

    chave_previa_serie = _chave_como_serie(df_previa_alinhada, chave)
    linha_previa_serie = _chave_como_serie(df_previa_alinhada, None)

    # Separa em "genuinamente nova" (chave nunca vista) e "editada" (chave
    # já existe, mas algum dado mudou) - uma chave existente sem nenhuma
    # mudança não entra em nenhum dos dois grupos (fica como está, sem cor).
    indices_novos = []
    edicoes = []  # (posição em df_existente, índice em df_previa_alinhada)
    for i, (k, linha_atual) in enumerate(zip(chave_previa_serie, linha_previa_serie)):
        if k not in posicao_por_chave:
            indices_novos.append(i)
        elif linha_existente_por_chave[k] != linha_atual:
            edicoes.append((posicao_por_chave[k], i))

    df_realmente_novo = df_previa_alinhada.iloc[indices_novos]

    if df_realmente_novo.empty and not edicoes and linha_totais_excel is None:
        logger.info(
            "Nenhum registro novo/alterado para '%s' nesta execução - planilha mantida como está (%d linhas).",
            nome_planilha, len(df_existente),
        )
        return 0, len(df_existente)

    wb = load_workbook(path)
    ws = wb.active

    # Referência de formatação: sempre a 1ª linha de dado real (linha 2) -
    # continua válida mesmo depois de remover a linha de totais de
    # qualquer posição (deletar uma linha mais abaixo não afeta a linha 2).
    linha_referencia = 2 if ws.max_row >= 2 else None

    if linha_totais_excel is not None:
        ws.delete_rows(linha_totais_excel, 1)

    # Limpa a cor de execuções anteriores antes de recolorir - só o que
    # mudar NESTA execução deve ficar destacado. Feito à mão (diferente das
    # outras 4 bases) porque esta planilha não é reescrita do zero a cada
    # execução via `DataFrame.to_excel` - é editada célula a célula pra
    # preservar a formatação herdada.
    sem_fill = PatternFill(fill_type=None)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.fill = sem_fill

    colunas_texto = _colunas_como_texto(ws, linha_referencia, df_existente.columns) if linha_referencia else set()

    # Atualiza no lugar as linhas cuja chave já existia e teve algum dado alterado
    df_existente_atualizada = df_existente.copy()
    for posicao_existente, idx_novo in edicoes:
        linha_excel = posicao_existente + 2  # +1 cabeçalho, +1 0-based -> 1-based
        linha_dados = df_previa_alinhada.iloc[idx_novo]
        for col, (nome, v) in enumerate(linha_dados.items(), start=1):
            valor = _valor_como_texto(v) if nome in colunas_texto else _valor_para_excel(v)
            ws.cell(row=linha_excel, column=col).value = valor
        if linha_referencia:
            for col in range(1, ws.max_column + 1):
                _formatar_celula_como(ws.cell(row=linha_excel, column=col), ws.cell(row=linha_referencia, column=col))
        for cell in ws[linha_excel]:
            cell.fill = AMARELO_LINHA_EDITADA
        df_existente_atualizada.iloc[posicao_existente] = linha_dados.values

    # Acrescenta as linhas genuinamente novas no final
    if not df_realmente_novo.empty:
        primeira_linha_nova = ws.max_row + 1
        for _, linha in df_realmente_novo.iterrows():
            valores = [
                _valor_como_texto(v) if nome in colunas_texto else _valor_para_excel(v)
                for nome, v in linha.items()
            ]
            ws.append(valores)
        if linha_referencia:
            for row_idx in range(primeira_linha_nova, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    _formatar_celula_como(ws.cell(row=row_idx, column=col), ws.cell(row=linha_referencia, column=col))
                for cell in ws[row_idx]:
                    cell.fill = VERDE_LINHA_NOVA

    df_tudo = pd.concat([df_existente_atualizada, df_realmente_novo], ignore_index=True)
    totais = _calcular_linha_totais(df_tudo)
    ws.append([_valor_para_excel(totais.get(c)) for c in df_existente.columns])
    if linha_referencia:
        for col in range(1, ws.max_column + 1):
            _formatar_celula_como(ws.cell(row=ws.max_row, column=col), ws.cell(row=linha_referencia, column=col))
        for cell in ws[ws.max_row]:
            cell.fill = sem_fill  # linha de totais nunca é colorida, mesmo que a linha 2 esteja

    _com_retry_arquivo_bloqueado(path, lambda: wb.save(path))
    if aplicar_autofiltro:
        _apply_excel_autofilter(path)

    total = len(df_tudo)
    logger.info(
        "Planilha '%s' atualizada: +%d linhas novas, %d linhas editadas (%d no total, linha de totais recalculada e mantida no final)",
        nome_planilha, len(df_realmente_novo), len(edicoes), total,
    )
    return len(df_realmente_novo) + len(edicoes), total


def _process_comissao_a_vista(downloaded_path: Path, base: dict) -> Path:
    """
    Fluxo específico da base "Comissão à Vista - Analítico": duas
    planilhas são atualizadas a cada execução, com regras diferentes:
      - Prévia (`config.caminho_previa_comissao_a_vista`): reescrita por
        inteiro a cada execução, mesclando com o que já existia (uma
        chave já existente é ATUALIZADA com os dados mais recentes
        baixados, não só ignorada) e recalculando as cores do zero -
        🟩 verde para chave nova, 🟨 amarelo para chave existente com
        algum dado alterado - exatamente o mesmo padrão das outras 4
        bases (`_marcar_linhas_novas_e_editadas`).
      - Planilha de origem oficial
        (`config.caminho_planilha_origem_comissao_a_vista`, que já vinha
        sendo mantida pelo time com dados de meses anteriores antes desta
        base existir aqui): recebe a mesma Prévia já acumulada - chave
        nova é acrescentada depois da última linha já preenchida (🟩
        verde) e chave existente é atualizada onde já estava se algum
        dado mudou (🟨 amarelo); nunca remove uma linha
        (`_acumular_origem_comissao_a_vista`).
    Chave em ambas: `regras["chave_comparacao"]`, config.py.

    Validada contra um download real e contra a planilha de origem oficial
    real (6 meses de histórico) em 17/08/2026: o arquivo baixado traz uma
    linha de totais no final (todas as colunas de identificação vêm
    vazias, só os valores em R$ aparecem somados) - essa linha é descartada
    antes de qualquer comparação, filtrando por "Cd Contrato" vazio (não é
    um registro de comissão de verdade, é só o rodapé do relatório).

    IMPORTANTE (bug real encontrado e corrigido em 17/08/2026): pelo menos
    uma coluna do relatório vem do Looker com espaços/quebra de linha ao
    redor do nome (ex: "\n    R$ Comissão À Vista Bruto - Master\n    "),
    enquanto a planilha de origem oficial (mantida à parte pelo time) tem
    o mesmo nome sem esses espaços. Sem normalizar, o alinhamento de
    colunas trata como duas colunas diferentes e a coluna correspondente
    fica vazia nas linhas novas - por isso os nomes de coluna do download
    são sempre limpos (`str.strip()`) antes de qualquer comparação.
    """
    previa_path = config.caminho_previa_comissao_a_vista()
    previa_path.parent.mkdir(parents=True, exist_ok=True)
    origem_path = config.caminho_planilha_origem_comissao_a_vista()
    origem_path.parent.mkdir(parents=True, exist_ok=True)
    chave = base["regras"].get("chave_comparacao")
    subset = list(chave) if isinstance(chave, (list, tuple)) else ([chave] if chave else None)
    autofiltro = bool(base["regras"].get("aplicar_autofiltro_excel"))

    df_novo = pd.read_excel(downloaded_path)
    df_novo.columns = df_novo.columns.str.strip()
    if "Cd Contrato" in df_novo.columns:
        linha_totais = df_novo["Cd Contrato"].isna()
        if linha_totais.any():
            logger.info(
                "Removendo %d linha(s) de totais/resumo (sem 'Cd Contrato') do relatório 'Comissão à Vista' baixado.",
                int(linha_totais.sum()),
            )
            df_novo = df_novo[~linha_totais].reset_index(drop=True)
    linhas_baixadas = len(df_novo)
    if df_novo.empty:
        logger.warning("Relatório 'Comissão à Vista' baixado veio vazio - nada para adicionar.")
        registrar_historico(base["nome"], linhas_baixadas, 0, None, "Sem dados no período")
        return origem_path

    # --- Prévia: reescrita + cor recalculada a cada execução (igual as outras 4 bases) ---
    df_previa_anterior = pd.read_excel(previa_path) if previa_path.exists() else None
    if df_previa_anterior is not None and not df_previa_anterior.empty:
        faltando = [c for c in df_previa_anterior.columns if c not in df_novo.columns]
        extras = [c for c in df_novo.columns if c not in df_previa_anterior.columns]
        if faltando or extras:
            logger.warning(
                "Colunas do relatório baixado divergem da Prévia 'Comissão à Vista' existente "
                "(faltando=%s, extras=%s) - alinhando pelas colunas já existentes na planilha.",
                faltando, extras,
            )
        df_novo_alinhado = df_novo.reindex(columns=df_previa_anterior.columns)
        df_previa_final = pd.concat([df_previa_anterior, df_novo_alinhado], ignore_index=True)
        # mantém sempre a versão mais recente baixada de cada chave (permite detectar "editada")
        df_previa_final = df_previa_final.drop_duplicates(subset=subset, keep="last")
    else:
        df_previa_final = df_novo

    _com_retry_arquivo_bloqueado(previa_path, lambda: df_previa_final.to_excel(previa_path, index=False))
    if autofiltro:
        _apply_excel_autofilter(previa_path)
    _marcar_linhas_novas_e_editadas(previa_path, df_previa_final, chave, df_previa_anterior)
    logger.info(
        "Planilha 'Comissão à Vista - Analítico (Prévia)' atualizada: %s (%d linhas)",
        previa_path, len(df_previa_final),
    )

    # --- Planilha de origem oficial: mescla a Prévia já acumulada (nunca perde linha) ---
    linhas_novas_origem, linhas_total_origem = _acumular_origem_comissao_a_vista(
        origem_path, df_previa_final, chave, subset, "Comissão à Vista - Analítico (origem oficial)",
        aplicar_autofiltro=autofiltro,
    )

    observacao = "Sem dados no período" if linhas_baixadas == 0 else ""
    registrar_historico(base["nome"], linhas_baixadas, linhas_novas_origem, linhas_total_origem, observacao)
    return origem_path


def process_base(downloaded_path: Path, base: dict) -> Path:
    """Pipeline completo para uma base: despacha para o fluxo dedicado conforme o 'modo' configurado."""
    modo = base["regras"].get("modo")

    if modo == "planilha_origem_local":
        return _process_numero_contratos(downloaded_path, base)

    if modo == "planilha_origem_local_dias_sem_producao":
        return _process_dias_sem_producao(downloaded_path, base)

    if modo == "planilha_origem_local_meta_financiamento_seguro":
        return _process_meta_financiamento_seguro(downloaded_path, base)

    if modo == "planilha_origem_local_carteira_parceiros":
        return _process_carteira_parceiros(downloaded_path, base)

    if modo == "planilha_origem_local_comissao_a_vista":
        return _process_comissao_a_vista(downloaded_path, base)

    raise ValueError(f"Base '{base['id']}' não tem 'modo' de tratamento reconhecido em config.py")
