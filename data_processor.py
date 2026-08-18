"""
Tratamento de dados das bases baixadas do Looker.

Cada uma das 5 bases (numero_contratos, dias_sem_producao,
meta_financiamento_seguro, carteira_parceiros, comissao_a_vista) tem seu
próprio fluxo de tratamento dedicado (`_process_*`), pois cada uma tem
regras próprias de seleção de colunas, identificação do "período atual" e
destino final - ver `process_base` para o dispatch entre elas.
"""

import logging
from copy import copy
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, PatternFill

import config

logger = logging.getLogger("data_processor")

DATE_COLUMN_NUMERO_CONTRATOS = "Dt Relatório"

VERDE_LINHA_NOVA = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
AMARELO_LINHA_EDITADA = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
VERMELHO_SEM_DADOS = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

HISTORICO_COLUNAS = ["Data/Hora", "Base", "Linhas baixadas", "Linhas novas", "Linhas totais (destino)", "Observação"]


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
    wb.save(path)


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
        wb.save(path)


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
    wb.save(path)


CHAVE_UNICA_NUMERO_CONTRATOS = "ID Proposta"


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
         recente baixada.
      4. Copia para a planilha de origem oficial do ano correspondente
         apenas os contratos que ainda não estão lá, preservando o histórico.

    Em ambas as planilhas, o resultado final fica ordenado por data
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

    df_previa.to_excel(previa_path, index=False)
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(previa_path)
    _marcar_linhas_novas_e_editadas(previa_path, df_previa, chave, df_previa_existente)
    logger.info("Prévia atualizada (sem duplicar '%s'): %s (%d linhas)", chave, previa_path, len(df_previa))

    # --- 2. Copia para a planilha de origem oficial só os contratos novos ---
    ano = date.today().year
    origem_path = config.caminho_planilha_origem_numero_contratos(ano)
    origem_path.parent.mkdir(parents=True, exist_ok=True)

    if origem_path.exists():
        df_origem = pd.read_excel(origem_path)
        ids_existentes = set(df_origem[chave])
        df_novos = df_previa[~df_previa[chave].isin(ids_existentes)]
    else:
        df_origem = pd.DataFrame(columns=df_previa.columns)
        df_novos = df_previa

    df_final = pd.concat([df_origem, df_novos], ignore_index=True)
    df_final = _ordenar_por_data(df_final)
    df_final.to_excel(origem_path, index=False)
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(origem_path)

    logger.info(
        "Planilha de origem atualizada: %s (+%d contratos novos, %d no total)",
        origem_path, len(df_novos), len(df_final),
    )
    observacao = "Sem dados no período" if linhas_baixadas == 0 else ""
    registrar_historico(base["nome"], linhas_baixadas, len(df_novos), len(df_final), observacao)
    return origem_path


CHAVE_UNICA_DIAS_SEM_PRODUCAO = ["Cd Loja", "Safra Mes"]


def _safra_mes_atual() -> int:
    """Mês/ano atual no formato AAAAMM, igual à coluna 'Safra Mes' do relatório -
    usa o período de referência centralizado em `config.periodo_referencia_atual`."""
    ano, mes = config.periodo_referencia_atual()
    return ano * 100 + mes


def _process_dias_sem_producao(downloaded_path: Path, base: dict) -> Path:
    """
    Fluxo específico da base "Dias sem Produção":
      1. Seleciona as colunas certas (essa base não tem filtro de status).
      2. Salva o resultado tratado na pasta "Prévia".
      3. Remove da planilha de origem oficial os registros do mês atual
         (identificado pela coluna "Safra Mes", formato AAAAMM) e cola os
         dados tratados no final, preservando o histórico. Diferente de
         Número de Contratos, o relatório já vem filtrado para "Este mês"
         (sem janela rolante), então o arquivo baixado não precisa ser
         restrito por mês antes de colar.
    """
    chave = CHAVE_UNICA_DIAS_SEM_PRODUCAO

    df_tratado = pd.read_excel(downloaded_path)
    linhas_baixadas = len(df_tratado)
    df_tratado = _select_columns(df_tratado, base)

    previa_path = config.caminho_previa_dias_sem_producao()
    previa_path.parent.mkdir(parents=True, exist_ok=True)
    df_previa_anterior = pd.read_excel(previa_path) if previa_path.exists() else None

    df_tratado.to_excel(previa_path, index=False)
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(previa_path)
    _marcar_linhas_novas_e_editadas(previa_path, df_tratado, chave, df_previa_anterior)
    logger.info("Base '%s' tratada (prévia): %s (%d linhas)", base["nome"], previa_path, len(df_tratado))

    origem_path = config.caminho_planilha_origem_dias_sem_producao()
    origem_path.parent.mkdir(parents=True, exist_ok=True)
    safra_atual = _safra_mes_atual()

    if origem_path.exists():
        df_origem = pd.read_excel(origem_path)
        removidos = int((df_origem["Safra Mes"] == safra_atual).sum())
        if removidos:
            logger.info(
                "Removendo %d linhas do mês atual (Safra Mes=%d) na planilha de origem",
                removidos, safra_atual,
            )
        df_origem = df_origem[df_origem["Safra Mes"] != safra_atual]
    else:
        df_origem = pd.DataFrame(columns=df_tratado.columns)

    df_final = pd.concat([df_origem, df_tratado], ignore_index=True)
    # segurança extra contra duplicados, mantendo a versão mais recente baixada
    df_final = df_final.drop_duplicates(subset=chave, keep="last")

    df_final.to_excel(origem_path, index=False)
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(origem_path)

    logger.info(
        "Planilha de origem atualizada: %s (+%d linhas do mês atual, %d no total)",
        origem_path, len(df_tratado), len(df_final),
    )
    observacao = "Sem dados no período" if linhas_baixadas == 0 else ""
    registrar_historico(base["nome"], linhas_baixadas, len(df_tratado), len(df_final), observacao)
    return origem_path


CHAVE_UNICA_META_FINANCIAMENTO_SEGURO = ["Anomes Apuracao", "Filial"]


def _process_meta_financiamento_seguro(downloaded_path: Path, base: dict) -> Path:
    """
    Fluxo específico da base "Meta Financiamento e Seguro":
      1. Seleciona as colunas certas (essa base não tem filtro de status).
      2. Salva o resultado tratado na pasta "Prévia" (sobrescrita a cada
         execução - essa base roda 1x por mês, sem acúmulo entre execuções
         como em Número de Contratos).
      3. Remove da planilha de origem oficial os registros dos meses
         presentes no arquivo baixado (identificados pela coluna "Anomes
         Apuracao", formato AAAAMM) e cola os dados tratados no final,
         preservando o histórico. Normalmente só um mês aparece no
         download, mas na "janela curta" (virada de mês sem dia útil antes
         do dia 01 - ver looker_automation.deve_usar_janela_curta_safra_mes)
         o mês anterior também pode aparecer.

    Cada ano tem seu próprio arquivo de origem (não subpasta, como em
    Número de Contratos): "Meta Financiamento Seguro - {ano}.xlsx". Se o
    download abranger mais de um ano (ex: janela curta na virada de
    dezembro/janeiro), cada ano é roteado para o arquivo correto.
    """
    chave = CHAVE_UNICA_META_FINANCIAMENTO_SEGURO

    df_tratado = pd.read_excel(downloaded_path)
    linhas_baixadas = len(df_tratado)
    df_tratado = _select_columns(df_tratado, base)

    previa_path = config.caminho_previa_meta_financiamento_seguro()
    previa_path.parent.mkdir(parents=True, exist_ok=True)
    df_previa_anterior = pd.read_excel(previa_path) if previa_path.exists() else None

    df_tratado.to_excel(previa_path, index=False)
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(previa_path)
    _marcar_linhas_novas_e_editadas(previa_path, df_tratado, chave, df_previa_anterior)
    logger.info("Base '%s' tratada (prévia): %s (%d linhas)", base["nome"], previa_path, len(df_tratado))

    if df_tratado.empty:
        # Sem "Anomes Apuracao" nenhum para identificar o(s) ano(s) - não dá
        # pra saber qual planilha de origem tocar. Não é um erro (só não
        # havia dados no período) - loga e segue em frente sem mexer na
        # origem oficial, em vez de tentar `origem_paths[-1]` numa lista
        # vazia (o que travaria a execução).
        logger.warning(
            "Relatório '%s' baixado veio vazio - nada para atualizar na planilha de origem oficial.",
            base["nome"],
        )
        registrar_historico(base["nome"], linhas_baixadas, 0, None, "Sem dados no período")
        return previa_path

    anos_presentes = sorted(df_tratado["Anomes Apuracao"].astype(str).str[:4].unique())
    origem_paths = []
    linhas_total_todos_anos = 0

    for ano_str in anos_presentes:
        ano = int(ano_str)
        df_ano = df_tratado[df_tratado["Anomes Apuracao"].astype(str).str[:4] == ano_str]
        meses_presentes = set(df_ano["Anomes Apuracao"])

        origem_path = config.caminho_planilha_origem_meta_financiamento_seguro(ano)
        origem_path.parent.mkdir(parents=True, exist_ok=True)

        if origem_path.exists():
            df_origem = pd.read_excel(origem_path)
            removidos = int(df_origem["Anomes Apuracao"].isin(meses_presentes).sum())
            if removidos:
                logger.info(
                    "Removendo %d linhas dos meses %s na planilha de origem %s",
                    removidos, sorted(meses_presentes), origem_path,
                )
            df_origem = df_origem[~df_origem["Anomes Apuracao"].isin(meses_presentes)]
        else:
            df_origem = pd.DataFrame(columns=df_ano.columns)

        df_final = pd.concat([df_origem, df_ano], ignore_index=True)
        # segurança extra contra duplicados, mantendo a versão mais recente baixada
        df_final = df_final.drop_duplicates(subset=chave, keep="last")

        df_final.to_excel(origem_path, index=False)
        if base["regras"].get("aplicar_autofiltro_excel"):
            _apply_excel_autofilter(origem_path)

        logger.info(
            "Planilha de origem atualizada: %s (+%d linhas, %d no total)",
            origem_path, len(df_ano), len(df_final),
        )
        origem_paths.append(origem_path)
        linhas_total_todos_anos += len(df_final)

    observacao = "Sem dados no período" if linhas_baixadas == 0 else ""
    registrar_historico(base["nome"], linhas_baixadas, len(df_tratado), linhas_total_todos_anos, observacao)
    return origem_paths[-1]


CHAVE_UNICA_CARTEIRA_PARCEIROS = ["Cnpj Da Loja", "Filial", "Anomes"]


def _process_carteira_parceiros(downloaded_path: Path, base: dict) -> Path:
    """
    Fluxo específico da base "Carteira e Parceiros":
      1. Não há filtro de colunas nem de status - o arquivo baixado é usado
         como está (todas as colunas).
      2. Substitui a "Prévia" por inteiro pelo arquivo recém-baixado (o
         filtro "Referência" é "Este Ano", então a Prévia sempre reflete o
         ano corrente completo, não só o dia/mês atual).
      3. Remove da planilha de origem oficial do ano corrente os registros
         do mês atual (identificado pela coluna "Anomes", formato AAAAMM)
         e cola os dados do mês atual no final, preservando o histórico dos
         meses anteriores.

    Os meses fechados (anteriores ao atual) não mudam mais uma vez
    registrados (confirmado comparando execuções: 0 diferenças) - só o mês
    em andamento tem métricas (Mercado/Retorno/Acordo) recalculadas dia a
    dia, por isso ele é sempre substituído por inteiro, como nas outras
    bases, em vez de só acrescentar linhas novas (o que geraria uma cópia
    quase-duplicada do mês atual a cada execução).

    Na primeira execução (planilha de origem ainda não existe), usa o ano
    inteiro baixado para não perder os meses anteriores já disponíveis.

    Em ambas as planilhas (Prévia e origem oficial), o resultado final fica
    ordenado por "Anomes" crescente (do mês mais antigo para o mais
    recente) - a ordenação é estável, então dentro de um mesmo Anomes os
    dados incluídos por último continuam por último, não embaralha nada.
    """
    def _ordenar_por_anomes(df: pd.DataFrame) -> pd.DataFrame:
        if "Anomes" in df.columns:
            return df.sort_values(by="Anomes", ascending=True, kind="stable").reset_index(drop=True)
        return df

    df_tratado = pd.read_excel(downloaded_path)
    linhas_baixadas = len(df_tratado)
    df_tratado = _ordenar_por_anomes(df_tratado)

    previa_path = config.caminho_previa_carteira_parceiros()
    previa_path.parent.mkdir(parents=True, exist_ok=True)
    df_previa_anterior = pd.read_excel(previa_path) if previa_path.exists() else None

    df_tratado.to_excel(previa_path, index=False)
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(previa_path)
    _marcar_linhas_novas_e_editadas(previa_path, df_tratado, CHAVE_UNICA_CARTEIRA_PARCEIROS, df_previa_anterior)
    logger.info("Prévia substituída: %s (%d linhas)", previa_path, len(df_tratado))

    ano = date.today().year
    anomes_atual = ano * 100 + date.today().month

    origem_path = config.caminho_planilha_origem_carteira_parceiros(ano)
    origem_path.parent.mkdir(parents=True, exist_ok=True)

    if origem_path.exists():
        df_origem = pd.read_excel(origem_path)
        removidos = int((df_origem["Anomes"] == anomes_atual).sum())
        if removidos:
            logger.info(
                "Removendo %d linhas do mês atual (Anomes=%d) na planilha de origem",
                removidos, anomes_atual,
            )
        df_origem = df_origem[df_origem["Anomes"] != anomes_atual]
        df_novo_mes = df_tratado[df_tratado["Anomes"] == anomes_atual]
    else:
        logger.info("Planilha de origem não existe ainda, criando com o ano inteiro baixado: %s", origem_path)
        df_origem = pd.DataFrame(columns=df_tratado.columns)
        df_novo_mes = df_tratado

    df_final = pd.concat([df_origem, df_novo_mes], ignore_index=True)
    # `pd.concat` mantém a ordem de colunas de `df_origem` (o 1º argumento) -
    # como essa base não tem uma lista fixa de `colunas_manter` (diferente
    # das outras 3), se o Looker já tiver reordenado alguma coluna desde a
    # última vez que a origem foi criada/salva (ex: "Loja Nova"), a origem
    # ia divergir da Prévia (que sempre reflete a ordem atual do Looker) e
    # ficar "presa" na ordem antiga para sempre. Reordena para bater com a
    # ordem atual (`df_tratado.columns`) antes de salvar - qualquer coluna
    # legada que não exista mais no download vai para o final, sem perder
    # nenhum dado.
    colunas_atuais = [c for c in df_tratado.columns if c in df_final.columns]
    colunas_legado = [c for c in df_final.columns if c not in df_tratado.columns]
    df_final = df_final[colunas_atuais + colunas_legado]
    df_final = _ordenar_por_anomes(df_final)
    df_final.to_excel(origem_path, index=False)
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(origem_path)

    logger.info(
        "Planilha de origem atualizada: %s (+%d linhas do mês atual, %d no total)",
        origem_path, len(df_novo_mes), len(df_final),
    )
    observacao = "Sem dados no período" if linhas_baixadas == 0 else ""
    registrar_historico(base["nome"], linhas_baixadas, len(df_novo_mes), len(df_final), observacao)
    return origem_path


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


def _formatar_celula_como(celula, referencia, alinhar_direita: bool = False):
    """Copia número/fonte/preenchimento/borda da célula `referencia` para
    `celula` - usado ao anexar linhas na planilha de origem oficial de
    "Comissão à Vista" para seguir o modelo que a planilha já tinha antes
    desta base existir no RPA (`ws.append()` sozinho não herda nenhuma
    formatação). `alinhar_direita=True` força o texto à direita em vez de
    copiar o alinhamento da referência."""
    celula.number_format = referencia.number_format
    celula.font = copy(referencia.font)
    celula.fill = copy(referencia.fill)
    celula.border = copy(referencia.border)
    if alinhar_direita:
        vertical_ref = referencia.alignment.vertical if referencia.alignment else None
        celula.alignment = Alignment(horizontal="right", vertical=vertical_ref)
    else:
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
    path: Path, df_novo: pd.DataFrame, chave, subset, nome_planilha: str, aplicar_autofiltro: bool,
) -> tuple[int, int]:
    """
    Acumula `df_novo` na planilha de origem oficial de "Comissão à Vista"
    sem duplicar por `chave` - cria a planilha do zero se ainda não
    existir, ou só ANEXA ao final os registros que ainda não existem lá
    (nunca remove/sobrescreve linhas de dado já existentes, nunca
    colorida). Diferente da Prévia dessa mesma base (ver
    `_process_comissao_a_vista`), aqui uma chave já existente NÃO é
    atualizada mesmo que algum dado tenha mudado no download - só chaves
    genuinamente novas entram. Cada célula nova segue a formatação
    (número/fonte/preenchimento/borda) da 1ª linha de dado já existente,
    com o texto sempre alinhado à direita.

    A planilha já vinha, antes desta base existir no RPA, com uma linha de
    totais no final (identificação vazia, soma nas colunas "R$..." e
    média nas colunas "%..." - ver `_calcular_linha_totais`). Ela é
    SEMPRE removida de onde estiver, recalculada com todos os dados atuais
    (antigos + novos) e recolocada como a última linha - assim nunca fica
    "presa" no meio conforme mais dados forem adicionados nas próximas
    execuções.

    Retorna (linhas_novas_adicionadas, linhas_totais_depois - sem contar a
    própria linha de totais).
    """
    if not path.exists():
        df_novo.to_excel(path, index=False)
        wb = load_workbook(path)
        ws = wb.active
        if len(df_novo) >= 1:
            totais = _calcular_linha_totais(df_novo)
            ws.append([_valor_para_excel(totais.get(c)) for c in df_novo.columns])
            linha_referencia = ws.cell(row=2, column=1).row  # 1ª linha de dado real
            for col in range(1, ws.max_column + 1):
                _formatar_celula_como(
                    ws.cell(row=ws.max_row, column=col), ws.cell(row=linha_referencia, column=col),
                )
        wb.save(path)
        if aplicar_autofiltro:
            _apply_excel_autofilter(path)
        logger.info("Planilha '%s' criada (primeira execução): %s (%d linhas)", nome_planilha, path, len(df_novo))
        return len(df_novo), len(df_novo)

    df_existente_completo = pd.read_excel(path)
    df_existente, posicao_totais = _separar_linha_totais(df_existente_completo)
    linha_totais_excel = posicao_totais + 2 if posicao_totais is not None else None  # +2: 0-based -> 1-based, +1 pelo cabeçalho

    faltando = [c for c in df_existente.columns if c not in df_novo.columns]
    extras = [c for c in df_novo.columns if c not in df_existente.columns]
    if faltando or extras:
        logger.warning(
            "Colunas do relatório baixado divergem da planilha '%s' existente "
            "(faltando=%s, extras=%s) - alinhando pelas colunas já existentes na planilha.",
            nome_planilha, faltando, extras,
        )
    df_novo_alinhado = df_novo.reindex(columns=df_existente.columns)

    if not df_existente.empty:
        chaves_existentes = set(_chave_como_serie(df_existente, chave))
        chave_novo_serie = _chave_como_serie(df_novo_alinhado, chave)
        df_realmente_novo = df_novo_alinhado[~chave_novo_serie.isin(chaves_existentes)]
    else:
        df_realmente_novo = df_novo_alinhado
    df_realmente_novo = df_realmente_novo.drop_duplicates(subset=subset, keep="last")

    if df_realmente_novo.empty and linha_totais_excel is None:
        logger.info(
            "Nenhum registro novo para '%s' nesta execução - planilha mantida como está (%d linhas).",
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

    if not df_realmente_novo.empty:
        colunas_texto = _colunas_como_texto(ws, linha_referencia, df_existente.columns) if linha_referencia else set()
        primeira_linha_nova = ws.max_row + 1
        for _, linha in df_realmente_novo.iterrows():
            valores = [
                _valor_como_texto(v) if nome in colunas_texto else _valor_para_excel(v)
                for nome, v in linha.items()
            ]
            ws.append(valores)
        if linha_referencia:
            for row in range(primeira_linha_nova, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    _formatar_celula_como(
                        ws.cell(row=row, column=col), ws.cell(row=linha_referencia, column=col), alinhar_direita=True,
                    )

    df_tudo = pd.concat([df_existente, df_realmente_novo], ignore_index=True)
    totais = _calcular_linha_totais(df_tudo)
    ws.append([_valor_para_excel(totais.get(c)) for c in df_existente.columns])
    if linha_referencia:
        for col in range(1, ws.max_column + 1):
            _formatar_celula_como(ws.cell(row=ws.max_row, column=col), ws.cell(row=linha_referencia, column=col))

    wb.save(path)
    if aplicar_autofiltro:
        _apply_excel_autofilter(path)

    total = len(df_tudo)
    logger.info(
        "Planilha '%s' atualizada: +%d linhas novas (%d no total, linha de totais recalculada e mantida no final)",
        nome_planilha, len(df_realmente_novo), total,
    )
    return len(df_realmente_novo), total


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
        base existir aqui): só recebe registros com chave genuinamente
        nova, anexados no final - nunca remove, sobrescreve ou recolore
        linhas já existentes (`_acumular_origem_comissao_a_vista`).
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

    df_previa_final.to_excel(previa_path, index=False)
    if autofiltro:
        _apply_excel_autofilter(previa_path)
    _marcar_linhas_novas_e_editadas(previa_path, df_previa_final, chave, df_previa_anterior)
    logger.info(
        "Planilha 'Comissão à Vista - Analítico (Prévia)' atualizada: %s (%d linhas)",
        previa_path, len(df_previa_final),
    )

    # --- Planilha de origem oficial: só anexa chave genuinamente nova, sem cor ---
    linhas_novas_origem, linhas_total_origem = _acumular_origem_comissao_a_vista(
        origem_path, df_novo, chave, subset, "Comissão à Vista - Analítico (origem oficial)",
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
