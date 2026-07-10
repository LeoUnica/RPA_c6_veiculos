"""
Tratamento de dados das bases baixadas do Looker.

Cada uma das 4 bases (numero_contratos, dias_sem_producao,
meta_financiamento_seguro, carteira_parceiros) tem seu próprio fluxo de
tratamento dedicado (`_process_*`), pois cada uma tem regras próprias de
seleção de colunas, identificação do "período atual" e destino final -
ver `process_base` para o dispatch entre elas.
"""

import logging
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

import config

logger = logging.getLogger("data_processor")

DATE_COLUMN_NUMERO_CONTRATOS = "Dt Relatório"


def _current_month_mask(df: pd.DataFrame, date_col: str) -> pd.Series:
    hoje = date.today()
    dt = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
    return (dt.dt.month == hoje.month) & (dt.dt.year == hoje.year)


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
    return origem_path


CHAVE_UNICA_DIAS_SEM_PRODUCAO = ["Cd Loja", "Safra Mes"]


def _safra_mes_atual() -> int:
    """Mês/ano atual no formato AAAAMM, igual à coluna 'Safra Mes' do relatório."""
    hoje = date.today()
    return hoje.year * 100 + hoje.month


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
    df_tratado = _select_columns(df_tratado, base)

    previa_path = config.caminho_previa_dias_sem_producao()
    previa_path.parent.mkdir(parents=True, exist_ok=True)
    df_tratado.to_excel(previa_path, index=False)
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(previa_path)
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
    df_tratado = _select_columns(df_tratado, base)

    previa_path = config.caminho_previa_meta_financiamento_seguro()
    previa_path.parent.mkdir(parents=True, exist_ok=True)
    df_tratado.to_excel(previa_path, index=False)
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(previa_path)
    logger.info("Base '%s' tratada (prévia): %s (%d linhas)", base["nome"], previa_path, len(df_tratado))

    anos_presentes = sorted(df_tratado["Anomes Apuracao"].astype(str).str[:4].unique())
    origem_paths = []

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

    return origem_paths[-1]


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
    """
    df_tratado = pd.read_excel(downloaded_path)

    previa_path = config.caminho_previa_carteira_parceiros()
    previa_path.parent.mkdir(parents=True, exist_ok=True)
    df_tratado.to_excel(previa_path, index=False)
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(previa_path)
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
    df_final.to_excel(origem_path, index=False)
    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(origem_path)

    logger.info(
        "Planilha de origem atualizada: %s (+%d linhas do mês atual, %d no total)",
        origem_path, len(df_novo_mes), len(df_final),
    )
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

    raise ValueError(f"Base '{base['id']}' não tem 'modo' de tratamento reconhecido em config.py")
