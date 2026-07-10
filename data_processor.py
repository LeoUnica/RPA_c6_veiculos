"""
Tratamento de dados das bases baixadas do Looker.

Implementa as regras descritas pela equipe de análise:
  1. (Se aplicável) filtrar por status da proposta.
  2. Alinhar as colunas da base baixada com as colunas da base original -
     isso substitui o passo manual de "comparar via PROCX" descrito para a
     base Dias sem Produção: qualquer coluna que exista no arquivo baixado
     mas não exista na base original é descartada automaticamente.
  3. Remover, na base original, os dados do mês atual (se já existirem).
  4. Colar/concatenar os dados novos na base original.

Assume que existe uma coluna de data em cada base para identificar "o mês
atual". Ajuste DATE_COLUMN_BY_BASE conforme o nome real da coluna em cada
relatório.
"""

import logging
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

import config

logger = logging.getLogger("data_processor")

# TODO: confirmar o nome exato da coluna de data em cada base
DATE_COLUMN_BY_BASE = {
    "meta_financiamento_seguro": "Data",
    "numero_contratos": "Dt Relatório",
    "carteira_parceiros": None,  # essa base é substituição total, não usa data
    # "dias_sem_producao" não usa esse mecanismo genérico - o mês é
    # identificado pela coluna "Safra Mes" (formato AAAAMM), tratada em
    # `_process_dias_sem_producao`.
}


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


def _align_columns_with_original(df_novo: pd.DataFrame, df_original: pd.DataFrame | None, base: dict) -> pd.DataFrame:
    """
    Mantém, na base baixada, apenas as colunas que também existem na base
    original (na mesma ordem da base original). Automatiza o que o material
    da equipe descreve como "comparar via PROCX" para saber quais colunas
    excluir do arquivo baixado.

    Se a base original ainda não existir (primeira execução), não há
    referência para alinhar - nesse caso cai de volta na lista manual
    `remover_colunas` de config.py, se houver alguma.
    """
    if df_original is None:
        colunas_remover = base["regras"].get("remover_colunas") or []
        if colunas_remover:
            df_novo = df_novo.drop(columns=[c for c in colunas_remover if c in df_novo.columns])
        return df_novo

    colunas_originais = list(df_original.columns)
    colunas_baixadas = set(df_novo.columns)

    faltando = [c for c in colunas_originais if c not in colunas_baixadas]
    if faltando:
        logger.warning("Colunas presentes na base original mas ausentes no arquivo baixado: %s", faltando)

    extras = [c for c in df_novo.columns if c not in colunas_originais]
    if extras:
        logger.info("Removendo colunas do arquivo baixado que não existem na base original: %s", extras)

    colunas_finais = [c for c in colunas_originais if c in colunas_baixadas]
    return df_novo[colunas_finais]


def _select_columns(df_novo: pd.DataFrame, df_original: pd.DataFrame | None, base: dict) -> pd.DataFrame:
    """
    Define quais colunas vão para a planilha final.

    Se a base tiver uma lista explícita `colunas_manter` em config.py (ex:
    numero_contratos), essa lista é a fonte da verdade e é aplicada
    diretamente - as demais colunas do arquivo baixado são descartadas.
    Caso contrário, cai no alinhamento automático com a base original.
    """
    colunas_manter = base["regras"].get("colunas_manter")
    if colunas_manter:
        faltando = [c for c in colunas_manter if c not in df_novo.columns]
        if faltando:
            logger.warning("Colunas esperadas não encontradas no arquivo baixado: %s", faltando)
        colunas_presentes = [c for c in colunas_manter if c in df_novo.columns]
        return df_novo[colunas_presentes]

    return _align_columns_with_original(df_novo, df_original, base)


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


def merge_into_original(df_novo: pd.DataFrame, original_path: Path, base: dict) -> pd.DataFrame:
    """
    Alinha colunas com a base original, remove o mês atual da base original
    (se existir) e concatena os dados novos. Se a base ainda não tiver o mês
    atual, apenas adiciona os novos registros.

    Quando `remover_mes_atual_antes_de_colar` está ativo, o arquivo baixado
    também é restrito ao mês atual antes de colar - isso evita duplicar
    meses anteriores quando o relatório usa uma janela "rolante" (ex: "Last
    30 Days"), que sempre inclui alguns dias do mês anterior junto com o
    mês atual.
    """
    date_col = DATE_COLUMN_BY_BASE.get(base["id"])
    remover_mes_atual = base["regras"].get("remover_mes_atual_antes_de_colar")

    if not original_path.exists():
        logger.info("Base original não existe ainda, criando: %s", original_path)
        return _select_columns(df_novo, None, base)

    df_original = pd.read_excel(original_path)
    df_novo = _select_columns(df_novo, df_original, base)

    if remover_mes_atual and date_col and date_col in df_original.columns:
        mask_mes_atual_original = _current_month_mask(df_original, date_col)
        removidos = mask_mes_atual_original.sum()
        if removidos:
            logger.info("Removendo %d linhas do mês atual na base original", removidos)
        df_original = df_original[~mask_mes_atual_original]

        if date_col in df_novo.columns:
            mask_mes_atual_novo = _current_month_mask(df_novo, date_col)
            fora_do_mes = int((~mask_mes_atual_novo).sum())
            if fora_do_mes:
                logger.info(
                    "Descartando %d linhas do arquivo baixado que não são do mês atual "
                    "(já devem existir na base original de meses anteriores)",
                    fora_do_mes,
                )
            df_novo = df_novo[mask_mes_atual_novo]

    df_final = pd.concat([df_original, df_novo], ignore_index=True)
    return df_final


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
    date_col = DATE_COLUMN_BY_BASE.get(base["id"])

    def _ordenar_por_data(df: pd.DataFrame) -> pd.DataFrame:
        if date_col and date_col in df.columns:
            return df.sort_values(by=date_col, ascending=True, kind="stable").reset_index(drop=True)
        return df

    def _apenas_mes_atual(df: pd.DataFrame) -> pd.DataFrame:
        if date_col and date_col in df.columns and not df.empty:
            return df[_current_month_mask_com_virada(df, date_col)]
        return df

    df_tratado = pd.read_excel(downloaded_path)
    df_tratado = _apply_row_filters(df_tratado, base)
    df_tratado = _select_columns(df_tratado, None, base)
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
    df_tratado = _select_columns(df_tratado, None, base)

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


def process_base(downloaded_path: Path, base: dict) -> Path:
    """
    Pipeline completo para uma base: filtra linhas, funde com a base
    original em staging (alinhando colunas automaticamente) e retorna o
    caminho do arquivo pronto para subir ao SharePoint.

    A base "Carteira e Parceiros" (modo "substituir_arquivo") não passa por
    esse pipeline - veja sharepoint_sync.upload_processed_base, que faz uma
    cópia direta do arquivo baixado sem reprocessar via pandas, conforme a
    instrução da equipe ("cole o arquivo e renomeie").

    A base "Número de Contratos" (modo "planilha_origem_local") também não
    usa SharePoint: os dados tratados vão para a pasta "Prévia" e depois
    são mesclados com a planilha de origem oficial local, organizada por
    ano - ver `_process_numero_contratos`.
    """
    modo = base["regras"].get("modo")

    if modo == "substituir_arquivo":
        return downloaded_path

    if modo == "planilha_origem_local":
        return _process_numero_contratos(downloaded_path, base)

    if modo == "planilha_origem_local_dias_sem_producao":
        return _process_dias_sem_producao(downloaded_path, base)

    original_path = config.STAGING_DIR / f"{base['id']}_original.xlsx"

    df_novo = pd.read_excel(downloaded_path)
    df_novo = _apply_row_filters(df_novo, base)
    df_final = merge_into_original(df_novo, original_path, base)

    df_final.to_excel(original_path, index=False)

    if base["regras"].get("aplicar_autofiltro_excel"):
        _apply_excel_autofilter(original_path)

    logger.info("Base '%s' atualizada: %s (%d linhas)", base["nome"], original_path, len(df_final))
    return original_path
