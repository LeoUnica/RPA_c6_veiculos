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
from datetime import date
from pathlib import Path

import pandas as pd

import config

logger = logging.getLogger("data_processor")

# TODO: confirmar o nome exato da coluna de data em cada base
DATE_COLUMN_BY_BASE = {
    "meta_financiamento_seguro": "Data",
    "numero_contratos": "Data Proposta",
    "dias_sem_producao": "Data",
    "carteira_parceiros": None,  # essa base é substituição total, não usa data
}


def _current_month_mask(df: pd.DataFrame, date_col: str) -> pd.Series:
    hoje = date.today()
    dt = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
    return (dt.dt.month == hoje.month) & (dt.dt.year == hoje.year)


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


def _apply_row_filters(df: pd.DataFrame, base: dict) -> pd.DataFrame:
    """Aplica filtros de linha (ex: STATUS PROPOSTA) antes de qualquer corte de coluna."""
    status_filtro = base["regras"].get("filtro_status_proposta")
    if status_filtro:
        # TODO: confirmar o nome exato da coluna ("STATUS PROPOSTA")
        df = df[df["STATUS PROPOSTA"] == status_filtro]
    return df


def merge_into_original(df_novo: pd.DataFrame, original_path: Path, base: dict) -> pd.DataFrame:
    """
    Alinha colunas com a base original, remove o mês atual da base original
    (se existir) e concatena os dados novos. Se a base ainda não tiver o mês
    atual, apenas adiciona os novos registros.
    """
    date_col = DATE_COLUMN_BY_BASE.get(base["id"])

    if not original_path.exists():
        logger.info("Base original não existe ainda, criando: %s", original_path)
        return _align_columns_with_original(df_novo, None, base)

    df_original = pd.read_excel(original_path)
    df_novo = _align_columns_with_original(df_novo, df_original, base)

    if base["regras"].get("remover_mes_atual_antes_de_colar") and date_col and date_col in df_original.columns:
        mask_mes_atual = _current_month_mask(df_original, date_col)
        removidos = mask_mes_atual.sum()
        if removidos:
            logger.info("Removendo %d linhas do mês atual na base original", removidos)
        df_original = df_original[~mask_mes_atual]

    df_final = pd.concat([df_original, df_novo], ignore_index=True)
    return df_final


def process_base(downloaded_path: Path, base: dict) -> Path:
    """
    Pipeline completo para uma base: filtra linhas, funde com a base
    original em staging (alinhando colunas automaticamente) e retorna o
    caminho do arquivo pronto para subir ao SharePoint.

    A base "Carteira e Parceiros" (modo "substituir_arquivo") não passa por
    esse pipeline - veja sharepoint_sync.upload_processed_base, que faz uma
    cópia direta do arquivo baixado sem reprocessar via pandas, conforme a
    instrução da equipe ("cole o arquivo e renomeie").
    """
    original_path = config.STAGING_DIR / f"{base['id']}_original.xlsx"

    if base["regras"].get("modo") == "substituir_arquivo":
        return downloaded_path

    df_novo = pd.read_excel(downloaded_path)
    df_novo = _apply_row_filters(df_novo, base)
    df_final = merge_into_original(df_novo, original_path, base)

    df_final.to_excel(original_path, index=False)
    logger.info("Base '%s' atualizada: %s (%d linhas)", base["nome"], original_path, len(df_final))
    return original_path
