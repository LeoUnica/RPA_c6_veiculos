"""
Configuração central da RPA - Bases C6 Veículos.

Cada item em BASES representa um relatório do Looker que precisa ser
baixado, tratado e consolidado em uma base "original".

Para adicionar uma nova base, basta adicionar um novo dicionário nesta
lista - nenhum outro arquivo precisa ser alterado.
"""

import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

# --------------------------------------------------------------------------
# Pastas locais de trabalho (staging antes de subir pro SharePoint)
# --------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
DOWNLOAD_DIR = BASE_DIR / "downloads"      # onde o Playwright salva o Excel baixado
STAGING_DIR = BASE_DIR / "staging"         # onde ficam as bases originais durante o processamento
LOG_DIR = BASE_DIR / "logs"

for d in (DOWNLOAD_DIR, STAGING_DIR, LOG_DIR):
    d.mkdir(parents=True, exist_ok=True)

# --------------------------------------------------------------------------
# Credenciais / URLs (via variáveis de ambiente - ver .env.example)
# --------------------------------------------------------------------------
LOOKER_URL = os.getenv("LOOKER_URL", "https://c6.c6consig.com.br/WebAutorizador/Login/AC.UI.LOGIN.aspx")
LOOKER_USER = os.getenv("LOOKER_USER")
LOOKER_PASSWORD = os.getenv("LOOKER_PASSWORD")

SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL")  # ex: https://empresa.sharepoint.com/sites/DadosC6
SHAREPOINT_CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
SHAREPOINT_CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")
SHAREPOINT_TENANT_ID = os.getenv("SHAREPOINT_TENANT_ID")

# Caminho da biblioteca de documentos no SharePoint onde ficam as pastas
# (Carteira de parceiros, Dias sem produção, Meta Financiamento e Seguro, etc.)
SHAREPOINT_ROOT_FOLDER = os.getenv("SHAREPOINT_ROOT_FOLDER", "/Documentos Compartilhados/BI C6 Auto")

# Pasta "Prévia" onde fica o arquivo baixado do Looker já tratado (colunas
# selecionadas + filtro de Status Proposta), antes do merge com a planilha
# de origem oficial.
PREVIA_NUMERO_CONTRATOS_DIR = os.getenv(
    "PREVIA_NUMERO_CONTRATOS_DIR",
    r"C:\Users\leonardo.mudrik\Desktop\C6 Bank\Número de Contratos - Previa",
)

# Pasta raiz onde fica a planilha de origem oficial "Número de Contratos" -
# organizada por ano, ex: ".../Numero de Contratos - 2026/Digitação
# Analítico - 2026.xlsx".
PLANILHA_ORIGEM_NUMERO_CONTRATOS_DIR = os.getenv(
    "PLANILHA_ORIGEM_NUMERO_CONTRATOS_DIR",
    r"C:\Users\leonardo.mudrik\Desktop\Setor Dados\Ana Price\Número de Contratos",
)


def caminho_previa_numero_contratos() -> Path:
    """Caminho do arquivo 'prévia' (tratado, antes do merge) de Número de Contratos."""
    return Path(PREVIA_NUMERO_CONTRATOS_DIR) / "Número de Contratos - Previa.xlsx"


def caminho_planilha_origem_numero_contratos(ano: int) -> Path:
    """Caminho da planilha de origem oficial 'Número de Contratos' de um ano específico."""
    return (
        Path(PLANILHA_ORIGEM_NUMERO_CONTRATOS_DIR)
        / f"Numero de Contratos - {ano}"
        / f"Digitação Analítico - {ano}.xlsx"
    )

# Pasta "Prévia" e planilha de origem oficial de "Dias sem Produção" (essa
# base não é organizada por ano - é um único arquivo acumulando tudo).
PREVIA_DIAS_SEM_PRODUCAO_DIR = os.getenv(
    "PREVIA_DIAS_SEM_PRODUCAO_DIR",
    r"C:\Users\leonardo.mudrik\Desktop\C6 Bank\Dias sem produção - Previa",
)
PLANILHA_ORIGEM_DIAS_SEM_PRODUCAO_DIR = os.getenv(
    "PLANILHA_ORIGEM_DIAS_SEM_PRODUCAO_DIR",
    r"C:\Users\leonardo.mudrik\Desktop\Setor Dados\Ana Price\Dias sem produção",
)


def caminho_previa_dias_sem_producao() -> Path:
    """Caminho do arquivo 'prévia' (tratado, antes do merge) de Dias sem Produção."""
    return Path(PREVIA_DIAS_SEM_PRODUCAO_DIR) / "Dias sem produção - Previa.xlsx"


def caminho_planilha_origem_dias_sem_producao() -> Path:
    """Caminho da planilha de origem oficial de Dias sem Produção."""
    return Path(PLANILHA_ORIGEM_DIAS_SEM_PRODUCAO_DIR) / "DIAS SEM PRODUCAO.xlsx"

# Pasta "Prévia" e pasta raiz da planilha de origem oficial de "Meta
# Financiamento e Seguro". Diferente de Número de Contratos, o ano fica no
# NOME do arquivo (não em subpasta): "Meta Financiamento Seguro - {ano}.xlsx",
# todos na mesma pasta.
PREVIA_META_FINANCIAMENTO_SEGURO_DIR = os.getenv(
    "PREVIA_META_FINANCIAMENTO_SEGURO_DIR",
    r"C:\Users\leonardo.mudrik\Desktop\C6 Bank\Meta Financiamento e Seguro - Previa",
)
PLANILHA_ORIGEM_META_FINANCIAMENTO_SEGURO_DIR = os.getenv(
    "PLANILHA_ORIGEM_META_FINANCIAMENTO_SEGURO_DIR",
    r"C:\Users\leonardo.mudrik\Desktop\Setor Dados\Ana Price\Meta Financiamento e Seguro",
)


def caminho_previa_meta_financiamento_seguro() -> Path:
    """Caminho do arquivo 'prévia' (tratado, antes do merge) de Meta Financiamento e Seguro."""
    return Path(PREVIA_META_FINANCIAMENTO_SEGURO_DIR) / "Meta Financiamento e Seguro - Previa.xlsx"


def caminho_planilha_origem_meta_financiamento_seguro(ano: int) -> Path:
    """Caminho da planilha de origem oficial de Meta Financiamento e Seguro de um ano específico."""
    return Path(PLANILHA_ORIGEM_META_FINANCIAMENTO_SEGURO_DIR) / f"Meta Financiamento Seguro - {ano}.xlsx"

# --------------------------------------------------------------------------
# Definição das 4 bases (extraído do material da equipe)
# --------------------------------------------------------------------------
BASES = [
    {
        "id": "meta_financiamento_seguro",
        "nome": "Meta Financiamento e Seguro",
        # Caminho até abrir o painel "Auto" (Relatórios > Relatórios
        # Gerenciais > Auto). A partir daí, clica no link "Resumo Apuração
        # Parceiro 2.0" (dentro do card "Apuração Parceiro 2.0") - tratado
        # à parte em looker_automation.py.
        "looker_path": ["Relatórios", "Relatórios Gerenciais", "Auto"],
        "link_relatorio": "Resumo Apuração Parceiro 2.0",
        "secao_tabela": "Bloco de Metas - Por Filial",
        "pasta_sharepoint": "Meta Financiamento e Seguro",
        "frequencia": "mensal",
        "regras": {
            # Esta base não usa SharePoint: os dados tratados são salvos na
            # pasta "Prévia" e depois mesclados com a planilha de origem
            # oficial local, por ano (ver
            # data_processor._process_meta_financiamento_seguro).
            "modo": "planilha_origem_local_meta_financiamento_seguro",
            # Colunas a manter na planilha baixada - o restante é excluído.
            "colunas_manter": [
                "Anomes Apuracao",
                "Filial",
                "R$ Meta",
                "R$ Produção",
                "R$ Meta Seguros",
                "R$ Seguros",
            ],
            "remover_colunas": [],
            "filtro_status_proposta": None,
            "aplicar_autofiltro_excel": True,
        },
    },
    {
        "id": "numero_contratos",
        "nome": "Número de Contratos",
        # Caminho de menu até abrir o painel "Auto" (Relatórios > Relatórios
        # Gerenciais > Auto). A partir daí, o card "Acompanhamento" e o botão
        # "Analítico" são tratados à parte em looker_automation.py, pois o
        # fluxo dessa base tem passos próprios (filtros, update, download
        # avançado) que não se aplicam às outras bases.
        "looker_path": ["Relatórios", "Relatórios Gerenciais", "Auto"],
        "card_acompanhamento": "Acompanhamento Veículos",  # dentro do card "Acompanhamento"
        "aba_relatorio": "Analítico",
        "filtro_valor": "este_mes",
        "bloco": None,
        # Filtros aplicados no painel lateral direito, específicos dessa base
        "filtros": {
            "tipo_exibicao": "Valor",              # manter somente "Valor" em Tipo Exibição
            "periodo_dt_relatorio": "Last 30 Days",  # Dt Relatorio Date -> Last 30 Days
        },
        "pasta_sharepoint": "Número de Contratos",
        "frequencia": "diaria",
        "regras": {
            # Esta base não usa SharePoint: os dados tratados são salvos na
            # pasta "Prévia" e depois mesclados com a planilha de origem
            # oficial local, organizada por ano (ver
            # data_processor._process_numero_contratos).
            "modo": "planilha_origem_local",
            # Colunas a manter na planilha "Analítico" baixada - o restante é excluído.
            "colunas_manter": [
                "ID Proposta",
                "Dt Relatório",
                "Lojista",
                "GP",
                "Status Proposta",
                "Cd Contrato",
                "Vl Principal",
                "Vl Financiamento",
                "(R$) Seguro Prestamista",
                "Filial",
            ],
            "remover_colunas": [],
            "filtro_status_proposta": "PROPOSTA PAGA",   # filtrar coluna Status Proposta
            "aplicar_autofiltro_excel": True,
        },
    },
    {
        "id": "dias_sem_producao",
        "nome": "Dias sem Produção",
        # Caminho até abrir o painel "Auto" (Relatórios > Relatórios
        # Gerenciais > Auto). A partir daí, clica no link "SLA - Última
        # atuação comercial - Analítico" (dentro do card "SLA - Última
        # atuação da loja") - tratado à parte em looker_automation.py.
        "looker_path": ["Relatórios", "Relatórios Gerenciais", "Auto"],
        "link_relatorio": "SLA - Última atuação comercial - Analítico",
        "bloco": None,
        "pasta_sharepoint": "Dias sem produção",
        "frequencia": "semanal_segunda",         # normalmente às segundas-feiras
        "regras": {
            # Esta base não usa SharePoint: os dados tratados são salvos na
            # pasta "Prévia" e depois mesclados com a planilha de origem
            # oficial local (ver data_processor._process_dias_sem_producao).
            "modo": "planilha_origem_local_dias_sem_producao",
            # Colunas a manter na planilha baixada - o restante é excluído.
            "colunas_manter": [
                "Cnpj Da Loja",
                "Cd Loja",
                "Nm Loja",
                "Safra Mes",
                "Faixa Qtde Meses Ult Simulacao",
                "Data Ultima Simulacao Date",
                "Faixa Qtde Meses Ult Proposta",
                "Data Ultima Proposta Date",
                "Faixa Qtde Meses Ult Contrato",
                "Data Ultimo Contrato Date",
                "Valor Financiamento",
                "Qtde. Financiamento",
            ],
            "remover_colunas": [],
            "filtro_status_proposta": None,
            "aplicar_autofiltro_excel": True,
        },
    },
    {
        "id": "carteira_parceiros",
        "nome": "Carteira e Parceiros",
        "looker_path": ["Auto", "Carteira e Parceiros"],
        "filtro_valor": "este_ano",               # "In this Year"
        "bloco": None,
        "pasta_sharepoint": "Carteira de parceiros e filiais",
        "frequencia": "diaria",
        "regras": {
            # Esta base não faz merge nem passa pelo pandas: o arquivo
            # baixado é copiado como está e sobe para o SharePoint com o
            # nome da base já existente ("cole e renomeie", conforme o
            # material da equipe) - ver data_processor.process_base.
            "modo": "substituir_arquivo",
            "remover_colunas": [],
            "filtro_status_proposta": None,
            "remover_mes_atual_antes_de_colar": False,
        },
    },
]


def get_base_by_id(base_id: str) -> dict:
    for base in BASES:
        if base["id"] == base_id:
            return base
    raise ValueError(f"Base '{base_id}' não encontrada em config.py")
