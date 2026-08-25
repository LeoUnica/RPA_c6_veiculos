"""
Configuração central da RPA - Bases C6 Veículos.

Cada item em BASES representa um relatório do Looker que precisa ser
baixado, tratado e consolidado em uma base "original".

Para adicionar uma nova base, basta adicionar um novo dicionário nesta
lista - nenhum outro arquivo precisa ser alterado.
"""

import os
from datetime import date
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


def periodo_referencia_atual() -> tuple[int, int]:
    """(ano, mês) atuais - período de referência "corrente" usado por todas as
    bases mensais/de mês corrente. Centralizado aqui para garantir que todas
    usem exatamente a mesma noção de "mês atual" (ex: Comissão à Vista sempre
    bate com o mês/ano de Número de Contratos)."""
    hoje = date.today()
    return hoje.year, hoje.month

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
    """Caminho da planilha 'Digitação Analítico - {ano} - Trimestre', que
    mantém uma JANELA MÓVEL dos últimos 90 dias (uma linha sai da planilha
    quando fica velha demais) - ver `data_processor._process_numero_contratos`.
    O acumulado do ano inteiro fica por conta da planilha anual abaixo."""
    return (
        Path(PLANILHA_ORIGEM_NUMERO_CONTRATOS_DIR)
        / f"Numero de Contratos - {ano}"
        / f"Digitação Analítico - {ano} - Trimestre.xlsx"
    )


def caminho_planilha_origem_numero_contratos_anual(ano: int) -> Path:
    """Caminho da planilha 'Digitação Analítico - {ano}' (sem sufixo), que
    acumula todo o histórico do ano (nunca remove uma linha) na mesma pasta
    da planilha "Trimestre" acima - ver `data_processor._process_numero_contratos`.
    Um ano diferente do corrente (ex: 2025 depois de fechado) nunca é
    escrito de novo pela automação - ver `data_processor._eh_ano_corrente`."""
    return (
        Path(PLANILHA_ORIGEM_NUMERO_CONTRATOS_DIR)
        / f"Numero de Contratos - {ano}"
        / f"Digitação Analítico - {ano}.xlsx"
    )

# Pasta "Prévia" e planilha de origem oficial de "Dias sem Produção" - por
# ano (mesmo padrão de Meta Financiamento e Seguro/Carteira e Parceiros),
# ex: "DIAS SEM PRODUCAO - 2026.xlsx".
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


def caminho_planilha_origem_dias_sem_producao(ano: int) -> Path:
    """Caminho da planilha de origem oficial de Dias sem Produção de um ano
    específico. Um ano diferente do corrente (ex: 2025 depois de fechado)
    nunca é escrito de novo pela automação - ver `data_processor._eh_ano_corrente`."""
    return Path(PLANILHA_ORIGEM_DIAS_SEM_PRODUCAO_DIR) / f"DIAS SEM PRODUCAO - {ano}.xlsx"

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


# Pasta "Prévia" e pasta raiz da planilha de origem oficial de "Carteira e
# Parceiros" - também um arquivo por ano na mesma pasta (sem subpasta),
# ex: "CARTEIRA- 2026.xlsx".
PREVIA_CARTEIRA_PARCEIROS_DIR = os.getenv(
    "PREVIA_CARTEIRA_PARCEIROS_DIR",
    r"C:\Users\leonardo.mudrik\Desktop\C6 Bank\Carteira de parceiros e filiais - Previa",
)
PLANILHA_ORIGEM_CARTEIRA_PARCEIROS_DIR = os.getenv(
    "PLANILHA_ORIGEM_CARTEIRA_PARCEIROS_DIR",
    r"C:\Users\leonardo.mudrik\Desktop\Setor Dados\Ana Price\Carteira de parceiros e filiais",
)


def caminho_previa_carteira_parceiros() -> Path:
    """Caminho do arquivo 'prévia' (substituído a cada execução) de Carteira e Parceiros."""
    return Path(PREVIA_CARTEIRA_PARCEIROS_DIR) / "Carteira de parceiros e filiais - Previa.xlsx"


def caminho_planilha_origem_carteira_parceiros(ano: int) -> Path:
    """Caminho da planilha de origem oficial de Carteira e Parceiros de um ano específico."""
    return Path(PLANILHA_ORIGEM_CARTEIRA_PARCEIROS_DIR) / f"CARTEIRA- {ano}.xlsx"

# Pasta "Prévia" de "Comissão à Vista - Analítico" - acumulada indefinidamente
# entre execuções (nunca sobrescrita, só recebe linhas novas - ver
# data_processor._process_comissao_a_vista), mesma lógica da planilha de
# origem oficial abaixo, só em arquivo diferente.
PREVIA_COMISSAO_A_VISTA_DIR = os.getenv(
    "PREVIA_COMISSAO_A_VISTA_DIR",
    r"C:\Users\leonardo.mudrik\Desktop\C6 Bank\Comissão à Vista - Analitico - Previa",
)


def caminho_previa_comissao_a_vista() -> Path:
    """Caminho da planilha (única, acumulativa) de Comissão à Vista - Analítico."""
    return Path(PREVIA_COMISSAO_A_VISTA_DIR) / "Comissão à Vista - Analitico - Previa.xlsx"

# Planilha de origem oficial de "Comissão à Vista - Analítico" - por ano
# (mesmo padrão das outras bases), ex: "Comissão à Vista - Analítico - 2026.xlsx".
PLANILHA_ORIGEM_COMISSAO_A_VISTA_DIR = os.getenv(
    "PLANILHA_ORIGEM_COMISSAO_A_VISTA_DIR",
    r"C:\Users\leonardo.mudrik\Desktop\Setor Dados\Ana Price\Comissao à Vista",
)


def caminho_planilha_origem_comissao_a_vista(ano: int) -> Path:
    """Caminho da planilha de origem oficial de Comissão à Vista - Analítico
    de um ano específico. Um ano diferente do corrente (ex: 2025 depois de
    fechado) nunca é escrito de novo pela automação - ver
    `data_processor._eh_ano_corrente`."""
    return Path(PLANILHA_ORIGEM_COMISSAO_A_VISTA_DIR) / f"Comissão à Vista - Analítico - {ano}.xlsx"

# --------------------------------------------------------------------------
# Definição das 5 bases.
# --------------------------------------------------------------------------
BASES = [
    {
        "id": "meta_financiamento_seguro",
        "nome": "Meta Financiamento e Seguro",
        # Relatórios > Relatórios Gerenciais > Auto > link "Resumo Apuração
        # Parceiro 2.0" (card "Apuração Parceiro 2.0") - ver looker_automation.py.
        "looker_path": ["Relatórios", "Relatórios Gerenciais", "Auto"],
        "link_relatorio": "Resumo Apuração Parceiro 2.0",
        "secao_tabela": "Bloco de Metas - Por Filial",
        "pasta_sharepoint": "Meta Financiamento e Seguro",
        "frequencia": "mensal",
        "regras": {
            # Não usa SharePoint: dados vão para a "Prévia" e depois são
            # mesclados na planilha de origem local, por ano (ver
            # data_processor._process_meta_financiamento_seguro). Busca o
            # ANO CORRENTE inteiro a cada execução (ver
            # looker_automation.download_meta_financiamento_seguro_report).
            "modo": "planilha_origem_local_meta_financiamento_seguro",
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
        # Relatórios > Relatórios Gerenciais > Auto > card "Acompanhamento" >
        # botão "Analítico" - fluxo próprio (filtros, update, download
        # avançado) em looker_automation.py.
        "looker_path": ["Relatórios", "Relatórios Gerenciais", "Auto"],
        "card_acompanhamento": "Acompanhamento Veículos",  # dentro do card "Acompanhamento"
        "aba_relatorio": "Analítico",
        "filtro_valor": "este_mes",
        "bloco": None,
        # Filtros aplicados no painel lateral direito, específicos dessa base
        "filtros": {
            "tipo_exibicao": "Valor",  # manter somente "Valor" em Tipo Exibição
            # Dt Relatorio Date -> Year To Date: baixa o ano inteiro (1º de
            # janeiro até hoje). A planilha "Trimestre" continua recortada
            # para os últimos 90 dias DEPOIS do merge (ver
            # data_processor.DIAS_JANELA_TRIMESTRE_NUMERO_CONTRATOS), mas a
            # "_Anual" recebe o ano inteiro a cada execução - ver
            # looker_automation._alterar_periodo_dt_relatorio.
            "periodo_dt_relatorio": "Year To Date",
        },
        "pasta_sharepoint": "Número de Contratos",
        "frequencia": "diaria",
        "regras": {
            # Não usa SharePoint: dados vão para a "Prévia" e depois são
            # mesclados na planilha de origem local por ano (ver
            # data_processor._process_numero_contratos).
            "modo": "planilha_origem_local",
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
        # Relatórios > Relatórios Gerenciais > Auto > link "SLA - Última
        # atuação comercial - Analítico" (card "SLA - Última atuação da loja").
        "looker_path": ["Relatórios", "Relatórios Gerenciais", "Auto"],
        "link_relatorio": "SLA - Última atuação comercial - Analítico",
        "bloco": None,
        "pasta_sharepoint": "Dias sem produção",
        "frequencia": "semanal_segunda",  # normalmente às segundas-feiras
        "regras": {
            # Não usa SharePoint: dados vão para a "Prévia" e depois são
            # mesclados na planilha de origem local, por ano (ver
            # data_processor._process_dias_sem_producao).
            "modo": "planilha_origem_local_dias_sem_producao",
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
        # Relatórios > Relatórios Gerenciais > Auto > link "Carteira" (card "Carteira").
        "looker_path": ["Relatórios", "Relatórios Gerenciais", "Auto"],
        "link_relatorio": "Carteira",
        "pasta_sharepoint": "Carteira de parceiros e filiais",
        "frequencia": "diaria",
        "regras": {
            # Não usa SharePoint: dados baixados (todas as colunas, sem
            # filtro) substituem a "Prévia" por inteiro a cada execução
            # (filtro "Referência" = "Este Ano") e só as linhas novas são
            # coladas na planilha de origem oficial daquele ano - ver
            # data_processor._process_carteira_parceiros.
            "modo": "planilha_origem_local_carteira_parceiros",
            "remover_colunas": [],
            "filtro_status_proposta": None,
        },
    },
    {
        "id": "comissao_a_vista",
        "nome": "Comissão à Vista - Analítico",
        # Relatórios > Relatórios Gerenciais > One Page - Operacional > Auto >
        # link "Apuração Comissão À Vista" (mesmo card de Meta Financiamento
        # e Seguro, "Apuração Parceiro 2.0", link diferente). O "À" vem
        # maiúsculo no portal - `exact=True` exige bater exatamente.
        "looker_path": ["Relatórios", "Relatórios Gerenciais", "Auto"],
        "link_relatorio": "Apuração Comissão À Vista",
        # Texto de referência (faixa de título acima da tabela) usado para
        # achar o botão "Tile actions" no download - mesmo padrão de
        # "Analítico" em Número de Contratos.
        "secao_tabela": "Analítico",
        "pasta_sharepoint": "Comissão à Vista - Analítico",
        "frequencia": "mensal",
        "regras": {
            # Não usa SharePoint - dados vão direto pra Prévia e pra
            # planilha de origem oficial local, por ano (ver
            # data_processor._process_comissao_a_vista). Nenhuma das duas
            # planilhas tem período removido a cada execução: as duas só
            # RECEBEM linhas novas no final (nunca sobrescritas).
            "modo": "planilha_origem_local_comissao_a_vista",
            "remover_colunas": [],
            "filtro_status_proposta": None,
            "aplicar_autofiltro_excel": True,
            # Chave composta para identificar registros já existentes e
            # evitar duplicação ao acumular (ver
            # data_processor._process_comissao_a_vista): "Cd Contrato"
            # sozinho tem duplicatas reais (contrato reajustado num mês
            # seguinte) - "Anomes Apuracao" junto elimina isso.
            "chave_comparacao": ["Cd Contrato", "Anomes Apuracao"],
        },
    },
]


def get_base_by_id(base_id: str) -> dict:
    for base in BASES:
        if base["id"] == base_id:
            return base
    raise ValueError(f"Base '{base_id}' não encontrada em config.py")


def validar_ambiente() -> None:
    """
    Confere que as variáveis de ambiente obrigatórias (credenciais do
    portal C6) estão presentes antes de a execução começar - falha rápido e
    claro aqui em vez de um erro confuso do Playwright mais tarde. Chamada
    uma vez no início de `main.main()`, antes de abrir qualquer navegador.
    """
    faltando = [
        nome for nome, valor in {
            "LOOKER_USER": LOOKER_USER,
            "LOOKER_PASSWORD": LOOKER_PASSWORD,
        }.items()
        if not valor
    ]
    if faltando:
        raise RuntimeError(
            "Variável(is) de ambiente obrigatória(s) ausente(s) no .env: "
            f"{', '.join(faltando)}. Veja .env.example para o formato esperado."
        )
