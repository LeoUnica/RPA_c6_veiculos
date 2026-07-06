# RPA - Bases C6 Veículos

## 📌 Objetivo

Automatizar o processo manual de atualização das bases de dados utilizadas
pela equipe de análise de dados do time de Veículos (C6).

Hoje, esse processo é feito manualmente: uma pessoa entra no Looker, navega
até o relatório certo, aplica um filtro, baixa o Excel, limpa colunas
desnecessárias, remove os dados do mês atual da planilha "original" (para
evitar duplicidade) e cola os dados novos por cima — repetindo isso para 4
bases diferentes, algumas diariamente e outras semanalmente.

Este projeto automatiza esse fluxo por completo:

1. Login e navegação automática no Looker
2. Aplicação do filtro correto (mês atual / ano atual)
3. Download do relatório em Excel
4. Limpeza dos dados (remoção de colunas, filtro de status)
5. Atualização da base original (remove mês atual duplicado + concatena os
   dados novos)
6. Upload do resultado final de volta para o SharePoint/OneDrive

## 🗂️ Estrutura do Projeto

```
rpa_c6_veiculos/
├── config.py              # Definição das 4 bases: caminho no Looker, filtro,
│                           # pasta de destino, frequência e regras de negócio
├── looker_automation.py   # Automação de login, navegação e download no Looker (Playwright)
├── data_processor.py      # Limpeza, filtro e merge dos dados (pandas)
├── sharepoint_sync.py     # Download/upload dos arquivos no SharePoint (Office365-REST-Python-Client)
├── main.py                # Orquestrador: roda uma base específica, todas, ou por frequência
├── requirements.txt       # Dependências do projeto
├── .env.example           # Modelo de variáveis de ambiente (credenciais)
├── downloads/             # Pasta de trabalho: arquivos baixados do Looker (gerada em runtime)
├── staging/               # Pasta de trabalho: bases originais durante o processamento (gerada em runtime)
└── logs/                  # Logs de execução (gerada em runtime)
```

### Fluxo entre os módulos

```
main.py
  ├──> sharepoint_sync.py   (baixa a base original atual)
  ├──> looker_automation.py (baixa o relatório novo do Looker)
  ├──> data_processor.py    (limpa e mescla os dados)
  └──> sharepoint_sync.py   (sobe o resultado final)
```

## 🛠️ Linguagens e Ferramentas Utilizadas

| Categoria | Tecnologia | Uso no projeto |
|---|---|---|
| Linguagem | **Python 3.11+** | Linguagem principal de todo o projeto |
| Automação de navegador (RPA) | **Playwright** | Login, navegação por menus e download dos relatórios no Looker |
| Tratamento de dados | **pandas** | Limpeza de colunas, filtros e merge das planilhas |
| Leitura/escrita de Excel | **openpyxl** | Suporte ao pandas para arquivos `.xlsx` |
| Integração com SharePoint | **Office365-REST-Python-Client** | Download/upload de arquivos nas pastas do SharePoint/OneDrive |
| Autenticação SharePoint | **Azure AD App Registration** | Client ID / Client Secret / Tenant ID para acesso via API |
| Configuração de ambiente | **python-dotenv** | Carrega credenciais do arquivo `.env` sem expor no código |
| Agendamento | **Windows Task Scheduler** (ou `cron` em Linux) | Executa `main.py` automaticamente nos horários definidos |
| Logging | **módulo `logging`** (nativo do Python) | Registro de execução e falhas em `logs/rpa.log` |

## ▶️ Como Rodar

```bash
# 1. Criar e ativar ambiente virtual
python -m venv venv
venv\Scripts\activate          # Windows
source venv/bin/activate       # Mac/Linux

# 2. Instalar dependências
pip install -r requirements.txt
playwright install chromium

# 3. Configurar credenciais
copy .env.example .env         # preencher com dados reais

# 4. Testar uma base isolada
python main.py --base numero_contratos

# 5. Rodar todas as bases de uma frequência (uso em agendamento)
python main.py --frequencia diaria
python main.py --frequencia semanal
```

## 📋 Bases Automatizadas

| Base | Frequência | Filtro |
|---|---|---|
| Meta Financiamento e Seguro | Semanal | Este mês |
| Número de Contratos | Diária | Este mês |
| Dias sem Produção | Semanal (segundas) | Este mês |
| Carteira e Parceiros | Diária | Este ano |

## ⚠️ Status Atual

- ✅ Lógica de tratamento de dados (`data_processor.py`) implementada e testada
- ⚠️ Seletores do Playwright em `looker_automation.py` são um esqueleto —
  precisam ser ajustados com `playwright codegen` olhando o Looker real
- ⚠️ Integração com SharePoint requer configuração de App Registration no
  Azure AD antes do primeiro uso

