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

1. Login (único, reaproveitado para todas as bases da execução) e navegação
   automática no Looker
2. Aplicação do filtro correto (mês atual / ano atual)
3. Download do relatório em Excel
4. Limpeza dos dados (remoção de colunas, filtro de status)
5. Atualização da planilha "Prévia" e da base "de origem" oficial (remove o
   período duplicado + concatena os dados novos), com marcação de cor
   (verde = linha nova, amarelo = linha editada)

> As 4 bases hoje configuradas usam pastas locais (normalmente
> sincronizadas por OneDrive) como destino - **nenhuma delas usa
> SharePoint via API**. O módulo `sharepoint_sync.py` existe no repositório
> para uma eventual base futura que precise disso, mas não é chamado hoje.

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
main.py (run_bases)
  ├──> looker_automation.py  (1 login só, baixa o relatório de cada base
  │                           na mesma sessão - se uma base falhar na
  │                           navegação, loga o erro e segue para a próxima)
  └──> data_processor.py     (limpa, mescla e marca de cor cada arquivo baixado)
```

`sharepoint_sync.py` só entraria no fluxo (download da base original antes
e upload do resultado depois) para uma base configurada com um `modo`
diferente de `planilha_origem_local*` - nenhuma das 4 bases atuais se
encaixa nisso.

## 🛠️ Linguagens e Ferramentas Utilizadas

| Categoria | Tecnologia | Uso no projeto |
|---|---|---|
| Linguagem | **Python 3.11+** | Linguagem principal de todo o projeto |
| Automação de navegador (RPA) | **Playwright** | Login, navegação por menus e download dos relatórios no Looker |
| Tratamento de dados | **pandas** | Limpeza de colunas, filtros e merge das planilhas |
| Leitura/escrita de Excel | **openpyxl** | Suporte ao pandas para arquivos `.xlsx` |
| Integração com SharePoint | **Office365-REST-Python-Client** | Presente em `sharepoint_sync.py`, mas **não usada por nenhuma das 4 bases atuais** (todas gravam em pasta local) |
| Autenticação SharePoint | **Azure AD App Registration** | Só seria necessário se alguma base futura usar `sharepoint_sync.py` |
| Configuração de ambiente | **python-dotenv** | Carrega credenciais do arquivo `.env` sem expor no código |
| Agendamento | **Windows Task Scheduler** (ou `cron` em Linux) | Executa `main.py` automaticamente nos horários definidos |
| Logging | **módulo `logging`** (nativo do Python) | Registro de execução e falhas em `logs/rpa.log` |
| Feriados/dias úteis | **holidays** | Usada por Meta Financiamento e Seguro para tratar a virada de mês (ver `looker_automation.deve_usar_janela_curta_safra_mes`) |

## ▶️ Como Rodar

> ⚠️ Em um computador diferente do que foi usado no desenvolvimento, os
> caminhos de pasta padrão em `config.py` (`Desktop\C6 Bank\...`) não vão
> existir. Antes de rodar em outra máquina, siga o **[SETUP.md](SETUP.md)**,
> que cobre a configuração obrigatória de caminhos por `.env`.

```bash
# 1. Criar e ativar ambiente virtual
python -m venv venv
venv\Scripts\activate          # Windows
source venv/bin/activate       # Mac/Linux

# 2. Instalar dependências
pip install -r requirements.txt
playwright install chromium

# 3. Configurar credenciais e caminhos de pasta
copy .env.example .env         # preencher com dados reais - ver SETUP.md

# 4. Testar uma base isolada
python main.py --base numero_contratos

# 5. Rodar todas as bases de uma frequência (uso em agendamento)
python main.py --frequencia diaria
python main.py --frequencia semanal_segunda
python main.py --frequencia mensal
```

## 📋 Bases Automatizadas

| Base | Frequência | Filtro |
|---|---|---|
| Meta Financiamento e Seguro | Mensal | Safra Mês = Este mês |
| Número de Contratos | Diária | Dt Relatório = Last 30 Days (tratamento restringe ao mês atual) |
| Dias sem Produção | Semanal (segundas) | Referência Month = Este mês |
| Carteira e Parceiros | Diária | Referência = Este ano |

> ⚠️ **Dias sem Produção (SLA)** tem apresentado falha intermitente de
> navegação no portal (timeout esperando a tabela carregar). O pipeline já
> trata isso automaticamente - loga o erro e segue para as próximas bases -
> mas o download dessa base específica pode não completar em algumas
> execuções até o problema ser investigado a fundo. **Importante:** quando
> isso acontece, a falha é técnica (navegação/carregamento da página) -
> **não significa que não há dados disponíveis** para o período.

## Status Atual

- ✅ As 4 bases (Número de Contratos, Dias sem Produção, Meta Financiamento e
  Seguro, Carteira e Parceiros) estão implementadas com fluxo dedicado em
  `looker_automation.py` e `data_processor.py`, validadas contra o portal e
  as planilhas reais.
- ✅ Login único no portal, reaproveitado para todas as bases de uma mesma
  execução (`--all`/`--frequencia`) - falha de navegação em uma base não
  impede as demais de rodar.
- ⚠️ Dias sem Produção (SLA) com falha intermitente de navegação conhecida
  (ver tabela acima) - ainda não corrigida na raiz.
- ❌ `sharepoint_sync.py` não é usado por nenhuma base hoje - não é preciso
  configurar Azure AD/App Registration para rodar o projeto como ele está.

Para instalar e rodar este projeto em outro computador, ver [SETUP.md](SETUP.md).
Para uma explicação mais completa de arquitetura, manutenção e como
adicionar uma base nova, ver [GUIA_TIME_DADOS.md](GUIA_TIME_DADOS.md).

