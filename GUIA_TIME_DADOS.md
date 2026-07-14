# Guia do Time de Dados - RPA C6 Veículos

Este documento explica o projeto `RPA_c6_veiculos` para quem vai colocá-lo
para rodar no próprio VS Code e/ou dar manutenção nele. Cobre o que o
projeto faz, como o código é organizado, como configurar e rodar, e como
estender para uma base nova no futuro.

## 1. O que esse projeto faz

Automatiza (via [Playwright](https://playwright.dev/python/), controlando
um navegador Chromium de verdade) o processo manual de atualização de 4
planilhas usadas pelo time de dados de Veículos (C6):

1. Loga no portal C6 Consig e navega até o Looker embutido (Relatórios >
   Relatórios Gerenciais > Auto).
2. Abre o relatório certo para cada base, aplica os filtros necessários e
   baixa a planilha em Excel.
3. Trata os dados baixados (seleciona colunas, filtra linhas, remove
   duplicidade).
4. Atualiza duas planilhas locais por base: uma "Prévia" (o que acabou de
   ser baixado/tratado) e a planilha "de origem" oficial, que acumula o
   histórico completo.
5. Pinta de **verde** as linhas novas e de **amarelo** as linhas
   editadas na Prévia, para facilitar a conferência visual do que mudou
   naquela execução (ver seção 9).

As 4 bases hoje configuradas:

| Base (`id` em `config.py`) | Nome | Frequência | Chave única (deduplicação) |
|---|---|---|---|
| `numero_contratos` | Número de Contratos | Diária | `ID Proposta` |
| `meta_financiamento_seguro` | Meta Financiamento e Seguro | Mensal | `Anomes Apuracao` + `Filial` |
| `dias_sem_producao` | Dias sem Produção | Semanal (segundas) | `Cd Loja` + `Safra Mes` |
| `carteira_parceiros` | Carteira e Parceiros | Diária | `Cnpj Da Loja` + `Filial` + `Anomes` |

Nenhuma das 4 bases usa SharePoint hoje - tudo é lido/gravado em pastas
locais (normalmente sincronizadas por OneDrive). O módulo
`sharepoint_sync.py` existe no repositório mas não é chamado por nenhuma
delas; pode ser ignorado.

## 2. Estrutura do código

```
RPA_c6_veiculos/
├── config.py              # As 4 bases (looker_path, colunas, regras) + caminhos das pastas locais
├── looker_automation.py   # Login, navegação e download no Looker (Playwright)
├── data_processor.py      # Tratamento, merge e marcação de cores (pandas + openpyxl)
├── sharepoint_sync.py     # Upload/download SharePoint - existe, mas não usado hoje
├── main.py                # Orquestrador (CLI: --base / --all / --frequencia)
├── requirements.txt
├── .env.example           # Modelo de variáveis de ambiente
├── downloads/             # Arquivos brutos baixados do Looker (gerada em runtime)
├── staging/                # Não usada pelas 4 bases atuais (reservada para o fluxo SharePoint)
└── logs/rpa.log            # Log de cada execução
```

### `config.py`

Fonte única de verdade de cada base: qual caminho de menu seguir no
Looker (`looker_path`, `link_relatorio`), quais colunas manter
(`colunas_manter`), qual filtro de linha aplicar
(`filtro_status_proposta`), qual `modo` de tratamento usar em
`data_processor.py`, e os caminhos das pastas "Prévia"/"origem" (via
variável de ambiente, com um valor padrão hardcoded como fallback). Para
adicionar uma base nova, normalmente só se mexe aqui e em
`looker_automation.py`/`data_processor.py` (ver seção 11).

### `looker_automation.py`

Um `download_base(base)` genérico que faz login e despacha para uma
função `download_<base>_report` específica de cada base (cada uma abre o
card certo, aplica filtros próprios e baixa a planilha). O Looker é
carregado dentro de popups (`context.expect_page()`), os menus usam hover
antes de clicar, e vários elementos vêm com emoji no texto - ver os
comentários no início de cada função para o motivo de cada `.hover()`,
`.nth(1)`, etc. Se o Looker mudar de layout, é aqui que os seletores
quebram e precisam ser reconferidos (a extensão "Playwright Test for
VSCode", citada na seção 7, ajuda a inspecionar isso).

### `data_processor.py`

Um `process_base(downloaded_path, base)` genérico que despacha, pelo
campo `regras["modo"]` de cada base, para uma função `_process_<base>`
dedicada. Cada uma dessas funções:

1. Lê o Excel baixado, filtra linhas e seleciona colunas conforme
   `config.py`.
2. Salva o resultado na pasta "Prévia" (algumas bases acumulam ao longo
   do mês, outras sobrescrevem a cada execução - ver o docstring de cada
   `_process_*`).
3. Mescla o resultado na planilha de origem oficial, removendo o período
   que está sendo atualizado (mês/ano) antes de colar os dados novos, para
   nunca duplicar.
4. Chama `_marcar_linhas_novas_e_editadas` para colorir a Prévia (seção 9).

### `main.py`

CLI e orquestrador:

```powershell
python main.py --base numero_contratos     # uma base específica
python main.py --all                        # as 4 de uma vez
python main.py --frequencia diaria          # usado no agendamento (Task Scheduler)
```

Para cada base, chama nessa ordem: `looker_automation.download_base` →
`data_processor.process_base`. Se uma base falhar, a exceção sobe e
**interrompe as próximas bases** no mesmo `--all`/`--frequencia` (não há
"pular e continuar" hoje) - o motivo mais comum de falha é a planilha de
destino estar aberta no Excel (ver seção 10).

## 3. Pré-requisitos

| Item | Observação |
|---|---|
| Windows 10/11 | `config.py` usa caminhos no formato Windows |
| Python 3.11+ | Instalar de [python.org](https://www.python.org/downloads/) - **marcar a opção "Add python.exe to PATH"** na instalação, senão os comandos abaixo não funcionam |
| Acesso ao portal C6 Consig | Usuário/senha com acesso aos relatórios "Auto" no Looker |
| VS Code | [code.visualstudio.com](https://code.visualstudio.com/) - ver extensões na seção 7 |

Não precisa de Git nem de conta no GitHub para isso - baixar o ZIP (passo
a seguir) é suficiente.

## 4. Baixar o projeto e abrir no VS Code

1. Na página do repositório no GitHub, clicar no botão verde **"Code"** →
   **"Download ZIP"**.
2. Extrair o ZIP baixado para uma pasta permanente no seu computador (ex:
   `C:\Projetos\RPA_c6_veiculos`) - **não deixar dentro da pasta
   Downloads** nem de uma pasta sincronizada por OneDrive/SharePoint, para
   evitar conflito de sincronização com os arquivos que o Python vai criar
   (`venv`, `downloads`, `logs`).
3. Abrir o VS Code → **File > Open Folder...** → selecionar a pasta
   extraída (a que contém o arquivo `main.py` direto dentro dela, sem uma
   subpasta no meio).
4. Abrir um terminal dentro do próprio VS Code: **Terminal > New
   Terminal** (ou `` Ctrl+` ``). Todos os comandos das próximas seções são
   digitados nesse terminal, com o próprio VS Code já "dentro" da pasta do
   projeto.

## 5. Instalar dependências

No terminal do VS Code (dentro da pasta do projeto):

```powershell
python -m venv venv
venv\Scripts\activate

pip install -r requirements.txt
playwright install chromium
```

`playwright install chromium` baixa o navegador usado pela automação -
sem isso o script falha ao tentar abrir o navegador. Depois de criar o
`venv`, selecionar esse interpretador no VS Code (ver seção 7) para o
autocomplete/execução funcionarem corretamente.

## 6. Configuração (credenciais e caminhos)

### 6.1 Credenciais do portal

```powershell
copy .env.example .env
```

Editar `.env` e preencher:

```
LOOKER_URL=https://c6.c6consig.com.br/WebAutorizador/Login/AC.UI.LOGIN.aspx
LOOKER_USER=02245542630_000367
LOOKER_PASSWORD=Unica@2027
```

O `.env` nunca é commitado (está no `.gitignore`) - cada pessoa/computador
tem o seu. As variáveis `SHAREPOINT_*` do `.env.example` podem ficar em
branco (não são usadas hoje).

### 6.2 Caminhos das pastas locais (passo mais importante - lê com atenção)

`config.py` tem, para cada uma das 4 bases, um caminho de pasta "padrão"
com o endereço das pastas do computador onde este projeto foi
desenvolvido (ex: `C:\Users\leonardo.mudrik\Desktop\C6 Bank\...`). **Esse
caminho não existe no seu computador** - é de outro setor, com outra
estrutura de pastas. Sem ajustar isso, o programa vai tentar criar as
planilhas dentro de uma pasta que não existe e o Python vai dar erro.

Para corrigir, **não precisa editar `config.py`** - basta adicionar 8
linhas no seu `.env` (o mesmo arquivo da seção 6.1), cada uma apontando
para onde **você** quer que aquela planilha fique salva no seu setor.
São 2 pastas por base (a "Prévia" e a "origem oficial"), 4 bases = 8
variáveis:

```
PREVIA_NUMERO_CONTRATOS_DIR=C:\Users\seu.usuario\Desktop\C6 Bank\Número de Contratos - Previa
PLANILHA_ORIGEM_NUMERO_CONTRATOS_DIR=C:\Users\seu.usuario\Desktop\Setor Dados\Ana Price\Número de Contratos

PREVIA_DIAS_SEM_PRODUCAO_DIR=C:\Users\seu.usuario\Desktop\C6 Bank\Dias sem produção - Previa
PLANILHA_ORIGEM_DIAS_SEM_PRODUCAO_DIR=C:\Users\seu.usuario\Desktop\Setor Dados\Ana Price\Dias sem produção

PREVIA_META_FINANCIAMENTO_SEGURO_DIR=C:\Users\seu.usuario\Desktop\C6 Bank\Meta Financiamento e Seguro - Previa
PLANILHA_ORIGEM_META_FINANCIAMENTO_SEGURO_DIR=C:\Users\seu.usuario\Desktop\Setor Dados\Ana Price\Meta Financiamento e Seguro

PREVIA_CARTEIRA_PARCEIROS_DIR=C:\Users\seu.usuario\Desktop\C6 Bank\Carteira de parceiros e filiais - Previa
PLANILHA_ORIGEM_CARTEIRA_PARCEIROS_DIR=C:\Users\seu.usuario\Desktop\Setor Dados\Ana Price\Carteira de parceiros e filiais
```

**Como decidir o caminho certo:** cada `PREVIA_..._DIR` é só uma pasta
qualquer onde você quer que a planilha "Prévia" daquela base seja salva
(pode ser em qualquer lugar do seu computador/OneDrive - crie uma pasta
nova se preferir). Já cada `PLANILHA_ORIGEM_..._DIR` deve apontar para a
pasta onde já fica (ou vai ficar) a planilha oficial daquela base no seu
setor - a que acumula o histórico e é usada pelo time.

Passo a passo para pegar o caminho certo de uma pasta já existente:

1. Abrir a pasta desejada no Explorador de Arquivos do Windows.
2. Clicar uma vez na barra de endereço (a faixa cinza onde aparece o
   "caminho" no topo da janela) - o texto vira editável e mostra o
   caminho completo.
3. Selecionar tudo (`Ctrl+A`) e copiar (`Ctrl+C`).
4. Colar no `.env`, na frente da variável correspondente.

**Importante:** não é preciso criar os arquivos `.xlsx` manualmente - o
código cria sozinho a pasta e a planilha (incluindo a planilha de origem
do ano corrente) na primeira execução, se ainda não existirem. Só a pasta
apontada em `PLANILHA_ORIGEM_..._DIR` precisa já existir de fato (ou você
já ter decidido onde ela vai ficar) - se a planilha oficial daquela base
já existe hoje em outro lugar/formato diferente do esperado (ver tabela
abaixo), o ideal é conversar com quem já mantém aquela planilha antes de
apontar o caminho, para não perder o histórico já acumulado.

| Base | Nome do arquivo "Prévia" | Nome do arquivo "origem" |
| --- | --- | --- |
| Número de Contratos | `Número de Contratos - Previa.xlsx` | `<pasta>\Numero de Contratos - {ano}\Digitação Analítico - {ano}.xlsx` (uma subpasta por ano) |
| Dias sem Produção | `Dias sem produção - Previa.xlsx` | `<pasta>\DIAS SEM PRODUCAO.xlsx` (arquivo único) |
| Meta Financiamento e Seguro | `Meta Financiamento e Seguro - Previa.xlsx` | `<pasta>\Meta Financiamento Seguro - {ano}.xlsx` (um arquivo por ano) |
| Carteira e Parceiros | `Carteira de parceiros e filiais - Previa.xlsx` | `<pasta>\CARTEIRA- {ano}.xlsx` (um arquivo por ano) |

## 7. VS Code - extensões recomendadas

| Extensão | Para quê |
|---|---|
| Python (Microsoft) | Rodar/depurar os `.py`, selecionar o interpretador do `venv` |
| Pylance (Microsoft) | Autocomplete e checagem de tipos |
| Playwright Test for VSCode (Microsoft) | `playwright codegen` e Trace Viewer - útil se o Looker mudar de layout e for preciso reconferir seletores em `looker_automation.py` |
| Excel Viewer (ou similar) | Inspecionar rapidamente os `.xlsx` gerados sem precisar abrir o Excel |

Depois de instalar a extensão Python: `Ctrl+Shift+P` → "Python: Select
Interpreter" → escolher `venv\Scripts\python.exe`.

## 8. Rodando

```powershell
venv\Scripts\activate

python main.py --base numero_contratos      # ids: numero_contratos, dias_sem_producao,
                                              # meta_financiamento_seguro, carteira_parceiros
python main.py --all                         # todas de uma vez
python main.py --frequencia diaria           # todas de uma frequência (uso em agendamento)
```

Para depurar visualmente uma base (abre o navegador em vez de rodar
escondido):

```powershell
python looker_automation.py --base numero_contratos --debug
```

Logs de cada execução ficam em `logs/rpa.log` e também aparecem no
console.

**Antes de rodar:** feche no Excel qualquer planilha (Prévia ou origem)
que a base for tocar - o pandas não consegue sobrescrever um `.xlsx`
aberto em outro programa e a execução para com `PermissionError`.

## 9. Marcação de cores na Prévia

A cada execução, `data_processor._marcar_linhas_novas_e_editadas` compara
a Prévia recém-gerada com a versão anterior dela (usando a chave única de
cada base, ver tabela da seção 1) e pinta:

- 🟩 **Verde**: linha nova (a chave não existia na Prévia anterior).
- 🟨 **Amarelo**: linha já existia, mas algum dado da linha mudou desde a
  última execução.
- Sem cor: linha idêntica à execução anterior.

Dois cuidados já tratados no código, caso precise mexer nessa lógica:
células vazias (`NaN`) são trocadas por um marcador fixo antes de
comparar (`NaN != NaN` faria toda linha com célula vazia parecer
"editada"), e colunas numéricas são arredondadas a 6 casas decimais antes
de comparar (o Excel perde um pouco de precisão de ponto flutuante ao
salvar/reabrir, o que geraria falsos "editada" em valores que na
prática não mudaram).

Essa marcação só existe na Prévia (`Desktop\C6 Bank\...`) - a planilha de
origem oficial não é colorida.

## 10. Problemas conhecidos

- **Arquivo de destino aberto:** ver seção 8 - fechar antes de rodar.
- **Timeout de download:** relatórios grandes podem demorar; o timeout já
  está ajustado por base em `looker_automation.py` (a maioria em 60s,
  Carteira e Parceiros em 120s por baixar o ano inteiro). Um timeout
  esporádico geralmente é só lentidão do portal - rodar de novo resolve.
- **Sessão já ativa:** se o usuário já estiver logado em outro lugar, o
  portal mostra um `confirm()` perguntando se quer continuar - o código já
  aceita esse diálogo sozinho, não precisa de ação manual.
- **`--all`/`--frequencia` para no meio:** se uma base falhar (ex:
  arquivo aberto), as bases seguintes daquela chamada não rodam. Rode a
  base que faltou individualmente depois de corrigir o problema.

## 11. Como adicionar uma base nova

1. Em `config.py`, adicionar um novo dicionário em `BASES` (copiar a
   estrutura de uma base parecida): `id`, `nome`, `looker_path`,
   `link_relatorio`, `pasta_sharepoint` (não usado, mas mantido por
   padrão), `frequencia`, e `regras` (`modo`, `colunas_manter`,
   `filtro_status_proposta`, `aplicar_autofiltro_excel`). Adicionar também
   as funções `caminho_previa_<base>()` / `caminho_planilha_origem_<base>()`
   e as variáveis `PREVIA_<BASE>_DIR` / `PLANILHA_ORIGEM_<BASE>_DIR`.
2. Em `looker_automation.py`, escrever uma função
   `download_<base>_report(context, page, base)` (navegação + filtros +
   download) e registrar o `elif base["id"] == "<base>"` em
   `download_base`.
3. Em `data_processor.py`, escrever uma função `_process_<base>` seguindo
   o padrão das existentes (selecionar colunas, salvar Prévia, mesclar na
   origem, chamar `_marcar_linhas_novas_e_editadas` com a chave única
   certa) e registrar o `if modo == "..."` em `process_base`.
4. Testar isoladamente com
   `python looker_automation.py --base <id> --debug` antes de rodar o
   fluxo completo com `python main.py --base <id>`.


