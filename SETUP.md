# Guia de Instalação em Outro Computador

Este documento existe para que qualquer pessoa consiga pegar este repositório
e rodar a RPA em um computador novo, do zero, sem precisar reconstruir nada
por tentativa e erro. Cobre pré-requisitos, instalação, configuração de
credenciais/pastas, execução manual, agendamento e os principais problemas
já conhecidos.

## 1. Pré-requisitos

| Item | Versão usada no desenvolvimento | Observação |
|---|---|---|
| Windows | 10/11 | Os caminhos de arquivo (`config.py`) usam sintaxe Windows (`C:\Users\...`) |
| Python | 3.14.6 | Qualquer 3.11+ deve funcionar |
| Git | qualquer recente | Para clonar o repositório |
| Navegador | Chromium (instalado pelo Playwright, não precisa instalar separado) | |
| Excel/LibreOffice | qualquer um | Só para abrir/conferir as planilhas geradas - **ver seção 6 sobre arquivos abertos** |
| Acesso ao portal | usuário e senha do Looker/WebAutorizador (C6 Consig) | Precisa ser um usuário com acesso aos relatórios "Auto" |

## 2. Clonar o repositório

```bash
git clone https://github.com/LeoUnica/RPA_c6_veiculos.git
cd RPA_c6_veiculos
```

## 3. Ambiente virtual e dependências

```powershell
python -m venv venv
venv\Scripts\activate

pip install -r requirements.txt
playwright install chromium
```

`requirements.txt` instala:

```
playwright==1.55.0
pandas==2.3.3
openpyxl==3.1.5
Office365-REST-Python-Client==2.5.9
python-dotenv==1.0.1
holidays==0.75
```

`playwright install chromium` baixa o navegador que o Playwright usa para
automação (obrigatório - sem isso o script falha ao abrir o navegador).

## 4. Configurar credenciais (`.env`)

```powershell
copy .env.example .env
```

Editar o `.env` e preencher pelo menos:

```
LOOKER_URL=https://c6.c6consig.com.br/WebAutorizador/Login/AC.UI.LOGIN.aspx
LOOKER_USER=<usuario_do_portal>
LOOKER_PASSWORD=<senha_do_portal>
```

As variáveis `SHAREPOINT_*` no `.env.example` **não precisam ser
preenchidas** - o `sharepoint_sync.py` existe no repositório mas não é usado
por nenhuma das 5 bases atuais (todas usam planilha local, ver seção 5).

O `.env` nunca é commitado (está no `.gitignore`) - cada computador tem o seu
próprio.

## 5. Ajustar os caminhos das pastas locais (passo mais importante)

Todas as pastas de destino ("Prévia" e planilha de origem oficial de cada
base) estão em `config.py` com um valor padrão fixo, apontando para o
computador onde o projeto foi desenvolvido:

```python
PREVIA_NUMERO_CONTRATOS_DIR = os.getenv(
    "PREVIA_NUMERO_CONTRATOS_DIR",
    r"C:\Users\leonardo.mudrik\Desktop\C6 Bank\Número de Contratos - Previa",
)
```

Em um computador novo esses caminhos **não vão existir**. Existem duas formas
de resolver, escolha uma:

**Opção A - variáveis de ambiente (recomendado, não precisa mexer no código):**
adicionar no `.env` os caminhos corretos para o computador novo:

```
PREVIA_NUMERO_CONTRATOS_DIR=D:\Meu Onedrive\C6 Bank\Número de Contratos - Previa
PLANILHA_ORIGEM_NUMERO_CONTRATOS_DIR=D:\Meu Onedrive\Setor Dados\Ana Price\Número de Contratos

PREVIA_DIAS_SEM_PRODUCAO_DIR=...
PLANILHA_ORIGEM_DIAS_SEM_PRODUCAO_DIR=...

PREVIA_META_FINANCIAMENTO_SEGURO_DIR=...
PLANILHA_ORIGEM_META_FINANCIAMENTO_SEGURO_DIR=...

PREVIA_CARTEIRA_PARCEIROS_DIR=...
PLANILHA_ORIGEM_CARTEIRA_PARCEIROS_DIR=...

PREVIA_COMISSAO_A_VISTA_DIR=...
```

A 5ª base (Comissão à Vista - Analítico) só tem 1 variável, não 2 - ela não
tem planilha de origem oficial separada, a "Prévia" já é o destino final
(acumulado indefinidamente, ver README.md/GUIA_TIME_DADOS.md seção 1.1).

**Opção B:** editar diretamente os valores padrão em `config.py` (menos
recomendado, pois não é tão fácil versionar por computador).

Independente da opção escolhida, cada base espera o seguinte formato de
arquivo/pasta dentro do diretório configurado (o código cria a pasta e o
arquivo sozinho se não existirem - não é preciso criar nada manualmente,
apenas apontar para onde os arquivos devem ficar):

| Base | Prévia | Planilha de origem oficial |
|---|---|---|
| Número de Contratos | `Número de Contratos - Previa.xlsx` (sempre sobrescrita) | `<pasta>\Digitação Analítico - {ano}.xlsx` (um arquivo por ano, na mesma pasta) |
| Dias sem Produção | `Dias sem produção - Previa.xlsx` (sempre sobrescrita) | `<pasta>\DIAS SEM PRODUCAO - {ano}.xlsx` (um arquivo por ano, na mesma pasta) |
| Meta Financiamento e Seguro | `Meta Financiamento e Seguro - Previa.xlsx` (sempre sobrescrita) | `<pasta>\Meta Financiamento Seguro - {ano}.xlsx` (um arquivo por ano, na mesma pasta) |
| Carteira e Parceiros | `Carteira de parceiros e filiais - Previa.xlsx` (sempre sobrescrita) | `<pasta>\CARTEIRA- {ano}.xlsx` (um arquivo por ano, na mesma pasta) |
| Comissão à Vista - Analítico | `Comissão à Vista - Analitico - Previa.xlsx` (**acumulada**, nunca sobrescrita - só recebe linhas novas) | `<pasta>\Comissão à Vista - Analítico - {ano}.xlsx` (um arquivo por ano, na mesma pasta) |

Se a planilha de origem oficial de um ano ainda não existir, a primeira
execução cria o arquivo do zero. O ano corrente recebe atualização a cada
execução; um ano fechado (ex: 2025) nunca é escrito de novo depois de
carregado (ver `data_processor._eh_ano_corrente`).

## 6. Rodando manualmente

```powershell
venv\Scripts\activate

# uma base específica (ids: numero_contratos, dias_sem_producao,
# meta_financiamento_seguro, carteira_parceiros, comissao_a_vista)
python main.py --base numero_contratos

# todas as bases de uma vez
python main.py --all

# todas as bases de uma frequência (usado no agendamento, ver seção 7)
python main.py --frequencia diaria
python main.py --frequencia semanal_segunda
python main.py --frequencia mensal
```

Para depurar visualmente uma base isolada (abre o navegador visível em vez
de headless):

```powershell
python looker_automation.py --base numero_contratos --debug
```

Logs de cada execução ficam em `logs/rpa.log` (e também no console).

**Antes de rodar:** feche no Excel/LibreOffice qualquer planilha de destino
que a base for tocar (Prévia e planilha de origem oficial). O pandas não
consegue sobrescrever um arquivo `.xlsx` aberto em outro programa e a
execução falha com `PermissionError`.

## 7. Agendamento (Windows Task Scheduler)

Cada base tem uma frequência já definida em `config.py` (`"frequencia"`):

| Base | Frequência |
|---|---|
| Número de Contratos | diária |
| Dias sem Produção | semanal (segundas-feiras) |
| Meta Financiamento e Seguro | mensal |
| Carteira e Parceiros | diária |
| Comissão à Vista - Analítico | mensal |

No Task Scheduler, criar uma tarefa por frequência, apontando para o
Python do `venv` e passando o argumento correspondente, por exemplo:

- Programa/script: `C:\caminho\para\RPA_c6_veiculos\venv\Scripts\python.exe`
- Argumentos: `main.py --frequencia diaria`
- Iniciar em: `C:\caminho\para\RPA_c6_veiculos`
- Horário: `06:00`

Repetir para `--frequencia semanal_segunda` (só às segundas, horário
`06:20`) e `--frequencia mensal` (dia 1 de cada mês, horário `06:40`).

**Importante:** cada tarefa precisa de um horário próprio, escalonado por
alguns minutos (não as 3 no mesmo `06:00`). Duas tarefas disparando no
mesmíssimo segundo (ex: diária e semanal numa segunda-feira) fazem dois
processos `main.py` nascerem ao mesmo tempo e logarem simultaneamente na
mesma conta do portal C6 - o portal só permite uma sessão ativa por
usuário e derruba a mais antiga (ver `looker_automation.login`), o que
pode fazer uma das duas execuções falhar no meio da navegação mesmo com a
trava de execução única (`main._trava_execucao_unica`) funcionando
corretamente.

Em **todas as 3 tarefas**, marcar a opção "Executar assim que possível
após uma inicialização agendada perdida" (`StartWhenAvailable`) na aba
Condições - se a máquina estiver desligada/em suspensão no horário
programado, a tarefa dispara assim que ela voltar a ficar disponível, em
vez de simplesmente não rodar naquele dia/semana/mês sem gerar nenhum erro.

## 8. Problemas já conhecidos / pontos de atenção

- **Arquivo de destino aberto:** ver seção 6 - fechar antes de rodar.
- **Timeout de download:** o download de cada relatório pode demorar
  (relatórios grandes); o timeout interno já está ajustado por base em
  `looker_automation.py` (a maioria em 60s, Carteira e Parceiros em 240s
  por baixar o ano inteiro), normalmente não precisa mexer.
- **Login único por execução:** `main.py` loga uma única vez no portal e
  reaproveita essa sessão para todas as bases de um `--all`/`--frequencia`
  - não é preciso configurar nada extra para isso funcionar. Se uma base
  falhar na navegação, ela tenta de novo automaticamente (reabrindo do
  zero) até `looker_automation.MAX_TENTATIVAS_POR_BASE` (5) vezes antes de
  ser considerada pulada, e as demais bases continuam normalmente. A base
  "Dias sem Produção" tem uma falha intermitente conhecida, então esse
  retry é especialmente útil para ela.
- **Sessão já ativa:** se o usuário do Looker já estiver logado em outro
  lugar, o portal mostra um `confirm()` perguntando se quer continuar - o
  código já aceita esse diálogo automaticamente (`page.on("dialog", ...)`),
  não precisa de ação manual.
- **`sharepoint_sync.py`:** existe no repositório mas não é usado por
  nenhuma base hoje (todas usam planilha local) - mantido por decisão do
  usuário, pode ser ignorado ao configurar um computador novo (não precisa
  de Azure AD App Registration nem das variáveis `SHAREPOINT_*`).
- **Biblioteca `holidays`:** usada só pela base "Meta Financiamento e
  Seguro" para calcular dias úteis (feriados de Brasil + Minas Gerais) na
  regra de virada de mês - não precisa de nenhuma configuração adicional,
  já vem com os feriados embutidos na biblioteca.
- **Comissão à Vista - Analítico (5ª base):** duas planilhas locais, cada
  uma só recebendo linhas novas (nunca reescritas) - ver detalhes e os
  bugs reais já corrigidos (linha de totais, nomes de coluna com espaço)
  em GUIA_TIME_DADOS.md seção 1.1.

## 9. Editor / extensões recomendadas (VS Code)

Não é obrigatório usar o VS Code, mas se for usar, estas extensões ajudam a
trabalhar neste projeto:

| Extensão | Para quê |
|---|---|
| Python (Microsoft) | Rodar/depurar os `.py`, selecionar o interpretador do `venv` |
| Pylance (Microsoft) | Autocomplete e checagem de tipos no `config.py`/`data_processor.py` |
| Playwright Test for VSCode (Microsoft) | Rodar `playwright codegen` e o Trace Viewer ao mexer em `looker_automation.py` - útil para reconferir seletores caso o Looker mude o layout |
| Excel Viewer ou similar | Inspecionar rapidamente os `.xlsx` gerados sem abrir o Excel |

Depois de instalar a extensão Python, selecionar o interpretador do projeto
(`Ctrl+Shift+P` → "Python: Select Interpreter" → apontar para
`venv\Scripts\python.exe`).
