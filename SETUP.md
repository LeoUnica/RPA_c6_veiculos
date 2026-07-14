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
por nenhuma das 4 bases atuais (todas usam planilha local, ver seção 5).

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
```

**Opção B:** editar diretamente os valores padrão em `config.py` (menos
recomendado, pois não é tão fácil versionar por computador).

Independente da opção escolhida, cada base espera o seguinte formato de
arquivo/pasta dentro do diretório configurado (o código cria a pasta e o
arquivo sozinho se não existirem - não é preciso criar nada manualmente,
apenas apontar para onde os arquivos devem ficar):

| Base | Prévia (sempre sobrescrita) | Planilha de origem oficial |
|---|---|---|
| Número de Contratos | `Número de Contratos - Previa.xlsx` | `<pasta>\Numero de Contratos - {ano}\Digitação Analítico - {ano}.xlsx` (uma subpasta por ano) |
| Dias sem Produção | `Dias sem produção - Previa.xlsx` | `<pasta>\DIAS SEM PRODUCAO.xlsx` (arquivo único, sem separação por ano) |
| Meta Financiamento e Seguro | `Meta Financiamento e Seguro - Previa.xlsx` | `<pasta>\Meta Financiamento Seguro - {ano}.xlsx` (um arquivo por ano, na mesma pasta) |
| Carteira e Parceiros | `Carteira de parceiros e filiais - Previa.xlsx` | `<pasta>\CARTEIRA- {ano}.xlsx` (um arquivo por ano, na mesma pasta) |

Se a planilha de origem oficial de um ano ainda não existir, a primeira
execução cria o arquivo do zero (para Número de Contratos e Dias sem
Produção, com o que for baixado; para Carteira e Parceiros, semeando com o
ano inteiro disponível no Looker no momento).

## 6. Rodando manualmente

```powershell
venv\Scripts\activate

# uma base específica (ids: numero_contratos, dias_sem_producao,
# meta_financiamento_seguro, carteira_parceiros)
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

No Task Scheduler, criar uma tarefa por frequência, apontando para o
Python do `venv` e passando o argumento correspondente, por exemplo:

- Programa/script: `C:\caminho\para\RPA_c6_veiculos\venv\Scripts\python.exe`
- Argumentos: `main.py --frequencia diaria`
- Iniciar em: `C:\caminho\para\RPA_c6_veiculos`

Repetir para `--frequencia semanal_segunda` (só às segundas) e
`--frequencia mensal` (uma vez por mês).

## 8. Problemas já conhecidos / pontos de atenção

- **Arquivo de destino aberto:** ver seção 6 - fechar antes de rodar.
- **Timeout de download:** o download de cada relatório pode demorar
  (relatórios grandes); o timeout interno já está configurado para 60s,
  normalmente não precisa mexer.
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
