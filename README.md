# RPA - Bases C6 Veículos

Automação do fluxo manual descrito no material da equipe de análise de dados:
baixar 4 relatórios do Looker, tratar e consolidar nas bases originais
salvas no SharePoint/OneDrive.

## O que já está pronto vs. o que falta ajustar

| Módulo | Status |
|---|---|
| `config.py` | ✅ Pronto - define as 4 bases e as regras de cada uma |
| `data_processor.py` | ✅ Pronto - filtra status, **alinha colunas automaticamente com a base original** (substitui o passo manual "comparar via PROCX"), remove mês atual e concatena |
| `looker_automation.py` | ⚠️ Esqueleto - os seletores do Playwright são genéricos e **precisam ser ajustados** olhando o Looker de vocês (ver abaixo) |
| `sharepoint_sync.py` | ⚠️ Esqueleto - precisa do App Registration no Azure AD (ver abaixo) |
| `main.py` | ✅ Pronto - orquestra tudo e já tem logging em arquivo |

### O que mudou nesta rodada

- **Alinhamento automático de colunas**: em vez de manter listas fixas de
  colunas a remover, `data_processor.py` agora compara o arquivo baixado com
  a base original e mantém só as colunas que já existem nela (na mesma
  ordem). Isso é exatamente o que o material da equipe descreve como
  "comparar via PROCX" para a base **Dias sem Produção**, só que automático
  e válido para todas as bases em modo de merge. Se uma coluna esperada não
  vier no arquivo baixado, um aviso é logado em vez de quebrar silenciosamente.
- **Correção de bug no upload ao SharePoint**: `sharepoint_sync.upload_processed_base`
  estava subindo o arquivo com o nome do staging local (`{id}_original.xlsx`)
  em vez do nome real da base (`{nome da base}.xlsx`), o que criaria um
  arquivo novo em vez de sobrescrever o existente. Corrigido.
- **Carteira e Parceiros vira cópia direta**: conforme a instrução "cole o
  arquivo Analítico na pasta e renomeie com o nome do já existente", essa
  base agora pula o pandas por completo e sobe o arquivo baixado como
  bytes crus, sem risco de alterar formatação/valores.

## Passo 1 - Instalar dependências

```bash
pip install -r requirements.txt
playwright install chromium
```

## Passo 2 - Configurar credenciais

Copie `.env.example` para `.env` e preencha com os dados reais. Use
`python-dotenv` ou exporte as variáveis no ambiente antes de rodar.

## Passo 3 - Ajustar os seletores do Looker (a parte mais manual)

Como não tenho acesso ao Looker de vocês, os seletores em
`looker_automation.py` são um ponto de partida. Para ajustar:

1. Rode em modo visível para comparar com o site real:
   ```bash
   python looker_automation.py --base numero_contratos --debug
   ```
2. Abra o DevTools do navegador (F12) na página do Looker e veja o texto/role
   real de cada botão de menu, filtro e botão de download.
3. Substitua os `# TODO` em `looker_automation.py` pelos seletores reais.

Dica: o Playwright tem um "Inspector" que grava suas ações e gera o código
do seletor automaticamente:
```bash
playwright codegen https://c6.dtcamp.com.br/WebAutomator/
```
Isso é a forma mais rápida de acertar os seletores sem tentativa e erro.

> O portal real é o **WebAutomator da C6** (`c6.dtcamp.com.br/WebAutomator`),
> chamado de "Looker" apenas como apelido interno da equipe no material.
> A URL de sessão (`?PSession=...`) é gerada dinamicamente após o login,
> então não deve ser fixada em `LOOKER_URL` - use só a raiz do portal.

## Passo 4 - Configurar acesso ao SharePoint

Você vai precisar de um **App Registration** no Azure AD:

1. Portal Azure → Azure Active Directory → App registrations → New registration.
2. Anote o **Application (client) ID** e o **Directory (tenant) ID**.
3. Em "Certificates & secrets", crie um Client Secret e anote o valor.
4. Em "API permissions", adicione permissão de aplicativo
   `Sites.ReadWrite.All` (SharePoint) e conceda consentimento do admin.
5. Preencha essas informações no `.env`.

## Passo 5 - Confirmar nomes de colunas

Em `data_processor.py`, a constante `DATE_COLUMN_BY_BASE` e o filtro de
`STATUS PROPOSTA` assumem nomes de coluna que precisam ser confirmados
abrindo um dos Excels baixados manualmente hoje.

## Passo 6 - Testar uma base isolada antes de tudo

```bash
python main.py --base numero_contratos
```

Verifique o arquivo gerado em `staging/numero_contratos_original.xlsx`
antes de rodar as outras bases ou de agendar.

## Passo 7 - Agendar

Depois de validado, agende via:
- **Windows Task Scheduler**: rodar `python main.py --frequencia diaria`
  todo dia, e `python main.py --frequencia semanal` (ou `semanal_segunda`)
  toda segunda-feira.
- Ou um servidor Linux com `cron`.

## Observação sobre a base "Dias sem Produção"

O material menciona que a comparação de colunas é feita "via PROCX" contra
a base original. Isso foi modelado aqui como remoção de colunas fixas em
`config.py` (`remover_colunas`) - depois de identificar quais colunas batem
via PROCX manualmente uma vez, é só listar as que sobram/faltam ali.
