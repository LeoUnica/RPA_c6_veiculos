"""
Notificação por e-mail da RPA - Bases C6 Veículos.

Envia UM único e-mail consolidado por execução para o endereço de
monitoramento (`unica.tech@unicapromotora.com.br` por padrão), com:

  - data/hora de início e fim da execução e o modo chamado (`--all`,
    `--base X`, `--frequencia Y`);
  - uma linha por base, no mesmo espírito do
    `logs/historico_execucoes.xlsx`: nome da base, horário de conclusão,
    status, linhas baixadas/novas/totais e - em caso de falha - o motivo;
  - um veredito final dizendo se a execução terminou 100% OK ou se alguma
    etapa falhou.

Casos cobertos:
  - A automação NEM COMEÇOU (config inválida, outra execução em
    andamento, falha de login no portal): o e-mail deixa explícito que a
    automação não pôde ser iniciada e por quê.
  - A automação PAROU no meio de uma base: a base aparece como "Falha"
    com o motivo, e as bases seguintes que não chegaram a rodar aparecem
    como "Não executada".

Dois canais de envio (via `.env`, ver `.env.example`). Se nenhum estiver
configurado/funcional, a notificação é só registrada no log e a RPA segue
normalmente - a notificação nunca derruba o robô.

Canal 1 - Microsoft Graph (PREFERIDO, não depende de "Authenticated SMTP"):
    GRAPH_TENANT_ID      tenant do Azure AD (cai para SHAREPOINT_TENANT_ID)
    GRAPH_CLIENT_ID      App Registration (cai para SHAREPOINT_CLIENT_ID)
    GRAPH_CLIENT_SECRET  secret do app     (cai para SHAREPOINT_CLIENT_SECRET)
    GRAPH_SENDER         caixa remetente   (cai para SMTP_FROM)
  O App Registration precisa da permissão de APLICAÇÃO `Mail.Send` com
  consentimento do admin.

Canal 2 - Outlook Desktop via COM/pywin32 (usa o Outlook instalado e
logado na máquina - sem SMTP AUTH, sem App Registration). Ligado por
padrão; MAIL_OUTLOOK=false desliga.

Canal 3 - SMTP autenticado (fallback, usado se GRAPH_* vazio e o Outlook
não enviar):
    SMTP_HOST      (padrão: smtp.office365.com)
    SMTP_PORT      (padrão: 587)
    SMTP_USER / SMTP_PASSWORD   conta + senha/app password
    SMTP_FROM      remetente exibido (padrão: SMTP_USER)
    SMTP_STARTTLS  "true"/"false" (padrão: true)

Destinatário(s) (ambos os canais): MAIL_TO (ou SMTP_TO), separados por
vírgula - padrão: unica.tech@unicapromotora.com.br
"""

from __future__ import annotations

import base64
import json
import logging
import os
import smtplib
import socket
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from dataclasses import dataclass, field
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate

from dotenv import load_dotenv

load_dotenv()

logger = logging.getLogger("notifier")

DESTINATARIO_PADRAO = "unica.tech@unicapromotora.com.br"

# Status possíveis de uma base numa execução.
STATUS_SUCESSO = "sucesso"
STATUS_SEM_DADOS = "sem_dados"
STATUS_FALHA = "falha"
STATUS_NAO_EXECUTADA = "nao_executada"

_ROTULO_STATUS = {
    STATUS_SUCESSO: "Concluída com sucesso",
    STATUS_SEM_DADOS: "Concluída - sem dados no período",
    STATUS_FALHA: "Falha",
    STATUS_NAO_EXECUTADA: "Não executada",
}
_COR_STATUS = {
    STATUS_SUCESSO: "#1a7f37",
    STATUS_SEM_DADOS: "#9a6700",
    STATUS_FALHA: "#cf222e",
    STATUS_NAO_EXECUTADA: "#57606a",
}


def _fmt_dt(dt: datetime | None) -> str:
    return dt.strftime("%d/%m/%Y %H:%M:%S") if dt else "—"


def _fmt_duracao(inicio: datetime | None, fim: datetime | None) -> str:
    if not inicio or not fim:
        return "—"
    total = int((fim - inicio).total_seconds())
    if total < 0:
        return "—"
    h, resto = divmod(total, 3600)
    m, s = divmod(resto, 60)
    if h:
        return f"{h}h{m:02d}min{s:02d}s"
    if m:
        return f"{m}min{s:02d}s"
    return f"{s}s"


def _fmt_num(valor: int | None) -> str:
    if valor is None:
        return "—"
    return f"{valor:,}".replace(",", ".")


def _esc(texto: object) -> str:
    s = "" if texto is None else str(texto)
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


@dataclass
class ResultadoBase:
    """Resultado de uma base numa execução, para a linha do relatório."""

    nome: str
    status: str = STATUS_NAO_EXECUTADA
    concluida_em: datetime | None = None
    linhas_baixadas: int | None = None
    linhas_novas: int | None = None
    linhas_totais: int | None = None
    detalhe: str = ""

    @property
    def rotulo_status(self) -> str:
        return _ROTULO_STATUS.get(self.status, self.status)


@dataclass
class RelatorioExecucao:
    """
    Acumula tudo o que aconteceu numa execução da RPA e sabe se
    transformar num e-mail consolidado (assunto + corpo texto + corpo
    HTML).
    """

    modo: str
    inicio: datetime = field(default_factory=datetime.now)
    fim: datetime | None = None
    maquina: str = field(default_factory=socket.gethostname)
    resultados: list[ResultadoBase] = field(default_factory=list)
    # Preenchido só quando a automação NÃO chega a rodar as bases.
    erro_inicializacao: str | None = None

    # -- construção -------------------------------------------------------
    def base(self, nome: str) -> ResultadoBase:
        """Retorna (criando se preciso) o resultado da base com esse nome."""
        for r in self.resultados:
            if r.nome == nome:
                return r
        r = ResultadoBase(nome=nome)
        self.resultados.append(r)
        return r

    def registrar_bases_previstas(self, nomes: list[str]) -> None:
        for nome in nomes:
            self.base(nome)

    def finalizar(self) -> None:
        if self.fim is None:
            self.fim = datetime.now()

    # -- veredito --------------------------------------------------------
    @property
    def nao_iniciou(self) -> bool:
        return self.erro_inicializacao is not None

    @property
    def houve_falha(self) -> bool:
        return self.nao_iniciou or any(
            r.status in (STATUS_FALHA, STATUS_NAO_EXECUTADA) for r in self.resultados
        )

    @property
    def veredito(self) -> str:
        if self.nao_iniciou:
            return "A automação NÃO pôde ser iniciada."
        falhas = [r for r in self.resultados if r.status == STATUS_FALHA]
        nao_exec = [r for r in self.resultados if r.status == STATUS_NAO_EXECUTADA]
        if not falhas and not nao_exec:
            return "Execução concluída com sucesso — todas as bases foram processadas."
        partes = []
        if falhas:
            partes.append(
                f"{len(falhas)} base(s) com falha: " + ", ".join(r.nome for r in falhas)
            )
        if nao_exec:
            partes.append(
                f"{len(nao_exec)} base(s) não executada(s): "
                + ", ".join(r.nome for r in nao_exec)
            )
        return "Execução concluída com pendências — " + "; ".join(partes) + "."

    # -- renderização ---------------------------------------------------
    def assunto(self) -> str:
        data = self.inicio.strftime("%d/%m/%Y %H:%M")
        if self.nao_iniciou:
            situacao = "NÃO INICIADA"
        elif self.houve_falha:
            situacao = "CONCLUÍDA COM FALHAS"
        else:
            situacao = "CONCLUÍDA COM SUCESSO"
        return f"[RPA C6 Veículos] {situacao} — {data} ({self.maquina})"

    def render_txt(self) -> str:
        linhas = [
            "RPA - Bases C6 Veículos - Relatório de execução",
            "=" * 52,
            f"Máquina .......: {self.maquina}",
            f"Modo ..........: {self.modo}",
            f"Início ........: {_fmt_dt(self.inicio)}",
            f"Fim ...........: {_fmt_dt(self.fim)}",
            f"Duração .......: {_fmt_duracao(self.inicio, self.fim)}",
            "",
            f"VEREDITO: {self.veredito}",
            "",
        ]
        if self.nao_iniciou:
            linhas += [
                "MOTIVO DE NÃO TER INICIADO:",
                self.erro_inicializacao or "",
                "",
            ]
        if self.resultados:
            linhas.append("BASES:")
            for r in self.resultados:
                linhas.append(f"  - {r.nome}")
                linhas.append(f"      Status ......: {r.rotulo_status}")
                linhas.append(f"      Concluída em : {_fmt_dt(r.concluida_em)}")
                linhas.append(
                    f"      Linhas ......: baixadas={_fmt_num(r.linhas_baixadas)} | "
                    f"novas={_fmt_num(r.linhas_novas)} | totais={_fmt_num(r.linhas_totais)}"
                )
                if r.detalhe:
                    linhas.append(f"      Detalhe .....: {r.detalhe}")
            linhas.append("")
        linhas.append(
            "Mensagem automática — não responda. Detalhes técnicos em logs/rpa.log "
            "e logs/historico_execucoes.xlsx na máquina de automação."
        )
        return "\n".join(linhas)

    def render_html(self) -> str:
        cor_veredito = "#cf222e" if self.houve_falha else "#1a7f37"

        cabecalho = f"""
          <table role="presentation" cellpadding="0" cellspacing="0" style="font:14px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#1f2328;">
            <tr><td style="padding:2px 16px 2px 0;color:#57606a;">Máquina</td><td style="padding:2px 0;">{_esc(self.maquina)}</td></tr>
            <tr><td style="padding:2px 16px 2px 0;color:#57606a;">Modo</td><td style="padding:2px 0;">{_esc(self.modo)}</td></tr>
            <tr><td style="padding:2px 16px 2px 0;color:#57606a;">Início</td><td style="padding:2px 0;">{_esc(_fmt_dt(self.inicio))}</td></tr>
            <tr><td style="padding:2px 16px 2px 0;color:#57606a;">Fim</td><td style="padding:2px 0;">{_esc(_fmt_dt(self.fim))}</td></tr>
            <tr><td style="padding:2px 16px 2px 0;color:#57606a;">Duração</td><td style="padding:2px 0;">{_esc(_fmt_duracao(self.inicio, self.fim))}</td></tr>
          </table>
        """

        bloco_erro = ""
        if self.nao_iniciou:
            bloco_erro = f"""
              <div style="margin:18px 0;padding:12px 14px;background:#fff1f0;border:1px solid #ffccc7;border-radius:6px;font:14px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#1f2328;">
                <strong style="color:#cf222e;">A automação não pôde ser iniciada.</strong><br>
                {_esc(self.erro_inicializacao)}
              </div>
            """

        linhas_tabela = ""
        for r in self.resultados:
            cor = _COR_STATUS.get(r.status, "#57606a")
            linhas_tabela += f"""
              <tr>
                <td style="padding:8px 10px;border-bottom:1px solid #eaeef2;">{_esc(r.nome)}</td>
                <td style="padding:8px 10px;border-bottom:1px solid #eaeef2;color:{cor};font-weight:600;">{_esc(r.rotulo_status)}</td>
                <td style="padding:8px 10px;border-bottom:1px solid #eaeef2;white-space:nowrap;">{_esc(_fmt_dt(r.concluida_em))}</td>
                <td style="padding:8px 10px;border-bottom:1px solid #eaeef2;text-align:right;">{_esc(_fmt_num(r.linhas_baixadas))}</td>
                <td style="padding:8px 10px;border-bottom:1px solid #eaeef2;text-align:right;">{_esc(_fmt_num(r.linhas_novas))}</td>
                <td style="padding:8px 10px;border-bottom:1px solid #eaeef2;text-align:right;">{_esc(_fmt_num(r.linhas_totais))}</td>
                <td style="padding:8px 10px;border-bottom:1px solid #eaeef2;color:#57606a;">{_esc(r.detalhe) or "—"}</td>
              </tr>
            """

        bloco_tabela = ""
        if self.resultados:
            bloco_tabela = f"""
              <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;margin-top:8px;font:13px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#1f2328;">
                <thead>
                  <tr style="background:#f6f8fa;text-align:left;">
                    <th style="padding:8px 10px;border-bottom:2px solid #d0d7de;">Base</th>
                    <th style="padding:8px 10px;border-bottom:2px solid #d0d7de;">Status</th>
                    <th style="padding:8px 10px;border-bottom:2px solid #d0d7de;">Concluída em</th>
                    <th style="padding:8px 10px;border-bottom:2px solid #d0d7de;text-align:right;">Baixadas</th>
                    <th style="padding:8px 10px;border-bottom:2px solid #d0d7de;text-align:right;">Novas</th>
                    <th style="padding:8px 10px;border-bottom:2px solid #d0d7de;text-align:right;">Totais</th>
                    <th style="padding:8px 10px;border-bottom:2px solid #d0d7de;">Detalhe</th>
                  </tr>
                </thead>
                <tbody>{linhas_tabela}</tbody>
              </table>
            """

        return f"""\
<!DOCTYPE html>
<html lang="pt-BR">
<body style="margin:0;padding:24px;background:#f6f8fa;">
  <div style="max-width:760px;margin:0 auto;background:#ffffff;border:1px solid #d0d7de;border-radius:8px;padding:24px;">
    <h1 style="margin:0 0 4px;font:600 18px/1.4 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#1f2328;">
      RPA — Bases C6 Veículos
    </h1>
    <p style="margin:0 0 16px;font:13px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#57606a;">
      Relatório automático de execução
    </p>
    {cabecalho}
    <p style="margin:18px 0 0;font:600 15px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:{cor_veredito};">
      {_esc(self.veredito)}
    </p>
    {bloco_erro}
    {bloco_tabela}
    <p style="margin:24px 0 0;font:12px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#8b949e;">
      Mensagem automática — não responda a este e-mail. Detalhes técnicos completos em
      <code>logs/rpa.log</code> e <code>logs/historico_execucoes.xlsx</code> na máquina de automação.
    </p>
  </div>
</body>
</html>
"""


MAX_ANEXO_LOG_BYTES = 2 * 1024 * 1024  # 2 MB - anexa só a "cauda" do rpa.log se for maior

_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@dataclass
class _Anexo:
    filename: str
    mimetype: str
    dados: bytes


def _coletar_destinatarios() -> list[str]:
    """Destinatários da notificação: `MAIL_TO` (ou `SMTP_TO`), separados por vírgula."""
    bruto = os.getenv("MAIL_TO") or os.getenv("SMTP_TO") or DESTINATARIO_PADRAO
    return [e.strip() for e in bruto.split(",") if e.strip()] or [DESTINATARIO_PADRAO]


def _coletar_anexos() -> list[_Anexo]:
    """
    Logs da execução para anexar: `logs/rpa.log` (só os últimos ~2 MB, se
    for grande) e `logs/historico_execucoes.xlsx`. Falha ao ler um anexo
    nunca impede o envio do e-mail.
    """
    log_dir = Path(__file__).parent / "logs"
    anexos: list[_Anexo] = []

    rpa_log = log_dir / "rpa.log"
    try:
        if rpa_log.exists():
            dados = rpa_log.read_bytes()
            truncado = len(dados) > MAX_ANEXO_LOG_BYTES
            if truncado:
                dados = (
                    b"[... inicio do log omitido - anexo limitado aos ultimos 2 MB ...]\n"
                    + dados[-MAX_ANEXO_LOG_BYTES:]
                )
            anexos.append(_Anexo(
                "rpa_ultimos_2MB.log" if truncado else "rpa.log", "text/plain", dados,
            ))
    except Exception:
        logger.warning("Não foi possível ler rpa.log para anexar.", exc_info=True)

    historico = log_dir / "historico_execucoes.xlsx"
    try:
        if historico.exists():
            anexos.append(_Anexo("historico_execucoes.xlsx", _XLSX_MIME, historico.read_bytes()))
    except Exception:
        logger.warning("Não foi possível ler historico_execucoes.xlsx para anexar.", exc_info=True)

    return anexos


# ==========================================================================
# Canal 1 (preferido): Microsoft Graph API - NÃO depende de SMTP AUTH.
# Precisa de um App Registration no Azure AD com a permissão de aplicação
# `Mail.Send` (com consentimento do admin). Envia como a caixa `GRAPH_SENDER`.
# ==========================================================================
@dataclass
class _ConfigGraph:
    tenant_id: str
    client_id: str
    client_secret: str
    remetente: str
    destinatarios: list[str]


def _carregar_config_graph() -> _ConfigGraph | None:
    tenant = (os.getenv("GRAPH_TENANT_ID") or os.getenv("SHAREPOINT_TENANT_ID") or "").strip()
    client = (os.getenv("GRAPH_CLIENT_ID") or os.getenv("SHAREPOINT_CLIENT_ID") or "").strip()
    secret = (os.getenv("GRAPH_CLIENT_SECRET") or os.getenv("SHAREPOINT_CLIENT_SECRET") or "").strip()
    remetente = (os.getenv("GRAPH_SENDER") or os.getenv("SMTP_FROM") or "").strip()
    if not (tenant and client and secret and remetente) or tenant.startswith("x"):
        return None
    return _ConfigGraph(tenant, client, secret, remetente, _coletar_destinatarios())


def _graph_token(cfg: _ConfigGraph) -> str:
    dados = urllib.parse.urlencode({
        "client_id": cfg.client_id,
        "client_secret": cfg.client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }).encode()
    url = f"https://login.microsoftonline.com/{cfg.tenant_id}/oauth2/v2.0/token"
    req = urllib.request.Request(url, data=dados, method="POST")
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read())["access_token"]


def _enviar_via_graph(relatorio: RelatorioExecucao, cfg: _ConfigGraph, anexos: list[_Anexo]) -> bool:
    token = _graph_token(cfg)
    corpo = {
        "message": {
            "subject": relatorio.assunto(),
            "body": {"contentType": "HTML", "content": relatorio.render_html()},
            "toRecipients": [{"emailAddress": {"address": e}} for e in cfg.destinatarios],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": a.filename,
                    "contentType": a.mimetype,
                    "contentBytes": base64.b64encode(a.dados).decode(),
                }
                for a in anexos
            ],
        },
        "saveToSentItems": True,
    }
    url = f"https://graph.microsoft.com/v1.0/users/{urllib.parse.quote(cfg.remetente)}/sendMail"
    req = urllib.request.Request(
        url, data=json.dumps(corpo).encode("utf-8"), method="POST",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=90):  # sendMail responde 202 sem corpo
        pass
    logger.info(
        "Notificação enviada via Microsoft Graph de %s para %s (assunto: %s)",
        cfg.remetente, ", ".join(cfg.destinatarios), relatorio.assunto(),
    )
    return True


# ==========================================================================
# Canal 2 (fallback): SMTP autenticado (smtp.office365.com). Exige que o
# admin habilite "Authenticated SMTP" na caixa remetente.
# ==========================================================================
@dataclass
class _ConfigSMTP:
    host: str
    port: int
    user: str
    password: str
    remetente: str
    destinatarios: list[str]
    starttls: bool


def _carregar_config_smtp() -> _ConfigSMTP | None:
    user = os.getenv("SMTP_USER", "").strip()
    password = os.getenv("SMTP_PASSWORD", "").strip()
    if not user or not password:
        return None
    return _ConfigSMTP(
        host=os.getenv("SMTP_HOST", "smtp.office365.com").strip(),
        port=int(os.getenv("SMTP_PORT", "587")),
        user=user,
        password=password,
        remetente=os.getenv("SMTP_FROM", user).strip() or user,
        destinatarios=_coletar_destinatarios(),
        starttls=os.getenv("SMTP_STARTTLS", "true").strip().lower() != "false",
    )


def _enviar_via_smtp(relatorio: RelatorioExecucao, cfg: _ConfigSMTP, anexos: list[_Anexo]) -> bool:
    msg = EmailMessage()
    msg["Subject"] = relatorio.assunto()
    msg["From"] = formataddr(("RPA C6 Veículos", cfg.remetente))
    msg["To"] = ", ".join(cfg.destinatarios)
    msg["Date"] = formatdate(localtime=True)
    msg.set_content(relatorio.render_txt())
    msg.add_alternative(relatorio.render_html(), subtype="html")
    for a in anexos:
        maintype, _, subtype = a.mimetype.partition("/")
        msg.add_attachment(a.dados, maintype=maintype, subtype=subtype, filename=a.filename)

    with smtplib.SMTP(cfg.host, cfg.port, timeout=60) as smtp:
        smtp.ehlo()
        if cfg.starttls:
            smtp.starttls()
            smtp.ehlo()
        smtp.login(cfg.user, cfg.password)
        smtp.send_message(msg, from_addr=cfg.remetente, to_addrs=cfg.destinatarios)
    logger.info(
        "Notificação enviada via SMTP para %s (assunto: %s)",
        ", ".join(cfg.destinatarios), msg["Subject"],
    )
    return True


# ==========================================================================
# Canal 3: Outlook Desktop (COM/pywin32). Não depende de SMTP AUTH nem de
# App Registration - usa o Outlook já instalado e com perfil configurado na
# máquina de automação (o mesmo usuário que roda a RPA). O e-mail sai pela
# conta padrão do perfil do Outlook. Só funciona no Windows, com Outlook
# clássico (desktop) instalado. Ligado por padrão; desligue com
# MAIL_OUTLOOK=false.
# ==========================================================================
def _outlook_habilitado() -> bool:
    return os.getenv("MAIL_OUTLOOK", "true").strip().lower() != "false"


def _enviar_via_outlook(relatorio: RelatorioExecucao, anexos: list[_Anexo]) -> bool:
    import tempfile

    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    destinatarios = _coletar_destinatarios()
    remetente = (os.getenv("GRAPH_SENDER") or os.getenv("SMTP_FROM") or "").strip()

    pythoncom.CoInitialize()
    tmp_dir = Path(tempfile.mkdtemp(prefix="rpa_notif_"))
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = "; ".join(destinatarios)
        mail.Subject = relatorio.assunto()
        mail.HTMLBody = relatorio.render_html()
        # Enviar pela conta certa, se ela existir no perfil do Outlook.
        if remetente:
            for conta in outlook.Session.Accounts:
                if str(conta.SmtpAddress).lower() == remetente.lower():
                    try:
                        mail.SendUsingAccount = conta
                    except Exception:
                        logger.warning(
                            "Não foi possível fixar a conta remetente '%s' no Outlook - "
                            "o e-mail sairá pela conta padrão do perfil.", remetente,
                        )
                    break
        for a in anexos:
            p = tmp_dir / a.filename
            p.write_bytes(a.dados)
            mail.Attachments.Add(str(p))
        mail.Send()
        logger.info(
            "Notificação enviada via Outlook Desktop para %s (assunto: %s)",
            ", ".join(destinatarios), relatorio.assunto(),
        )
        return True
    finally:
        try:
            for p in tmp_dir.iterdir():
                p.unlink()
            tmp_dir.rmdir()
        except OSError:
            pass
        pythoncom.CoUninitialize()


def enviar_relatorio(relatorio: RelatorioExecucao) -> bool:
    """
    Envia o e-mail consolidado da execução, tentando os canais nesta ordem:
      1. Microsoft Graph  (GRAPH_* - não depende de SMTP AUTH)
      2. Outlook Desktop  (pywin32/COM - usa o Outlook da máquina)
      3. SMTP autenticado (SMTP_* - exige "Authenticated SMTP" habilitado)

    Nunca levanta exceção - qualquer falha é só logada, a notificação não
    pode derrubar a RPA. Retorna True se o e-mail foi efetivamente entregue.
    """
    relatorio.finalizar()
    anexos = _coletar_anexos()

    cfg_graph = _carregar_config_graph()
    if cfg_graph is not None:
        try:
            return _enviar_via_graph(relatorio, cfg_graph, anexos)
        except urllib.error.HTTPError as e:
            detalhe = e.read().decode("utf-8", "replace")[:500]
            logger.error("Falha ao enviar via Microsoft Graph (HTTP %s): %s", e.code, detalhe)
        except Exception:
            logger.exception("Falha ao enviar via Microsoft Graph.")

    if _outlook_habilitado():
        try:
            return _enviar_via_outlook(relatorio, anexos)
        except ImportError:
            logger.warning("Canal Outlook indisponível: pywin32 não instalado (pip install pywin32).")
        except Exception:
            logger.exception("Falha ao enviar via Outlook Desktop.")

    cfg_smtp = _carregar_config_smtp()
    if cfg_smtp is not None:
        try:
            return _enviar_via_smtp(relatorio, cfg_smtp, anexos)
        except Exception:
            logger.exception("Falha ao enviar via SMTP.")

    logger.warning(
        "Notificação por e-mail NÃO enviada: nenhum canal funcional "
        "(Graph via GRAPH_*, Outlook Desktop, ou SMTP via SMTP_*). "
        "Veredito da execução: %s",
        relatorio.veredito,
    )
    return False
