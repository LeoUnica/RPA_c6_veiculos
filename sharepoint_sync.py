"""
Integração com SharePoint/OneDrive usando a biblioteca Office365-REST-Python-Client.

Requer um App Registration no Azure AD (client id/secret) com permissão
Sites.ReadWrite.All (ou permissão delegada equivalente) no site do SharePoint.
Veja o README.md para o passo a passo de configuração no Azure.
"""

import logging
from pathlib import Path

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

import config

logger = logging.getLogger("sharepoint_sync")


def _get_context() -> ClientContext:
    credentials = ClientCredential(config.SHAREPOINT_CLIENT_ID, config.SHAREPOINT_CLIENT_SECRET)
    ctx = ClientContext(config.SHAREPOINT_SITE_URL).with_credentials(credentials)
    return ctx


def download_original_base(base: dict, local_dest: Path) -> Path:
    """Baixa a base original do SharePoint antes de fazer o merge local."""
    ctx = _get_context()
    remote_folder = f"{config.SHAREPOINT_ROOT_FOLDER}/{base['pasta_sharepoint']}"

    # TODO: confirmar o nome exato do arquivo original em cada pasta
    remote_file = f"{remote_folder}/{base['nome']}.xlsx"

    with open(local_dest, "wb") as f:
        ctx.web.get_file_by_server_relative_url(remote_file).download(f).execute_query()

    logger.info("Base original baixada do SharePoint: %s", remote_file)
    return local_dest


def upload_processed_base(local_path: Path, base: dict):
    """
    Sobe o arquivo final (já tratado e mesclado) de volta para o SharePoint.

    O nome do arquivo no SharePoint é sempre "{nome da base}.xlsx" (o mesmo
    nome usado em download_original_base), independente do nome do arquivo
    local em staging/downloads - isso garante que o arquivo existente seja
    sobrescrito em vez de criar uma cópia nova com outro nome.
    """
    ctx = _get_context()
    remote_folder_url = f"{config.SHAREPOINT_ROOT_FOLDER}/{base['pasta_sharepoint']}"
    remote_file_name = f"{base['nome']}.xlsx"

    target_folder = ctx.web.get_folder_by_server_relative_url(remote_folder_url)

    with open(local_path, "rb") as f:
        content = f.read()

    target_folder.upload_file(remote_file_name, content).execute_query()
    logger.info("Arquivo enviado para o SharePoint: %s/%s", remote_folder_url, remote_file_name)
