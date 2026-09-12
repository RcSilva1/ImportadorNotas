import os
import sys
import json
import urllib.request
import subprocess
import threading
import tkinter as tk
from tkinter import messagebox
from packaging.version import Version

from version import (
    VERSAO,
    GITHUB_OWNER,
    GITHUB_REPO,
    NOME_EXE,
    NOME_ATUALIZADOR
)


URL_RELEASE = (
    f"https://api.github.com/repos/"
    f"{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
)


def obter_release():
    try:
        requisicao = urllib.request.Request(
            URL_RELEASE,
            headers={
                "User-Agent": "ImportadorNotas"
            }
        )

        with urllib.request.urlopen(requisicao, timeout=10) as resposta:
            dados = json.loads(resposta.read().decode("utf-8"))

        return dados

    except Exception as erro:
        print(f"Erro ao verificar atualização: {erro}")
        return None


def verificar_nova_versao():
    dados = obter_release()

    if not dados:
        return None

    tag = dados.get("tag_name", "")

    if not tag:
        return None

    nova_versao = tag.lstrip("vV")

    try:
        if Version(nova_versao) <= Version(VERSAO):
            return None
    except Exception:
        print("Versão inválida recebida do GitHub.")
        return None

    url_download = None

    for asset in dados.get("assets", []):
        if asset.get("name") == NOME_EXE:
            url_download = asset.get("browser_download_url")
            break

    if not url_download:
        print("Executável não encontrado na Release.")
        return None

    return {
        "versao": nova_versao,
        "url": url_download
    }


def iniciar_verificacao(janela):
    def verificar():
        atualizacao = verificar_nova_versao()

        if atualizacao:
            janela.after(
                0,
                lambda: perguntar_atualizacao(
                    janela,
                    atualizacao
                )
            )

    threading.Thread(
        target=verificar,
        daemon=True
    ).start()


def perguntar_atualizacao(janela, atualizacao):

    resposta = messagebox.askyesno(
        "Atualização disponível",
        (
            f"Uma nova versão do Importador de Notas está disponível.\n\n"
            f"Versão atual: {VERSAO}\n"
            f"Nova versão: {atualizacao['versao']}\n\n"
            f"Deseja atualizar agora?"
        ),
        parent=janela
    )

    if not resposta:
        return

    iniciar_atualizador(
        atualizacao["url"]
    )


def iniciar_atualizador(url_download):

    pasta = os.path.dirname(
        os.path.abspath(sys.executable)
    )

    caminho_atualizador = os.path.join(
        pasta,
        NOME_ATUALIZADOR
    )

    if not os.path.exists(caminho_atualizador):
        messagebox.showerror(
            "Erro",
            "O arquivo Atualizador.exe não foi encontrado."
        )
        return

    caminho_app = os.path.abspath(
        sys.executable
    )

    processo_atual = os.getpid()

    try:
        subprocess.Popen([
            caminho_atualizador,
            str(processo_atual),
            caminho_app,
            url_download
        ])

        sys.exit(0)

    except Exception as erro:
        messagebox.showerror(
            "Erro",
            f"Não foi possível iniciar o atualizador.\n\n{erro}"
        )