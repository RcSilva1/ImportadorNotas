import os
import sys
import time
import urllib.request
import tempfile
import ctypes
import subprocess


def esperar_processo(pid):

    SYNCHRONIZE = 0x00100000

    handle = ctypes.windll.kernel32.OpenProcess(
        SYNCHRONIZE,
        False,
        pid
    )

    if not handle:
        return

    ctypes.windll.kernel32.WaitForSingleObject(
        handle,
        0xFFFFFFFF
    )

    ctypes.windll.kernel32.CloseHandle(handle)


def baixar_arquivo(url, destino):

    requisicao = urllib.request.Request(
        url,
        headers={
            "User-Agent": "ImportadorNotas-Updater"
        }
    )

    with urllib.request.urlopen(
        requisicao,
        timeout=60
    ) as resposta:

        with open(destino, "wb") as arquivo:

            while True:

                bloco = resposta.read(1024 * 1024)

                if not bloco:
                    break

                arquivo.write(bloco)


def atualizar():

    if len(sys.argv) != 4:
        return

    pid = int(sys.argv[1])
    caminho_app = os.path.abspath(sys.argv[2])
    url_download = sys.argv[3]

    pasta = os.path.dirname(caminho_app)

    nome_temp = os.path.basename(caminho_app) + ".new"

    caminho_temp = os.path.join(
        pasta,
        nome_temp
    )

    esperar_processo(pid)

    time.sleep(1)

    try:

        if os.path.exists(caminho_temp):
            os.remove(caminho_temp)

        baixar_arquivo(
            url_download,
            caminho_temp
        )

        if not os.path.exists(caminho_temp):
            raise Exception(
                "O download não foi concluído."
            )

        if os.path.getsize(caminho_temp) < 1024 * 1024:
            raise Exception(
                "O arquivo baixado parece inválido."
            )

        os.replace(
            caminho_temp,
            caminho_app
        )

        subprocess.Popen([
            caminho_app
        ])

    except Exception as erro:

        print(
            f"Erro durante atualização: {erro}"
        )

        try:
            if os.path.exists(caminho_temp):
                os.remove(caminho_temp)
        except:
            pass


if __name__ == "__main__":
    atualizar()