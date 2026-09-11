"""
Módulo de atualização automática do Gerador de Etiquetas Pimaco.

Como funciona:
1. Cada release no GitHub recebe uma tag de versão (ex: "v1.1.0") e tem o
   executável (.exe) anexado como "asset".
2. Ao abrir, o programa consulta a API do GitHub e compara a tag mais
   recente com a VERSION embutida no próprio script.
3. Se houver versão nova, baixa o novo .exe para uma pasta temporária.
4. Como o Windows não deixa sobrescrever um .exe que está rodando, um
   pequeno script .bat é criado: ele espera o processo atual fechar,
   substitui o arquivo antigo pelo novo e reabre o programa.

Este módulo não depende de bibliotecas externas (só a stdlib), então não
precisa estar no requirements.txt nem ser reempacotado à parte.
"""

import json
import os
import sys
import subprocess
import tempfile
import urllib.request
import urllib.error

# ======================================================
# CONFIGURAÇÃO — AJUSTE ESTES DOIS VALORES PARA O SEU CASO
# ======================================================
GITHUB_REPO = "Delabeneta/geradorEtiquetas"     
NOME_ASSET = "GeradorEtiquetas.exe"      
TIMEOUT = 10


def _versao_para_tupla(versao: str):
    """Converte 'v1.2.10' ou '1.2.10' em (1, 2, 10) para comparar como número, não como texto."""
    versao = versao.strip().lstrip("vV")
    partes = []
    for p in versao.split("."):
        numero = "".join(ch for ch in p if ch.isdigit())
        partes.append(int(numero) if numero else 0)
    return tuple(partes) if partes else (0,)


def verificar_atualizacao(versao_atual: str):
    """
    Consulta a Release mais recente no GitHub.
    Retorna um dict {"versao", "url", "tamanho", "notas"} se houver
    versão mais nova, ou None se já estiver atualizado / erro de rede.
    Nunca levanta exceção — falha de internet não deve travar o programa.
    """
    url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
    req = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json"})

    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            dados = json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, ValueError, OSError):
        return None

    tag = dados.get("tag_name", "")
    if not tag or _versao_para_tupla(tag) <= _versao_para_tupla(versao_atual):
        return None

    asset_url, tamanho = None, 0
    for asset in dados.get("assets", []):
        if asset.get("name") == NOME_ASSET:
            asset_url = asset.get("browser_download_url")
            tamanho = asset.get("size", 0)
            break

    if not asset_url:
        return None  # release existe mas ainda não tem o .exe anexado

    return {"versao": tag, "url": asset_url, "tamanho": tamanho, "notas": dados.get("body", "")}


def baixar_atualizacao(url: str, callback_progresso=None) -> str:
    """
    Baixa o novo executável para um arquivo temporário e devolve o caminho.
    callback_progresso(bytes_baixados, bytes_totais) é chamado a cada bloco lido
    (pode ser chamado de uma thread separada da interface — trate isso do lado
    de quem chama, por ex. agendando com `widget.after(...)`).
    """
    destino = os.path.join(tempfile.gettempdir(), NOME_ASSET + ".novo")

    with urllib.request.urlopen(url, timeout=TIMEOUT) as resp:
        total = int(resp.headers.get("Content-Length", 0))
        baixado = 0
        with open(destino, "wb") as f:
            while True:
                bloco = resp.read(65536)
                if not bloco:
                    break
                f.write(bloco)
                baixado += len(bloco)
                if callback_progresso:
                    callback_progresso(baixado, total)

    return destino


def aplicar_atualizacao_e_reiniciar(caminho_novo_exe: str):
    """
    Gera um .bat que espera o processo atual encerrar, troca o .exe antigo
    pelo novo e reabre o programa. Depois encerra o processo atual.
    Só funciona quando rodando como .exe empacotado (sys.frozen).
    """
    if not getattr(sys, "frozen", False):
        raise RuntimeError(
            "A troca automática de arquivo só funciona no .exe empacotado. "
            "Rodando como script Python, baixe e instale manualmente."
        )

    exe_atual = sys.executable
    pasta = os.path.dirname(exe_atual)
    nome_exe = os.path.basename(exe_atual)
    bat_path = os.path.join(tempfile.gettempdir(), "atualizar_etiquetas.bat")

    script_bat = f"""@echo off
:esperar
tasklist /FI "IMAGENAME eq {nome_exe}" 2>NUL | find /I "{nome_exe}" >NUL
if "%ERRORLEVEL%"=="0" (
    timeout /t 1 /nobreak >NUL
    goto esperar
)
move /Y "{caminho_novo_exe}" "{exe_atual}" >NUL
start "" "{exe_atual}"
del "%~f0"
"""
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(script_bat)

    subprocess.Popen(
        ["cmd", "/c", bat_path],
        creationflags=subprocess.CREATE_NO_WINDOW,
        cwd=pasta,
    )
    sys.exit(0)