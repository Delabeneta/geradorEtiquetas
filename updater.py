"""
Módulo de atualização automática do Gerador de Etiquetas Pimaco.

Como funciona:
1. Cada release no GitHub recebe uma tag de versão (ex: "v3.0.1") e tem o
   INSTALADOR (GeradorEtiquetas-Setup.exe, gerado pelo Inno Setup) anexado
   como "asset".
2. Ao abrir, o programa consulta a API do GitHub e compara a tag mais
   recente com a VERSION embutida no próprio script.
3. Se houver versão nova, baixa o instalador para a pasta temporária,
   confere o tamanho e (quando o GitHub informa) o hash SHA-256.
4. O instalador é executado em modo "silencioso" (com barra de progresso),
   fecha o programa, atualiza os arquivos na mesma pasta de instalação e
   reabre o programa. Nenhum .bat é criado.

Este módulo não depende de bibliotecas externas (só a stdlib).
"""

import hashlib
import json
import os
import sys
import subprocess
import tempfile
import urllib.request
import urllib.error

# ======================================================
# CONFIGURAÇÃO — AJUSTE ESTES VALORES PARA O SEU CASO
# ======================================================
GITHUB_REPO = "Delabeneta/geradorEtiquetas"
NOME_ASSET = "GeradorEtiquetas-Setup.exe"   # nome EXATO do instalador anexado à release
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
    Retorna um dict {"versao", "url", "tamanho", "sha256", "notas"} se houver
    versão mais nova, ou None se já estiver atualizado / erro de rede.
    ("sha256" pode ser None se o GitHub não informar o hash do arquivo.)
    Nunca levanta exceção — falha de internet não deve travar o programa.
    """
    url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
    req = urllib.request.Request(url, headers={
        "Accept": "application/vnd.github+json",
        "User-Agent": f"GeradorEtiquetas/{versao_atual}",
    })

    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            dados = json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, ValueError, OSError):
        return None

    tag = dados.get("tag_name", "")
    if not tag or _versao_para_tupla(tag) <= _versao_para_tupla(versao_atual):
        return None

    asset_url, tamanho, sha256 = None, 0, None
    for asset in dados.get("assets", []):
        if asset.get("name") == NOME_ASSET:
            asset_url = asset.get("browser_download_url")
            tamanho = asset.get("size", 0)
            digest = asset.get("digest") or ""
            if digest.lower().startswith("sha256:"):
                sha256 = digest.split(":", 1)[1].strip()
            break

    if not asset_url:
        return None  # release existe mas ainda não tem o instalador anexado

    return {
        "versao": tag,
        "url": asset_url,
        "tamanho": tamanho,
        "sha256": sha256,
        "notas": dados.get("body") or "",
    }


def baixar_atualizacao(url: str, callback_progresso=None, sha256_esperado=None) -> str:
    """
    Baixa o instalador para a pasta temporária e devolve o caminho.
    O arquivo só ganha o nome final depois de baixado e conferido; se algo
    falhar (rede, tamanho, hash), o arquivo parcial é apagado e a exceção sobe.

    callback_progresso(bytes_baixados, bytes_totais) é chamado a cada bloco lido
    (pode ser chamado de uma thread separada da interface — trate isso do lado
    de quem chama, por ex. agendando com `widget.after(...)`).
    """
    destino = os.path.join(tempfile.gettempdir(), NOME_ASSET)
    parcial = destino + ".part"
    hasher = hashlib.sha256()

    try:
        with urllib.request.urlopen(url, timeout=TIMEOUT) as resp:
            total = int(resp.headers.get("Content-Length", 0))
            baixado = 0
            with open(parcial, "wb") as f:
                while True:
                    bloco = resp.read(65536)
                    if not bloco:
                        break
                    f.write(bloco)
                    hasher.update(bloco)
                    baixado += len(bloco)
                    if callback_progresso:
                        callback_progresso(baixado, total)

        if total and baixado != total:
            raise RuntimeError("Download incompleto. Verifique a conexão e tente novamente.")

        if sha256_esperado and hasher.hexdigest().lower() != sha256_esperado.lower():
            raise RuntimeError(
                "O arquivo baixado não confere com o publicado no GitHub "
                "(verificação SHA-256 falhou). Tente novamente mais tarde."
            )

        os.replace(parcial, destino)
    except Exception:
        try:
            os.remove(parcial)
        except OSError:
            pass
        raise

    return destino


def aplicar_atualizacao_e_reiniciar(caminho_instalador: str):
    """
    Executa o instalador baixado (modo silencioso, com barra de progresso) e
    encerra o processo atual. O instalador fecha o programa, atualiza os
    arquivos na pasta onde ele já está instalado e o reabre.
    Só funciona quando rodando como programa empacotado (sys.frozen).
    """
    if not getattr(sys, "frozen", False):
        raise RuntimeError(
            "A atualização automática só funciona no programa instalado. "
            "Rodando como script Python, baixe e instale manualmente."
        )

    subprocess.Popen(
        [caminho_instalador, "/SILENT", "/NOCANCEL", "/CLOSEAPPLICATIONS"],
        creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
        close_fds=True,
    )
    sys.exit(0)