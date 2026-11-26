"""
Gera o instalador do Gerador de Etiquetas Pimaco.

Uso (num ambiente virtual limpo, só com o necessário):

    python -m venv .venv-build
    .venv-build\\Scripts\\activate
    pip install openpyxl reportlab pyinstaller
    python build.py

Passos:
  1. lê VERSION de gerador_etiquetas.py (única fonte da versão);
  2. gera version_info.txt (informações de versão embutidas no .exe);
  3. roda o PyInstaller em modo --onedir  ->  dist\\GeradorEtiquetas\\
  4. roda o Inno Setup (ISCC)             ->  installer\\GeradorEtiquetas-Setup.exe

Assinatura digital (opcional, recomendada): defina a variável de ambiente
ASSINAR_CMD com o comando de assinatura, usando {arquivo} como marcador. Ex.:
    set ASSINAR_CMD=signtool sign /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 /a "{arquivo}"
O build assina o GeradorEtiquetas.exe antes de montar o instalador e depois assina o instalador.
"""

import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
NOME = "GeradorEtiquetas"
DESCRICAO = "Gerador de Etiquetas Pimaco"
EMPRESA = "Delabeneta"

EXCLUIR_MODULOS = ["pandas", "numpy", "matplotlib", "scipy", "IPython", "pytest"]


def ler_versao() -> str:
    fonte = (RAIZ / "gerador_etiquetas.py").read_text(encoding="utf-8")
    achou = re.search(r'^VERSION\s*=\s*"(\d+(?:\.\d+)*)"', fonte, re.M)
    if not achou:
        sys.exit('VERSION = "x.y.z" não encontrada em gerador_etiquetas.py')
    return achou.group(1)


def gerar_version_info(versao: str) -> str:
    partes = [int(p) for p in versao.split(".")][:4]
    partes += [0] * (4 - len(partes))
    tupla = ", ".join(str(p) for p in partes)
    ponto = ".".join(str(p) for p in partes)
    return f"""VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({tupla}),
    prodvers=({tupla}),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '041604B0',
        [StringStruct('CompanyName', '{EMPRESA}'),
         StringStruct('FileDescription', '{DESCRICAO}'),
         StringStruct('FileVersion', '{ponto}'),
         StringStruct('InternalName', '{NOME}'),
         StringStruct('OriginalFilename', '{NOME}.exe'),
         StringStruct('ProductName', '{DESCRICAO}'),
         StringStruct('ProductVersion', '{ponto}')])
    ]),
    VarFileInfo([VarStruct('Translation', [1046, 1200])])
  ]
)
"""


def assinar(arquivo: Path):
    comando = os.environ.get("ASSINAR_CMD")
    if not comando:
        print(f"  (sem ASSINAR_CMD: {arquivo.name} NÃO foi assinado)")
        return
    print(f"  assinando {arquivo.name} ...")
    subprocess.run(comando.format(arquivo=arquivo), shell=True, check=True)


def achar_iscc():
    """Procura o ISCC.exe do Inno Setup 7 ou 6 (a versão 7 tem prioridade)."""
    # 1) caminho informado manualmente (variável de ambiente ISCC_PATH)
    manual = os.environ.get("ISCC_PATH")
    if manual and Path(manual).exists():
        return manual
    # 2) no PATH
    caminho = shutil.which("ISCC")
    if caminho:
        return caminho
    # 3) pastas de instalação padrão (para todos os usuários ou só para o usuário atual)
    for versao in ("Inno Setup 7", "Inno Setup 6"):
        for var in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
            base = os.environ.get(var)
            if not base:
                continue
            for sub in (versao, os.path.join("Programs", versao)):
                candidato = Path(base) / sub / "ISCC.exe"
                if candidato.exists():
                    return str(candidato)
    return None


def main():
    versao = ler_versao()
    print(f"== Versão {versao} ==")

    (RAIZ / "version_info.txt").write_text(gerar_version_info(versao), encoding="utf-8")

    cmd = [
        sys.executable, "-m", "PyInstaller", "gerador_etiquetas.py",
        "--onedir", "--windowed", "--clean", "--noconfirm", "--noupx",
        f"--name={NOME}",
        "--version-file=version_info.txt",
        "--add-data=ComicNeue-Regular.ttf;.",
        "--add-data=ComicNeue-Bold.ttf;.",
    ]
    if (RAIZ / "icone.ico").exists():
        cmd.append("--icon=icone.ico")
    else:
        print("  (icone.ico não encontrado: o .exe sairá com o ícone padrão)")
    for modulo in EXCLUIR_MODULOS:
        cmd.append(f"--exclude-module={modulo}")

    print("== PyInstaller ==")
    subprocess.run(cmd, check=True, cwd=RAIZ)

    exe = RAIZ / "dist" / NOME / f"{NOME}.exe"
    assinar(exe)

    iscc = achar_iscc()
    if not iscc:
        sys.exit("Inno Setup (ISCC.exe) não encontrado. Instale o Inno Setup 6 ou 7, "
                 "adicione-o ao PATH ou defina a variável ISCC_PATH com o caminho do ISCC.exe.")

    print("== Inno Setup ==")
    subprocess.run([iscc, f"/DMyAppVersion={versao}", str(RAIZ / f"{NOME}.iss")],
                   check=True, cwd=RAIZ)

    instalador = RAIZ / "installer" / f"{NOME}-Setup.exe"
    assinar(instalador)

    print(f"\nPronto: {instalador}")
    print(f"Próximo passo: criar a release v{versao} no GitHub e anexar esse arquivo.")


if __name__ == "__main__":
    main()