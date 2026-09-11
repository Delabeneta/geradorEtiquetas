from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import simpleSplit
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import time
import sys
import os
import tempfile
from datetime import datetime

from updater import verificar_atualizacao, baixar_atualizacao, aplicar_atualizacao_e_reiniciar

VERSION = "1.0.0"

# =========================
# LISTA DE CAPELAS
# =========================
CAPELAS = [
    "CAPELA NOSSA SRA. DE FÁTIMA",
    "CAPELA SAGRADO CORAÇÃO DE JESUS",
    "CAPELA SANT'ANNA",
    "CAPELA SANTA CATARINA",
    "CAPELA SÃO JERÔNIMO",
    "CAPELA SÃO JORGE",
    "CAPELA SÃO JOSÉ",
    "CAPELA S. SEBASTIÃO - MEIO DA SERRA",
    "CAPELA SÃO SEBASTIAO RUA J",
    "MATRIZ",
    "VAZIO"
]

# =========================
# CONFIGURAÇÕES PIMACO 3x11
# =========================

MM = 2.83465  # Conversão

ETIQUETA_LARGURA_CM = 6.35
ETIQUETA_ALTURA_CM = 2.54
COLUNAS = 3
LINHAS = 11
POSICOES_TOTAIS = COLUNAS * LINHAS  # 33 — deriva da grade, em vez de número fixo solto no código

MARGEM_ESQ_CM = 0.8
MARGEM_SUP_CM = 0.9
ESPACO_H_CM = 0.3

ETIQUETA_LARGURA = ETIQUETA_LARGURA_CM * 10 * MM
ETIQUETA_ALTURA = ETIQUETA_ALTURA_CM * 10 * MM
MARGEM_ESQ = MARGEM_ESQ_CM * 10 * MM
MARGEM_SUP = MARGEM_SUP_CM * 10 * MM
ESPACO_H = ESPACO_H_CM * 10 * MM

# Variáveis globais da interface
entrada_excel = None
texto_manual = None
combo_capela = None
modo = None
status_label = None
dados_atual = None
frame_manual = None
botao_procurar = None
root = None
tipo_publico = None


# =========================
# TELA DE INÍCIO (SPLASH SCREEN)
# =========================
def mostrar_tela_inicio():
    """Mostra tela de início enquanto o sistema abre."""

    splash = tk.Tk()
    splash.title("Iniciando...")
    splash.geometry("400x250")
    splash.resizable(False, False)
    splash.overrideredirect(True)

    splash.update_idletasks()
    width = splash.winfo_width()
    height = splash.winfo_height()
    x = (splash.winfo_screenwidth() // 2) - (width // 2)
    y = (splash.winfo_screenheight() // 2) - (height // 2)
    splash.geometry(f"+{x}+{y}")

    tk.Label(
        splash,
        text="Gerador de Etiquetas Pimaco",
        font=("Arial", 16, "bold")
    ).pack(pady=30)

    status_text = tk.StringVar(value="Iniciando sistema...")
    tk.Label(splash, textvariable=status_text, fg="gray").pack(pady=5)

    canvas = tk.Canvas(splash, width=300, height=4, bg='white', highlightthickness=0)
    canvas.pack(pady=10)
    progress_bar = canvas.create_rectangle(0, 0, 0, 4, fill='blue', outline='')

    def atualizar_progresso_ui(percentual, texto):
        """Só mexe na UI — deve ser chamada sempre pela thread principal (via splash.after)."""
        largura = 300 * percentual / 100
        canvas.coords(progress_bar, 0, 0, largura, 4)
        status_text.set(texto)

    def carregar_sistema():
        """
        Roda em thread separada. NUNCA toca em widgets diretamente aqui —
        toda atualização visual é agendada com splash.after(0, ...) para
        rodar na thread principal do Tkinter.
        """
        try:
            splash.after(0, atualizar_progresso_ui, 10, "Preparando ambiente...")
            time.sleep(0.1)

            # Pré-carrega os módulos pesados de verdade (antes eram só textos decorativos)
            splash.after(0, atualizar_progresso_ui, 40, "Carregando módulos de PDF...")
            import reportlab  # noqa: F401
            from reportlab.pdfgen import canvas as _c  # noqa: F401

            splash.after(0, atualizar_progresso_ui, 70, "Carregando módulos de planilha...")
            import pandas  # noqa: F401

            splash.after(0, atualizar_progresso_ui, 100, "Sistema pronto!")
            time.sleep(0.2)

            splash.after(0, lambda: (splash.destroy(), iniciar_sistema_principal()))

        except Exception as e:
            erro = str(e)
            def mostrar_erro():
                splash.destroy()
                messagebox.showerror("Erro", f"Falha ao iniciar sistema:\n{erro}")
                sys.exit(1)
            splash.after(0, mostrar_erro)

    threading.Thread(target=carregar_sistema, daemon=True).start()
    splash.bind('<Escape>', lambda e: sys.exit())
    splash.mainloop()


# =========================
# LEITURA DO EXCEL
# =========================
def ler_excel(caminho):
    import pandas as pd

    df_bruto = pd.read_excel(caminho, header=None)

    linha_cabecalho = None
    for i in range(min(20, len(df_bruto))):
        linha = df_bruto.iloc[i].astype(str).str.upper().tolist()
        if any("NOME" in cel for cel in linha):
            linha_cabecalho = i
            break

    if linha_cabecalho is None:
        raise Exception("Cabeçalho com 'NOME' não encontrado.")

    df = pd.read_excel(caminho, header=linha_cabecalho)
    df.columns = [str(c).strip().upper() for c in df.columns]

    col_nome = col_codigo = col_comunidade = None
    for c in df.columns:
        if "NOME" in c:
            col_nome = c
        if "CÓDIGO" in c:
            col_codigo = c
        if "COMUNIDADE" in c or "CAPELA" in c:
            col_comunidade = c

    if not all([col_nome, col_codigo, col_comunidade]):
        raise Exception("Colunas obrigatórias não encontradas.")

    df[col_nome] = df[col_nome].fillna("").astype(str).str.upper()
    df[col_codigo] = (
        df[col_codigo]
        .fillna("")
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
    )
    df[col_comunidade] = df[col_comunidade].fillna("").astype(str).str.upper()

    return df[[col_nome, col_codigo, col_comunidade]].rename(columns={
        col_nome: "NOME",
        col_codigo: "CÓDIGO DIZIMISTA",
        col_comunidade: "COMUNIDADE"
    })


# =========================
# ENTRADA MANUAL
# =========================
def ler_manual(texto, capela):
    """
    Formato de cada linha: Posicao;NOME;Codigo (Codigo é opcional).
    Corrigido: preenche etiquetas em branco não só antes da primeira posição,
    mas também nos "buracos" entre posições não consecutivas, para que cada
    nome caia exatamente na posição da folha que o usuário pediu.
    """
    import pandas as pd

    linhas = texto.strip().splitlines()
    registros = []

    for numero_linha, linha in enumerate(linhas, start=1):
        partes = [p.strip() for p in linha.split(";")]

        if len(partes) < 2:
            continue  # linha vazia ou incompleta: ignora

        try:
            posicao = int(partes[0].strip())
        except ValueError:
            raise Exception(f"Linha {numero_linha}: Posição deve ser um número")

        if not (1 <= posicao <= POSICOES_TOTAIS):
            raise Exception(f"Linha {numero_linha}: Posição deve ser entre 1 e {POSICOES_TOTAIS}")

        nome = partes[1].strip().upper()
        codigo = partes[2].strip() if len(partes) >= 3 else ""

        registros.append({
            "POSICAO": posicao,
            "NOME": nome,
            "CÓDIGO DIZIMISTA": codigo,
            "COMUNIDADE": capela
        })

    if not registros:
        raise Exception("Nenhuma linha válida encontrada.")

    registros.sort(key=lambda x: x["POSICAO"])

    dados = []
    pos_esperada = 1
    for r in registros:
        # preenche em branco cada posição pulada até chegar na posição do registro
        while pos_esperada < r["POSICAO"]:
            dados.append({"NOME": "", "CÓDIGO DIZIMISTA": "", "COMUNIDADE": ""})
            pos_esperada += 1

        dados.append({
            "NOME": r["NOME"],
            "CÓDIGO DIZIMISTA": r["CÓDIGO DIZIMISTA"],
            "COMUNIDADE": r["COMUNIDADE"]
        })
        pos_esperada += 1

    return pd.DataFrame(dados)


# =========================
# GERAÇÃO DO PDF
# =========================
def gerar_pdf(dados):
    try:
        from reportlab.pdfgen import canvas
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.pdfbase import pdfmetrics

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        documents_path = os.path.join(os.path.expanduser("~"), "Documents")

        if os.access(desktop_path, os.W_OK):
            save_dir = desktop_path
        elif os.access(documents_path, os.W_OK):
            save_dir = documents_path
        else:
            save_dir = tempfile.gettempdir()

        if tipo_publico.get() == "infantil":
            if hasattr(sys, "_MEIPASS"):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))

            fonte_regular = os.path.join(base_path, "ComicNeue-Regular.ttf")
            fonte_bold = os.path.join(base_path, "ComicNeue-Bold.ttf")

            if not (os.path.exists(fonte_regular) and os.path.exists(fonte_bold)):
                raise Exception(
                    "Fontes ComicNeue não encontradas junto ao programa.\n"
                    "Use o modo Adulto ou reinstale o programa com os arquivos "
                    "ComicNeue-Regular.ttf e ComicNeue-Bold.ttf na mesma pasta."
                )

            pdfmetrics.registerFont(TTFont("ComicNeue", fonte_regular))
            pdfmetrics.registerFont(TTFont("ComicNeue-Bold", fonte_bold))

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pdf_filename = f"etiquetas_{timestamp}.pdf"
        pdf_path = os.path.join(save_dir, pdf_filename)

        c = canvas.Canvas(pdf_path, pagesize=A4)
        largura_pagina, altura_pagina = A4

        x0 = MARGEM_ESQ
        y0 = altura_pagina - MARGEM_SUP - ETIQUETA_ALTURA
        x, y = x0, y0
        contador = 0

        for _, row in dados.iterrows():
            nome = row["NOME"]
            codigo = row["CÓDIGO DIZIMISTA"]
            comunidade = row["COMUNIDADE"]

            if not nome and not codigo and not comunidade:
                contador += 1
                x += ETIQUETA_LARGURA + ESPACO_H
            else:
                linha_base = y + ETIQUETA_ALTURA - (6 * MM)

                if tipo_publico.get() == "infantil":
                    c.setFont("ComicNeue-Bold", 14)
                    c.drawCentredString(x + ETIQUETA_LARGURA / 2, linha_base, f"Nº {codigo}")
                    c.setFont("ComicNeue", 7.5)
                    c.drawCentredString(x + ETIQUETA_LARGURA / 2, linha_base - 10, comunidade)
                    c.setFont("ComicNeue-Bold", 10.5)
                    linhas = simpleSplit(nome, "ComicNeue-Bold", 10.5, ETIQUETA_LARGURA - 15)
                else:
                    c.setFont("Helvetica-Bold", 14)
                    c.drawCentredString(x + ETIQUETA_LARGURA / 2, linha_base, f"Nº {codigo}")
                    c.setFont("Helvetica", 7.5)
                    c.drawCentredString(x + ETIQUETA_LARGURA / 2, linha_base - 10, comunidade)
                    c.setFont("Helvetica-Bold", 10.5)
                    linhas = simpleSplit(nome, "Helvetica-Bold", 10.5, ETIQUETA_LARGURA - 15)

                for i, linha in enumerate(linhas[:2]):
                    c.drawCentredString(
                        x + ETIQUETA_LARGURA / 2,
                        linha_base - 23 - (i * 11),
                        linha
                    )

                contador += 1
                x += ETIQUETA_LARGURA + ESPACO_H

            if contador % COLUNAS == 0:
                x = x0
                y -= ETIQUETA_ALTURA

            if contador % POSICOES_TOTAIS == 0:
                c.showPage()
                x, y = x0, y0

        c.save()
        return pdf_path

    except PermissionError:
        raise Exception(
            "Permissão negada para salvar o arquivo.\n"
            "Tente fechar o arquivo PDF anterior ou escolher outra pasta."
        )
    except Exception as e:
        raise Exception(f"Erro ao gerar PDF: {str(e)}")


# =========================
# FUNÇÕES DA INTERFACE
# =========================
def gerar():
    try:
        global dados_atual

        if status_label:
            status_label.config(text="Processando...")
        root.update()

        if modo.get() == "excel":
            caminho = entrada_excel.get().strip()
            if not caminho:
                raise Exception("Selecione um arquivo Excel.")
            dados = ler_excel(caminho)
        else:
            capela = combo_capela.get()
            if not capela or capela == "VAZIO":
                capela = ""
            dados = ler_manual(texto_manual.get("1.0", tk.END), capela)

        dados_atual = dados

        pdf_path = gerar_pdf(dados)
        if status_label:
            status_label.config(text="")

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        if pdf_path.startswith(desktop_path):
            local = "na Área de Trabalho"
        elif "Documents" in pdf_path:
            local = "na pasta Documentos"
        elif "temp" in pdf_path.lower():
            local = "na pasta temporária"
        else:
            local = f"em: {os.path.dirname(pdf_path)}"

        resposta = messagebox.askyesno(
            "Sucesso",
            f"PDF gerado com sucesso!\n\n"
            f"Arquivo salvo {local}:\n{os.path.basename(pdf_path)}\n\n"
            f"Deseja abrir o arquivo?"
        )

        if resposta:
            try:
                if sys.platform == "win32":
                    os.startfile(pdf_path)
                elif sys.platform == "darwin":
                    subprocess_open(["open", pdf_path])
                else:
                    subprocess_open(["xdg-open", pdf_path])
            except Exception:
                messagebox.showinfo(
                    "Abrir Arquivo",
                    f"Não foi possível abrir o arquivo automaticamente.\n"
                    f"Abra manualmente: {pdf_path}"
                )

    except Exception as e:
        if status_label:
            status_label.config(text="")
        messagebox.showerror("Erro", str(e))
        import traceback
        traceback.print_exc()


def subprocess_open(args):
    import subprocess
    subprocess.run(args, check=False)


def abrir_crud_comunidades():
    indice_em_edicao = {"valor": None}

    janela = tk.Toplevel(root)
    janela.title("Gerenciar Comunidades")
    janela.geometry("420x420")
    janela.resizable(False, False)

    tk.Label(janela, text="Comunidades cadastradas:", font=("Arial", 10, "bold")).pack(pady=8)

    lista = tk.Listbox(janela, width=45, height=12)
    lista.pack(pady=5)

    tk.Label(janela, text="Nome da comunidade:", font=("Arial", 9)).pack(pady=(15, 3))
    entrada = tk.Entry(janela, width=45)
    entrada.pack()

    def atualizar_listas():
        lista.delete(0, tk.END)
        for c in CAPELAS:
            lista.insert(tk.END, c)

        combo_capela["values"] = CAPELAS
        if combo_capela.get() not in CAPELAS:
            combo_capela.set("")

    def selecionar_lista(event):
        selecao = lista.curselection()
        if selecao:
            indice = selecao[0]
            indice_em_edicao["valor"] = indice
            entrada.delete(0, tk.END)
            entrada.insert(0, CAPELAS[indice])

    def adicionar():
        nome = entrada.get().strip().upper()
        if not nome:
            messagebox.showerror("Erro", "Digite o nome da comunidade.")
            return
        if nome in CAPELAS:
            messagebox.showerror("Erro", "Comunidade já existe.")
            return

        CAPELAS.append(nome)
        atualizar_listas()
        entrada.delete(0, tk.END)

    def editar():
        indice = indice_em_edicao["valor"]
        if indice is None:
            messagebox.showerror("Erro", "Selecione uma comunidade para editar.")
            return

        novo_nome = entrada.get().strip().upper()
        antigo_nome = CAPELAS[indice]

        if not novo_nome:
            messagebox.showerror("Erro", "Digite o novo nome.")
            return

        if novo_nome in CAPELAS and antigo_nome != novo_nome:
            messagebox.showerror("Erro", "Já existe uma comunidade com esse nome.")
            return

        CAPELAS[indice] = novo_nome

        global dados_atual
        if dados_atual is not None:
            dados_atual.loc[dados_atual["COMUNIDADE"] == antigo_nome, "COMUNIDADE"] = novo_nome

        indice_em_edicao["valor"] = None
        atualizar_listas()
        entrada.delete(0, tk.END)

    def excluir():
        selecao = lista.curselection()
        if not selecao:
            messagebox.showerror("Erro", "Selecione uma comunidade para excluir.")
            return

        indice = selecao[0]
        nome = CAPELAS[indice]

        if not messagebox.askyesno("Confirmar", f"Excluir '{nome}'?"):
            return

        del CAPELAS[indice]
        indice_em_edicao["valor"] = None
        entrada.delete(0, tk.END)
        atualizar_listas()

    lista.bind("<<ListboxSelect>>", selecionar_lista)

    frame_botoes = tk.Frame(janela)
    frame_botoes.pack(pady=25)

    tk.Button(frame_botoes, text="Adicionar", width=14, command=adicionar).grid(row=0, column=0, padx=10, pady=5)
    tk.Button(frame_botoes, text="Confirmar edição", width=14, command=editar).grid(row=0, column=1, padx=10, pady=5)
    tk.Button(frame_botoes, text="Excluir", width=30, command=excluir).grid(row=1, column=0, columnspan=2, pady=10)

    atualizar_listas()


def selecionar_excel():
    arquivo = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
    if arquivo:
        entrada_excel.delete(0, tk.END)
        entrada_excel.insert(0, arquivo)


def atualizar_modo():
    if modo.get() == "excel":
        frame_manual.grid_remove()
        entrada_excel.config(state="normal")
        botao_procurar.config(state="normal")
        combo_capela.config(state="disabled")
        combo_capela.set("")
    else:
        frame_manual.grid()
        entrada_excel.config(state="disabled")
        entrada_excel.delete(0, tk.END)
        botao_procurar.config(state="disabled")
        combo_capela.config(state="readonly")
        if not combo_capela.get():
            combo_capela.set(CAPELAS[0])


# =========================
# ATUALIZAÇÃO AUTOMÁTICA
# =========================
def verificar_e_notificar_atualizacao():
    """
    Chamada logo após a janela principal abrir. Faz a checagem de rede em
    uma thread separada (para não travar a interface) e só volta para a
    thread principal via root.after(...) na hora de mostrar diálogos.
    """
    def tarefa():
        info = verificar_atualizacao(VERSION)
        if info:
            root.after(0, lambda: _perguntar_atualizacao(info))

    threading.Thread(target=tarefa, daemon=True).start()


def _perguntar_atualizacao(info):
    aceitar = messagebox.askyesno(
        "Atualização disponível",
        f"Uma nova versão está disponível: {info['versao']}\n\n"
        f"{info['notas'][:300]}\n\n"
        f"Deseja baixar e instalar agora?"
    )
    if aceitar:
        _baixar_com_barra_progresso(info["url"])


def _baixar_com_barra_progresso(url):
    janela = tk.Toplevel(root)
    janela.title("Baixando atualização...")
    janela.geometry("360x110")
    janela.resizable(False, False)
    janela.grab_set()
    janela.protocol("WM_DELETE_WINDOW", lambda: None)  # não deixa fechar no meio do download

    tk.Label(janela, text="Baixando atualização, aguarde...").pack(pady=(15, 5))
    barra = ttk.Progressbar(janela, length=300, mode="determinate", maximum=100)
    barra.pack(pady=5)
    label_pct = tk.Label(janela, text="0%")
    label_pct.pack()

    resultado = {"caminho": None, "erro": None}

    def progresso(baixado, total):
        # Chamada pela thread de download: nunca mexe direto no widget,
        # só agenda a atualização real na thread principal.
        def atualizar_ui():
            pct = int(baixado * 100 / total) if total else 0
            barra["value"] = pct
            label_pct.config(text=f"{pct}%")
        janela.after(0, atualizar_ui)

    def tarefa():
        try:
            resultado["caminho"] = baixar_atualizacao(url, callback_progresso=progresso)
        except Exception as e:
            resultado["erro"] = str(e)
        janela.after(0, finalizar)

    def finalizar():
        janela.destroy()
        if resultado["erro"]:
            messagebox.showerror("Erro na atualização", resultado["erro"])
            return
        try:
            root.destroy()
            aplicar_atualizacao_e_reiniciar(resultado["caminho"])
        except Exception as e:
            messagebox.showerror("Erro na atualização", str(e))

    threading.Thread(target=tarefa, daemon=True).start()


# =========================
# INTERFACE PRINCIPAL
# =========================
def criar_interface():
    global root, entrada_excel, texto_manual, \
        combo_capela, modo, status_label, frame_manual, botao_procurar, tipo_publico

    root = tk.Tk()
    root.title(f"Gerador de Etiquetas Pimaco — v{VERSION}")

    container = tk.Frame(root, padx=15, pady=15)
    container.grid(row=0, column=0, sticky="nsew")

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    frame_titulo = tk.Frame(container)
    frame_titulo.grid(row=0, column=0, sticky="w", pady=(0, 10))

    tk.Label(
        frame_titulo,
        text="Gerador de Etiquetas Pimaco",
        font=("Arial", 14, "bold")
    ).grid(row=0, column=0, sticky="w")

    frame_fonte = tk.LabelFrame(container, text="Fonte dos dados", padx=10, pady=10)
    frame_fonte.grid(row=1, column=0, sticky="ew", pady=5)

    modo = tk.StringVar(value="excel")

    tk.Radiobutton(frame_fonte, text="Ler Excel",
                   variable=modo, value="excel",
                   command=atualizar_modo).grid(row=0, column=0, sticky="w")

    tk.Radiobutton(frame_fonte, text="Entrada Manual",
                   variable=modo, value="manual",
                   command=atualizar_modo).grid(row=0, column=1, sticky="w")

    entrada_excel = tk.Entry(frame_fonte, width=35)
    entrada_excel.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5)

    botao_procurar = tk.Button(frame_fonte, text="Procurar", command=selecionar_excel)
    botao_procurar.grid(row=1, column=2, padx=5)

    frame_config = tk.LabelFrame(container, text="Configurações", padx=10, pady=10)
    frame_config.grid(row=2, column=0, sticky="ew", pady=5)

    tk.Label(frame_config, text="Capela:").grid(row=0, column=0, sticky="w")

    combo_capela = ttk.Combobox(frame_config, values=CAPELAS, state="readonly", width=40)
    combo_capela.grid(row=0, column=1, pady=3, sticky="w")

    tk.Button(
        frame_config,
        text="Gerenciar Comunidades",
        command=abrir_crud_comunidades
    ).grid(row=1, column=0, columnspan=2, pady=5)

    frame_tipo = tk.LabelFrame(container, text="Tipo de Público", padx=10, pady=10)
    frame_tipo.grid(row=3, column=0, sticky="ew", pady=5)

    tipo_publico = tk.StringVar(value="adulto")

    tk.Radiobutton(frame_tipo, text="Adulto", variable=tipo_publico, value="adulto").grid(row=0, column=0, sticky="w")
    tk.Radiobutton(frame_tipo, text="Infantil", variable=tipo_publico, value="infantil").grid(row=0, column=1, sticky="w")

    frame_manual = tk.LabelFrame(container, text="Entrada Manual - (Posição;NOME;Código)", padx=10, pady=10)
    frame_manual.grid(row=3, column=0, sticky="ew", pady=5)

    tk.Label(
        frame_manual,
        text="Exemplo: 5; JOÃO DA SILVA; 1234",
        justify="left",
        fg="gray"
    ).grid(row=0, column=0, sticky="w", pady=(0, 8))

    texto_manual = scrolledtext.ScrolledText(frame_manual, width=55, height=6)
    texto_manual.grid(row=1, column=0, sticky="ew")

    frame_acao = tk.Frame(container)
    frame_acao.grid(row=4, column=0, pady=15)

    tk.Button(
        frame_acao,
        text="GERAR PDF",
        font=("Arial", 11, "bold"),
        width=25,
        command=gerar
    ).grid(row=0, column=0)

    status_label = tk.Label(container, text="", fg="blue")
    status_label.grid(row=5, column=0, pady=(5, 0))

    atualizar_modo()

    return root


def iniciar_sistema_principal():
    """Inicia o sistema principal após a tela de início."""
    global root
    root = criar_interface()

    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'+{x}+{y}')

    # Checa atualização depois que a janela já está visível
    root.after(800, verificar_e_notificar_atualizacao)

    root.mainloop()


if __name__ == "__main__":
    mostrar_tela_inicio()