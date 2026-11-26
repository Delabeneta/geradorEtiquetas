import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import sys
import os
import re
import tempfile
import unicodedata
from datetime import datetime

# Obs.: reportlab, openpyxl e updater são importados só quando necessários
# (dentro das funções), para a janela de início aparecer o mais rápido possível.

VERSION = "3.0.0"

# =========================
# TIPOS DE PÚBLICO
# =========================
TIPO_ADULTO = "adulto"
TIPO_INFANTIL = "infantil"
ROTULO_TIPO = {TIPO_ADULTO: "Adulto", TIPO_INFANTIL: "Infantil"}

# Nomes aceitos para a coluna de tipo no Excel (comparados sem acento e em maiúsculas)
NOMES_COLUNA_TIPO = {"TIPO", "PUBLICO", "TIPO DE PUBLICO", "TIPO PUBLICO", "ADULTO/INFANTIL"}

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
combo_capela = None
modo = None
status_label = None
dados_atual = None       # lista de dicts da última geração
frame_manual = None
frame_tipo = None
botao_procurar = None
root = None
tipo_publico = None      # tipo padrão (modo Excel)
tipo_manual = None       # tipo da etiqueta que está sendo adicionada (modo manual)
var_posicao = None
entrada_nome = None
entrada_codigo = None
tree_manual = None

# Lista das etiquetas do modo manual. Cada item é um dict:
# {"POSICAO", "NOME", "CÓDIGO DIZIMISTA", "COMUNIDADE", "TIPO"}
entradas_manuais = []


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
            splash.after(0, atualizar_progresso_ui, 15, "Carregando módulos de PDF...")
            # Pré-carrega os módulos pesados de verdade, enquanto a splash está na tela
            import reportlab  # noqa: F401
            from reportlab.pdfgen import canvas as _c  # noqa: F401
            from reportlab.lib.utils import simpleSplit as _s  # noqa: F401

            splash.after(0, atualizar_progresso_ui, 65, "Carregando módulos de planilha...")
            import openpyxl  # noqa: F401

            splash.after(0, atualizar_progresso_ui, 100, "Sistema pronto!")

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
# TIPO DE PÚBLICO (ADULTO / INFANTIL)
# =========================
def _sem_acentos(texto):
    return "".join(
        ch for ch in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(ch) != "Mn"
    )


def normalizar_tipo(valor):
    """
    Converte o texto de uma célula em 'adulto' ou 'infantil'.
    Retorna None se a célula estiver vazia (aí vale o tipo padrão da interface).
    Levanta ValueError se o valor não for reconhecido.
    Aceita: ADULTO / A  e  INFANTIL / I / INF / CRIANÇA / C.
    """
    v = _sem_acentos(valor).strip().upper()
    if v in ("", "NAN", "NONE"):
        return None
    if v == "A" or v.startswith("ADU"):
        return TIPO_ADULTO
    if v in ("I", "C") or v.startswith("INF") or v.startswith("CRI"):
        return TIPO_INFANTIL
    raise ValueError(valor)


# =========================
# LEITURA DO EXCEL (openpyxl — bem mais leve que o pandas)
# =========================
def _texto(valor):
    return "" if valor is None else str(valor).strip()


def _texto_codigo(valor):
    """Código como texto, sem o '.0' que números inteiros ganham ao virar float."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return re.sub(r"\.0$", "", str(valor).strip())


def ler_excel(caminho, tipo_padrao=TIPO_ADULTO):
    """
    Lê a primeira aba da planilha (.xlsx / .xlsm).
    Colunas obrigatórias: NOME, CÓDIGO e COMUNIDADE (ou CAPELA).
    Coluna opcional: TIPO (ADULTO ou INFANTIL), definida linha a linha.
    Linhas sem tipo (ou planilhas sem a coluna) usam `tipo_padrao`.
    Retorna uma lista de dicts: NOME, CÓDIGO DIZIMISTA, COMUNIDADE, TIPO.
    """
    from openpyxl import load_workbook

    if caminho.lower().endswith(".xls"):
        raise Exception(
            "O formato .xls (Excel antigo) não é suportado.\n"
            "Abra o arquivo no Excel e salve como .xlsx."
        )

    try:
        wb = load_workbook(caminho, read_only=True, data_only=True)
    except Exception as e:
        raise Exception(f"Não foi possível abrir o Excel:\n{e}")

    try:
        linhas = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()

    # Procura o cabeçalho (linha que contém "NOME") nas 20 primeiras linhas
    linha_cabecalho = None
    for i, linha in enumerate(linhas[:20]):
        if any("NOME" in _texto(cel).upper() for cel in linha):
            linha_cabecalho = i
            break

    if linha_cabecalho is None:
        raise Exception("Cabeçalho com 'NOME' não encontrado.")

    cabecalho = [_texto(c).upper() for c in linhas[linha_cabecalho]]

    idx_nome = idx_codigo = idx_comunidade = idx_tipo = None
    for i, c in enumerate(cabecalho):
        sem_acento = _sem_acentos(c)
        if "NOME" in c:
            idx_nome = i
        if "CODIGO" in sem_acento:
            idx_codigo = i
        if "COMUNIDADE" in c or "CAPELA" in c:
            idx_comunidade = i
        if sem_acento.strip() in NOMES_COLUNA_TIPO:
            idx_tipo = i

    if None in (idx_nome, idx_codigo, idx_comunidade):
        raise Exception("Colunas obrigatórias não encontradas.")

    def celula(linha, i):
        return linha[i] if i < len(linha) else None

    registros = []
    invalidos = []
    for j, linha in enumerate(linhas[linha_cabecalho + 1:]):
        tipo = None
        if idx_tipo is not None:
            bruto = celula(linha, idx_tipo)
            try:
                tipo = normalizar_tipo(bruto)
            except ValueError:
                # número da linha como aparece no Excel (1 = primeira linha da planilha)
                invalidos.append(f"Linha {linha_cabecalho + j + 2}: '{bruto}'")

        registros.append({
            "NOME": _texto(celula(linha, idx_nome)).upper(),
            "CÓDIGO DIZIMISTA": _texto_codigo(celula(linha, idx_codigo)),
            "COMUNIDADE": _texto(celula(linha, idx_comunidade)).upper(),
            "TIPO": tipo or tipo_padrao,
        })

    if invalidos:
        raise Exception(
            "Valores inválidos na coluna de tipo (use ADULTO ou INFANTIL):\n"
            + "\n".join(invalidos[:10])
            + ("\n..." if len(invalidos) > 10 else "")
        )

    # Remove linhas totalmente vazias no final da planilha (formatação sobrando)
    while registros and not (registros[-1]["NOME"] or registros[-1]["CÓDIGO DIZIMISTA"]
                             or registros[-1]["COMUNIDADE"]):
        registros.pop()

    if not registros:
        raise Exception("Nenhuma linha de dados encontrada no Excel.")

    return registros


# =========================
# ENTRADA MANUAL
# =========================
def montar_dados_manual(entradas):
    """
    Converte a lista de etiquetas do modo manual em lista de dicts, na ordem das posições.
    Posições puladas viram etiquetas em branco, para que cada nome caia exatamente
    na posição da folha que o usuário pediu. Cada etiqueta tem sua própria
    comunidade e seu próprio tipo (adulto/infantil).
    """
    if not entradas:
        raise Exception("Adicione ao menos uma etiqueta à lista.")

    ordenadas = sorted(entradas, key=lambda e: e["POSICAO"])

    dados = []
    pos_esperada = 1
    for e in ordenadas:
        while pos_esperada < e["POSICAO"]:
            dados.append({"NOME": "", "CÓDIGO DIZIMISTA": "", "COMUNIDADE": "", "TIPO": TIPO_ADULTO})
            pos_esperada += 1

        dados.append({
            "NOME": e["NOME"],
            "CÓDIGO DIZIMISTA": e["CÓDIGO DIZIMISTA"],
            "COMUNIDADE": e["COMUNIDADE"],
            "TIPO": e["TIPO"]
        })
        pos_esperada += 1

    return dados


def proxima_posicao_livre():
    usadas = {e["POSICAO"] for e in entradas_manuais}
    for p in range(1, POSICOES_TOTAIS + 1):
        if p not in usadas:
            return p
    return None


def atualizar_lista_manual():
    """Redesenha a tabela do modo manual a partir de `entradas_manuais`."""
    tree_manual.delete(*tree_manual.get_children())
    for e in sorted(entradas_manuais, key=lambda e: e["POSICAO"]):
        tree_manual.insert(
            "", tk.END, iid=str(e["POSICAO"]),
            values=(
                e["POSICAO"],
                e["NOME"],
                e["CÓDIGO DIZIMISTA"],
                e["COMUNIDADE"] or "—",
                ROTULO_TIPO[e["TIPO"]],
            )
        )


def adicionar_manual():
    try:
        posicao = int(var_posicao.get().strip())
    except ValueError:
        messagebox.showerror("Erro", "A posição deve ser um número.")
        return

    if not (1 <= posicao <= POSICOES_TOTAIS):
        messagebox.showerror("Erro", f"A posição deve estar entre 1 e {POSICOES_TOTAIS}.")
        return

    if any(e["POSICAO"] == posicao for e in entradas_manuais):
        messagebox.showerror(
            "Erro",
            f"A posição {posicao} já está na lista.\n"
            "Remova a etiqueta existente ou escolha outra posição."
        )
        return

    nome = entrada_nome.get().strip().upper()
    if not nome:
        messagebox.showerror("Erro", "Digite o nome.")
        return

    capela = combo_capela.get().strip()
    if capela == "VAZIO":
        capela = ""

    entradas_manuais.append({
        "POSICAO": posicao,
        "NOME": nome,
        "CÓDIGO DIZIMISTA": entrada_codigo.get().strip(),
        "COMUNIDADE": capela,
        "TIPO": tipo_manual.get()
    })
    atualizar_lista_manual()

    # Mantém capela e tipo selecionados (facilita lançar vários da mesma comunidade);
    # limpa só nome/código e avança para a próxima posição livre.
    entrada_nome.delete(0, tk.END)
    entrada_codigo.delete(0, tk.END)
    proxima = proxima_posicao_livre()
    var_posicao.set(str(proxima) if proxima else "")
    entrada_nome.focus_set()


def remover_manual():
    selecao = tree_manual.selection()
    if not selecao:
        messagebox.showerror("Erro", "Selecione uma etiqueta da lista para remover.")
        return

    posicao = int(selecao[0])
    entradas_manuais[:] = [e for e in entradas_manuais if e["POSICAO"] != posicao]
    atualizar_lista_manual()

    proxima = proxima_posicao_livre()
    var_posicao.set(str(proxima) if proxima else "")


def limpar_manual():
    if not entradas_manuais:
        return
    if not messagebox.askyesno("Confirmar", "Remover todas as etiquetas da lista?"):
        return
    entradas_manuais.clear()
    atualizar_lista_manual()
    var_posicao.set("1")


def editar_manual(event):
    """Duplo clique numa linha: devolve a etiqueta ao formulário para ser corrigida."""
    iid = tree_manual.identify_row(event.y)
    if not iid:
        return

    posicao = int(iid)
    entrada = next((e for e in entradas_manuais if e["POSICAO"] == posicao), None)
    if entrada is None:
        return

    entradas_manuais.remove(entrada)
    atualizar_lista_manual()

    var_posicao.set(str(posicao))
    entrada_nome.delete(0, tk.END)
    entrada_nome.insert(0, entrada["NOME"])
    entrada_codigo.delete(0, tk.END)
    entrada_codigo.insert(0, entrada["CÓDIGO DIZIMISTA"])
    combo_capela.set(entrada["COMUNIDADE"] or "VAZIO")
    tipo_manual.set(entrada["TIPO"])
    entrada_nome.focus_set()


# =========================
# GERAÇÃO DO PDF
# =========================
def gerar_pdf(dados):
    """
    Gera o PDF. Cada item de `dados` (lista de dicts) traz seu próprio TIPO
    ('adulto'/'infantil'), então uma mesma folha pode misturar os dois.
    """
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.utils import simpleSplit
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

        # Só exige/registra as fontes ComicNeue se houver alguma etiqueta infantil
        if any(d["TIPO"] == TIPO_INFANTIL for d in dados):
            if hasattr(sys, "_MEIPASS"):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))

            fonte_regular = os.path.join(base_path, "ComicNeue-Regular.ttf")
            fonte_bold = os.path.join(base_path, "ComicNeue-Bold.ttf")

            if not (os.path.exists(fonte_regular) and os.path.exists(fonte_bold)):
                raise Exception(
                    "Há etiquetas do tipo Infantil, mas as fontes ComicNeue não foram "
                    "encontradas junto ao programa.\n"
                    "Reinstale o programa com os arquivos ComicNeue-Regular.ttf e "
                    "ComicNeue-Bold.ttf na mesma pasta, ou marque as etiquetas como Adulto."
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

        for row in dados:
            nome = row["NOME"]
            codigo = row["CÓDIGO DIZIMISTA"]
            comunidade = row["COMUNIDADE"]

            if not nome and not codigo and not comunidade:
                contador += 1
                x += ETIQUETA_LARGURA + ESPACO_H
            else:
                linha_base = y + ETIQUETA_ALTURA - (6 * MM)

                if row["TIPO"] == TIPO_INFANTIL:
                    fonte, fonte_negrito = "ComicNeue", "ComicNeue-Bold"
                else:
                    fonte, fonte_negrito = "Helvetica", "Helvetica-Bold"

                c.setFont(fonte_negrito, 14)
                c.drawCentredString(x + ETIQUETA_LARGURA / 2, linha_base, f"Nº {codigo}")
                c.setFont(fonte, 7.5)
                c.drawCentredString(x + ETIQUETA_LARGURA / 2, linha_base - 10, comunidade)
                c.setFont(fonte_negrito, 10.5)
                linhas = simpleSplit(nome, fonte_negrito, 10.5, ETIQUETA_LARGURA - 15)

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
            dados = ler_excel(caminho, tipo_publico.get())
        else:
            dados = montar_dados_manual(entradas_manuais)

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

        qtd_adulto = sum(1 for d in dados if d["NOME"] and d["TIPO"] == TIPO_ADULTO)
        qtd_infantil = sum(1 for d in dados if d["NOME"] and d["TIPO"] == TIPO_INFANTIL)

        resposta = messagebox.askyesno(
            "Sucesso",
            f"PDF gerado com sucesso!\n\n"
            f"Etiquetas: {qtd_adulto} adulto(s) e {qtd_infantil} infantil(is)\n\n"
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

        # Propaga o novo nome para tudo que já usa a comunidade antiga
        if dados_atual is not None:
            for d in dados_atual:
                if d["COMUNIDADE"] == antigo_nome:
                    d["COMUNIDADE"] = novo_nome

        for e in entradas_manuais:
            if e["COMUNIDADE"] == antigo_nome:
                e["COMUNIDADE"] = novo_nome
        atualizar_lista_manual()

        if combo_capela.get() == antigo_nome:
            combo_capela.set(novo_nome)

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
    arquivo = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
    if arquivo:
        entrada_excel.delete(0, tk.END)
        entrada_excel.insert(0, arquivo)


def atualizar_modo():
    if modo.get() == "excel":
        frame_manual.grid_remove()
        frame_tipo.grid()
        entrada_excel.config(state="normal")
        botao_procurar.config(state="normal")
    else:
        frame_tipo.grid_remove()
        frame_manual.grid()
        entrada_excel.config(state="disabled")
        botao_procurar.config(state="disabled")
        if not combo_capela.get() and CAPELAS:
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
        try:
            import updater
            info = updater.verificar_atualizacao(VERSION)
        except Exception:
            return  # falha na checagem nunca deve incomodar o usuário
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
        _baixar_com_barra_progresso(info)


def _baixar_com_barra_progresso(info):
    import updater

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
            resultado["caminho"] = updater.baixar_atualizacao(
                info["url"],
                callback_progresso=progresso,
                sha256_esperado=info.get("sha256"),
            )
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
            updater.aplicar_atualizacao_e_reiniciar(resultado["caminho"])
        except Exception as e:
            messagebox.showerror("Erro na atualização", str(e))

    threading.Thread(target=tarefa, daemon=True).start()


# =========================
# INTERFACE PRINCIPAL
# =========================
def criar_interface():
    global root, entrada_excel, combo_capela, modo, status_label, frame_manual, \
        frame_tipo, botao_procurar, tipo_publico, tipo_manual, var_posicao, \
        entrada_nome, entrada_codigo, tree_manual

    root = tk.Tk()
    root.title(f"Gerador de Etiquetas Pimaco — v{VERSION}")

    container = tk.Frame(root, padx=15, pady=15)
    container.grid(row=0, column=0, sticky="nsew")

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    # ---------- Título ----------
    frame_titulo = tk.Frame(container)
    frame_titulo.grid(row=0, column=0, sticky="w", pady=(0, 10))

    tk.Label(
        frame_titulo,
        text="Gerador de Etiquetas Pimaco",
        font=("Arial", 14, "bold")
    ).grid(row=0, column=0, sticky="w")

    # ---------- Fonte dos dados ----------
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

    # ---------- Modo Excel: tipo padrão ----------
    frame_tipo = tk.LabelFrame(container, text="Tipo de Público padrão", padx=10, pady=10)
    frame_tipo.grid(row=2, column=0, sticky="ew", pady=5)

    tipo_publico = tk.StringVar(value=TIPO_ADULTO)

    tk.Radiobutton(frame_tipo, text="Adulto", variable=tipo_publico,
                   value=TIPO_ADULTO).grid(row=0, column=0, sticky="w")
    tk.Radiobutton(frame_tipo, text="Infantil", variable=tipo_publico,
                   value=TIPO_INFANTIL).grid(row=0, column=1, sticky="w")

    tk.Label(
        frame_tipo,
        text="Se o Excel tiver uma coluna TIPO (ADULTO ou INFANTIL), ela vale linha a linha.\n"
             "Este tipo padrão só é usado nas linhas sem tipo informado.",
        justify="left",
        fg="gray"
    ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

    # ---------- Modo Manual: formulário + lista ----------
    frame_manual = tk.LabelFrame(container, text="Entrada Manual", padx=10, pady=10)
    frame_manual.grid(row=2, column=0, sticky="ew", pady=5)

    form = tk.Frame(frame_manual)
    form.grid(row=0, column=0, sticky="ew")
    form.columnconfigure(3, weight=1)

    var_posicao = tk.StringVar(value="1")
    tipo_manual = tk.StringVar(value=TIPO_ADULTO)

    tk.Label(form, text="Posição:").grid(row=0, column=0, sticky="w", pady=2)
    tk.Spinbox(form, from_=1, to=POSICOES_TOTAIS, width=5,
               textvariable=var_posicao).grid(row=0, column=1, sticky="w", padx=(5, 15))

    tk.Label(form, text="Código:").grid(row=0, column=2, sticky="w")
    entrada_codigo = tk.Entry(form, width=12)
    entrada_codigo.grid(row=0, column=3, sticky="w", padx=5)

    tk.Label(form, text="Nome:").grid(row=1, column=0, sticky="w", pady=2)
    entrada_nome = tk.Entry(form, width=50)
    entrada_nome.grid(row=1, column=1, columnspan=3, sticky="ew", padx=5)

    tk.Label(form, text="Capela:").grid(row=2, column=0, sticky="w", pady=2)
    combo_capela = ttk.Combobox(form, values=CAPELAS, state="readonly", width=47)
    combo_capela.grid(row=2, column=1, columnspan=3, sticky="ew", padx=5)
    if CAPELAS:
        combo_capela.set(CAPELAS[0])

    tk.Label(form, text="Tipo:").grid(row=3, column=0, sticky="w", pady=2)
    frame_radios = tk.Frame(form)
    frame_radios.grid(row=3, column=1, columnspan=2, sticky="w")
    tk.Radiobutton(frame_radios, text="Adulto", variable=tipo_manual,
                   value=TIPO_ADULTO).grid(row=0, column=0, sticky="w", padx=(5, 0))
    tk.Radiobutton(frame_radios, text="Infantil", variable=tipo_manual,
                   value=TIPO_INFANTIL).grid(row=0, column=1, sticky="w")

    tk.Button(form, text="Adicionar etiqueta", command=adicionar_manual).grid(
        row=3, column=3, sticky="e", padx=5)

    entrada_nome.bind("<Return>", lambda e: adicionar_manual())
    entrada_codigo.bind("<Return>", lambda e: adicionar_manual())

    # Lista das etiquetas já adicionadas
    frame_lista = tk.Frame(frame_manual)
    frame_lista.grid(row=1, column=0, sticky="ew", pady=(10, 0))
    frame_lista.columnconfigure(0, weight=1)

    colunas = ("pos", "nome", "codigo", "capela", "tipo")
    tree_manual = ttk.Treeview(frame_lista, columns=colunas, show="headings",
                               height=8, selectmode="browse")
    for col, titulo, largura in [
        ("pos", "Pos.", 40),
        ("nome", "Nome", 200),
        ("codigo", "Código", 65),
        ("capela", "Capela", 200),
        ("tipo", "Tipo", 70),
    ]:
        tree_manual.heading(col, text=titulo)
        tree_manual.column(col, width=largura, anchor="w")

    scroll = ttk.Scrollbar(frame_lista, orient="vertical", command=tree_manual.yview)
    tree_manual.configure(yscrollcommand=scroll.set)
    tree_manual.grid(row=0, column=0, sticky="ew")
    scroll.grid(row=0, column=1, sticky="ns")
    tree_manual.bind("<Double-1>", editar_manual)

    frame_botoes_lista = tk.Frame(frame_manual)
    frame_botoes_lista.grid(row=2, column=0, sticky="ew", pady=(8, 0))

    tk.Button(frame_botoes_lista, text="Remover selecionada",
              command=remover_manual).grid(row=0, column=0, padx=(0, 8))
    tk.Button(frame_botoes_lista, text="Limpar lista",
              command=limpar_manual).grid(row=0, column=1)
    tk.Label(frame_botoes_lista, text="Dois cliques numa linha para editar.",
             fg="gray").grid(row=0, column=2, padx=12)

    # ---------- Comunidades / ação ----------
    tk.Button(
        container,
        text="Gerenciar Comunidades",
        command=abrir_crud_comunidades
    ).grid(row=3, column=0, pady=5)

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