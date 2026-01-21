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
MM = 2.83465 # Conversão 

ETIQUETA_LARGURA_CM = 6.35
ETIQUETA_ALTURA_CM = 2.54
COLUNAS = 3
LINHAS = 11

MARGEM_ESQ_CM = 0.8
MARGEM_SUP_CM = 0.9
ESPACO_H_CM = 0.3

ETIQUETA_LARGURA = ETIQUETA_LARGURA_CM * 10 * MM
ETIQUETA_ALTURA = ETIQUETA_ALTURA_CM * 10 * MM
MARGEM_ESQ = MARGEM_ESQ_CM * 10 * MM
MARGEM_SUP = MARGEM_SUP_CM * 10 * MM
ESPACO_H = ESPACO_H_CM * 10 * MM

# Variáveis globais para lazy loading
pandas_loaded = False

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

# =========================
# TELA DE INÍCIO (com SPLASH SCREEN)
# =========================
def mostrar_tela_inicio():
    """Mostra tela de início enquanto carrega o sistema"""
    
    splash = tk.Tk()
    splash.title("Iniciando...")
    splash.geometry("400x250")
    splash.resizable(False, False)
    
    # Remover bordas para carregamento mais rápido
    splash.overrideredirect(True)
    
    # Centralizar
    splash.update_idletasks()
    width = splash.winfo_width()
    height = splash.winfo_height()
    x = (splash.winfo_screenwidth() // 2) - (width // 2)
    y = (splash.winfo_screenheight() // 2) - (height // 2)
    splash.geometry(f"+{x}+{y}")
    
    # Conteúdo da tela de início
    tk.Label(
        splash,
        text="Gerador de Etiquetas Pimaco",
        font=("Arial", 16, "bold")
    ).pack(pady=30)
    
    # Mensagem de status
    status_text = tk.StringVar(value="Iniciando sistema...")
    status_label_splash = tk.Label(
        splash,
        textvariable=status_text,
        fg="gray"
    )
    status_label_splash.pack(pady=5)
    
    # Barra de progresso simples
    canvas = tk.Canvas(splash, width=300, height=4, bg='white', highlightthickness=0)
    canvas.pack(pady=10)
    progress_bar = canvas.create_rectangle(0, 0, 0, 4, fill='blue', outline='')
    
    def atualizar_progresso(percentual, texto):
        """Atualiza a barra de progresso"""
        largura = 300 * percentual / 100
        canvas.coords(progress_bar, 0, 0, largura, 4)
        status_text.set(texto)
        splash.update()
    
    def carregar_sistema():
        """Carrega o sistema em segundo plano"""
        try:
            
            atualizar_progresso(10, "Preparando ambiente...")
            time.sleep(0.1)
            
          
            atualizar_progresso(30, "Carregando módulos PDF...")
          
            atualizar_progresso(60, "Preparando interface...")
            time.sleep(0.2)
            
           
            atualizar_progresso(100, "Sistema pronto!")
            time.sleep(0.3)
            
            # Fechar splash e abrir sistema principal
            splash.after(0, lambda: (splash.destroy(), iniciar_sistema_principal()))
            
        except Exception as e:
            print(f"Erro ao carregar: {e}")
            splash.after(0, lambda: (
                splash.destroy(),
                messagebox.showerror("Erro", f"Falha ao iniciar sistema:\n{e}"),
                sys.exit(1)
            ))
    
    # Iniciar carregamento em thread separada
    threading.Thread(target=carregar_sistema, daemon=True).start()
    
    # Permitir fechar com ESC
    splash.bind('<Escape>', lambda e: sys.exit())
    
    splash.mainloop()

# =========================
# LEITURA DO EXCEL (com lazy loading)
# =========================
def ler_excel(caminho):
    global pandas_loaded
    if not pandas_loaded:
        import pandas as pd
        pandas_loaded = True
    else:
        import pandas as pd
    
    df_bruto = pd.read_excel(caminho, header=None)

    linha_cabecalho = None
    for i in range(min(20, len(df_bruto))):  # Limita a busca
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
    linhas = texto.strip().splitlines()
    registros = []

    for numero_linha, linha in enumerate(linhas, start=1):
        partes = [p.strip() for p in linha.split(";")]
        
        if len(partes) < 2:
            continue
        
        elif len(partes) < 1:
           raise Exception(f"Linha {numero_linha}: formato inválido.")

        try:
            posicao = int(partes[0].strip())
        except ValueError:
            raise Exception(f"Linha {numero_linha}: Posicao deve ser um numero")    

        if not(1 <= posicao <= 33):   
           raise Exception(f"linha {numero_linha}: Posicao deve ser entre 1 e 33")
        
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
    posicao_inicial = registros[0]["POSICAO"]

    dados = []

    for _ in range(posicao_inicial - 1):
        dados.append({
            "NOME": "",
            "CÓDIGO DIZIMISTA": "",
            "COMUNIDADE": ""
        })

    for r in registros:
        dados.append({
            "NOME": r["NOME"],
            "CÓDIGO DIZIMISTA": r["CÓDIGO DIZIMISTA"],
            "COMUNIDADE": r["COMUNIDADE"]
        })

    # Importar pandas apenas quando necessário para otimizar 
    global pandas_loaded
    if not pandas_loaded:
        import pandas as pd
        pandas_loaded = True
    else:
        import pandas as pd
    
    return pd.DataFrame(dados)

# =========================
# GERAÇÃO DO PDF
# =========================
def gerar_pdf(dados):
    try:
        # Importação local para evitar o erro. Usuário tem salvado em locais indevido. 
        from reportlab.pdfgen import canvas
        
        # Definir local de salvamento (Desktop ou pasta do usuário)
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        documents_path = os.path.join(os.path.expanduser("~"), "Documents")
        
        # Verificar qual pasta está acessível
        if os.access(desktop_path, os.W_OK):
            save_dir = desktop_path
        elif os.access(documents_path, os.W_OK):
            save_dir = documents_path
        else:
            # Usar pasta temporária como último recurso
            save_dir = tempfile.gettempdir()
        
        # Gerar nome de arquivo com timestamp para evitar conflitos
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

            if contador % (COLUNAS * LINHAS) == 0:
                c.showPage()
                x, y = x0, y0

        c.save()
        return pdf_path
        
    except PermissionError as e:
        raise Exception(f"Permissão negada para salvar o arquivo.\nTente fechar o arquivo PDF anterior ou escolher outra pasta.")
    except Exception as e:
        raise Exception(f"Erro ao gerar PDF: {str(e)}")

# =========================
# FUNÇÕES DA INTERFACE
# =========================
def gerar():
    try:
        global dados_atual
        
        # Mostrar mensagem de processamento
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
        
        # Mostrar mensagem informando onde o arquivo foi salvo
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
                    os.system(f"open '{pdf_path}'")
                else:
                    os.system(f"xdg-open '{pdf_path}'")
            except:
                messagebox.showinfo("Abrir Arquivo", 
                    f"Não foi possível abrir o arquivo automaticamente.\n"
                    f"Abra manualmente: {pdf_path}")

    except Exception as e:
        if status_label:
            status_label.config(text="")
        messagebox.showerror("Erro", str(e))
        import traceback
        traceback.print_exc()  # Para debug

def abrir_crud_comunidades():
    indice_em_edicao = {"valor": None}

    janela = tk.Toplevel(root)
    janela.title("Gerenciar Comunidades")
    janela.geometry("420x420")
    janela.resizable(False, False)

    # Lista
    tk.Label(janela, text="Comunidades cadastradas:", font=("Arial", 10, "bold")).pack(pady=8)

    lista = tk.Listbox(janela, width=45, height=12)
    lista.pack(pady=5)

    # Entrada
    tk.Label(janela, text="Nome da comunidade:", font=("Arial", 9)).pack(pady=(15, 3))
    entrada = tk.Entry(janela, width=45)
    entrada.pack()

    # ===== Funções =====
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
        
        #atualizar a lista
        CAPELAS[indice] = novo_nome
        
        global dados_atual
        if dados_atual is not None:
                dados_atual.loc[
                    dados_atual["COMUNIDADE"] == antigo_nome,
                    "COMUNIDADE"
                    ] = novo_nome
                
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

    # Bind da lista
    lista.bind("<<ListboxSelect>>", selecionar_lista)

    # Botões
    frame_botoes = tk.Frame(janela)
    frame_botoes.pack(pady=25)

    tk.Button(frame_botoes, text="Adicionar", width=14, command=adicionar)\
        .grid(row=0, column=0, padx=10, pady=5)

    tk.Button(frame_botoes, text="Confirmar edição", width=14, command=editar)\
        .grid(row=0, column=1, padx=10, pady=5)

    tk.Button(frame_botoes, text="Excluir", width=30, command=excluir)\
        .grid(row=1, column=0, columnspan=2, pady=10)

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
# INTERFACE PRINCIPAL
# =========================
def criar_interface():
    global root, entrada_excel, texto_manual,\
    combo_capela, modo, status_label, frame_manual, botao_procurar

    root = tk.Tk()
    root.title("Gerador de Etiquetas Pimaco")

    container = tk.Frame(root, padx=15, pady=15)
    container.grid(row=0, column=0, sticky="nsew")

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
  
    #titulo
    frame_titulo = tk.Frame(container)
    frame_titulo.grid(row=0, column=0, sticky="w", pady=(0, 10))

    tk.Label(
        frame_titulo,
        text="Gerador de Etiquetas Pimaco",
        font=("Arial", 14, "bold")
    ).grid(row=0, column=0, sticky="w")

    # excel/manual:

    frame_fonte = tk.LabelFrame(container, text="Fonte dos dados", padx=10, pady=10)
    frame_fonte.grid(row=1, column=0, sticky="ew", pady=5)

    modo = tk.StringVar(value="excel")

    tk.Radiobutton(frame_fonte, text="Ler Excel", 
                   variable=modo, value="excel",
                   command=atualizar_modo).grid(
                   row=0, column=0, sticky="w")

    tk.Radiobutton(frame_fonte, text="Entrada Manual",
                    variable=modo, value="manual", \
                    command=atualizar_modo)\
        .grid(row=0, column=1, sticky="w")

    entrada_excel = tk.Entry(frame_fonte, width=35)
    entrada_excel.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5)

    botao_procurar = tk.Button(
        frame_fonte, text="Procurar",
        command=selecionar_excel)
    
    botao_procurar.grid(row=1, column=2, padx=5)


    #config

    frame_config = tk.LabelFrame(container, text="Configurações", padx=10, pady=10)
    frame_config.grid(row=2, column=0, sticky="ew", pady=5)

    tk.Label(frame_config, text="Capela:")\
        .grid(row=0, column=0, sticky="w")

    combo_capela = ttk.Combobox(frame_config, values=CAPELAS, state="readonly", width=40)
    combo_capela.grid(row=0, column=1, pady=3, sticky="w")

    tk.Button(
        frame_config,
        text="Gerenciar Comunidades",
        command=abrir_crud_comunidades
    ).grid(row=1, column=0, columnspan=2, pady=5)

    #entrada manual: 
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

    #acao principal:
    frame_acao = tk.Frame(container)
    frame_acao.grid(row=4, column=0, pady=15)

    tk.Button(
        frame_acao,
        text="GERAR PDF",
        font=("Arial", 11, "bold"),
        width=25,
        command=gerar
    ).grid(row=0, column=0)

    # status final:
        
    status_label = tk.Label(container, text="", fg="blue")
    status_label.grid(row=5, column=0, pady=(5, 0))
    
    # Configurar estado inicial
    atualizar_modo()
    
    return root

def iniciar_sistema_principal():
    """Inicia o sistema principal após a tela de início"""
    global root
    root = criar_interface()
    
    # Centralizar a janela principal
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'+{x}+{y}')
    
    # Dica sobre permissões
    print("Dica: O arquivo será salvo na Área de Trabalho ou Documentos para evitar problemas de permissão.")
    
    root.mainloop()

if __name__ == "__main__":
    mostrar_tela_inicio()