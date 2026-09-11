# 🏷️ Gerador de Etiquetas

Aplicação desktop desenvolvida em **Python** para gerar **etiquetas em PDF** a partir de planilhas Excel, facilitando a organização e impressão de etiquetas para uso administrativo (ex.: dízimo, correspondências, cadastro de fiéis, etc.).

O projeto foi pensado para ser **simples de usar**, mesmo para pessoas sem conhecimento técnico.

---

## ✨ Funcionalidades

- 📄 Leitura de dados a partir de arquivos **Excel (.xlsx)**
- 🏷️ Geração automática de **etiquetas em PDF** (layout padrão Pimaco)
- 🖨️ Arquivo pronto para **impressão**
- 🪟 Interface gráfica amigável (**Tkinter**)
- 🚀 Executável para Windows (gerado com PyInstaller)

---

## 🛠️ Tecnologias utilizadas

- **Python 3.11+**
- **Tkinter** (interface gráfica)
- **Pandas** (leitura e tratamento de dados)
- **OpenPyXL** (manipulação de Excel)
- **ReportLab** (geração de PDF)
- **PyInstaller** (empacotamento para Windows)

---

## 📁 Estrutura do projeto

```
geradorEtiquetas/
├── app.py              # Arquivo principal da aplicação
├── requirements.txt    # Dependências do projeto
├── README.md           # Documentação
├── .gitignore          # Arquivos ignorados pelo Git
```

> ⚠️ Pastas como `dist/`, `build/`, `venv/` e arquivos `.exe` **não são versionados**.

---

## ▶️ Como executar o projeto (modo desenvolvimento)

1. Clone o repositório:

```bash
git clone https://github.com/Delabeneta/geradorEtiquetas.git
cd geradorEtiquetas
```

2. Crie e ative um ambiente virtual:

```bash
python -m venv venv
venv\Scripts\activate
```

3. Instale as dependências:

```bash
pip install -r requirements.txt
```

4. Execute a aplicação:

```bash
python app.py
```

---

## 📦 Gerando o executável (Windows)

Com o ambiente ativado, execute:

```powershell
pyinstaller --onedir --windowed --name="GeradorEtiquetas" --clean --noconfirm --noupx --hidden-import=pandas --hidden-import=openpyxl --hidden-import=reportlab app.py
```

O executável será gerado em:

```
dist/GeradorEtiquetas/
```

> ⚠️ **Importante:**
>
> - Não mova o `.exe` para fora da pasta
> - Utilize um **atalho** para executar

---

## 🔐 Observação sobre antivírus

Executáveis gerados com PyInstaller podem gerar **falso positivo** em antivírus.

Isso ocorre devido ao empacotamento do interpretador Python e **não indica vírus**.

Recomendações:

- Utilize a versão `--onedir`
- Compacte a pasta em `.zip`
- Oriente o usuário a extrair tudo antes de executar

---

## 📌 Possíveis melhorias futuras

- Versão **web** para uso em celular
- Personalização de layouts de etiquetas
- Salvamento de configurações do usuário
- Integração com banco de dados

---

## 👤 Autor

Desenvolvido por **Rafael Santos**

📧 Projeto acadêmico / prático para uso administrativo e aprendizado em Python.

---

## 📄 Licença

Este projeto é de uso livre para fins educacionais e comunitários.
