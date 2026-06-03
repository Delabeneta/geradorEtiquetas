# Gerador de Etiquetas para Dizimistas

Sistema desktop desenvolvido para automatizar a geração de etiquetas de identificação de fiéis e dizimistas a partir de planilhas exportadas da plataforma paroquial.

## Visão Geral

Antes da criação deste sistema, as etiquetas eram produzidas manualmente utilizando modelos do Word, exigindo edição individual e consumindo um tempo significativo da equipe responsável pelo cadastro dos fiéis.

O projeto foi desenvolvido para simplificar esse processo. Agora basta exportar a planilha do sistema de cadastro da paróquia, importar o arquivo no programa e gerar automaticamente um PDF pronto para impressão em folhas de etiquetas Pimaco.

O sistema também oferece um modo de entrada manual para situações específicas, permitindo a criação rápida de etiquetas sem depender de planilhas.

## Principais Funcionalidades

* Importação automática de arquivos Excel (.xlsx e .xls)
* Identificação automática do cabeçalho da planilha
* Leitura e tratamento de dados de dizimistas
* Geração de etiquetas compatíveis com folhas Pimaco 3x11
* Quebra automática de nomes longos
* Geração de PDF pronto para impressão
* Entrada manual de registros
* Cadastro e gerenciamento de comunidades/capelas
* Interface gráfica desktop intuitiva
* Tela de carregamento (Splash Screen)
* Salvamento automático do PDF na Área de Trabalho ou Documentos

## Tecnologias Utilizadas

* Python
* Tkinter
* Pandas
* OpenPyXL
* ReportLab

## Aprendizados

* Desenvolvimento de aplicações desktop com Tkinter
* Manipulação e tratamento de dados em Excel com Pandas
* Geração dinâmica de PDFs utilizando ReportLab
* Estruturação de interfaces gráficas em Python
* Validação de entradas e tratamento de erros
* Distribuição de aplicações Python para usuários finais

## Como Executar o Projeto

### 1. Clonar o repositório

```bash
git clone https://github.com/Delabeneta/geradorEtiquetas.git
cd geradorEtiquetas
```

### 2. Criar ambiente virtual

Windows:

```bash
python -m venv venv
venv\Scripts\activate
```

Linux/Mac:

```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Instalar dependências

```bash
pip install -r requirements.txt
```

### 4. Executar aplicação

```bash
python app.py
```

## Gerando Executável (.exe)

Instale o PyInstaller:

```bash
pip install pyinstaller
```

Gere o executável:

```bash
pyinstaller --onefile --windowed app.py
```

O executável será criado na pasta:

```text
dist/
```

Basta distribuir o arquivo `.exe` para utilização em computadores Windows sem necessidade de instalar Python.

## Estrutura do Projeto

```text
.
├── app.py
├── requirements.txt
├── README.md
└── dist/
```

## Autor

Desenvolvido por Rafael Delabeneta para auxiliar o processo de cadastro e organização dos dizimistas em ambiente paroquial.
