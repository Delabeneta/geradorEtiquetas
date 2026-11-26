# Gerador de Etiquetas

Aplicação desktop desenvolvida em Python para automatizar a geração de etiquetas de identificação de fiéis e dizimistas a partir de planilhas Excel.

O sistema foi desenvolvido para simplificar um processo que anteriormente era realizado manualmente utilizando modelos do Word, reduzindo o tempo necessário para preparar e organizar as etiquetas.

A partir da importação de uma planilha exportada do sistema de cadastro da paróquia, o programa realiza o tratamento dos dados e gera automaticamente um arquivo PDF pronto para impressão em folhas de etiquetas Pimaco.

O sistema também permite a entrada manual de registros para situações específicas em que não seja necessário utilizar uma planilha.

---

## Funcionalidades

- Importação de arquivos Excel (`.xlsx` e `.xls`)
- Identificação automática do cabeçalho da planilha
- Leitura e tratamento dos dados dos dizimistas
- Geração de etiquetas compatíveis com folhas Pimaco 3x11
- Quebra automática de nomes longos
- Geração de PDF pronto para impressão
- Entrada manual de registros
- Cadastro e gerenciamento de comunidades e capelas
- Interface gráfica desktop desenvolvida com Tkinter
- Tela de carregamento (Splash Screen)
- Salvamento automático do PDF na Área de Trabalho ou na pasta Documentos

---

## Tecnologias utilizadas

- Python 3.11+
- Tkinter
- Pandas
- OpenPyXL
- ReportLab
- PyInstaller

---

## Estrutura do projeto

```text
geradorEtiquetas/
├── app.py
├── requirements.txt
├── README.md
└── .gitignore
```

Arquivos e pastas gerados durante o desenvolvimento, como `dist/`, `build/`, `venv/` e executáveis `.exe`, não são versionados no Git.

---

## Como executar o projeto

### 1. Clonar o repositório

```bash
git clone https://github.com/Delabeneta/geradorEtiquetas.git
cd geradorEtiquetas
```

### 2. Criar o ambiente virtual

No Windows:

```bash
python -m venv venv
venv\Scripts\activate
```

No Linux ou macOS:

```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Instalar as dependências

```bash
pip install -r requirements.txt
```

### 4. Executar a aplicação

```bash
python app.py
```

---

## Gerando o executável para Windows

Instale o PyInstaller:

```bash
pip install pyinstaller
```

Para gerar o executável em formato de pasta:

```powershell
pyinstaller --onedir --windowed --name="GeradorEtiquetas" --clean --noconfirm --noupx --hidden-import=pandas --hidden-import=openpyxl --hidden-import=reportlab app.py
```

O programa será gerado em:

```text
dist/GeradorEtiquetas/
```

O executável estará dentro dessa pasta:

```text
dist/GeradorEtiquetas/GeradorEtiquetas.exe
```

### Observações

- Não mova o arquivo `.exe` para fora da pasta `GeradorEtiquetas`.
- Para facilitar o uso, recomenda-se criar um atalho para o executável.
- Toda a pasta gerada pelo PyInstaller deve ser mantida ao distribuir o programa.

---

## Antivírus

Executáveis gerados com PyInstaller podem eventualmente ser identificados por alguns antivírus como uma possível ameaça, caracterizando um falso positivo.

Isso pode ocorrer devido à forma como o PyInstaller empacota o interpretador Python e as dependências da aplicação.

Para distribuição, recomenda-se utilizar o formato `--onedir` e compactar a pasta gerada em `.zip`, mantendo todos os arquivos necessários para a execução do programa.

---

## Aprendizados

O desenvolvimento do projeto proporcionou experiência em:

- Desenvolvimento de aplicações desktop com Tkinter
- Manipulação e tratamento de dados em Excel utilizando Pandas
- Geração dinâmica de arquivos PDF com ReportLab
- Estruturação de interfaces gráficas em Python
- Validação de dados e tratamento de erros
- Automatização de processos administrativos
- Distribuição de aplicações Python para usuários finais
- Empacotamento de aplicações utilizando PyInstaller

---

## Possíveis melhorias futuras

- Personalização dos layouts das etiquetas
- Suporte a diferentes modelos de folhas de etiquetas
- Salvamento de configurações do usuário
- Histórico de arquivos gerados
- Melhorias na importação e validação das planilhas
- Versão web para utilização em diferentes dispositivos
- Integração com banco de dados

---

## Autor

Desenvolvido por **Rafael Delabeneta** para auxiliar o processo de cadastro e organização dos dizimistas em ambiente paroquial.

---

## Licença

Este projeto é de uso livre para fins educacionais e comunitários.
