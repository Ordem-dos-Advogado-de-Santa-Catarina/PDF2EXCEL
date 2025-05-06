# 🧾 Extrator de Informações Específicas em Boletos Não Pesquisáveis

Este programa em Python realiza a extração automática de informações específicas (como valor, CNPJ e linha digitável) de boletos bancários em PDF que **não são pesquisáveis**, ou seja, cujo conteúdo é armazenado como imagem. Utiliza OCR para reconhecer o texto e organiza os dados em uma planilha Excel para facilitar o tratamento posterior.

## 📌 Funcionalidades

- Realiza OCR em arquivos PDF de boletos (mesmo sem texto selecionável).
- Extrai as seguintes informações:
  - **Nome do arquivo**
  - **Valor do boleto**
  - **CNPJ do beneficiado**
  - **Linha digitável**
- Gera uma planilha `.xlsx` com os dados extraídos.
- Cria opcionalmente um arquivo `.csv` separado por ponto e vírgula.
- Interface gráfica amigável (Tkinter).
- Permite configurar CNPJs a serem ignorados.
- Identifica guias de custas e boletos com múltiplas páginas.

> Essas pastas podem ser ajustadas conforme necessário na interface gráfica do programa.

## ⚙️ Tecnologias Utilizadas

- `Python 3`
- `pytesseract` — Para OCR (reconhecimento de texto)
- `pdf2image` — Conversão de PDFs para imagens
- `openpyxl` — Geração de planilhas Excel
- `tkinter` — Interface gráfica
- `poppler` — Utilitário necessário para `pdf2image`

## ✅ Requisitos

- Python 3 instalado.
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) instalado:
  - Caminho padrão: `C:\Program Files\Tesseract-OCR\tesseract.exe`
- [Poppler](http://blog.alivate.com.au/poppler-windows/) instalado:
  - Caminho padrão: `C:\Program Files\poppler\bin`
- Instale as bibliotecas Python necessárias com:
