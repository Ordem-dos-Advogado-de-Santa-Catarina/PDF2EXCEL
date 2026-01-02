# Extrator de Boletos & Custas (OCR)

Ferramenta desktop para extração em lote de dados de boletos bancários e guias de custas judiciais em formato PDF (imagem/scanned). Utiliza OCR (Tesseract) para converter o conteúdo visual em dados estruturados, gerando relatórios em Excel e CSV.

O sistema ignora automaticamente CNPJs configurados em *blacklist* e valida integridade de linhas digitáveis.

## Funcionalidades

- **OCR em Lote:** Processamento de múltiplos arquivos PDF simultaneamente.
- **Extração de Dados:**
  - Linha Digitável (Código de Barras)
  - Valor Monetário
  - CNPJ do Beneficiário
  - Número da Guia (para Guias de Custas)
- **Filtros:** Ignora CNPJs específicos configuráveis via interface (ex: OAB).
- **Validação:** Alerta visual para valores acima de R$ 2.000,00 ou falhas de leitura.
- **Output:** Gera planilha `.xlsx` formatada e opcionalmente um arquivo `.csv` (separador ponto e vírgula).
- **Logs:** Sistema de log detalhado para debug (`%APPDATA%/PDF2EXCEL`).

## Dependências do Sistema

Para execução do código fonte ou do executável, as seguintes ferramentas devem estar instaladas no Windows:

1. **Tesseract OCR:**
   - Caminho padrão esperado: `C:\Program Files\Tesseract-OCR\tesseract.exe`
2. **Poppler (para pdf2image):**
   - Caminho padrão esperado: `C:\Program Files\poppler\bin`

> Caso utilize a versão compilada (.exe), o Poppler geralmente é empacotado junto, mas o Tesseract deve estar instalado na máquina host.

## Instalação (Source)

```bash
pip install -r requirements.txt
```

**Bibliotecas principais:**
- `pytesseract`
- `pdf2image`
- `openpyxl`
- `Pillow`
- `tkinter` (bult-in)

## Estrutura de Pastas

```text
📂 PDF2EXCEL
├── 📄 main.py               # Código fonte principal
├── 📄 correios_icon.ico     # Ícone da aplicação
├── 📂 logs                  # (Gerado em %APPDATA%)
│   ├── 📄 PDF2EXCEL.log
│   └── 📄 Filtro.config     # Lista de CNPJs ignorados
└── 📂 output                # Local selecionado pelo usuário para salvar relatórios
```

## Utilização

1. Execute o script/aplicação.
2. **Selecionar PDFs:** Escolha os arquivos ou a pasta contendo os boletos.
3. **Planilha de Saída:** Defina o nome e local do arquivo Excel.
4. **Parâmetros:**
   - *Ordem de Custas:* Identificador sequencial para organização interna.
   - *CSV:* Marque se desejar uma cópia em texto simples.
5. **Configuração de Filtro:**
   - Clique no botão "i" (Informações) -> "Filtro" para adicionar/remover CNPJs da blacklist.
