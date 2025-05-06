import os
import subprocess
import sys
import re
from tkinter import *
from tkinter import filedialog, messagebox, ttk, font
from pdf2image import convert_from_path
import pytesseract
import openpyxl
from openpyxl import Workbook
import webbrowser
from threading import Thread
import logging
import tkinter as tk
import glob  # Importe o módulo glob para listar arquivos PDF
from openpyxl.styles import Alignment, PatternFill, Font
import csv # Importe o módulo csv

# Definir o caminho do Poppler (apenas para a versão .exe)
poppler_path = os.path.join(sys._MEIPASS, 'poppler', 'bin') if getattr(sys, 'frozen', False) else r"C:\Program Files\poppler\bin"

# Se o Poppler não estiver na pasta padrão do executável, verifica se está instalado em C:\Program Files\poppler\bin
if getattr(sys, 'frozen', False) and not os.path.exists(poppler_path):
    poppler_path = r"C:\Program Files\poppler\bin"  # Define o caminho alternativo

# Definir o caminho do Tesseract OCR (apenas para a versão .exe)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Desabilitar log no console/cmd (quase ignorado em .exe)
logging.basicConfig(level=logging.CRITICAL)

# Centraliza o programa na tela
def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    window.geometry(f'+{x}+{y}') # Removed fixed size and let window adapt to content

# Define no path do windows o poppler
def ocr_pdf(pdf_path):
    try:
        # Verifica a quantidade de páginas do PDF
        images = convert_from_path(pdf_path, poppler_path=poppler_path)
        num_pages = len(images)

        text = ""
        for image in images:
            if not processing: # Check processing flag inside the loop for faster cancellation
                break
            text += pytesseract.image_to_string(image, lang='por')
        return text, num_pages  # Retorna o texto e o número de páginas
    except Exception as e:
        logging.exception("Erro ao processar OCR do PDF")
        return None, 0  # Retorna None para o texto e o número de páginas

def extract_info(text):
    cnpj = None
    linhas_digitaveis = []
    valores_monetarios = []
    numero_guia = None  # Adicionado
    valor = None  # Adicionado

    # Busca o CNPJ do Beneficiário (modificado para excluir 82.519.190/0001-12)
    cnpj_matches = re.findall(r'(\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2})', text)
    valid_cnpjs = [cnp for cnp in cnpj_matches if cnp != '82.519.190/0001-12']
    if valid_cnpjs:
        cnpj = valid_cnpjs[0] # Pega o primeiro CNPJ válido encontrado
    else:
        cnpj = 'N/A'

    # Verifica se é uma guia de custas
    if "GUIA ÚNICA DE CUSTAS" in text:
        # Extrai o número da guia
        numero_guia_match = re.search(r"Nº da Guia\s*([\d\.]+/\d+)", text)
        if numero_guia_match:
            numero_guia = numero_guia_match.group(1)

        # Extrai o valor
        valor_match = re.search(r"R\$\s*([\d,.]+)", text)
        if valor_match:
            valor = valor_match.group(1)

        # Retorna os valores da guia de custas
        return {
            'cnpj': cnpj,
            'numero_guia': numero_guia,
            'valor': valor,
            'linhas_digitaveis': [],
            'valores': [],
            'tipo': 'guia_custas'
        }
    else:
        # Remove números de agência bancária antes de procurar a linha digitável
        text = re.sub(r'\d{3}-\d', '', text)

        for line in text.splitlines():
            cleaned_line = re.sub(r'[^0-9]', '', line)  # Remove tudo que não é número
            if 47 <= len(cleaned_line) <= 48:  # Verifica se a linha tem entre 47 e 48 dígitos
                linhas_digitaveis.append(cleaned_line)  # Adiciona a linha digitável à lista
                valor_monetario = f"{cleaned_line[-10:-2]},{cleaned_line[-2:]}"  # Formata como valor monetário
                valores_monetarios.append(valor_monetario)

        # Retorno dos valores obtidos do PDF escaneado
        return {
            'cnpj': cnpj,
            'linhas_digitaveis': linhas_digitaveis,
            'valores': valores_monetarios,
            'numero_guia': None,
            'valor': None,
            'tipo': 'boleto'
        }

# Criando uma variável global para armazenar a lista de arquivos com divergências
arquivos_com_paginas_a_mais = set()  # Usando set para evitar repetição
arquivos_com_dados_incompletos = set()  # Lista para armazenar arquivos com dados incompletos

def save_to_csv(result_file, ws):
    csv_file = os.path.splitext(result_file)[0] + ".csv"
    with open(csv_file, 'w', newline='', encoding='utf-8-sig') as file: # utf-8-sig para corrigir encoding
        writer = csv.writer(file, delimiter=';')
        for row in ws.iter_rows(values_only=True):
            writer.writerow(row)

def process_pdfs(input_files, output_dir, result_file, custas, save_csv): # Adicionado save_csv
    global arquivos_com_paginas_a_mais, arquivos_com_dados_incompletos, processing

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Formatação de planilha
    if not os.path.exists(result_file):
        wb = Workbook()
        ws = wb.active
        ws.title = "Boletos"
        ws.append(['Obeservação', 'Fornecedor', 'Código de Barras', 'Valor', 'Nome do Titulo'])
    else:
        try:
            wb = openpyxl.load_workbook(result_file)
            ws = wb.active
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo Excel: {e}", icon="error")
            return

    pdf_count = 0
    error_messages = []
    linhas_digitaveis_processadas = set()  # Usando um conjunto para rastrear linhas digitáveis processadas
    total_lines = 0  # Variável para controlar a numeração das linhas na coluna 'Nome do Titulo'

    try:
        for pdf_path in input_files:
            n_processo = os.path.basename(pdf_path) # Use o nome do arquivo com a extensão
            if not n_processo.lower().endswith('.pdf'):
                continue

            pdf_count += 1
            # pdf_path = os.path.join(input_dir, n_processo) # Não precisa mais do join, pois pdf_path já é o caminho completo
            popup_file_label.config(text=f"Processando: {n_processo}")
            root.update_idletasks()

            if not processing:  # Verifique se o processamento não foi cancelado
                break

            ocr_text, num_pages = ocr_pdf(pdf_path)  # Obtem o texto e o número de páginas
            if not processing:  # Check again after OCR for faster cancellation
                break

            if ocr_text:
                info = extract_info(ocr_text)
                nome_sem_extensao = os.path.splitext(n_processo)[0] #Remove a extensão .pdf

                total_lines += 1
                #Arquivo não encontrado
                if not any(info.values()):
                    arquivos_com_dados_incompletos.add(nome_sem_extensao)
                    ws.append([nome_sem_extensao, '', '', '', f"Custas{custas}:{total_lines:02}"])
                    # Marca a linha como amarela
                    for cell in ws[ws.max_row]:
                        cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                        cell.font = Font(color='000000')
                    error_messages.append(f"Arquivo {n_processo}: Nenhuma informação encontrada.")

                # Guia de custas com CNPJ, mas sem linha digitável
                elif info['tipo'] == 'guia_custas' and info['cnpj'] != 'N/A':
                    arquivos_com_dados_incompletos.add(nome_sem_extensao)
                    ws.append([nome_sem_extensao, info['cnpj'], '', info['valor'] if info['valor'] else '', f"Custas{custas}:{total_lines:02}"])  # Adiciona os dados à planilha, deixando "Código de Barras" e "Valor" em branco
                    # Marca a linha como amarela
                    for cell in ws[ws.max_row]:
                        cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                        cell.font = Font(color='000000')
                    error_messages.append(f"Arquivo {n_processo}: CNPJ encontrado, mas sem linha digitável (Guia de Custas?).")
                # Boletos Normais
                elif info['cnpj'] == 'N/A': # CNPJ não encontrado ou é o CNPJ indesejado
                    arquivos_com_dados_incompletos.add(nome_sem_extensao)
                    num_linhas = len(info['linhas_digitaveis'])
                    if num_linhas > 0:
                        for i in range(num_linhas):
                            linha_digitavel = info['linhas_digitaveis'][i]
                            valor_monetario = info['valores'][i]

                            if linha_digitavel in linhas_digitaveis_processadas:
                                continue
                            linhas_digitaveis_processadas.add(linha_digitavel)

                            try:
                                valor_float = float(valor_monetario.replace(',', '.'))
                                valor_formatado = "{:,.2f}".format(valor_float).replace(',', '*').replace('.', ',').replace('*', '.')
                            except ValueError:
                                valor_formatado = valor_monetario

                            if i == 0:
                                ws.append([nome_sem_extensao, '', linha_digitavel, valor_formatado, f"Custas{custas}:{total_lines:02}"]) # CNPJ fica vazio
                            else:
                                ws.append([f"{nome_sem_extensao} - Boleto página {i + 1}", '', linha_digitavel, valor_formatado, f"Custas{custas}:{total_lines:02}"]) # CNPJ fica vazio
                    else:
                        ws.append([nome_sem_extensao, '', '', '', f"Custas{custas}:{total_lines:02}"]) # CNPJ e outros dados ficam vazios
                    # Marca a linha como amarela
                    for cell in ws[ws.max_row]:
                        cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                        cell.font = Font(color='000000')
                    error_messages.append(f"Arquivo {n_processo}: CNPJ não encontrado ou inválido.")


                # Boletos Normais COM CNPJ válido
                else:
                    num_linhas = len(info['linhas_digitaveis'])

                    # Itera pelas linhas digitáveis encontradas no arquivo
                    for i in range(num_linhas):
                        linha_digitavel = info['linhas_digitaveis'][i]
                        valor_monetario = info['valores'][i]

                        # Verifica se a linha digitável já foi processada para este arquivo
                        if linha_digitavel in linhas_digitaveis_processadas:
                            continue

                        linhas_digitaveis_processadas.add(linha_digitavel)  # Adiciona a linha digitável ao conjunto de processados

                        # Converte o valor para float e formata com vírgula
                        try:
                            valor_float = float(valor_monetario.replace(',', '.'))
                            valor_formatado = "{:,.2f}".format(valor_float).replace(',', '*').replace('.', ',').replace('*', '.')  # Formatação para o formato brasileiro
                        except ValueError:
                            valor_formatado = valor_monetario  # Mantém o valor original se não for possível converter

                        # Adiciona os dados do boleto atual à planilha
                        if i == 0:  # Adiciona a linha digitável com o nome original do processo
                            ws.append([nome_sem_extensao, info['cnpj'], linha_digitavel, valor_formatado, f"Custas{custas}:{total_lines:02}"])
                        else:  # Adiciona as linhas digitáveis seguintes com a numeração da página
                            ws.append([f"{nome_sem_extensao} - Boleto página {i + 1}", info['cnpj'], linha_digitavel, valor_formatado, f"Custas{custas}:{total_lines:02}"])

                        # Verifica se algum campo é N/A (exceto CNPJ que já foi tratado antes)
                        if 'N/A' in [v for k, v in info.items() if k != 'cnpj']: # Check N/A in other fields except cnpj
                            error_messages.append(f"Arquivo {n_processo}: Dados inválidos ou ausentes (exceto CNPJ).")
                            arquivos_com_dados_incompletos.add(nome_sem_extensao)  # Adiciona à lista de arquivos com dados incompletos

                # Ajusta a largura das colunas la no excel
                for col in range(1, ws.max_column + 1):
                    column_letter = openpyxl.utils.get_column_letter(col)
                    column_width = max(len(str(cell.value)) if cell.value else 0 for cell in ws[column_letter])
                    ws.column_dimensions[column_letter].width = max(column_width, 10)

                # Se o arquivo tiver mais de uma página, adiciona-o à lista de arquivos com divergências
                if num_pages > 1:
                    arquivos_com_paginas_a_mais.add(nome_sem_extensao)

            else:
                error_messages.append(f"Arquivo {n_processo}: Falha no processamento do OCR.")
                arquivos_com_dados_incompletos.add(nome_sem_extensao)
                ws.append([nome_sem_extensao, '', '', '', f"Custas{custas}:{total_lines:02}"])
                # Marca a linha como amarela
                for cell in ws[ws.max_row]:
                    cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                    cell.font = Font(color='000000')

            if processing: # Only update progress bar if not cancelled
                popup_progress_bar['value'] += 1
                root.update_idletasks()
            if not processing: # Check again inside the loop for faster cancellation
                break

        # Marca as linhas com problemas após processar todos os arquivos
        for row in range(2, ws.max_row + 1):
            nome_arquivo = ws.cell(row, 1).value
            if nome_arquivo in arquivos_com_dados_incompletos:
                for cell in ws[row]:
                    cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                    cell.font = Font(color='000000')

        # Formata para ser valor... mas na pratica ta meio estranho não me aparece os R$
        # Alinha o conteudo da coluna D para a Direita, MENOS o titulo que é na linha 1
        for cell in ws['D']:
            cell.number_format = 'R$ #,##0.00'
            cell.alignment = Alignment(horizontal='right')
        ws['D1'].alignment = Alignment(horizontal='left')

        # Verifica valores acima de 2000 reais e marca na planilha e dá alerta
        for row in range(2, ws.max_row + 1):
            valor_boleto = ws.cell(row, 4).value
            if valor_boleto and isinstance(valor_boleto, str):
                valor_boleto = float(valor_boleto.replace('.', '').replace(',', '.')) # Transforma para float removendo formatação indesejada.
            if valor_boleto and valor_boleto > 2000:
                error_messages.append(f"Arquivo {ws.cell(row, 1).value}: Valor do boleto acima de R$ 2000. Verificar manual.")
                for cell in ws[row]:
                    cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                    cell.font = Font(color='000000')

    except Exception as e:
        logging.exception("Erro durante o processamento dos PDFs")
        messagebox.showerror("Erro", f"Erro durante o processamento: {e}", icon="error")
    finally:
        if processing: # Only save if not cancelled
            try:
                wb.save(result_file)
                if save_csv and not error_messages and not arquivos_com_paginas_a_mais and not arquivos_com_dados_incompletos: # salva CSV se não tiver erros e checkbox marcado
                    save_to_csv(result_file, ws)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar o arquivo Excel: {e}", icon="error")

    # Converte os sets para listas antes de passar para a função de aviso
    arquivos_com_paginas_a_mais_list = list(arquivos_com_paginas_a_mais)
    arquivos_com_dados_incompletos_list = list(arquivos_com_dados_incompletos)

    if processing: # Only show popups if not cancelled
        # Mostra o aviso de erro caso haja mensagens
        if error_messages or arquivos_com_paginas_a_mais_list or arquivos_com_dados_incompletos_list: # Verifica se há divergências para mostrar popup correto
            show_divergencia_popup(error_messages, result_file, arquivos_com_paginas_a_mais_list, arquivos_com_dados_incompletos_list, save_csv) # Pass save_csv para o popup
        else:
            show_success_popup(result_file, save_csv) # Mostra popup de sucesso.
    else:
        show_cancelled_popup() # Show cancelled popup

    return arquivos_com_paginas_a_mais, arquivos_com_dados_incompletos

def select_input_files():
    global input_files
    input_files = filedialog.askopenfilenames(
        filetypes=[("Arquivos PDF", "*.pdf")],
        title="Selecione os arquivos PDF"
    )
    input_dir_label.config(text=f"{len(input_files)} arquivos selecionados")  # Mostra quantos arquivos foram selecionados

def select_result_file():
    global result_file
    result_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Planilhas Excel", "*.xlsx")],
        title="Salve a planilha de resultados",
    )
    result_file_label.config(text=f"{result_file}")

# Certifica das dependencias corretas para rodar o programa
def process_pdfs_thread(input_files, output_dir, result_file, custas, save_csv_var): # Adicionado save_csv_var
    global processing, popup_progress_bar

    # Verificação do tipo de 'output_dir' antes de usá-lo
    if not isinstance(output_dir, str):
        messagebox.showerror("Erro", "O caminho de saída (output_dir) não é uma string válida.", icon="error")
        processing = False
        start_button['state'] = NORMAL
        return

    processing = True
    try:
        os.environ["PATH"] = poppler_path + os.path.sep + os.environ["PATH"]
        popup_progress_bar['value'] = 0
        num_pdfs = len(input_files) # Conta a lista de arquivos selecionados
        popup_progress_bar['maximum'] = num_pdfs
        if num_pdfs == 0:
            messagebox.showinfo("Aviso", "Nenhum arquivo PDF selecionado.", icon="warning")
            processing = False
            return

        save_csv = save_csv_var.get() # Pega o valor do checkbox
        arquivos_com_paginas_a_mais, arquivos_com_dados_incompletos = process_pdfs(input_files, output_dir, result_file, custas, save_csv) # Passa save_csv para process_pdfs
        # Não precisa mais do webbrowser aqui.
        processing = False # Garante que o processing vai virar False após o loop

    except Exception as e:
        messagebox.showerror("Erro", f"Erro durante o processamento: {e}", icon="error")
        logging.exception("Erro no thread de processamento")
        processing = False # Garante que o processing vai virar False mesmo em caso de erro
    finally:
        processing = False
        if popup and popup.winfo_exists(): # Check if popup exists and is not destroyed before destroying
            popup.destroy()
        # **Reativa o botão de iniciar processamento após o término do thread**
        start_button['state'] = NORMAL

# Função para criar botões com estilo uniforme
def create_button(parent, text, command, is_default=False, fg_color="#FF0000"): # Adicionado is_default and fg_color
    # Cria o botão com a fonte, cor de fundo, e borda padrão da janela popup
    button = Button(parent, text=text, command=command,
                     font=("Segoe UI Bold", 10),
                     bg=parent.cget("bg"),
                     fg=fg_color, # Use fg_color here
                     relief=FLAT,
                     borderwidth=1,
                     padx=10,
                     pady=5)
    if is_default: # Se for o botão padrão, define como default
        button.bind("<Return>", lambda event: command())
        button.focus_set() # Define o foco para que Enter funcione imediatamente
    return button

# Mostra a janela de aviso de erro
def show_divergencia_popup(error_messages, result_file, arquivos_com_paginas_a_mais, arquivos_com_dados_incompletos, save_csv): # Adicionado save_csv
    global processing
    popup = Toplevel(root)
    popup.title("Processamento Concluído com Divergências")
    popup.transient(root)
    popup.grab_set()
    popup.resizable(False, False) # Impede redimensionar

    label = Label(popup, text=f"Processamento concluído! Divergências encontradas:", font=("Segoe UI Bold", 10), fg="#FF0000")
    label.pack(pady=10)

    # Seção para arquivos com várias páginas
    if arquivos_com_paginas_a_mais:
        label_paginas_a_mais = Label(popup, text=f"Arquivos com mais de uma página: {len(arquivos_com_paginas_a_mais)}", font=("Segoe UI Bold", 10))
        label_paginas_a_mais.pack(pady=5)
        lista_paginas_a_mais = Label(popup, text="\n".join(arquivos_com_paginas_a_mais), font=("Segoe UI Bold", 10))
        lista_paginas_a_mais.pack(pady=5)

    # Seção para arquivos com dados incompletos
    if arquivos_com_dados_incompletos:
        label_dados_incompletos = Label(popup, text=f"Arquivos com informações faltando: {len(arquivos_com_dados_incompletos)}", font=("Segoe UI Bold", 10))
        label_dados_incompletos.pack(pady=5)
        lista_dados_incompletos = Label(popup, text="\n".join(arquivos_com_dados_incompletos), font=("Segoe UI Bold", 10))
        lista_dados_incompletos.pack(pady=5)

    # Seção para mensagens de erro gerais
    if error_messages:
        label_erros = Label(popup, text="Erros gerais:", font=("Segoe UI Bold", 10))
        label_erros.pack(pady=5)
        lista_erros = Label(popup, text="\n".join(error_messages), font=("Segoe UI Bold", 10))
        lista_erros.pack(pady=5)

    # Mensagem sobre CSV não criado por falta de confiabilidade
    if save_csv and (error_messages or arquivos_com_paginas_a_mais or arquivos_com_dados_incompletos):
        label_csv_nao_criado = Label(popup, text="(CSV automático não criado, por falta de confiabilidade)", font=("Segoe UI Bold", 10), fg="#FF0000")
        label_csv_nao_criado.pack(pady=5)

    button_frame = Frame(popup)
    button_frame.pack(pady=10)

    ok_button = create_button(button_frame, "OK", lambda: [popup.destroy(), webbrowser.open_new_tab(f"file://{result_file}")], is_default=True) # OK is default
    ok_button.pack(side=LEFT, padx=10)

    # Abre todos os arquivos com várias páginas e arquivos com dados incompletos em uma única lista
    def get_pdf_path(arquivo_nome):
        # Procura o arquivo PDF pelo nome (sem extensão) na lista de arquivos selecionados
        for file_path in input_files:
            if os.path.splitext(os.path.basename(file_path))[0] == arquivo_nome:
                return file_path
        return None  # Retorna None se não encontrar

    # Abre todos os arquivos com várias páginas e arquivos com dados incompletos em uma única lista
    conferir_pdfs_button = Button(button_frame, text="Conferir PDFs", command=lambda: [popup.destroy(), abrir_arquivos_pdf(arquivos_com_paginas_a_mais + arquivos_com_dados_incompletos), webbrowser.open_new_tab(f"file://{result_file}")])
    conferir_pdfs_button.pack(side=LEFT, padx=10)

    center_window(popup)

# Mostra a janela de sucesso
def show_success_popup(result_file, save_csv): # Adicionado save_csv
    popup = Toplevel(root)
    popup.title("Processamento Concluído com Sucesso")
    popup.transient(root)
    popup.grab_set()
    popup.resizable(False, False) # Impede redimensionar

    label = Label(popup, text="Processamento concluído com sucesso!", font=("Segoe UI Bold", 10), fg="green")
    label.pack(pady=10)

    if save_csv:
        label_csv_criado = Label(popup, text="(CSV automático também foi criado)", font=("Segoe UI Bold", 10), fg="green")
        label_csv_criado.pack(pady=5)

    button_frame = Frame(popup)
    button_frame.pack(pady=10)

    ok_button = create_button(button_frame, "OK", lambda: [popup.destroy(), webbrowser.open_new_tab(f"file://{result_file}")], is_default=True) # OK is default
    ok_button.pack(side=LEFT, padx=10)

    center_window(popup)

# Mostra a janela de cancelamento
def show_cancelled_popup():
    popup = Toplevel(root)
    popup.title("Processamento Cancelado!")
    popup.transient(root)
    popup.grab_set()
    popup.resizable(False, False) # Impede redimensionar

    label = Label(popup, text="Processamento Cancelado!", font=("Segoe UI Bold", 10), fg="red") # Red text
    label.pack(pady=10)

    button_frame = Frame(popup)
    button_frame.pack(pady=10)

    ok_button = create_button(button_frame, "OK", popup.destroy, is_default=True, fg_color="#FF0000") # Red OK button, matching other buttons
    ok_button.pack(side=LEFT, padx=10)

    center_window(popup)


# Função para abrir todos os PDFs com várias páginas
def abrir_arquivos_pdf(arquivos):
    for arquivo in arquivos:
        # Procura o caminho completo do arquivo PDF
        caminho_completo = next((f for f in input_files if os.path.splitext(os.path.basename(f))[0] == arquivo), None)
        if caminho_completo:
            webbrowser.open_new_tab(f"file://{caminho_completo}")  # Abre o arquivo se encontrado

# Inicia o processo, se tiver rodando ignore a solicitação (era para o botão Enter dar start processing, mas ele desconfigura e de erros estranhos)
def start_processing():
    global processing, popup, popup_file_label, popup_progress_bar, custas_entry, input_files, result_file, save_csv_var  # Make sure save_csv_var is global

    # **Se o botão estiver desabilitado, não faz nada**
    if start_button['state'] == DISABLED:
        return

    processing = True

    # Check if 'input_files' is defined
    try:
        input_files
    except NameError:
        messagebox.showerror("Erro", "Arquivos não Selecionados: Selecione os arquivos PDF.", icon="error")
        start_button['state'] = NORMAL
        return

    # Check if 'result_file' is defined
    try:
        result_file
    except NameError:
        messagebox.showerror("Erro", "Planilha não Selecionada: Selecione a planilha de resultados.", icon="error")
        start_button['state'] = NORMAL
        return

    if not input_files or not result_file:
        # **Cria a janela de aviso de erro**
        error_popup = Toplevel(root)
        error_popup.title("Erro")
        error_popup.transient(root)
        error_popup.grab_set()
        error_popup.resizable(False, False) # Impede redimensionar

        # Verifica qual campo está vazio e monta a mensagem de erro
        if not input_files:
            error_message = "Arquivos PDF não selecionados. Selecione os arquivos."
        else:
            error_message = "Campo 'Local do Resultado' vazio. Preencha o campo."

        error_label = Label(error_popup, text=error_message, font=("Segoe UI Bold", 10))
        error_label.pack(pady=10)

        ok_button = create_button(error_popup, "OK", error_popup.destroy, is_default=True)  # Usa a função create_button aqui, OK is default
        ok_button.pack(pady=10)

        center_window(error_popup)
        # **Reativa o botão**
        start_button['state'] = NORMAL
        return

    # **Verificação do Poppler movida para aqui, pois é realizada antes de iniciar o thread**
    if not os.path.exists(poppler_path):
        messagebox.showerror("Erro", f"Pasta do Poppler não encontrada em: {poppler_path}. O programa não poderá funcionar corretamente", icon="error")
        # **Reativa o botão**
        start_button['state'] = NORMAL
        return

    custas = custas_entry.get()
    if not re.match(r'^[0-9\.\:\/\\]{1,5}$', custas):  # Permite de 1 a 5 caracteres
        messagebox.showerror("Erro", "Digite um valor válido para as custas (apenas números, '.', ':', '/', '\\) com até 5 caracteres.", icon="error")
        # **Reativa o botão**
        start_button['state'] = NORMAL
        return

    # **Desabilita o botão de iniciar processamento**
    start_button['state'] = DISABLED

    # Janelas pop-up que deram trabalho
    popup = Toplevel(root)
    popup.title("Processando PDFs")
    popup.transient(root)  # Garante a janela pop-up sempre na frente da janela principal (erros em .exe podem ocorrer por causa do alt+tab)
    popup.grab_set()  # Bloqueia mexer nos parametros da janela principal, usuario não pode mexer nesse momento
    popup.resizable(False, False) # Impede redimensionar
    popup_file_label = Label(popup, text="Processando arquivo...", font=("Segoe UI Bold", 10))
    popup_file_label.pack(pady=10)
    popup_progress_bar = ttk.Progressbar(popup, orient="horizontal", length=300, mode="determinate")
    popup_progress_bar.pack(pady=10)
    cancel_button = Button(popup, text="Cancelar", command=cancel_processing,
                           font=("Segoe UI Bold", 10),
                           bg=popup.cget("bg"),
                           fg="#FF0000",
                           relief=FLAT,
                           borderwidth=1,
                           padx=10,
                           pady=5)
    cancel_button.pack(pady=10)
    center_window(popup)
    # Correção aqui: output_dir precisa ser uma string válida (um caminho de diretório)
    # Certifique-se de que 'input_dir' está definido e é uma string
    output_dir = input_dir_label.cget("text")  # Ou qualquer diretório que você queira usar
    if not output_dir or output_dir == "Nenhum arquivo selecionado":
        output_dir = os.getcwd()  # Use o diretório atual como padrão
        messagebox.showinfo("Informação", "Pasta de saída não definida. Usando o diretório atual.", icon="warning")

    save_csv = save_csv_var.get() # Get checkbox value here
    thread = Thread(target=process_pdfs_thread, args=(input_files, output_dir, result_file, custas, save_csv_var)) # Pass save_csv_var
    thread.start()

def cancel_processing():
    global processing, popup

    processing = False # Set processing to false immediately

    if popup and popup.winfo_exists(): # Check if popup exists and is not destroyed
        popup.destroy() # Destroy the processing popup
    show_cancelled_popup() # Show the cancellation popup
    start_button['state'] = NORMAL # Re-enable start button

def toggle_csv_save():
    save_csv_var.set(not save_csv_var.get())

# Função para criar botões arredondados
def create_rounded_button(parent, text, command, width=100, height=50):
    canvas = Canvas(parent, width=width, height=height, bd=0, highlightthickness=0, relief='ridge')
    canvas.create_oval(5, 5, width-5, height-5, outline="#0000FF", fill="#0000FF")
    canvas.create_text(width/2, height/2, text=text, fill="#FFFFFF", font=("Segoe UI Bold", 10))  # centralizado
    canvas.bind("<Button-1>", lambda event: command())
    return canvas

# Janela principal do programa
root = Tk()
root.title("Extrator de Dados Do Boleto (1.3.0a)")
font_style = font.Font(family="Segoe UI Bold", size=15)
root.option_add("*Font", font_style)
root.resizable(False, False) # Impede redimensionar
frame = Frame(root)
frame.pack(padx=20, pady=20)

# Botões entrada
input_dir_button = Button(frame, text=" Selecionar PDFs ", command=select_input_files)  # Alterado para arquivos
input_dir_button.grid(row=0, column=0, padx=5, pady=5, sticky="e")
input_dir_label = Label(frame, text="PDF não selecionado")
input_dir_label.grid(row=0, column=1, padx=5, pady=5, sticky="w")

# Botões saida
result_file_button = Button(frame, text="Local do Resultado", command=select_result_file)
result_file_button.grid(row=1, column=0, padx=5, pady=5, sticky="e")
result_file_label = Label(frame, text="Arquivo não selecionado")
result_file_label.grid(row=1, column=1, padx=5, pady=5, sticky="w")

# Checkbox para salvar em CSV e label como botão
save_csv_var = BooleanVar()
save_csv_label = Label(frame, text="CSV ponto e vírgula:", cursor="hand2") # Cursor hand2 para indicar que é clicável
save_csv_label.grid(row=2, column=0, pady=5, sticky="e") # Posiciona ACIMA do Custas
save_csv_label.bind("<Button-1>", lambda event: toggle_csv_save()) # Liga o clique do label para a função
save_csv_check = Checkbutton(frame, variable=save_csv_var, command=toggle_csv_save) # Checkbox agora também chama a função
save_csv_check.grid(row=2, column=1, pady=5, sticky="w") # Checkbox na coluna 1, alinhado à esquerda

# Campo digitavel Matheus
custas_label = Label(frame, text="Custas:", anchor='e', justify='left') # Alinhamento a direita e justificado a esquerda
custas_label.grid(row=3, column=0, padx=5, pady=5, sticky="e") # Posiciona ABAIXO do CSV e alinhado a direita
custas_entry = Entry(frame, width=10)
custas_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w") # Posiciona ABAIXO do CSV

# Limita a entrada do campo Custas a 5 caracteres
def limit_custas_entry(event):
    if len(custas_entry.get()) >= 5:
        event.widget.delete(5, END)

custas_entry.bind("<KeyRelease>", limit_custas_entry)


# **Botão de Iniciar processamento**
start_button = Button(root, text="Iniciar Processamento", command=start_processing)
start_button.pack(pady=10)

# Botão de Informação
def show_info():
    info_popup = Toplevel(root)
    info_popup.title("Informação")
    info_popup.transient(root)
    info_popup.grab_set()
    info_popup.resizable(False, False) # Impede redimensionar
    version_label = Label(info_popup, text="1.3.0a - by Elias", font=("Segoe UI Bold", 10), bg=info_popup.cget("bg"), fg="#00FF00")
    pix_label = Label(info_popup, text="Pix de Apoio: eliasgkersten@gmail.com", font=("Segoe UI Bold", 10), bg=info_popup.cget("bg"), fg="#00FF00")
    version_label.pack(pady=10)
    pix_label.pack(pady=5)
    # Cria o botão "Sair" normal
    exit_button = create_button(info_popup, "Sair", info_popup.destroy, is_default=True) # Exit is default
    exit_button.pack(pady=10)
    center_window(info_popup)

# **Criando o botão redondo de informação**
show_info_button = create_rounded_button(root, "i", show_info, width=25, height=25)
show_info_button.place(relx=1.0, rely=1.0, anchor="se")  # Coloca na posição SE
show_info_button = create_rounded_button(root, "i", show_info, width=50, height=50)

center_window(root)

# Corrigindo o problema do botão Enter
root.bind("<Return>", lambda event: start_processing() if not root.grab_current() else "break") # Only start processing if no popup is active

# Inicialize a variável processing como False antes de qualquer função que a use.
processing = False

# **Parte para esconder a janela do CMD**
if getattr(sys, 'frozen', False):
    # Se o programa estiver empacotado como .exe, esconde a janela do CMD
    subprocess.run('start "" /B cmd /c @echo off', shell=True)

root.mainloop()