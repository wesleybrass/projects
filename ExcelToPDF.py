import os
import win32com.client as win32

# Função para converter arquivo Excel para PDF
def convert_excel_to_pdf(excel_file, output_pdf):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(excel_file)
    wb.ExportAsFixedFormat(0, output_pdf)
    wb.Close()
    excel.Quit()

# Diretório contendo os arquivos Excel
diretorio_excel = "C:\\Users\\TD CONSTRUCOES\\OneDrive\\Área de Trabalho\\Transfer\\Script - EXCEL to PDF"

# Diretório de saída para os arquivos PDF
diretorio_pdf = "C:\\Users\\TD CONSTRUCOES\\OneDrive\\Área de Trabalho\\Transfer"

# Loop pelos arquivos Excel no diretório
for nome_arquivo in os.listdir(diretorio_excel):
    if nome_arquivo.endswith(".xlsm") or nome_arquivo.endswith(".xlsx"):
        # Caminho completo para o arquivo Excel
        caminho_arquivo_excel = os.path.join(diretorio_excel, nome_arquivo)

        # Nome do arquivo PDF de saída
        nome_arquivo_pdf = os.path.splitext(nome_arquivo)[0] + ".pdf"

        # Caminho completo para o arquivo PDF de saída
        caminho_arquivo_pdf = os.path.join(diretorio_pdf, nome_arquivo_pdf)

        # Converter o arquivo Excel para PDF
        convert_excel_to_pdf(caminho_arquivo_excel, caminho_arquivo_pdf)

