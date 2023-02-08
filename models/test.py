# Guilherme TI

# pip install pypdf2

from PyPDF2 import PdfReader
import re
import os

file = r'C:\Users\TI\OneDrive - IG Guilherme TI\OneDrive - Irmaos Goncalves Com Ind Ltda\GIT\fatura_aws\Custo_AWS.xlsx'
pdf_past_aws = r'C:\Users\TI\OneDrive - IG Guilherme TI\OneDrive - Irmaos Goncalves Com Ind Ltda\GIT\fatura_aws\pdf\aws'
pdf_past_payment = r'C:\Users\TI\OneDrive - IG Guilherme TI\OneDrive - Irmaos Goncalves Com Ind Ltda\GIT\fatura_aws\pdf\payment'
pdf_file_aws = f'{pdf_past_aws}\Billing Management Console IGERP.pdf'
pdf_file_payment = f'{pdf_past_payment}\Billing Management Console IGERP.pdf'

# os.chdir(pdf_past_aws)

# Abre o arquivo pdf 
pdf_file_aws = PdfReader(pdf_file_aws)

# Pega o numero de páginas
number_of_pages = len(pdf_file_aws.pages)

# Lê a primeira página completa
page = pdf_file_aws.pages[0]

# Extrai apenas o texto
text = page.extract_text()
print()
print(text[222:])

conta_id = text[64:77]

mes = text[28:31]
ano = text[49:53]
mes_faturamento = mes + '/' + ano

fim = text.find('Amazon AWS Serviços Brasil Ltda.')
usd_market = float(text[238:fim].replace(' ','').replace(',', '.'))

usd_conta = float(text[280:fim].replace(' ','').replace(',', '.'))
print(usd_conta)

a = text.find('Ltda. USD')
print(a)




# print(text[172:])