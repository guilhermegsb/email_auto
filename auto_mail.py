from imap_tools import MailBox, AND, OR, MailMessage
from getpass import getpass




usuario= input("Insira seu email:")
senha= getpass("Insira sua Senha:")
    
meu_email= MailBox("imap-mail.outlook.com").login(usuario, senha)

import PyPDF2
import os


lista_emails = meu_email.fetch(AND(from_="financeiro@deltainf.com.br"))

for email in (lista_emails):
    if len(email.attachments) > 0:
        for anexo in email.attachments: 
            nome = anexo.filename
            if 'pdf' in nome:
                with open("{}.pdf".format(nome), "wb") as arquivo_pdf:
                    informacoes_anexo= anexo.payload
                    arquivo_pdf.write(informacoes_anexo) 
                
pdfiles = []

local= os.listdir('.')

for filename in local:
        if filename.endswith('.pdf'):
                if filename != 'merged.pdf':
                        pdfiles.append(filename)
                        
pdfiles.sort(key = str.lower)


pdfMerge = PyPDF2.PdfMerger()

for filename in pdfiles:
        pdfFile = open(filename, 'rb')
        pdfReader = PyPDF2.PdfReader(pdfFile)
        pdfMerge.append(pdfReader)

pdfFile.close()
pdfMerge.write('Boletosdomes.pdf')
            
