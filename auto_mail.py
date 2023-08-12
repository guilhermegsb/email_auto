import tkinter as tk
import tkinter.simpledialog as sd
from tkinter.simpledialog import askstring
from tkinter import filedialog
from imap_tools.mailbox import MailBox
from imap_tools.query import AND

import os
import datetime
import PyPDF4



class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Baixar e Mesclar Boletos")
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.btn = tk.Button(self)
        self.btn["text"] = "Baixar e Mesclar Boletos"
        self.btn["command"] = self.run_script
        self.btn.pack(side="top")

        self.text = tk.Text(self, state='disabled', height=5, wrap='word')
        self.text.pack(side="bottom", fill="both", expand=True)

    def run_script(self):
        try:
            usuario = sd.askstring("Email", "Insira seu email: ")
            senha = askstring("Senha", "Insira sua senha", show='*')

            meu_email = MailBox("imap-mail.outlook.com").login(usuario, senha)

            lista_emails = meu_email.fetch(AND(from_="financeiro@deltainf.com.br"))

            for email in lista_emails:
                if len(email.attachments) > 0:
                    for anexo in email.attachments:
                        nome = anexo.filename
                        if nome.endswith('.pdf'):
                            with open("{}.pdf".format(nome), "wb") as arquivo_pdf:
                                informacoes_anexo = anexo.payload
                                arquivo_pdf.write(informacoes_anexo)

            pdfiles = []
            local = os.listdir('.')

            for filename in local:
                if filename.endswith('.pdf') and filename != 'merged.pdf':
                    pdfiles.append(filename)

            pdfiles.sort(key=str.lower)

            pdfMerge = PyPDF4.PdfFileMerger()

            for filename in pdfiles:
                with open(filename, 'rb') as pdfFile:
                    pdfReader = PyPDF4.PdfFileReader(pdfFile)
                    pdfMerge.append(pdfReader)

            with open('Boletosdomes.pdf', 'wb') as pdfOutputFile:
                pdfMerge.write(pdfOutputFile)

            data_atual = datetime.datetime.now().strftime('%Y-%m-%d')
            os.mkdir(data_atual)

            for filename in pdfiles + ['Boletosdomes.pdf']:
                os.rename(filename, os.path.join(data_atual, filename))

            self.show_message("Boletos mesclados com sucesso!")
        except Exception as e:
            self.show_message("Erro ao mesclar boletos: " + str(e))

    def show_message(self, message):
        self.text.config(state='normal')
        self.text.insert(tk.END, message + '\n')
        self.text.config(state='disabled')


if __name__ == '__main__':
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()
