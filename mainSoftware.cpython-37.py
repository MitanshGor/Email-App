import tkinter as tk
from tkinter import ttk
from tkinter import Frame
from tkinter import *
from tkinter.messagebox import showinfo
import tkinter.font as font
import tkinter.filedialog as filedialog
from tkinter import messagebox as mb
import pandas as pd, smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from smtplib import SMTPAuthenticationError

class App(tk.Tk):

    def __init__(self):
        self.root = tk.Tk.__init__(self)
        self.title('Mailing App')
        self.message = MIMEMultipart('alternative')
        self.password = open('creditenials.txt').read().splitlines()[1]
        self.email = open('creditenials.txt').read().splitlines()[0]
        self.geometry('500x400')
        self.resizable(False, False)
        self.top = tk.Toplevel(self.root)
        self.top.withdraw()
        self.top.resizable(False, False)
        self.top.title('Attachment Selection')
        self.top.geometry('575x600')
        self.path = {}
        self.ctop = tk.Toplevel(self.root)
        self.ctop.withdraw()
        self.ctop.resizable(False, False)
        self.ctop.title('Mailing Process')
        self.ctop.geometry('300x400')
        self.frame = Frame(self.top)
        self.ctop.ctitle = tk.Label((self.ctop), text='Total Mail : ', font=('Arial',
                                                                             18))
        self.ctop.ctitle.grid(row=2, column=2, pady=(25, 25))
        self.ctop.ctitle1 = tk.Label((self.ctop), text='Mail Left : ', font=('Arial',
                                                                             18))
        self.ctop.ctitle1.grid(row=3, column=2, pady=(25, 25))
        self.ctop.ctitle2 = tk.Label((self.ctop), text='Mail Sent : ', font=('Arial',
                                                                             18))
        self.ctop.ctitle2.grid(row=4, column=2, pady=(25, 25))
        self.i = 0
        self.excelFile = tk.Label(self, font=('Arial', 12)).grid(row=(self.i), column=1)
        self.i += 1
        self.title = tk.Label(self, text='Hello, To Mail!', font=('Arial', 25)).grid(row=(self.i), column=2)
        self.i += 1
        self.excelFile = tk.Label(self, font=('Arial', 12)).grid(row=(self.i), column=1)
        self.i += 1
        self.excelFile = tk.Label(self, text='Excel File: ', font=('Arial', 12)).grid(row=(self.i), column=1)
        self.excelPath = tk.Entry(self, bg='red', width=40, state='readonly')
        self.excelPath.grid(row=(self.i), column=2)
        self.excelPathButton = ttk.Button(self, text='Browse', command=(self.inputExcelFunction)).grid(row=(self.i), column=3)
        self.i += 1
        self.excelFile = tk.Label(self, font=('Arial', 12)).grid(row=(self.i), column=1)
        self.i += 1
        self.htmlFile = tk.Label(self, text='HTML File: ', font=('Arial', 12)).grid(row=(self.i), column=1)
        self.htmlFilePath = tk.Entry(self, width=40, state='readonly')
        self.htmlFilePath.grid(row=(self.i), column=2)
        self.htmlPathButton = ttk.Button(self, text='Browse', command=(self.inputHTMLFunction)).grid(row=(self.i), column=3)
        self.i += 1
        self.excelFile = tk.Label(self, font=('Arial', 12)).grid(row=(self.i), column=1)
        self.i += 1
        self.subject = tk.Label(self, text='Subject : ', font=('Arial', 12)).grid(row=(self.i), column=1)
        self.subject = tk.Entry(self, width=45)
        self.subject.grid(row=(self.i), column=2, columnspan=2)
        self.i += 1
        self.excelFile = tk.Label(self, font=('Arial', 12)).grid(row=(self.i), column=1)
        self.i += 1
        self.selectedColumnName = tk.StringVar()
        self.excelFile = tk.Label(self, text='Column Name: ', font=('Arial', 12)).grid(row=(self.i), column=1)
        self.columnName = ttk.Combobox(self, width=45, textvariable=(self.selectedColumnName), state='readonly', postcommand=(self.columnHead))
        num = []
        num.insert(0, 'Please select value')
        self.columnName['values'] = list(num)
        self.columnName.current(0)
        self.columnName.grid(row=(self.i), column=2, columnspan=2)
        self.i += 1
        self.excelFile = tk.Label(self, font=('Arial', 12)).grid(row=(self.i), column=1)
        self.i += 1
        self.maximumAttachmentN = tk.StringVar()
        self.excelFile = tk.Label(self, text='Number of Attachment : ', font=('Arial',
                                                                              12)).grid(row=(self.i), column=1)
        self.maximumAttachment = ttk.Combobox(self, width=45, textvariable=(self.maximumAttachmentN), state='readonly')
        self.maximumAttachment.bind('<<ComboboxSelected>>', self.attachmentBlock)
        num = list(range(0, 6))
        num.insert(0, 'Please select value')
        self.maximumAttachment['values'] = list(num)
        self.maximumAttachment.current(0)
        self.maximumAttachment.grid(row=(self.i), column=2, columnspan=2)
        self.i += 1
        self.path = {}
        self.entryList = []
        self.excelFile = tk.Label(self, font=('Arial', 12)).grid(row=(self.i), column=1)
        self.i += 1
        self.button1 = tk.Button(self, text='Attach', command=(self.attachMentButton), padx=40, pady=10).grid(row=(self.i), column=2, columnspan=2)
        self.button2 = tk.Button(self, text='Exit', command=(self.exitClicked), padx=40, pady=10).grid(row=(self.i), column=1, columnspan=2)

    def attachmentBlock(self, event):
        for widget in self.top.winfo_children():
            widget.destroy()

        self.path.clear()
        self.entryList.clear()
        if self.maximumAttachment.get() != 'Please select value':
            num = int(self.maximumAttachment.get())
            if num == 1:
                self.top.geometry('650x180')
            elif num == 2:
                self.top.geometry('650x250')
            elif num == 3:
                self.top.geometry('650x340')
            elif  num == 4:
                self.top.geometry('650x430')
            elif num == 5:
                self.top.geometry('650x480')
            i = 0
            self.top.title = tk.Label((self.top), text='Attachment List', font=('Arial',
                                                                                25)).grid(row=i, column=1)
            i += 1
            z = 1
            for j in range(0, num):
                labelName = 'self.top.labelName' + str(z)
                entryName = 'self.top.filePath' + str(z)
                buttonName = 'self.top.filePathButton' + str(z)
                functionToCalled = 'self.inputAttachment' + str(z)
                i += 1
                labelName = tk.Label((self.top), text='Attachment {z}: '.format(z=z), font=('Arial',
                                                                                            12)).grid(row=i, column=0, padx=(50,
                                                                                                                             0), pady=(25,
                                                                                                                                       25))
                entryName = tk.Entry((self.top), bg='red', width=60, state='readonly')
                entryName.grid(row=i, column=1)
                self.entryList.append(entryName)
                buttonName = ttk.Button((self.top), text='Browse', command=(eval(functionToCalled))).grid(row=i, column=2)
                i += 1
                z += 1

            self.button1 = tk.Button((self.top), text='Send-Mail', command=(self.sendMailValidation), padx=40, pady=10).grid(row=i, column=1, padx=(80,
                                                                                                                                                    0), columnspan=2)
            self.button2 = tk.Button((self.top), text='Home-Page', command=(self.homePage), padx=40, pady=10).grid(row=i, column=0, columnspan=2)

    def homePage(self):
        self.top.withdraw()
        self.iconify()

    def sendMailValidation(self):
        num = int(self.maximumAttachment.get())
        if len(self.path) == num:
            self.sendMailFunction()
        else:
            mb.showerror('Error', 'Please enter valid number of files. Number of files left = ' + str(num - len(self.path)))

    def messageLoading(self):
        html = open(self.htmlFilePath.get()).read()
        part2 = MIMEText(html, 'html')
        self.message['From'] = self.email
        self.message['Subject'] = self.subject.get()
        self.message.attach(part2)
        files = []
        files = list(self.path.values())
        if len(files) != 0:
            for path in files:
                part = MIMEBase('application', 'octet-stream')
                with open(path, 'rb') as (file):
                    part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename="{}"'.format(Path(path).name))
                self.message.attach(part)

    def mailFunction(self):
        recieverEmail = self.df[self.columnName.get()].tolist()
        recieverEmail = [x for x in recieverEmail if str(x) != 'nan']
        counter = 0
        self.top.withdraw()
        self.ctop.deiconify()
        self.ctop.title = tk.Label((self.ctop), text='Mailing In Process!', font=('Arial',25)).grid(row=1, column=2, pady=(25,25))
        self.ctop.ctitle.config(text=('Total Mail : ' + str(len(recieverEmail))))
        self.ctop.ctitle1.config(text=('Mail Left : ' + str(len(recieverEmail) - counter)))
        self.ctop.ctitle2.config(text=('Mail Sent : ' + str(counter)))
        self.ctop.update()
        flag = False
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
            try:
                server.login(self.email, self.password)
                for emailID in recieverEmail:
                    del self.message['To']
                    self.message['To'] = emailID
                    server.sendmail(self.message['From'], self.message['To'], self.message.as_string())
                    counter += 1
                    self.ctop.ctitle1.config(text=('Mail Left : ' + str(len(recieverEmail) - counter)))
                    self.ctop.ctitle2.config(text=('Mail Sent : ' + str(counter)))
                    self.ctop.ctitle1.update()
                    self.ctop.ctitle2.update()
                    

            except SMTPAuthenticationError:
                flag = True
                mb.showerror('Error', 'Please check that whether EMAILID or PASSWORD is valid or not')
            except:
                flag=True
                mb.showerror('Error', 'Some error Occured')

        if not flag:
            mb.showinfo('Success', 'Completed all mails succesfully')
        for widget in self.top.winfo_children():
            widget.destroy()

        self.path.clear()
        self.entryList.clear()
        self.excelPath.config(state='normal')
        self.excelPath.delete(0, END)
        self.excelPath.config(state='readonly')
        self.htmlFilePath.config(state='normal')
        self.htmlFilePath.delete(0, END)
        self.htmlFilePath.config(state='readonly')
        self.subject.delete(0, END)
        num = []
        num.insert(0, 'Please select value')
        self.columnName['values'] = list(num)
        self.columnName.current(0)
        self.maximumAttachment.current(0)
        self.update()
        self.top.withdraw()
        self.ctop.withdraw()
        self.deiconify()

    def sendMailFunction(self):
        self.messageLoading()
        self.mailFunction()

    def inputAttachment1(self):
        inputPath = tk.filedialog.askopenfilename(title='Select file')
        try:
            if len(inputPath) != 0:
                self.entryList[0].config(state='normal')
                self.entryList[0].delete(0, END)
                self.entryList[0].insert(0, inputPath)
                self.entryList[0].config(state='readonly')
                self.path[1] = inputPath
            else:
                if self.entryList[0].get() == '':
                    raise Exception
        except:
            mb.showerror('Error', 'Wrong file path')

    def inputAttachment2(self):
        inputPath = tk.filedialog.askopenfilename(title='Select file')
        try:
            if len(inputPath) != 0:
                self.entryList[1].config(state='normal')
                self.entryList[1].delete(0, END)
                self.entryList[1].insert(0, inputPath)
                self.entryList[1].config(state='readonly')
                self.path[2] = inputPath
            else:
                if self.entryList[1].get() == '':
                    raise Exception
        except:
            mb.showerror('Error', 'Wrong file path')

    def inputAttachment3(self):
        inputPath = tk.filedialog.askopenfilename(title='Select file')
        try:
            if len(inputPath) != 0:
                self.entryList[2].config(state='normal')
                self.entryList[2].delete(0, END)
                self.entryList[2].insert(0, inputPath)
                self.entryList[2].config(state='readonly')
                self.path[3] = inputPath
            else:
                if self.entryList[2].get() == '':
                    raise Exception
        except:
            mb.showerror('Error', 'Wrong file path')

    def inputAttachment4(self):
        inputPath = tk.filedialog.askopenfilename(title='Select file')
        try:
            if len(inputPath) != 0:
                self.entryList[3].config(state='normal')
                self.entryList[3].delete(0, END)
                self.entryList[3].insert(0, inputPath)
                self.entryList[3].config(state='readonly')
                self.path[4] = inputPath
            else:
                if self.entryList[3].get() == '':
                    raise Exception
        except:
            mb.showerror('Error', 'Wrong file path')

    def inputAttachment5(self):
        inputPath = tk.filedialog.askopenfilename(title='Select file')
        try:
            if len(inputPath) != 0:
                self.entryList[4].config(state='normal')
                self.entryList[4].delete(0, END)
                self.entryList[4].insert(0, inputPath)
                self.entryList[4].config(state='readonly')
                self.path[5] = inputPath
            else:
                if self.entryList[4].get() == '':
                    raise Exception
        except:
            mb.showerror('Error', 'Wrong file path')

    def columnHead(self):
        value = self.excelPath.get()
        if value != '':
            if value.endswith('.xlsx'):
                self.df = pd.read_excel(value)
                value = self.df.columns.values.tolist()
                value.insert(0, 'Please select value')
                self.columnName['values'] = list(value)
            else:
                if value.endswith('.csv'):
                    self.df = pd.read_csv(value)
                    value = self.df.columns.values.tolist()
                    value.insert(0, 'Please select value')
                    self.columnName['values'] = list(value)

    def inputExcelFunction(self):
        inputPath = tk.filedialog.askopenfilename(title='Select file', filetypes=(('XLSX Files', '*.xlsx'),
                                                                                  ('CSV Files', '*.csv')))
        try:
            if len(inputPath) != 0:
                self.excelPath.config(state='normal')
                self.excelPath.delete(0, END)
                self.excelPath.insert(0, inputPath)
                self.excelPath.config(state='readonly')
            else:
                if self.excelPath.get() == '':
                    raise Exception
        except:
            mb.showerror('Error', 'Wrong file path or wrong file extention')

    def inputHTMLFunction(self):
        inputPath = tk.filedialog.askopenfilename(title='Select file', filetypes=(('HTM Files', '*.htm'), ))
        try:
            if len(inputPath) != 0:
                self.htmlFilePath.config(state='normal')
                self.htmlFilePath.delete(0, END)
                self.htmlFilePath.insert(0, inputPath)
                self.htmlFilePath.config(state='readonly')
            else:
                if self.htmlFilePath.get() == '':
                    raise Exception
        except:
            mb.showerror('Error', 'Wrong file path or wrong file extention')

    def attachMentButton(self):
        if self.excelPath.get() == '':
            mb.showerror('Error', 'Please add Excel or CSV file.')
            return
        if self.htmlFilePath.get() == '':
            mb.showerror('Error', 'Please add HTML/HTM File')
            return
        if self.subject.get() == '':
            mb.showerror('Error', 'Please add valid Subject')
            return
        if self.columnName.get() == 'Please select value':
            mb.showerror('Error', 'Please select valid column name')
            return
        if self.maximumAttachment.get() == 'Please select value':
            mb.showerror('Error', 'Please select number of attachment')
            return
        if int(self.maximumAttachment.get()) != 0:
            self.withdraw()
            self.top.deiconify()
        else:
            self.withdraw()
            self.sendMailFunction()

    def exitClicked(self):
        self.destroy()


if __name__ == '__main__':
    app = App()
    app.mainloop()
