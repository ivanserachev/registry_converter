from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from logic import *

class Create_form():
    def __init__(self,
                 bank_lst,
                 address_dict,
                 ):
        self.bank_lst=bank_lst
        self.address_dict=address_dict
        self.bank=''
        self.address=''
        self.account=''
        self.path_pdf=''
        self.path_save=''

    def act_bank_get(self, bank_box):
        self.bank=bank_box.get()

    def act_address_get(self, address_box):
        self.address= address_box.get()

    def act_account_get(self, account_box):
        self.account= account_box.get()

    def open_file(self):
        self.path_pdf = filedialog.askopenfilename()
        if self.path_pdf!='':
            messagebox.showinfo('Уведомление', f'Выбран файл {self.path_pdf}')

    def save_folder(self):
        filepath = filedialog.askdirectory()
        self.path_save = filepath + '/'
        if self.path_save!='':
            messagebox.showinfo('Уведомление', f'Выбрана папка {filepath}')

    def convert(self):
        self.act_bank_get(self.bank_box)
        self.act_account_get(self.account_box)
        self.act_address_get(self.address_box)

        if self.path_pdf=='':
            messagebox.showerror(message='Не выбран файл PDF для конвертации')
        elif self.path_save=='':
            messagebox.showerror(message='Не выбрана папка для сохранения файла')
        elif self.bank == '':
            messagebox.showerror(message='Не выбран или не введен банк')
        elif self.address == '':
            messagebox.showerror(message='Не введен адрес банка')
        elif self.account == '':
            messagebox.showerror(message='Не введен номер счета')
        else:
            cv_engine = Cv_processing(export_pdf(path_pdf=self.path_pdf))
            path_save=self.path_save+self.bank.replace('"','')+'_выписка.xlsx'

            if self.bank=='АО "АЛЬФА-БАНК"':
                try:
                    data = cv_engine.alpha_bank()
                except:
                    messagebox.showerror(message='Проверьте выбранный PDF файл и запустите программу заново')

            elif self.bank=='ПАО "БАНК УРАЛСИБ"':
                try:
                    data = cv_engine.uralsib_bank()
                except:
                    messagebox.showerror(message='Проверьте выбранный PDF файл и запустите программу заново')

            import_to_excel(dframe=data, path=path_save)
            messagebox.showinfo('Info', 'Импорт в ' + path_save)

            excel_processing(bank_name=self.bank, address=self.address, account=self.account, path=path_save)
            messagebox.showinfo('Info', 'Обработка файла ' + path_save)





    def refresh_address(self, event):
        self.act_bank_get(self.bank_box)
        try:

            if self.address_dict[self.bank] != self.bank or self.bank == '':
                self.address_box.delete(0, END)
                self.address_box.insert(0, self.address_dict[self.bank])
        except:
            pass


    def create_interface(self):
        app_window = Tk()
        app_window.title('Конвертер банковских выписок из PDF в Excel')
        app_window.geometry('600x200')

        frame = Frame(app_window,
                      padx=20,
                      pady=20)

        frame.pack(expand=True)

        bank_label = Label(frame,
                           text='Выберете или введите название банка')
        bank_label.grid(row=0, column=1)

        self.bank_box = ttk.Combobox(frame,
                                values=self.bank_lst,
                                width=35)
        self.bank_box.grid(row=0, column=2)

        address_label = Label(frame,
                              text='Введите адрес банка')
        address_label.grid(row=1, column=1)

        self.address_box = Entry(frame,
                            width=38
                            )
        self.address_box.grid(row=1, column=2)

        account_label = Label(frame,
                              text='Введите номер счета')
        account_label.grid(row=2, column=1)

        self.account_box = Entry(frame,
                            width=38
                            )
        self.account_box.grid(row=2, column=2)

        pdf_file_label = Label(frame,
                               text='Выберете PDF файл')
        pdf_file_label.grid(row=3, column=1)

        source_btn = Button(frame,
                            text='Выбрать',
                            command=self.open_file)
        source_btn.grid(row=3, column=2)

        excel_file_label = Label(frame,
                                 text='Выберете папку для сохранения Excel файла')
        excel_file_label.grid(row=4, column=1)

        save_btn = Button(frame,
                          text='Выбрать',
                          command=self.save_folder)
        save_btn.grid(row=4, column=2)

        start_btn = Button(frame,
                           text='Конвертировать',
                           command=self.convert)
        start_btn.grid(row=5, column=2)

        app_window.bind('<Button-1>', self.refresh_address)

        app_window.mainloop()



