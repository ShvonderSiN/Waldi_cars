from tkinter.ttk import Frame, Label, Button, Entry
from tkinter import Tk, filedialog, Text, Menu, Toplevel, messagebox, PhotoImage, StringVar
from tkinter.constants import WORD, FLAT, END, CENTER, DISABLED, ACTIVE, NORMAL
from openpyxl import load_workbook
import sys
import os

__author__ = "Sergei Shekin"
__license__ = "GPL"
__version__ = "1.0.1"
__maintainer__ = "Sergei Shekin"
__email__ = "shekin.sergey@yandex.com"

basedir = os.path.abspath(os.path.dirname(__file__))


def center(window, dvx: float = 2.5, dvy: float = 2.5) -> None:
    """Opening window on center of the screen

    Args:
        dvx ():
        dvy ():
        window (tkTk): parent window of tk.Tk application
    """
    x = (window.winfo_screenwidth() - window.winfo_reqwidth()) / dvx
    y = (window.winfo_screenheight() - window.winfo_reqheight()) / dvy
    window.wm_geometry("+%d+%d" % (x, y))


def warning_box(text) -> None:
    """Warning window with Error message

    Args:
        text ([type]): Text provided to the message box.
    """
    messagebox.showinfo(title='Message', message=text)


class Popup(Toplevel):
    """Window for "About" information

    Args:
        Toplevel ([type]): child of main class window
    """

    def __init__(self, parent, text, title) -> None:

        self.text = text
        design = f'Version {__version__}\nCreated by {__maintainer__}\nemail: {__email__}'
        super().__init__(parent)
        center(self, dvx=2.3, dvy=2.3)
        self.label = Label(self, text=self.text, padding=15, font=('Arial', 10, 'normal'), width=50, wraplength=600)
        self.button = Button(self, text="Close", command=self.destroy)
        self.attributes('-topmost', True)  # On top over all windows

        self.title(title)
        self.label.pack(side='top', fill='y')
        if 'about' in title.lower():
            self.design = Label(self, text=design, justify='center', font=('Arial', 8, 'normal'), state="readonly")
            self.design.pack(side='top', fill='y')
        self.button.pack(fill='both', side='bottom')


class CarPartsApp(Tk):
    """Main class for main window

    Args:
        Tk ([type]): parent
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        icon = os.path.join(basedir, 'img', 'dodge_logo_2948.gif')
        self.iconphoto(True, PhotoImage(file=icon))
        self.title('FlashPCM')
        center(self)
        self.resizable(width=False, height=False)
        self.menu_ui()
        # self.geometry('900x200')
        for index in range(11):
            self.grid_rowconfigure(index, weight=1)
            self.grid_columnconfigure(index, weight=0)

        self.frameBottom = Frame(self)
        self.but_open_file = Button(self, text='Read file', command=self.open_file, state=ACTIVE)
        self.but_open_file.grid(row=0, column=0, columnspan=6, sticky='WE')
        self.search_field = Button(self, state=DISABLED)
        self.search_field.grid(row=0, column=6, columnspan=6, sticky='WE')

        self.showFrame = Frame(self)
        self.showFrame.grid(row=1, rowspan=5, column=0, columnspan=12, sticky='NSWE')
        self.textFile = Text(self.showFrame,
                             relief=FLAT,
                             height=20,
                             width=80,
                             font=('Arial', 8, 'normal'),
                             wrap=WORD,
                             exportselection=True,
                             state=DISABLED)
        self.textFile.grid(row=0, column=0, columnspan=12, rowspan=5, sticky='NSEW')

        self.butExit = Button(self, text='Exit', command=self.exit_app)
        self.butExit.grid(row=10, column=0, columnspan=12, sticky='WE')

        self.file_name = 'Chrysler PCM database.xlsx'

    def menu_ui(self):
        main_menu = Menu(self)
        self.config(menu=main_menu)
        file_menu = Menu(main_menu, tearoff=0)
        file_menu.add_command(label='Read file', command=self.open_file)
        file_menu.add_command(label='Exit', command=self.exit_app)

        about_text = '''
        The principle of operation is simple:
        1. Open the source exel file,
        2. Type in search field what you want
        3. Press <Enter> or <Return>
        '''
        contacts_text = '''email: flashpcm@icloud.com\nPhone: +48501937974'''
        main_menu.add_cascade(label='File', menu=file_menu)
        main_menu.add_command(label='Contacts', command=lambda: Popup(self, text=contacts_text, title='Contacts'))
        main_menu.add_command(label='About', command=lambda: Popup(self, text=about_text, title='About'))

    def open_file(self) -> (str, None):
        # self.search_field = Button(self, text='Loading...', state=DISABLED)
        # self.search_field.grid(row=0, column=6, columnspan=6, sticky='WE')
        try:
            ext = os.path.splitext(self.file_name)[1].lower()
            if 'xls' in ext:
                workbook = load_workbook(self.file_name)
                worksheet = workbook.active
                var = StringVar()
                self.search_field = Entry(self, textvariable=var, takefocus=True, justify=CENTER)
                self.search_field.bind('<Return>', lambda event, x=worksheet: self.get_much_rows(x, looking=var.get()))
                self.search_field.grid(row=0, column=6, columnspan=6, sticky='WE')
            else:
                warning_box(text=f'Error: file not defined or incorrect format\n \
                Make sure the file must be beside this software')
        except Exception as e:
            warning_box(text=f'Error: wrong file\n{e}')

    def get_much_rows(self, worksheet, looking: str = ''):
        self.textFile.delete(0.1, END)
        count = 0
        rows = []
        try:
            for row in worksheet.rows:
                for col in row:
                    value = str(col.value)
                    if looking in value:
                        rows.append(row)
                    continue
            strings = ''.strip()
            if rows:
                for row in rows:
                    strings += '#' * 50 + '\n'
                    for col in row:
                        value = str(col.value)
                        strings += value + '\n' + '-' * 25 + '\n'
                    count += 1
            text = f'Found matches: {len(rows)}\n\n{strings}'
            self.show_in_text_field(text=text)
        except Exception as e:
            warning_box(text=f'Error: file not defined or incorrect format:\n {e}')

    def show_in_text_field(self, text: str = ''):
        self.textFile.delete(0.1, END)
        if text:
            self.textFile['state'] = NORMAL
            self.textFile.insert(0.1, text)
            self.textFile['state'] = DISABLED
        else:
            pass

    def start_app(self):
        self.mainloop()

    def exit_app(self):
        self.destroy()
        sys.exit(0)


if __name__ == "__main__":
    a = CarPartsApp()
    a.start_app()
