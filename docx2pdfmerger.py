import tkinter as tk
import tkinter.font as tkFont
from tkinter import filedialog, messagebox, ttk
from docx2pdf import convert
import os
from PyPDF2 import PdfWriter, PdfReader
import zipfile
import shutil
import datetime

class Functions:
    def search_for_dirs(path):
        dirs = [path]
        for item in os.listdir(path):
            full_path = f'{path}/{item}'
            if os.path.isdir(full_path) and '.docx' in [item[-5::] for item in os.listdir(full_path)]:
                dirs += Functions.search_for_dirs(full_path)
        
        return dirs
    
    def search_for_docx(folder_path):
        dirs = Functions.search_for_dirs(folder_path)
        docx = []

        for dir in dirs:
            for file in os.listdir(dir):
                if file[-5::] == '.docx':
                    docx.append(f'{dir}/{file}')

        return docx

    def merge(folder_path, pdf_file_name, docx_selection):

        temp_dir = 'temp'

        os.mkdir(temp_dir)
        docs = Functions.search_for_docx(folder_path)
        
        for doc in docs:
            if os.path.basename(doc) in docx_selection:
                shutil.copyfile(doc, f'{temp_dir}/{os.path.basename(doc)}')

        Functions.print_status_msg('Converting . . .')
        convert(temp_dir)

        merger = PdfWriter()

        for file in os.listdir(temp_dir):
            if file[-4::] == '.pdf':
                full_path = f'{temp_dir}/{file}'

                pdf_length = 0
                with open(full_path, 'rb') as f:
                    reader = PdfReader(f)
                    
                    pdf_length = len(reader.pages)

                merger.append(full_path)
                if pdf_length % 2 == 1:
                    merger.add_blank_page()

                Functions.print_status_msg(f'Merged : {file}')
        
        Functions.print_status_msg(f'Done! > {pdf_file_name}')
        messagebox.showinfo("Merge", "Merge complete")
        merger.write(pdf_file_name)
        merger.close()

        shutil.rmtree(temp_dir)

    def print_status_msg(msg):
        print(f'[{str(datetime.datetime.now())[0:19]}] {msg}')

    def check_pdf_file_name(pdf_file_name):
        if pdf_file_name[-4::] != '.pdf':
            pdf_file_name += '.pdf'

        return pdf_file_name

class Page(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

class Folder_Merger(Page):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)

        self.folder_path_var = tk.StringVar()
        self.pdf_file_name_var = tk.StringVar()

        container = tk.Frame(self)
        container.pack(fill="both", expand=True)

        file_dir_scrollbar_y = tk.Scrollbar(container, orient='vertical')
        file_dir_scrollbar_y.pack(side='right', fill='y')
        file_dir_scrollbar_x = tk.Scrollbar(container, orient='horizontal')
        file_dir_scrollbar_x.pack(side='bottom', fill='x')
        
        self.file_dir = tk.Listbox(container, selectmode='multiple', yscrollcommand=file_dir_scrollbar_y.set, xscrollcommand=file_dir_scrollbar_x.set)
        self.file_dir.pack(side='right', fill='both')

        file_dir_scrollbar_y.config(command=self.file_dir.yview)
        file_dir_scrollbar_x.config(command=self.file_dir.xview)

        self.content_frame = tk.Frame(container)
        self.content_frame.pack(side="top", fill="both", expand=True, padx=10)
        self.create_content()

    def create_content(self):
        tk.Label(self.content_frame, text="Selected Folder:").pack(side="top", pady=5)
        tk.Entry(self.content_frame, textvariable=self.folder_path_var, width=30).pack(pady=5)

        tk.Button(self.content_frame, text="Browse", command=self.browse_folder).pack(pady=5)

        tk.Label(self.content_frame, text="PDF File Name:").pack(pady=5)
        tk.Entry(self.content_frame, textvariable=self.pdf_file_name_var, width=30).pack(pady=5)

        tk.Button(self.content_frame, text="Create File", command=self.to_pdf).pack(pady=20)

    def browse_folder(self):
        folder_path = filedialog.askdirectory(
            title="Select a folder"
        )
        if folder_path:
            self.folder_path_var.set(folder_path)

        self.file_dir.delete(0, 'end')
        for file in Functions.search_for_docx(folder_path):
            self.file_dir.insert('end', f'\'{os.path.basename(file)}\' ({file})')
            self.file_dir.select_set(0, 'end')

    def to_pdf(self):
        folder_path = self.folder_path_var.get()
        pdf_file_name = self.pdf_file_name_var.get()

        pdf_file_name = Functions.check_pdf_file_name(pdf_file_name)

        docx_selection = []
        for i in self.file_dir.curselection():
            docx = self.file_dir.get(i)
            docx_selection.append(docx[1:docx.find("'", 1)])

        Functions.merge(folder_path, pdf_file_name, docx_selection)



class Zip_Merger(Page):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)

        self.file_path_var = tk.StringVar()
        self.pdf_file_name_var = tk.StringVar()

        container = tk.Frame(self)
        container.pack(fill="both", expand=True)

        file_dir_scrollbar_y = tk.Scrollbar(container, orient='vertical')
        file_dir_scrollbar_y.pack(side='right', fill='y')
        file_dir_scrollbar_x = tk.Scrollbar(container, orient='horizontal')
        file_dir_scrollbar_x.pack(side='bottom', fill='x')
        
        self.file_dir = tk.Listbox(container, selectmode='multiple', yscrollcommand=file_dir_scrollbar_y.set, xscrollcommand=file_dir_scrollbar_x.set)
        self.file_dir.pack(side='right', fill='both')

        file_dir_scrollbar_y.config(command=self.file_dir.yview)
        file_dir_scrollbar_x.config(command=self.file_dir.xview)

        self.content_frame = tk.Frame(container)
        self.content_frame.pack(side="top", fill="both", expand=True, padx=10)
        self.create_content()

    def create_content(self):
        tk.Label(self.content_frame, text="Selected Zip File:").pack(side="top", pady=5)
        tk.Entry(self.content_frame, textvariable=self.file_path_var, width=30).pack(pady=5)

        tk.Button(self.content_frame, text="Browse", command=self.browse_zip).pack(pady=5)

        tk.Label(self.content_frame, text="PDF File Name:").pack(pady=5)
        tk.Entry(self.content_frame, textvariable=self.pdf_file_name_var, width=30).pack(pady=5)

        tk.Button(self.content_frame, text="Create File", command=self.to_pdf).pack(pady=20)

    def browse_zip(self):
        file_path = filedialog.askopenfilename(
            title="Select a zip file"
        )

        if file_path:
            self.file_path_var.set(file_path)

        temp_dir_name = 'unzipped'
        self.unzip(file_path, temp_dir_name)

        self.file_dir.delete(0, 'end')
        for file in Functions.search_for_docx(temp_dir_name):
            self.file_dir.insert('end', f'\'{os.path.basename(file)}\' ({file})')
            self.file_dir.select_set(0, 'end')
        shutil.rmtree(temp_dir_name)

    def unzip(self, file_path, temp_dir_name):
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir_name)

    def to_pdf(self):
        file_path = self.file_path_var.get()
        pdf_file_name = self.pdf_file_name_var.get()
        temp_dir_name = 'unzipped'

        pdf_file_name = Functions.check_pdf_file_name(pdf_file_name)

        self.unzip(file_path, temp_dir_name)

        docx_selection = []
        for i in self.file_dir.curselection():
            docx = self.file_dir.get(i)
            docx_selection.append(docx[1:docx.find("'", 1)])
        
        Functions.merge(temp_dir_name, pdf_file_name, docx_selection)

        shutil.rmtree(temp_dir_name)

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("DOCX to PDF Merger")
        self.geometry("560x520")
        try:
            self.iconbitmap('pdf.ico')
        except tk.TclError:
            Functions.print_status_msg('Warning: Ico not loaded')
        self.protocol('WM_DELETE_WINDOW', self.on_close)

        default_font = tkFont.nametofont('TkDefaultFont')
        default_font.configure(size=16)
        self.option_add('*Font', default_font)

        tabControl = ttk.Notebook(self)
        folder_frame = Folder_Merger(self, self)
        zip_frame = Zip_Merger(self, self)
        tabControl.add(folder_frame, text='Folder Merger')
        tabControl.add(zip_frame, text='Zip Merger')
        tabControl.pack(expand=1, fill='both')

    def on_close(self):
        print("closed")
        self.destroy()

if __name__ == "__main__":
    app = App()
    app.mainloop()