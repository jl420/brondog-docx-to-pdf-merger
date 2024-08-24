import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert
import os
from PyPDF2 import PdfWriter
import zipfile

class Page(tk.Frame):
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

        self.header_frame = tk.Frame(container)
        self.header_frame.pack(side="top", fill="both", expand=True, padx=10)
        self.create_header()

        self.content_frame = tk.Frame(container)
        self.content_frame.pack(side="top", fill="both", expand=True, padx=10)
        self.create_content()

    def create_header(self):
        tk.Button(self.header_frame, text="Zip Merger", command=self.go_to_zip_merger).pack(side="left", pady=5)
        tk.Button(self.header_frame, text="Folder Merger").pack(side="left", pady=5)
        tk.Label(self.header_frame, text="Folder Merger").pack(side="left", pady=5)

    def create_content(self):
        tk.Label(self.content_frame, text="Selected Folder:").pack(side="top", pady=5)
        tk.Entry(self.content_frame, textvariable=self.folder_path_var, width=30).pack(pady=5)

        tk.Button(self.content_frame, text="Browse", command=self.browse_folder).pack(pady=5)

        tk.Label(self.content_frame, text="PDF File Name:").pack(pady=5)
        tk.Entry(self.content_frame, textvariable=self.pdf_file_name_var, width=30).pack(pady=5)

        tk.Button(self.content_frame, text="Create File", command=self.to_pdf).pack(pady=20)

    
    def go_to_zip_merger(self):
        self.controller.show_frame(Zip_Merger)

    def browse_folder(self):
        folder_path = filedialog.askdirectory(
            title="Select a folder"
        )
        if folder_path:
            self.folder_path_var.set(folder_path)

    def get_files(self, path, file_paths):
        for file in os.listdir(path):
            full_path = f'{path}/{file}'
            if os.path.isdir(full_path):
                file_paths += self.get_files(full_path, [])
            else:
                file_paths.append(full_path)
        
        return file_paths

    def to_pdf(self):
        folder_path = self.folder_path_var.get()

        convert(folder_path)

        merger = PdfWriter()

        for file in os.listdir(folder_path):
            if file[-4::] == '.pdf':

                full_path = f'{folder_path}\\{file}'

                merger.append(full_path)
                os.remove(full_path)

                print(f'{file} : Done')

        merger.write(self.pdf_file_name_var.get())
        merger.close()

        if len(os.listdir(folder_path)) == 0:
            print('No files found')
        print('Done!')

class Zip_Merger(Page):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)

        self.file_path_var = tk.StringVar()
        self.pdf_file_name_var = tk.StringVar()

        container = tk.Frame(self)
        container.pack(fill="both", expand=True)

        self.header_frame = tk.Frame(container)
        self.header_frame.pack(side="top", fill="both", expand=True, padx=10)
        self.create_header()

        self.content_frame = tk.Frame(container)
        self.content_frame.pack(side="top", fill="both", expand=True, padx=10)
        self.create_content()

    def create_header(self):
        tk.Button(self.header_frame, text="Zip Merger").pack(side="left", pady=5)
        tk.Button(self.header_frame, text="Folder Merger", command=self.go_to_folder_merger).pack(side="left", pady=5)
        tk.Label(self.header_frame, text="Zip Merger").pack(side="left", pady=5)

    def create_content(self):
        tk.Label(self.content_frame, text="Selected Zip File:").pack(side="top", pady=5)
        tk.Entry(self.content_frame, textvariable=self.file_path_var, width=30).pack(pady=5)

        tk.Button(self.content_frame, text="Browse", command=self.browse_zip).pack(pady=5)

        tk.Label(self.content_frame, text="PDF File Name:").pack(pady=5)
        tk.Entry(self.content_frame, textvariable=self.pdf_file_name_var, width=30).pack(pady=5)

        tk.Button(self.content_frame, text="Create File", command=self.to_pdf).pack(pady=20)

    def go_to_folder_merger(self):
        self.controller.show_frame(Folder_Merger)

    def browse_zip(self):
        file_path = filedialog.askopenfilename(
            title="Select a zip file"
        )

        if file_path:
            self.file_path_var.set(file_path)

    def unzip(self, file_path):
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall('test')


    def to_pdf(self):
        file_path = self.file_path_var.get()

        self.unzip(file_path)
        # convert(file_path)

        # merger = PdfWriter()

        # for file in os.listdir(folder_path):
        #     if file[-4::] == '.pdf':

        #         full_path = f'{folder_path}\\{file}'

        #         merger.append(full_path)
        #         os.remove(full_path)

        # merger.write(self.pdf_file_name_var.get())
        # merger.close()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DOCX to PDF Merger")
        self.geometry("250x300")
        self.frames = {}
        self.create_frames()

    def create_frames(self):
        for F in (Folder_Merger, Zip_Merger):
            page_name = F.__name__
            frame = F(parent=self, controller=self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        self.show_frame(Folder_Merger)

    def show_frame(self, page_class):
        frame = self.frames[page_class]
        frame.tkraise()

if __name__ == "__main__":
    app = App()
    app.mainloop()