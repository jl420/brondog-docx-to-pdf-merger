import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert
import os
from PyPDF2 import PdfWriter

def browse_folder():
    folder_path = filedialog.askdirectory(
        title="Select a folder"
    )
    if folder_path:
        folder_path_var.set(folder_path)

def save_file():
    folder_path = folder_path_var.get()
    pdf_file_name = pdf_file_name_var.get()

    print(folder_path)
    
    convert(folder_path)

    merger = PdfWriter()

    for file in os.listdir(folder_path):
        if file[-4::] == '.pdf':

            full_path = f'{folder_path}\\{file}'

            merger.append(full_path)
            os.remove(full_path)

    merger.write(pdf_file_name)
    merger.close()

# Create the main window
root = tk.Tk()
root.title("BronDog Doc to PDF Merger")

# Create a StringVar to hold the folder path
folder_path_var = tk.StringVar()
pdf_file_name_var = tk.StringVar()

# Create and place widgets
# T = tk.Text(root, height=5, width=50)
# T.pack()
# T.insert(tk.END, "Select the folder which contains the unzipped .docx files")

tk.Label(root, text="Selected Folder:").pack(pady=5)
tk.Entry(root, textvariable=folder_path_var, width=50).pack(pady=5)

tk.Button(root, text="Browse", command=browse_folder).pack(pady=5)

tk.Label(root, text="PDF File Name:").pack(pady=5)
tk.Entry(root, textvariable=pdf_file_name_var, width=50).pack(pady=5)

tk.Button(root, text="Create File", command=save_file).pack(pady=20)

# Start the Tkinter event loop
root.mainloop()