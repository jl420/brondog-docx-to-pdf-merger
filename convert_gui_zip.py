import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert
import os
from PyPDF2 import PdfWriter
import zipfile

def browse_file():
    zip_file = filedialog.askopenfile(
        title="Select a zipfile"
    )
    if zip_file:
        zip_path_var.set(zip_file)

def save_file():
    zip_path = zip_path_var.get()
    pdf_file_name = pdf_file_name_var.get()

    
    
    convert(zip_path)

    merger = PdfWriter()

    for file in os.listdir(zip_path):
        if file[-4::] == '.pdf':

            full_path = f'{zip_path}\\{file}'

            merger.append(full_path)
            os.remove(full_path)

    merger.write(pdf_file_name)
    merger.close()

# Create the main window
root = tk.Tk()
root.title("BronDog Zip to PDF Merger")

# Create a StringVar to hold the folder path
zip_path_var = tk.StringVar()
pdf_file_name_var = tk.StringVar()

# Create and place widgets
T = tk.Text(root, height=5, width=50)
T.pack()
T.insert(tk.END, "Select the zip file which contains the downlaoded .docx files")

tk.Label(root, text="Selected Zip File:").pack(pady=5)
tk.Entry(root, textvariable=zip_path_var, width=50).pack(pady=5)

tk.Button(root, text="Browse", command=browse_file).pack(pady=5)

tk.Label(root, text="PDF File Name:").pack(pady=5)
tk.Entry(root, textvariable=pdf_file_name_var, width=50).pack(pady=5)

tk.Button(root, text="Create File", command=save_file).pack(pady=20)

# Start the Tkinter event loop
root.mainloop()