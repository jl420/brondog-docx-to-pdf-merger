from docx2pdf import convert
import os
from PyPDF2 import PdfWriter

folder = 'Docs'

convert(folder)

merger = PdfWriter()

for file in os.listdir(folder):
    if file[-4::] == '.pdf':

        full_path = f'{folder}/{file}'

        merger.append(full_path)
        os.remove(full_path)

merger.write('output.pdf')
merger.close()


        