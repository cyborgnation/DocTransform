import os
from bs4 import BeautifulSoup
from docx import Document

# Get all htm files in current directory
htm_files = [f for f in os.listdir(os.getcwd()) if f.endswith('.htm')]

for htm_file in htm_files:
    with open(htm_file, "r") as f:
        contents = f.read()

    soup = BeautifulSoup(contents, 'lxml')
    
    doc = Document()
    doc.add_paragraph(soup.get_text())
    doc.save(htm_file.replace('.htm', '.docx'))
