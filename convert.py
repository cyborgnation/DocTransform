import os
from bs4 import BeautifulSoup
from docx import Document

# Get all html files in current directory
html_files = [f for f in os.listdir(os.getcwd()) if f.endswith('.html')]

for html_file in html_files:
    with open(html_file, "r") as f:
        contents = f.read()

    soup = BeautifulSoup(contents, 'lxml')
    
    doc = Document()
    doc.add_paragraph(soup.get_text())
    doc.save(html_file.replace('.html', '.docx'))
