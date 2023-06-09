from docx import Document
from docx.shared import Inches
import os
from natsort import natsorted

def write():
    document = Document()

    document.add_heading('Title', 0)

    p = document.add_paragraph('Course name and code\n')

    p.add_run('Submitted to:\n').bold = True
    p.add_run('your information...\n\n\n ')

    p.add_run('Submitted by:\n').bold = True
    p.add_run('name:\n').bold = True
    p.add_run('ID:\n').bold = True
    p.add_run('Subject:\n').bold = True

    document.add_page_break()

    # assign directory
    directory="input"

    files_list = []
    for filename in os.listdir(directory):
        f = os.path.join(directory, filename)
        files_list.append(f)

	# if you want to sort them
    # files_list.natsorted()
    files_list = natsorted(files_list)

    for i in files_list:
        if i[-2:]=='.c':
            print("Found file -> ", i)

    for i in range(len(files_list)):
        if files_list[i][-2:]=='.c':
            file_content = open(files_list[i])
            document.add_heading(files_list[i][6:-2], level=1)
            document.add_paragraph(file_content.read())
            document.add_page_break() # page break
            # document.add_page_break() # page break for screenshots

    document.save('output.docx')

if __name__ == "__main__":
    write()
