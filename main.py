from datetime import datetime
from docx import Document
from docx.shared import Inches
import os
from natsort import natsorted

# reading ignore files
with open("blacklist.txt") as f:
    blacklisted_files = ["input/" + blacklist.removesuffix("\n") for blacklist in f.readlines()]
    blacklisted_files = list(set(filter(lambda x: '#' not in  x, blacklisted_files)))
files = []

def getfilesinfolder(basepath = "input"):    
    if not os.path.exists(basepath):
        print(f"Error: Directory '{basepath}' does not exist. Creating it now.")
        print("Please put your codes on the 'input' directory, and re-run the script to use!")
        os.makedirs(basepath)
    for filename in os.listdir(basepath):
        f = os.path.join(basepath,filename)
        if f not in blacklisted_files:
            if os.path.isdir(f):
                getfilesinfolder(f)
            else:
                files.append(f)
    return files

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

    # does dig in deeper
    files_list = getfilesinfolder()


	# if you want to sort them
    files_list = natsorted(files_list)
        
    # loop through the files
    for i in range(len(files_list)):
        # with open it can safely close the open() function
        with open(files_list[i]) as thisfile:
            document.add_heading(files_list[i].replace("input/",""), level=1)
            document.add_paragraph(thisfile.read(),)
            document.add_page_break() 

    # get time as filename
    now = datetime.now()
    outputfile = f"output {now}.docx"
    document.save(outputfile)
    print(f"saved in file name {outputfile}")

if __name__ == "__main__":
    write()
