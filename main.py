import glob
from docx import Document
from copydoc import copyDoc
from cleardata import clearData
#See https://python-docx.readthedocs.io/en/latest/api/document.html

def stripdata():
    for file in glob.glob("*.docx"):
        print('Copying', file)
        doc = Document(file)
        clearData(doc)
        copyDoc(file)

    print('Finished.')
