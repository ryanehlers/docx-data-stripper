import os
from docx import Document
#See https://python-docx.readthedocs.io/en/latest/api/document.html

def copyDoc(doc, path, filename):
    copy_filepath = path + 'stripped/' + filename
    print('Saving updated file to', copy_filepath)
    generateFilePath(copy_filepath)
    doc.save(copy_filepath)
    print('Update successful.')
    return Document(copy_filepath)

def generateFilePath(filepath):
    directory = os.path.dirname(filepath)
    if not os.path.exists(directory):
        os.makedirs(directory)
    return
