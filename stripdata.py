import glob, os, datetime
from docx import Document
from copydoc import copyDoc

blankDate = datetime.datetime(1,1,1)

def stripdata(path):
    for file in glob.glob(path + "*.docx"):
        filename = os.path.basename(file)
        print('Copying', filename)
        doc = Document(file)
        clearDataOnDoc(doc)
        copyDoc(doc, path, filename)

    print('Finished.')

def clearDataOnDoc(doc):
    print('Clearing data on copy.')
    clearAuthor(doc)
    clearCategory(doc)
    clearComments(doc)
    clearContentStatus(doc)
    clearCreated(doc)
    clearIdentifier(doc)
    clearKeywords(doc)
    clearLanguage(doc)
    clearLastModifiedBy(doc)
    clearLastPrinted(doc)
    clearModified(doc)
    clearRevision(doc)
    clearSubject(doc)
    clearTitle(doc)
    clearVersion(doc)
    return

def clearAuthor(doc):
    doc.core_properties.author = ''
    return

def clearCategory(doc):
    doc.core_properties.category = ''
    return

def clearComments(doc):
    doc.core_properties.comments = ''
    return    

def clearContentStatus(doc):
    doc.core_properties.content_status = ''
    return

def clearCreated(doc):
    doc.core_properties.created = blankDate
    return

def clearIdentifier(doc):
    doc.core_properties.identifier = ''
    return

def clearKeywords(doc):
    doc.core_properties.identifier = ''
    return

def clearLanguage(doc):
    doc.core_properties.language = ''
    return

def clearLastModifiedBy(doc):
    doc.core_properties.last_modified_by = ''
    return

def clearLastPrinted(doc):
    doc.core_properties.last_printed = blankDate
    return

def clearModified(doc):
    doc.core_properties.modified = blankDate
    return

def clearRevision(doc):
    doc.core_properties.revision = 1
    return

def clearSubject(doc):
    doc.core_properties.subject = ''
    return

def clearTitle(doc):
    doc.core_properties.title = ''
    return

def clearVersion(doc):
    doc.core_properties.version = ''
    return
    