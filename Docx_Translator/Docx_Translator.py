# Docx_Translator
from docx import Document

# assign the file document path to a variable and store the document to "doc" using module document
file_path = "C:\\Users\\USUARIO\\OneDrive\\Documentos\\My_Projects\\Projects\\AppTools\\Docx_Translator\\TestDocument.docx"
doc = Document(file_path)

# extract all paragraphs from the document
for paragraph in doc.paragraphs:
    print(paragraph.text)

# get complete text paragraph
full_text = paragraph.text

# translate text
translated_text = translate(full_text)

# distribute translated text back into runs keeping the formatting