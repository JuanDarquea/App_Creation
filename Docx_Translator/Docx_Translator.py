# Docx_Translator
import os
from tkinter import Tk
from tkinter import filedialog as fd

from docx import Document

# define a variable to select the file to be translated
def select_docx_file():
    """Open a file dialog and return the selected filepath to translate"""
    # create hidden root window
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    # create child window(file dialog)
    file_path = fd.askopenfilename( # assign a variable to open a dialog and select the file
        title="Choose a file to translate",
        filetypes=[
            ("Word Documents", "*.docx"), # shows only .docx files
            ("All FIles", "*.*") # show every type of file
        ],
        initialdir=os.path.expanduser("~") # sets the initial directory to user's home folder
    )
    
    # destroy the root dialog window
    root.destroy

    # Return None instead of empty string for better logic
    return file_path if file_path else None

def file_validation(file_path):
    """Validate if a file path was selected""" 
    if file_path is None:
        print("\nUser closed window before selecting a file.")
        return None
    elif not file_path.endswith(".docx"):
        print("\nThe file selected must be a '.docx' file.")
        return None
    else:
        return file_path
    
def main():
    """Main function to test file selection"""
    print("Select a .docx file to translate...")

    # get file selected from user
    chosen_file = select_docx_file()

    if file_validation(chosen_file) is None:
        return # exit the program if no valid file is selected
    else:
        print(f"\nFile selected to translate: {chosen_file}")


if __name__=="__main__":
    main()
    

# assign the file document path to a variable and store the document to "doc" using module document
#file_path = "C:\\Users\\USUARIO\\OneDrive\\Documentos\\My_Projects\\Projects\\AppTools\\Docx_Translator\\TestDocument.docx"
#doc = Document(file_path)

#extract all paragraphs from the document
#for paragraph in doc.paragraphs:
#    print(paragraph.text)

# get complete text paragraph
#full_text = paragraph.text

# translate text
#translated_text = translate(full_text)

# distribute translated text back into runs keeping the formatting