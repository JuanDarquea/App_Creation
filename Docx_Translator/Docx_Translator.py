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
        initialdir=os.path.expanduser("C:\\Users\\USUARIO\\OneDrive\\Documentos\\My_Projects\\Projects\\AppTools\\Docx_Translator") # sets the initial directory to user's home folder
    )
    
    # destroy the root dialog window
    root.destroy

    # Return None instead of empty string for better logic
    return file_path if file_path else None

def file_validation(file_path):
    """Validate if a file path was selected""" 
    if file_path is None:
        return False
    elif not os.path.exists(file_path):
        print(f"\nError!! The file {file_path} selected does not exist.")
        return False
    elif not file_path.endswith(".docx"):
        print("\nError!! The file selected must be a '.docx' file.")
        return False
    else:
        # When file is selected
        print(f"\nFile selected to translate: {file_path}", 
                f"\nPath directory: {os.path.dirname(file_path)}", 
                f"\nFile name: {os.path.basename(file_path)}", 
                f"\nFile size: {os.path.getsize(file_path)} KB\n")
        return True
    
def main():
    """Main function to test file selection"""
    print("Select a .docx file to translate...")

    # get file selected from user
    chosen_file = select_docx_file()

    # Close if cancelled and no file selected
    if not chosen_file:
        print("\nUser closed window before selecting a file. Goodbye!")
        return

    # Validate file
    if file_validation(chosen_file) is None:
        return

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