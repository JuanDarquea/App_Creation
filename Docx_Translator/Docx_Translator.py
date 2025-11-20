# Docx_Translator
import os
from tkinter import Tk
from tkinter import filedialog as fd
from dotenv import load_dotenv # to load environment variables from .env file

from docx import Document

# load environment variable from .env file
load_dotenv()

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
        # use the environment variable to set the initial directory and a spare default value
        initialdir=os.getenv("docx_translator_dir", os.getenv("app_tools_dir")) 
    )
    
    # destroy the root dialog window
    root.destroy

    # Return None instead of empty string for better logic
    return file_path if file_path else None

def file_validation(file_path):
    """Validate if a file path was selected""" 
    if file_path is None: # when no file is selected
        return 
    elif not file_path.lower().endswith(".docx"): # validate file extension
        print("\nError!! The file selected must be a '.docx' file.")
        return 
    else:
        try: # validate file existence
            # When file is selected
            print(f"\nFile selected to translate: {file_path}", 
                    f"\nFile path: {os.path.dirname(file_path)}", 
                    f"\nFile name: {os.path.basename(file_path)}", 
                    f"\nFile size: {os.path.getsize(file_path)} KB", sep="")
            return True
        except FileExistsError: # file does not exist
            print(f"\nError!! The file {file_path} selected does not exist.")
            return 
        except Exception as e: # other errors
            print(f"\nError validating the file: {e}")
        return

def read_document(file_path):
    """Read the .docx file and return it as an object"""    
    selected_document = Document(file_path)
    doc = [] # create empty list to store paragraphs

    # extract all text paragraphs from the document
    try: 
        print("\nReading document content...")
        for paragraph in selected_document.paragraphs:
#            if paragraph.text.strip() != "": # skip empty paragraphs
            doc.append(paragraph.text)
        # print document content read success message
        print("\nDocument content read successfully.")
    except Exception as e: # handle errors while reading document
        print(f"Error reading the document: {e}")
        return

    # print paragraph count
    total = len(selected_document.paragraphs)
    print(f"Paragraph count: {total}\n")

    # print first paragraph as test
#    print("First paragraph: \n", 
#          doc[0])  # print first paragraph of document
#    print()

    # print all paragraphs with a paragraph index as test
    for index, paragraph in enumerate(doc):
        if paragraph.strip() != "": # skip empty paragraphs
            print(index + 1, 
                  paragraph, selected_document.paragraphs[index].style.name, 
                  f"{len(selected_document.paragraphs[index].text)} characters", 
                  sep = " - ")
#            print(f"P{index + 1}: {paragraph}") # alternative print format
        else:
            print(index +1, 
                  "<Empty paragraph>", 
                  selected_document.paragraphs[index].style.name, 
                  f"{len(selected_document.paragraphs[index].text)} characters", 
                  sep = " - ")
#            index - 1 # do not count empty paragraphs

def main():
    """Main function to test file selection"""
    print("Select a .docx file to translate...")

    # get file selected from user
    chosen_file = select_docx_file()

    # Close if cancelled and no file selected
    if not chosen_file:
        print("\nUser closed the window before selecting a file.", 
              "\nGoodbye!")
        return

    # Validate file
    if file_validation(chosen_file) is None:
        return
    # Read document
    read_document(chosen_file)

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