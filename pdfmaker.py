import os
import tkinter as tk
from tkinter import filedialog
from tkinter import Tk
from comtypes.client import CreateObject
import shutil
from PyPDF2 import PdfMerger

def merge_pdfs(pdf_files, output):
    merger = PdfMerger()

    # Sort the files by their modification time (oldest first)
    pdf_files.sort(key=os.path.getmtime)

    # Check if a file named 'merge.pdf' already exists
    if os.path.exists(output):
        base_name = os.path.splitext(output)[0]
        extension = os.path.splitext(output)[1]
        i = 1
        while os.path.exists(output):
            output = f"{base_name}({i}){extension}"
            i += 1

    for pdf in pdf_files:
        # Ignore any existing 'merge.pdf' files when merging new ones
        if 'merge' not in pdf:
            print(f"Merging file: {pdf}")
            merger.append(pdf)

    print(f"merged file: {output}")
    merger.write(output)
    merger.close()


def select_folder():
    root = Tk()
    root.withdraw()  # Hide the main window
    folder_selected = filedialog.askdirectory()  # Open the dialog to select a folder
    return folder_selected

def find_ppt_files(folder):
    ppt_files = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith(".ppt") or file.endswith(".pptx"):
                ppt_files.append(os.path.join(root, file))
    return ppt_files

def convert_to_pdf(ppt_file, output_folder):
    ppt_file = os.path.normpath(ppt_file)
    if not os.path.exists(ppt_file):
        print(f"File does not exist: {ppt_file}")
        return

    print(f"Converting file: {ppt_file}")
    powerpoint = CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    try:
        deck = powerpoint.Presentations.Open(ppt_file)
    except Exception as e:
        print(f"Failed to open file: {ppt_file}. Error: {e}")
        return

    # Save the PDF in the same directory as the original PowerPoint file
    pdf_file = ppt_file.replace('.pptx', '.pdf').replace('.ppt', '.pdf')
    deck.SaveAs(pdf_file, 32)  # 32 stands for pdf format
    deck.Close()
    powerpoint.Quit()

def ask_user_choice():
    print("1. Convert only")
    print("2. Convert and merge")
    print("3. Merge only")
    choice = input("Enter your choice (1, 2, or 3): ")
    return choice

folder = select_folder()

# Check if the folder exists, create it if it doesn't
if not os.path.exists(folder):
    os.makedirs(folder)

choice = ask_user_choice()

if choice == '1' or choice == '2':
    ppt_files = find_ppt_files(folder)
    for ppt_file in ppt_files:
        convert_to_pdf(ppt_file, folder)
        # After conversion, move the file to the initial folder
        pdf_file = ppt_file.replace('.pptx', '.pdf').replace('.ppt', '.pdf')
        shutil.move(pdf_file, os.path.join(folder, os.path.basename(pdf_file)))

if choice == '2' or choice == '3':
    # After all conversions are complete, merge the PDFs
    pdf_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith('.pdf')]
    merge_pdfs(pdf_files, os.path.join(folder, 'merged.pdf'))