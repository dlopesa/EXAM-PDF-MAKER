from tkinter import filedialog, Tk, simpledialog, messagebox
from PyPDF2 import PdfMerger
import os
import shutil
import win32com.client

# Function to find Word files in a folder
def find_word_files(folder):
    word_files = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith('.doc') or file.endswith('.docx'):
                word_files.append(os.path.join(root, file))
    return word_files

# Function to convert Word files to PDF
def convert_word_to_pdf(input_file, output_folder):
    input_file = os.path.normpath(input_file)  # Normalize the file path
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = 0
    if input_file.endswith('.doc'):
        pdf_name = input_file.replace('.doc', '.pdf')
    if input_file.endswith('.docx'):
        pdf_name = input_file.replace('.docx', '.pdf')
    print(f"Opening word file: {input_file}")
    doc = word.Documents.Open(input_file)
    doc.SaveAs(pdf_name, FileFormat=17)  # formatType = 17 for pdf
    doc.Close()
    word.Quit()
    print(f"Converted word file: {pdf_name}")
    shutil.move(pdf_name, os.path.join(output_folder, os.path.basename(pdf_name)))
    print(f"Moved Word file to: {os.path.join(output_folder, os.path.basename(pdf_name))}")
# Function to find PowerPoint files in a folder
def find_ppt_files(folder):
    ppt_files = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith('.ppt') or file.endswith('.pptx'):
                ppt_files.append(os.path.join(root, file))
    return ppt_files

# Function to convert PowerPoint files to PDF
def convert_ppt_to_pdf(input_file, output_folder):
    input_file = os.path.normpath(input_file)  # Normalize the file path
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    powerpoint.Visible = 1
    if input_file.endswith('.ppt'):
        pdf_name = input_file.replace('.ppt', '.pdf')
    if input_file.endswith('.pptx'):
        pdf_name = input_file.replace('.pptx', '.pdf')
    print(f"Opening PowerPoint file: {input_file}")
    deck = powerpoint.Presentations.Open(input_file)
    deck.SaveAs(pdf_name, 32)  # formatType = 32 for pdf
    deck.Close()
    powerpoint.Quit()
    print(f"Converted PowerPoint file: {pdf_name}")
    shutil.move(pdf_name, os.path.join(output_folder, os.path.basename(pdf_name)))
    print(f"Moved PowerPoint file to: {os.path.join(output_folder, os.path.basename(pdf_name))}")

# Function to merge PDF files
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

    print(f"Merged file: {output}")
    merger.write(output)
    merger.close()

# Function to ask the user for a processing choice
def ask_user_choice(root):
    message = "Select an option:\n"
    message += "1: Convert PowerPoint files to PDF\n"
    message += "2: Convert PowerPoint files to PDF and merge them\n"
    message += "3: Merge PDF files\n"
    message += "4: Convert Word files to PDF\n"
    message += "5: Convert Word files to PDF and merge them\n"
    return simpledialog.askstring("Input", message, parent=root)

# Function to select a folder using Tkinter file dialog
def select_folder():
    root = Tk()
    root.withdraw()  # Hide the main window
    messagebox.showinfo("Information", "Select a folder to process files in it and its subfolders")
    folder_selected = filedialog.askdirectory()  # Open the dialog to select a folder
    return folder_selected
# Function to process the user's choice
def process_choice(choice, folder):
    if choice == '1':
        ppt_files = find_ppt_files(folder)
        for ppt_file in ppt_files:
            convert_ppt_to_pdf(ppt_file, folder)

    elif choice == '2':
        ppt_files = find_ppt_files(folder)
        for ppt_file in ppt_files:
            convert_ppt_to_pdf(ppt_file, folder)
        pdf_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith('.pdf')]
        merge_pdfs(pdf_files, os.path.join(folder, 'merged.pdf'))

    elif choice == '3':
        pdf_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith('.pdf')]
        merge_pdfs(pdf_files, os.path.join(folder, 'merged.pdf'))

    elif choice == '4':
        word_files = find_word_files(folder)
        for word_file in word_files:
            convert_word_to_pdf(word_file, folder)

    elif choice == '5':
        word_files = find_word_files(folder)
        for word_file in word_files:
            convert_word_to_pdf(word_file, folder)
        pdf_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith('.pdf')]
        merge_pdfs(pdf_files, os.path.join(folder, 'merged.pdf'))
    

# Main function
def main():
    root = Tk()
    root.withdraw()  # Hide the main window
    while True:
        folder = select_folder()
        if not folder:
            print("No folder selected. Exiting program.")
            return

        choice = ask_user_choice(root)
        process_choice(choice, folder)

        quit = simpledialog.askstring("Input", "Do you want to quit? (yes/no): ", parent=root)
        if quit.lower() == 'yes':
            break

if __name__ == "__main__":
    main()
