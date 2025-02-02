import os
import win32com.client
import tkinter as tk
from tkinter import filedialog
from pypdf import PdfReader, PdfWriter, PageObject

# Function to open a dialog box to select a folder
def select_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    folder_selected = filedialog.askdirectory()
    return folder_selected


def merge_pdf_pages(input_pdf, output_pdf, right_margin=50, left_margin = 50):
    # Create a PDF reader object
    pdf_reader = PdfReader(input_pdf)
    # Create a PDF writer object
    pdf_writer = PdfWriter()

    # Calculate total height and maximum width of the new single page PDF
    total_height = sum(page.mediabox.height for page in pdf_reader.pages)
    max_width = max(page.mediabox.width for page in pdf_reader.pages)

    # Adjust the width to include the right margin
    adjusted_width = left_margin + max_width + right_margin

    # Create a new blank page to add all other pages to
    merged_page = PageObject.create_blank_page(width=adjusted_width, height=total_height)

    # Iterate through each page and combine it on the merged page
    current_height = total_height
    for page in pdf_reader.pages:
        page_height = page.mediabox.height
        current_height -= page_height
        merged_page.merge_translated_page(page, left_margin, current_height)

    # Add the merged page to the writer
    pdf_writer.add_page(merged_page)

    # Write the merged PDF to a file
    with open(output_pdf, 'wb') as f:
        pdf_writer.write(f)

    print(f"Pages from {input_pdf} have been merged into one page in {output_pdf}, with a right margin of {right_margin} units.")

input_folder_path_old = select_folder()

input_folder_path = input_folder_path_old.replace("/", "\\")

if input_folder_path:
    output_folder_path = os.path.join(input_folder_path, "ConvertedMarg")

    # Create output directory if it doesn't exist
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    # Iterate over files in the input folder
    for file_name in os.listdir(input_folder_path):
        input_file_path = os.path.join(input_folder_path, file_name)
        file_base_name, file_extension = os.path.splitext(file_name)
        output_file_path = os.path.join(output_folder_path, file_base_name + "marg.pdf")

        if file_extension.lower().endswith(".pdf"):
            merge_pdf_pages(input_file_path, output_file_path, right_margin=600, left_margin = 110)