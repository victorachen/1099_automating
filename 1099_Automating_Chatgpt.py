import pandas as pd
import io
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def map_llc(parkname):
    d = {'Hitching Post':'Hitching Post Mobile Home Park, LLC',
    'Crestview':'Yucaipa Crestview, LLC',
    'Westwind':'Yucaipa Westwind Estates, LLC',
    'Holiday':'Holiday Rancho Park, LLC',
    'Wishing Well':'Wishing Well Mobile Home Park, LLC',
    'Patrician':'Patrician Mobile Home Park',
    'Mount Vista':'Mount Vista, LLC',
    'Jian Personal':'Jian Chen',
    'Banning':'Banning Wilson Gardens, LLC'
    }
    for i in d:
        if i in parkname:
            return d[i]
    return parkname

def map_data_to_existing_pdf(person, llc, payer_tin, recipient_tin, street, city_state_zip, compensation, template_pdf_path, output_path):
    # Read the existing PDF template
    existing_pdf = PdfReader(template_pdf_path)

    # Create a new PDF to write the modifications
    output_pdf = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        # Create a canvas for drawing on the PDF
        packet = io.BytesIO()

        can = canvas.Canvas(packet, pagesize=letter)
        llc = map_llc(llc)

        # Add the information to the existing PDF
        can.drawString(58, 722, f"{llc}")
        can.drawString(58,706,"11034 Deer Canyon Dr")
        can.drawString(58,690, "Rancho Cucamonga, CA 91737")
        can.drawString(58, 635, f"{person}")
        can.drawString(58, 662, f"{payer_tin}")
        can.drawString(175, 662, f"{recipient_tin}")
        can.drawString(58, 603, f"{street}")
        can.drawString(58, 579, f"{city_state_zip}")
        can.drawString(320, 662, f"{compensation}"+f"0")
        can.drawString(431, 690, "23")

        # Save the canvas to the packet
        can.save()
        packet.seek(0)

        # Merge the template PDF and the filled form fields
        new_pdf = PdfReader(packet)
        existing_page = existing_pdf.pages[page_num]
        existing_page.merge_page(new_pdf.pages[0])
        output_pdf.add_page(existing_page)

    # Save the modified PDF with mapped data
    output_filename = f"{output_path}\\{person}_mapped_template.pdf"
    with open(output_filename, 'wb') as output_file:
        output_pdf.write(output_file)

def main():
    # Read the Excel file into a DataFrame
    excel_file = r'C:\Users\Lenovo\Documents\pycharmprojects\1099_automating\input_template.xlsx'
    df = pd.read_excel(excel_file)

    # Path to existing PDF template
    template_pdf_path = r'C:\Users\Lenovo\Documents\pycharmprojects\1099_automating\empty_template.pdf'

    # Output path for saving PDFs
    output_path = r'C:\Users\Lenovo\Documents\pycharmprojects\1099_automating\chatgpt'

    # Iterate through rows and map data onto existing PDF for each
    for index, row in df.iterrows():
        person, llc, payer_tin, recipient_tin, street, city_state_zip, compensation = row
        map_data_to_existing_pdf(person, llc, payer_tin, recipient_tin, street, city_state_zip, compensation, template_pdf_path, output_path)

if __name__ == "__main__":
    main()

#Create new folder with only the third pages of everything
import os
import shutil
from PyPDF2 import PdfReader, PdfWriter

def extract_third_page(input_pdf_path, output_pdf_path):
    with open(input_pdf_path, 'rb') as input_file:
        pdf_reader = PdfReader(input_file)
        pdf_writer = PdfWriter()

        third_page = pdf_reader.pages[2]  # 0-based index, so the third page is index 2
        pdf_writer.add_page(third_page)

        with open(output_pdf_path, 'wb') as output_file:
            pdf_writer.write(output_file)

def main():
    # Set up paths
    folder1_path = r'C:\Users\Lenovo\Documents\pycharmprojects\1099_automating\chatgpt'
    folder2_path = r'C:\Users\Lenovo\Documents\pycharmprojects\1099_automating\chatgpt_foremployees'

    # Create Folder 2 if it doesn't exist
    if not os.path.exists(folder2_path):
        os.makedirs(folder2_path)

    # Iterate through each PDF in Folder 1
    for pdf_file in os.listdir(folder1_path):
        if pdf_file.endswith('.pdf'):
            input_pdf_path = os.path.join(folder1_path, pdf_file)
            output_pdf_path = os.path.join(folder2_path, f'third_page_{pdf_file}')

            # Extract the third page and save it to Folder 2
            extract_third_page(input_pdf_path, output_pdf_path)

    print("Extraction completed.")

if __name__ == "__main__":
    main()

#merge the third page of everyone's 1099

import os
from PyPDF2 import PdfMerger

def merge_pdfs(input_folder, output_path):
    pdf_merger = PdfMerger()

    # Iterate through each PDF in the input folder
    for pdf_file in os.listdir(input_folder):
        if pdf_file.endswith('.pdf'):
            pdf_path = os.path.join(input_folder, pdf_file)
            pdf_merger.append(pdf_path)

    # Write the merged PDF to the output path
    with open(output_path, 'wb') as output_file:
        pdf_merger.write(output_file)

    print("Merging completed.")

if __name__ == "__main__":
    # Set up paths
    input_folder_path = r'C:\Users\Lenovo\Documents\pycharmprojects\1099_automating\chatgpt_foremployees'
    output_pdf_path = r'C:\Users\Lenovo\Documents\pycharmprojects\1099_automating\output_merged.pdf'


    # Merge PDFs from the input folder
    merge_pdfs(input_folder_path, output_pdf_path)

