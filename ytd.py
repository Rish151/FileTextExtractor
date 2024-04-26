import os
import docx2txt
import re
import xlsxwriter
from PyPDF2 import PdfReader

def extract_email(text):
    pattern = re.compile(r'[a-zA-Z0-9-\.]+@[a-zA-Z-\.]*\.(com|edu|net)')
    matches = pattern.finditer(text)
    emails = [x.group(0) for x in matches]
    return emails

def extract_phone(text):
    pattern = re.compile(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]')
    matches = pattern.finditer(text)
    phones = [x.group(0) for x in matches]
    return phones

def extract_text_pdf(file_path):
    with open(file_path, 'rb') as file:
        reader = PdfReader(file)
        text = " "
        for page_num in range(len(reader.pages)):
            text += reader.pages[page_num].extract_text()
        return text

def extract_text_docx(file_path):
    return docx2txt.process(file_path)

if __name__ == "__main__":
    folder_path = r"C:\Users\Hp\Desktop\1"
    output_file = "Data.xlsx"

    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, "#")
    worksheet.write(0, 1, "Emails")
    worksheet.write(0, 2, "Phone Numbers")
    worksheet.write(0, 3, "Text")

    row = 1
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if filename.endswith(".pdf"):
            text = extract_text_pdf(file_path)
        elif filename.endswith(".docx"):
            text = extract_text_docx(file_path)
        else:
            continue  # Skip non-PDF and non-DOCX files

        emails = extract_email(text)
        phones = extract_phone(text)


        for email, phone in zip(emails, phones):
            worksheet.write(row, 0, row)
            worksheet.write(row, 1, email)
            worksheet.write(row, 2, phone)
            worksheet.write(row, 3, text)
            row += 1

    workbook.close()
