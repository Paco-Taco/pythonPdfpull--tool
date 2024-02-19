from pypdf import PdfReader
from openpyxl import Workbook
import re

def extract_pdf_codes(pdf_file):
    codes = []
    with open(pdf_file, "rb") as file:
        pdf_reader = PdfReader(file)
        startOn = int(input("Start on page: "))
        if startOn > 0: 
            for page_num in range(startOn - 1, len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                matches = re.findall(r'\b\d+\s*•\s*\d{7}\b', text) # Checar si todos los códigos de parte tengan la misma extensión
                codes.extend(matches)
                print(matches)
            return codes

        else: 
            print("Invalid page number. Try again") 
            exit()
    
def save_codes_xl(codes, excel_file):
    wb = Workbook()
    ws = wb.active

    headers = ["Part No.", "Kgs", "Rate No."]    # Header
    ws.append(headers) 

    for code in codes:
        ws.append([code])
    wb.save(excel_file)

def main():
    pdf_file = "hisense.pdf" # PDF file
    excel_file= "demo.xlsx" # Output .xlsx file
    codes = extract_pdf_codes(pdf_file)
    save_codes_xl(codes, excel_file)
    print("PDF codes saved successfully")

main()
