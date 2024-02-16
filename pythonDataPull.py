import PyPDF2
from openpyxl import Workbook

def extract_pdf_codes(pdf_file):
    codes = []
    with open(pdf_file, "rb") as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
            
            codes.append(text)

        return codes
    
def save_codes_xl(codes, excel_file):
    wb = Workbook()
    ws = wb.active
    ws.append(["Code"]) # Header

    for code in codes:
        ws.append([code])
    wb.save(excel_file)

def main():
    pdf_file = "hisense.pdf" # PDF file
    excel_file= "demo.xlsx"
    codes = extract_pdf_codes(pdf_file)
    save_codes_xl(codes, excel_file)
    print("PDF codes saved successfully")

if __name__ == "__main__":
    main()
