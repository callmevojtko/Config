from pdfminer.high_level import extract_text

def extract_text_from_pdf(pdf_path):
    # Extract text from the PDF file
    text = extract_text(pdf_path)
    return text

# Example usage
pdf_path = "Misc/Config_Sheets/scan_bbossler_2024-06-20-15-22-11.pdf"
extracted_text = extract_text_from_pdf(pdf_path)
print(extracted_text)
