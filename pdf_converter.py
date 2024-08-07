from fpdf import FPDF

def convert_txt_to_pdf(txt_path, pdf_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=9)
    with open(txt_path, 'r') as file:
        for line in file:
            pdf.cell(200, 10, txt=line, ln=True)
    pdf.output(pdf_path)
