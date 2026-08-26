import pdfplumber
import re

def read_pdf_regex(file_path):
    all_text = ""
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            all_text += page.extract_text() + "\n"
    
    # Regex pattern:
    # Description (anything before "(Código:")
    # (Código: code)
    # Vl. Total (optional)
    # Qtde.: quantity
    # UN.: unit
    # Vl. Unit.: price
    # Total price
    
    pattern = re.compile(
        r"(?P<descricao>.*?)\s*\(C[óo]digo:\s*(?P<codigo>[^)]+)\s*\).*?Qtde\.:\s*(?P<qtde>[\d,]+)\s*UN:\s*(?P<unidade>\w+)\s*Vl\. Unit\.:\s*(?P<vl_unit>[\d,]+)\s*(?P<vl_total>[\d,]+)",
        re.DOTALL | re.IGNORECASE
    )
    
    matches = pattern.finditer(all_text)
    count = 0
    for m in matches:
        count += 1
        d = m.groupdict()
        desc = d['descricao'].strip().replace('\n', ' ')
        print(f"{count:02d}. {desc} | Qtd: {d['qtde']} | Un: {d['unidade']} | Unit: {d['vl_unit']} | Total: {d['vl_total']}")

if __name__ == "__main__":
    read_pdf_regex("cupons_txt/DOCUMENTO AUXILIAR DA NOTA FISCAL DE CONSUMIDOR ELETRÔNICA.pdf")
