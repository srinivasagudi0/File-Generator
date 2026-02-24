# imports
from __future__ import annotations
try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4, letter
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from docx import Document as Doc
except ImportError:
    Doc = None
    canvas = None
    A4 = None
    letter = None
    pdfmetrics = None
    TTFont = None




import textwrap
import os



def convert(input_file: str, output_file: str, conversion_type: int) -> None:
    """Direct the conversion process based on the conversion type."""
    if conversion_type == 1:
        txt_to_docx(input_file, output_file)
    elif conversion_type == 2:
        docx_to_txt(input_file, output_file)
    elif conversion_type == 3:
        txt_to_pdf(input_file, output_file)
    else:
        raise ValueError ('It should never reach here since the conversion type ')


def txt_to_docx(input_file: str, output_file: str) -> str:
    """Convert a TXT file to a DOCX file."""
    with open(input_file, 'r') as txt_file:
        content = txt_file.read()
    # Create a new DOCX file and write the content to it
    doc = Doc()
    doc.add_paragraph(content)
    doc.save(output_file)
    return 'Conversion Successful!'

def docx_to_txt(input_file: str, output_file: str) -> str:
    """Convert a DOCX file to a TXT file."""
    with open(input_file, 'r') as file:
        doc = Doc(input_file)
        content = '\n'.join([para.text for para in doc.paragraphs])
    # Write the content to a new TXT file
    with open(output_file, 'w') as txt_file:
        txt_file.write(content)
    return 'Conversion Successful!'

def txt_to_pdf(input_txt: str, output_pdf: str, page_size=A4, font_name='Helvetica', font_size=12, margin=72):
    """
    TXT -> PDF
    """

    if not os.path.isfile(input_txt):
        return f"File not found: {input_txt}"

    page_width, page_height = page_size
    usable_width =  page_width - 2 * margin
    usable_height = page_height - 2 * margin
    line_height = font_size * 1.2
    max_lines_per_page = int(usable_height // line_height)

    c = canvas.Canvas(output_pdf, pagesize=page_size)
    c.setFont(font_name, font_size)

    # heuristic chars per line (approximate: works well for monospaced or typical fonts)
    avg_char_width = pdfmetrics.stringWidth("M", font_name, font_size)
    chars_per_line = max(20, int(usable_width // avg_char_width))

    with open(input_txt, "r", encoding="utf-8") as f:
        lines =[]
        for rae_line in f:
            stripped = rae_line.rstrip("\n")
            if stripped == "":
                lines.append("")  # preserve intentional blank lines
            else:
                wrapped = textwrap.wrap(stripped, width=chars_per_line, replace_whitespace=False)
                if not wrapped:
                    lines.append("")
                else:
                    lines.extend(wrapped)
    page_line_index = 0
    y_start = page_height - margin - font_size

    for i,line in enumerate(lines):
        if page_line_index == max_lines_per_page:
            c.showPage()
            c.setFont(font_name, font_size)
            page_line_index = 0
        x = margin
        y = y_start - page_line_index * line_height
        c.drawString(x, y, line)
        page_line_index +=1
    c.save()
    return 'Conversion Successful!'
