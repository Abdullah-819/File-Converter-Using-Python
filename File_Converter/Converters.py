# converters.py
import os
from docx2pdf import convert as docx_to_pdf
from pdf2docx import Converter as pdf2docx_conv
from pdf2image import convert_from_path
from pptx import Presentation
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import PyPDF2
from Utils import sanitize_text

# --- DOCX → PDF ---
def convert_docx_to_pdf(input_path, output_path, progress_callback=None):
    docx_to_pdf(input_path, output_path)
    if progress_callback:
        progress_callback(100)

# --- PDF → DOCX ---
def convert_pdf_to_docx(input_path, output_path, progress_callback=None):
    converter = pdf2docx_conv(input_path)
    num_pages = converter.doc.pages

    def update_progress(page):
        if progress_callback:
            percent = int((page / num_pages) * 100)
            progress_callback(percent)

    converter.convert(output_path, start=0, end=num_pages, callback=update_progress)
    converter.close()
    if progress_callback:
        progress_callback(100)

# --- PDF → PPTX ---
def convert_pdf_to_pptx(input_path, output_path, progress_callback=None):
    prs = Presentation()
    images = convert_from_path(input_path)
    total = len(images)
    for idx, img in enumerate(images):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        temp_img = f"temp_page_{idx}.jpg"
        img.save(temp_img)
        slide.shapes.add_picture(temp_img, 0, 0, width=prs.slide_width, height=prs.slide_height)
        os.remove(temp_img)
        if progress_callback:
            progress_callback(int((idx+1)/total*100))
    prs.save(output_path)
    if progress_callback:
        progress_callback(100)

# --- TXT → PDF ---
def convert_txt_to_pdf(input_path, output_path, progress_callback=None):
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(output_path)
    with open(input_path, "r", encoding="utf-8", errors="ignore") as f:
        text = sanitize_text(f.read())
    story = [Paragraph(text.replace("\n", "<br/>"), styles["Normal"])]
    doc.build(story)
    if progress_callback:
        progress_callback(100)

# --- PDF → TXT ---
def convert_pdf_to_txt(input_path, output_path, progress_callback=None):
    reader = PyPDF2.PdfReader(open(input_path, "rb"))
    text = ""
    total = len(reader.pages)
    for idx, page in enumerate(reader.pages):
        page_text = page.extract_text() or ""
        text += sanitize_text(page_text) + "\n"
        if progress_callback:
            progress_callback(int((idx+1)/total*100))
    with open(output_path, "w", encoding="utf-8", errors="ignore") as f:
        f.write(text)
    if progress_callback:
        progress_callback(100)

# --- DOCX → PPTX ---
def convert_docx_to_pptx(input_path, output_path, progress_callback=None):
    from docx import Document
    docx_file = Document(input_path)
    prs = Presentation()
    paragraphs = [p for p in docx_file.paragraphs if p.text.strip()]
    total = len(paragraphs)
    for idx, para in enumerate(paragraphs):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        if slide.placeholders:
            body = slide.placeholders[1]
            body.text_frame.text = sanitize_text(para.text)
        else:
            slide.shapes.add_textbox(0,0,prs.slide_width,prs.slide_height).text = sanitize_text(para.text)
        if progress_callback:
            progress_callback(int((idx+1)/total*100))
    prs.save(output_path)
    if progress_callback:
        progress_callback(100)

# --- PPTX → DOCX ---
def convert_pptx_to_docx(input_path, output_path, progress_callback=None):
    from docx import Document
    prs = Presentation(input_path)
    docx_file = Document()
    total = len(prs.slides)
    for idx, slide in enumerate(prs.slides):
        if slide.shapes.title and slide.shapes.title.text.strip():
            docx_file.add_heading(sanitize_text(slide.shapes.title.text), level=1)
        for shape in slide.shapes:
            if hasattr(shape,"text") and shape.text.strip() and shape is not slide.shapes.title:
                docx_file.add_paragraph(sanitize_text(shape.text))
        docx_file.add_paragraph("\n")
        if progress_callback:
            progress_callback(int((idx+1)/total*100))
    docx_file.save(output_path)
    if progress_callback:
        progress_callback(100)
