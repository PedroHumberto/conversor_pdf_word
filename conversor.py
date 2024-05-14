import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from docx import Document
from io import BytesIO

pytesseract.pytesseract.tesseract_cmd = r'D:\\Tesseract\\tesseract.exe'
def pdf_to_word(pdf_path, word_path):
    document = fitz.open(pdf_path)
    word_doc = Document()

    for page_num in range(len(document)):
        page = document[page_num]
        pix = page.get_pixmap(dpi = 300)
        img = Image.open(BytesIO(pix.tobytes()))

        text = pytesseract.image_to_string(img, lang='por')

        word_doc.add_paragraph(text)

    word_doc.save(word_path)

pdf_to_word('rascunho2.pdf', 'saida.docx')
