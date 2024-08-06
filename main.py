import pytesseract
from pdf2image import convert_from_path
from PIL import Image

# Defina o caminho para o executável do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def ocr_image(image_path):
    return pytesseract.image_to_string(Image.open(image_path))

def ocr_pdf(pdf_path):
    pages = convert_from_path(pdf_path)
    text = ""
    for page in pages:
        text += pytesseract.image_to_string(page)
    return text

if __name__ == "__main__":
    # Teste com uma imagem JPG ou PNG
    print(ocr_image('sample.jpg'))

    # Teste com um PDF
    print(ocr_pdf('sample.pdf'))
