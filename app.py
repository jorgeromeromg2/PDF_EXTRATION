import os
import time
import shutil
from flask import Flask, request, send_file, render_template
import pytesseract
from pytesseract import TesseractNotFoundError
from pdf2image import convert_from_path
from PIL import Image
import pandas as pd
import tempfile
import re

# Inicializando Flask
app = Flask(__name__)

# Caminho para o Poppler
poppler_path = r"C:\Program Files\poppler\Library\bin"

# Configurando o caminho do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def safe_remove(path, retries=5, delay=0.1):
    for i in range(retries):
        try:
            if os.path.isfile(path):
                os.remove(path)
            elif os.path.isdir(path):
                shutil.rmtree(path)
            return
        except PermissionError:
            if i < retries - 1:
                time.sleep(delay)
            else:
                print(f"Failed to remove {path} after {retries} attempts")

def extract_text_from_image(image_path):
    try:
        return pytesseract.image_to_string(Image.open(image_path))
    except TesseractNotFoundError:
        print("Tesseract not found. Make sure it's installed and in your PATH.")
        return "Error: Tesseract not found"
    except Exception as e:
        print(f"Error processing image: {str(e)}")
        return f"Error: {str(e)}"

def extract_text_from_pdf(pdf_path, max_retries=3, delay=1):
    temp_dir = tempfile.mkdtemp()
    try:
        for attempt in range(max_retries):
            try:
                images = convert_from_path(pdf_path, poppler_path=poppler_path, output_folder=temp_dir)
                text = ''
                for image in images:
                    try:
                        text += pytesseract.image_to_string(image)
                    except TesseractNotFoundError:
                        print("Tesseract not found. Make sure it's installed and in your PATH.")
                        return "Error: Tesseract not found"
                    except Exception as e:
                        print(f"Error processing image: {str(e)}")
                        continue
                    finally:
                        # Tentativa de remover o arquivo de imagem temporário
                        if hasattr(image, 'filename'):
                            safe_remove(image.filename)
                return text
            except PermissionError:
                if attempt < max_retries - 1:
                    time.sleep(delay)
                else:
                    raise
            except Exception as e:
                print(f"Error extracting text from PDF: {str(e)}")
                return f"Error: {str(e)}"
    finally:
        # Limpeza do diretório temporário
        safe_remove(temp_dir)

# Parte de edição do código para fazer a busca. Apenas apresente a informação do documento que precisa de extração.
def extract_specific_information(text):
    #Linha do excel com as informação que ficarão como cabeçalho do documento
    data = {
        "Escola/Instituição Educacional": "",
        "CNPJ:": "",
        "Representado por:": "",
        "Cargo de:": "",
        "E-mail:": "",
        "Supervisor do Estágio:": "",
        "que ocupa o Cargo de:": "",
        "Curso Superior:": "",
        "Fone:": "",
        "E-mail de contato:": ""
    }

    #Informações a serem extraídas do documento. Estas informações serão incrementadas nas linhas abaixo do cabeçalho.
    patterns = {
        "A Escola/Instituição Educacional:": r"A Escola/Instituição Educacional:\s*(.*)",
        ", CNPJ:": r", CNPJ:\s*(\d{14})",
        "Representado por:": r"Representado por:\s*(.*)",
        "no Cargo de:": r"no Cargo de:\s*(.*)",
        "E-mail:": r"E-mail:\s*([^\s@]+@[^\s@]+\.[^\s@]+)",
        "Supervisor do Estágio:": r"Supervisor do Estágio:\s*(.*)",
        "- que ocupa o Cargo de:": r"- que ocupa o Cargo de:\s*(.*)",
        ", e é formado no Curso Superior:": r", e é formado no Curso Superior:\s*(.*)",
        ", Fone:": r", Fone:\s*(\(\d{2}\) \d{4}-\d{4})",
        "- E-mail de contato:": r"- E-mail de contato:\s*([^\s@]+@[^\s@]+\.[^\s@]+)"
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            data[key] = match.group(1).strip()

    return data

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if file and file.filename.lower().endswith(('.jpg', '.png', '.pdf')):
            try:
                temp_dir = tempfile.mkdtemp()
                file_path = os.path.join(temp_dir, file.filename)
                file.save(file_path)

                if file.filename.lower().endswith('.pdf'):
                    text = extract_text_from_pdf(file_path)
                else:
                    text = extract_text_from_image(file_path)

                if text.startswith("Error:"):
                    return text, 400  # Retorna mensagem de erro com status 400

                data = extract_specific_information(text)

                # Salvando dados em um arquivo Excel
                df = pd.DataFrame([data])
                output_excel_path = os.path.join(temp_dir, 'extracted_data.xlsx')
                df.to_excel(output_excel_path, index=False)

                return send_file(output_excel_path, as_attachment=True, download_name='extracted_data.xlsx')
            except Exception as e:
                return f"An error occurred: {str(e)}", 500
            finally:
                # Limpeza do diretório temporário
                safe_remove(temp_dir)
        else:
            return "Invalid file type. Please upload a JPG, PNG, or PDF file.", 400

    return '''
    <!doctype html>
    <title>Upload File</title>
    <h1>Upload a JPG, PNG, or PDF file</h1>
    <form method=post enctype=multipart/form-data>
      <input type=file name=file>
      <input type=submit value=Upload>
    </form>
    '''

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)
