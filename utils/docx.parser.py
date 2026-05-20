import docx

def extract_text_from_docx(file_path):
    text = ""
    doc = docx.Document(file_path)

    for para in doc.paragraphs:
        text += para.text + "\n"

    return text