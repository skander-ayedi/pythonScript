import docx2txt
from docx import Document

def extract_text_from_docx(docx_file):
    main_text = docx2txt.process(docx_file)
    doc = Document(docx_file)
    text_boxes_text = []
    for shape in doc.inline_shapes:
        if shape.type == 3:  # code text box
            text_boxes_text.append(shape.text)
    return main_text, '\n'.join(text_boxes_text)
# fichier word 
main_text, text_boxes_text = extract_text_from_docx(r'C:\Users\user\Desktop\python\G25057685_PLMACTION11471781_CHECKREPORT_DA_BACKCHECK_converted.docx')

print("Main Text:")
print(main_text)

print("\nText Boxes Text:")
print(text_boxes_text)
