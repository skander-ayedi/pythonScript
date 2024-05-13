import os
import win32com.client

def convert_docm_to_docx(docm_file):
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(docm_file)
    new_filename = os.path.splitext(docm_file)[0] + "_converted.docx"
    doc.SaveAs(new_filename, FileFormat=16)  # FileFormat=16 indicates .docx format
    doc.Close()
    word.Quit()
    return new_filename

def main():
    docm_folder = input("Enter the folder path containing .docm files: ")
    output_folder = input("Enter the output folder path for converted .docx files: ")

    # Iterate .docm files in the folder
    for filename in os.listdir(docm_folder):
        if filename.endswith(".docm"):
            docm_file = os.path.join(docm_folder, filename)
            converted_docx_file = convert_docm_to_docx(docm_file)
            print(f"Converted {docm_file} to {converted_docx_file}")

if __name__ == "__main__":
    main()
