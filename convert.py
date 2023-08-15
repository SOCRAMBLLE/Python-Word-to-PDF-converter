import win32com.client as win32
import os


# change file permissions

def change_permissions(file_path, permissions):
    try:
        os.chmod(file_path, permissions)
        print(f"Permissions changed for {file_path}")
    except Exception as e:
        print(f"Error changing permissions: {e}")
        
# convert the file

def convert_docx_to_pdf(docx_path, pdf_path):
    word = win32.Dispatch("Word.Application")
    doc = word.Documents.Open(docx_path)
    doc.SaveAs(pdf_path, FileFormat=17) # 17 = pdf
    doc.Close()
    word.Quit()

if __name__ == "__main__":
    word_folder = os.path.abspath("./word/")
    pdf_folder = os.path.abspath("./pdf/")

    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)

    for filename in os.listdir(word_folder):
        if filename.endswith(".docx"):
            docx_path = os.path.join(word_folder, filename)
            pdf_filename = os.path.splitext(filename)[0] + "_converted.pdf"
            pdf_path = os.path.join(pdf_folder, pdf_filename)

            change_permissions(docx_path, 0o777) 
            convert_docx_to_pdf(docx_path, pdf_path)
            print(f"Converted {filename} to {pdf_filename}")
