import streamlit as st
from backend import *
import os
import aspose.words as aw


def main():
    st.title("Resume Data Extraction App")
    st.header("Upload Resumes")
    uploaded_files = st.file_uploader("Upload one or more resumes", type=["doc", "docx", "pdf"], accept_multiple_files=True)
    print(uploaded_files)
    if uploaded_files:
        data = []
        for uploaded_file in uploaded_files:
            file_extension = os.path.splitext(uploaded_file.name)[1]
            #path = os.path.dirname(str(uploaded_file))
            if file_extension == ".doc":
                #doc_path = os.path.join(cv_directory, filename)
                doc_path = uploaded_file.name
                doc = aw.Document(uploaded_file)
                doc.save(f"output.docx")
                #doc_path = uploaded_file
                #doc = spire.doc.Document(doc_path)
                #doc_path_ = doc_path.split('.')[0] + '.docx'
                #doc.SaveToFile(doc_path_, spire.doc.FileFormat.Docx)
                text = extract_text_from_docx("output.docx")
                emails = extract_email_ids(text)
                phone_numbers = extract_phone_numbers(text)
                data.append((emails, phone_numbers, text))
            elif file_extension == ".docx":
                #doc_path = os.path.join(cv_directory, filename)
                doc_path = uploaded_file
                text = extract_text_from_docx(doc_path)
                emails = extract_email_ids(text)
                phone_numbers = extract_phone_numbers(text)
                data.append((emails, phone_numbers, text))
            elif file_extension == ".pdf":
                #pdf_path = os.path.join(cv_directory, filename)
                pdf_path = uploaded_file
                text = extract_text_from_pdf(pdf_path)
                emails = extract_email_ids(text)
                phone_numbers = extract_phone_numbers(text)
                data.append((emails, phone_numbers, text))
            else:
                st.warning("Unsupported file format: {}".format(file_extension))

        output_file_path = "output_file.xlsx"
        save_to_excel(data, output_file_path)

        #st.markdown(f"### [Download Output Excel File](/{output_file_path})")

        # Provide a download button for the output file
        with open(output_file_path, "rb") as f:
            file_contents = f.read()
        st.download_button(label="Download Output Excel File", data=file_contents, file_name=output_file_path)


if __name__ == "__main__":
    main()
