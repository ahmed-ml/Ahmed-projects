import os
from docx import Document

def combine_word_documents(directory):
    # Get the list of all .docx files in the directory
    files = [os.path.join(directory, file) for file in os.listdir(directory) if file.endswith('.docx')]
    merged_document = Document()

    for index, file in enumerate(files):
        # Load the content of the Word document
        sub_doc = Document(file)
        print(f'Merging: {file}')

        # Combine the documents
        for element in sub_doc.element.body:
            merged_document.element.body.append(element)

        # Add a page break if it's not the last document
        if index != len(files) - 1:
            merged_document.add_page_break()

    # Save the merged document
    merged_document.save(os.path.join(directory, 'combined_document.docx'))
    print("Documents have been successfully merged.")

# Path to the directory containing the Word files
directory_path = 'C:/Users/QK615NU/OneDrive - EY/Desktop/Dossier de travail/DEV/BOFIP 1/BOFIP/Part 1'
combine_word_documents(directory_path)
