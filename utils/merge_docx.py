from docxcompose.composer import Composer
from docx import Document
import os

def merge_docx_files(docx_paths, output_path):
    base_doc = Document(docx_paths[0])
    composer = Composer(base_doc)

    for file_path in docx_paths[1:]:
        base_doc.add_page_break()
        doc = Document(file_path)
        composer.append(doc)

    composer.save(output_path)
    print(f"✅ fusion_test.docx avec sauts de page créé avec succès dans : {output_path}")

