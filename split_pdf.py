import os
from fitz import Page, Document

os.system('cls')

doc = Document("_pdfs/LUNA_all_in_one_annotated_and_checked.pdf")
toc = doc.get_toc()

for i, bm in enumerate(toc):
    print(f"Processing page {i+1}...")
    bm_title = bm[1]
    bm_page_number = bm[2]-1
    print(f"{bm_title=} {bm_page_number=}")
    new_doc = Document()
    new_doc.insert_pdf(doc, bm_page_number, bm_page_number, annots=True)
    new_doc.save(f"_pdfs/{bm_title}.pdf")
    new_doc.close()

doc.close()