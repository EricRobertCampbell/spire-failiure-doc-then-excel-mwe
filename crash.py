#!/usr/bin/env python3

from spire.xls import *
from spire.xls.common import *

from spire.doc import *
from spire.doc.common import *
from spire.doc import FileFormat as DocxFileFormat

from pathlib import Path
import threading
import time

results_folder = Path('./results')
docx_file_path = results_folder / Path("sample_doc.docx")
docx_pdf_path = results_folder / Path('sample_doc_pdf.pdf')

logo_path = Path('./MainLogo.png')

def main() -> None:
    for run in range(2):
        print(f"Generating documents #{run}")
        thread = threading.Thread(target=simulated_generate_documents)
        thread.start()
        thread.join()
        print(f"Done generating documents #{run}")

        time.sleep(1)

def simulated_generate_documents() -> None:
    simulated_generate_documents_fns = [
        # generate_spire_doc, generate_xlsx_docs, generate_spire_doc, # crash
        # generate_xlsx_docs, generate_spire_doc, generate_xlsx_docs, # crash
        generate_xlsx_docs, generate_spire_doc, # crash
        # generate_spire_doc, generate_xlsx_docs, # crash
        # generate_xlsx_docs # no crash
        # generate_spire_doc # no crash
    ]
    for f in simulated_generate_documents_fns:
        print(f"Running {f.__name__}")
        try:
            f()
        except:
            print(f"This is a bare exception. This should catch any error that is generated.")

        print(f"Done running {f.__name__}")

def generate_spire_doc() -> None:
    # Initialize a document
    doc = Document()

    # Add a section to the document
    section = doc.AddSection()
    paragraph = section.AddParagraph()
    paragraph.AppendPicture(str(logo_path))

    # Save the document
    doc.SaveToFile(str(docx_file_path), DocxFileFormat.Docx2013)
    doc.Close()
    doc.Dispose()

def generate_xlsx_docs() -> None:
    wbs = []
    for index in range(2):
        wb = Workbook()
        sheet = wb.Worksheets[0]
        sheet.Pictures.Add(1, 3, str(logo_path))
        wbs.append(wb)

    # now save them to pdf
    for wb_index, wb in enumerate(wbs):
        sheet = wb.Worksheets[0]
        pdf_filename = results_folder / f"excel_pdf_{wb_index}.pdf"
        sheet.SaveToPdf(str(pdf_filename))
        print(f"Done saving pdf to {pdf_filename}")

    for wb in wbs:
        wb.Dispose()

if __name__ == "__main__":
    main()
