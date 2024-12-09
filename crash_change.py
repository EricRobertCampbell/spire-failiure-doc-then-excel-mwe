#!/usr/bin/env python3

import traceback

from spire.xls import *
from spire.xls.common import *
from spire.doc import *
from spire.doc.common import *
from spire.doc import FileFormat as DocxFileFormat
from pathlib import Path
import threading
import time

logo_path = Path('./MainLogo.png')
output_dir = Path('./results')

def main() -> None:
    threads = []
    for run in range(2):
        print(f"Generating documents #{run}")
        thread = threading.Thread(target=simulated_generate_documents)
        thread.start()
        threads.append(thread)
        print(f"Done generating documents #{run}")

     # Waiting for all threads to complete
    for thread_index, thread in enumerate(threads):
        thread.join()
        print(f"word #{thread_index} is saved.")


def simulated_generate_documents() -> None:
    simulated_generate_documents_fns = [
        #generate_xlsx_docs
         generate_xlsx_docs, generate_spire_doc, generate_xlsx_docs

    ]
    for f in simulated_generate_documents_fns:
        print(f"Running {f.__name__}")
        try:
            f()
        except Exception as e:
            print(f"Error generating document:", e)
            traceback.print_exc()
            # print(f"This is a bare exception. This should catch any error that is generated.")

        print(f"Done running {f.__name__}")

def generate_spire_doc() -> None:
    # Initialize a document
    doc = Document()
    # Add a section to the document
    section = doc.AddSection()
    paragraph = section.AddParagraph()
    paragraph.AppendPicture(str(logo_path))
    while True:
        timestamp = int(time.time() * 1000)
        thread_id = threading.get_ident()
        doc_filename = Path(f"results/out_{thread_id}_{timestamp}.docx")
        if not doc_filename.exists():
            break
        time.sleep(0.01)   
    # Save the document
    doc.SaveToFile(str(doc_filename), DocxFileFormat.Docx2013)
    print(f"Document saved as {doc_filename}")
    doc.Close()

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
        while True:
            timestamp = int(time.time() * 1000)
            thread_id = threading.get_ident()
            pdf_filename = Path(f"results/out_{thread_id}_{timestamp}.pdf")
            if not pdf_filename.exists():
                break
            time.sleep(0.01)   
        sheet.SaveToPdf(str(pdf_filename))
        print(f"Done saving pdf to {pdf_filename}")

    for wb in wbs:
        wb.Dispose()
        
if __name__ == "__main__":
    main()
