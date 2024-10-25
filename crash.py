#!/usr/bin/env python3

from spire.xls import *
from spire.xls.common import *
from spire.xls import FileFormat as XlsxFileFormat

from spire.doc import *
from spire.doc.common import *
from spire.doc import FileFormat as DocxFileFormat

from pathlib import Path
import threading
import time

results_folder = Path('./results')
excel_file_path = results_folder / Path("sample_xlsx.xlsx")
docx_file_path = results_folder / Path("sample_doc.docx")
docx_pdf_path = results_folder / Path('sample_doc_pdf.pdf')

logo_path = Path('./MainLogo.png')
font_path = Path('./fonts')


def spire_doc_test() -> None:
    # Initialize a document
    doc = Document()

    # Add a section to the document
    section = doc.AddSection()

    doc.PrivateFontList.append(PrivateFontPath("Muli", "Muli-Regular.ttf"))

    format = CharacterFormat(doc)
    format.FontName = "Muli"
    format.FontSize = 12
    format.TextColor = Color.FromArgb(alpha=1, red=67, green=67, blue=67)
    format.Bold = True

    # Add a header and footer to the section
    header = section.HeadersFooters.Header

    # Append a paragraph to the header
    header_paragraph = header.AddParagraph()
    header_paragraph.AppendText("Header content")

    for i in range(header_paragraph.ChildObjects.Count):
        childObj = header_paragraph.ChildObjects.get_Item(i)
        if isinstance(childObj, TextRange):
            tr = childObj if isinstance(childObj, TextRange) else None
            tr.ApplyCharacterFormat(format)

    header.ChildObjects.Clear()

    # Create table in header
    header_table = header.AddTable(False)
    header_table.TableFormat.HorizontalAlignment = HorizontalAlignment.Center
    header_table.PreferredWidth = PreferredWidth(WidthType.Percentage, 100)
    header_table.TableFormat.Borders.BorderType = BorderStyle.none
    row = header_table.AddRow(False, 2)

    row.Height = 50

    # Set cell width
    cell = row.Cells[0]
    cell.SetCellWidth(50, CellWidthType.Percentage)

    # Add another paragraph to a cell in the table
    cell_paragraph = cell.AddParagraph()
    cell_paragraph.AppendText("Header content")
    cell_paragraph.Format.HorizontalAlignment = HorizontalAlignment.Right
    cell_paragraph.ChildObjects.Clear()

    appended_pic = cell_paragraph.AppendPicture(str(logo_path))

    # Apply format to the paragraph
    cell_paragraph.Format.HorizontalAlignment = HorizontalAlignment.Left

    # Add a horizontal line
    horizontal_line = section.AddParagraph().AppendHorizonalLine()

    # Set table borders
    table = section.AddTable(False)
    table.TableFormat.HorizontalAlignment = HorizontalAlignment.Center
    table.TableFormat.Borders.BorderType = BorderStyle.Single

    # Set a picture to the paragraph (stubbed with a test path)
    logo_doc_picture = section.AddParagraph().AppendPicture(str(logo_path))
    logo_doc_picture.Height = 80
    logo_doc_picture.Width = 115

    # Create a summary table and add rows
    summary_table = section.AddTable(False)
    row = summary_table.AddRow(False, 2)
    cell = row.Cells[0]
    cell.SetCellWidth(50, CellWidthType.Percentage)
    cell.AddParagraph().AppendText("Summary content")

    # Add a footer table
    footer = section.HeadersFooters.Footer
    footer_paragraph = footer.AddParagraph()
    footer.ChildObjects.Clear()
    footer_paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center

    footer_table = footer.AddTable(False)
    footer_table.TableFormat.HorizontalAlignment = HorizontalAlignment.Center
    footer_table.PreferredWidth = PreferredWidth(WidthType.Percentage, 100)
    footer_table.TableFormat.Borders.BorderType = BorderStyle.none

    row = footer_table.AddRow(False, 2)
    row.Height = 50

    cell = row.Cells[0]
    cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
    cell.CellFormat.BackColor = Color.FromArgb(alpha=1, red=216, green=216, blue=216)
    cell.SetCellWidth(100 / 2, CellWidthType.Percentage)
    cell_paragraph = cell.AddParagraph()
    logo_doc_picture = cell_paragraph.AppendPicture(str(logo_path))

    # Add signature image (stubbed with base64 data)
    signature_paragraph = section.AddParagraph()
    signature_image = signature_paragraph.AppendPicture(str(logo_path))
    signature_image.Height = 60

    # Save the document
    doc.SaveToFile(str(docx_file_path), DocxFileFormat.Docx2013)
    doc.Close()
    doc.Dispose()

    # Load the document and save it as PDF
    pdf_doc = Document()
    pdf_doc.LoadFromFile(str(docx_file_path))

    ppl = ToPdfParameterList()
    ppl.IsEmbeddedAllFonts = True
    ppl.PrivateFontPaths = [PrivateFontPath("Muli", "Muli-Regular.ttf")]
    pdf_doc.SaveToFile(str(docx_pdf_path), ppl)
    pdf_doc.Close()
    pdf_doc.Dispose()


def spire_xls_test() -> None:
    wb = Workbook()
    wb.CustomFontFileDirectory = [str(font_path)]
    wb.Worksheets.Clear()
    sheet = wb.Worksheets.Add("Test Sheet")
    sheet.Range["A1"].Text = "Test Text"
    sheet.Range["A1"].Style.Font.FontName = "Roboto"

    wb.SaveToFile(str(excel_file_path), ExcelVersion.Version2016)
    print(f"Done save to {excel_file_path}")

    wb.Save()
    wb.Dispose()

def simulated_generate_documents() -> None:
    simulated_generate_documents_fns = [spire_doc_test, spire_xls_test, spire_doc_test, spire_xls_test, spire_doc_test, spire_xls_test, spire_doc_test]
    for f in simulated_generate_documents_fns:
        print(f"Running {f.__name__}")
        f()
        print(f"Done running {f.__name__}")

def main() -> None:
    print(f"Generating documents 1")
    thread = threading.Thread(target=simulated_generate_documents)
    thread.start()
    thread.join()
    print(f"Done generating documents 1")

    time.sleep(1)

    print(f"Generating documents 2")
    thread = threading.Thread(target=simulated_generate_documents)
    thread.start()
    thread.join()
    print(f"Done generating documents 2")

    # spire_xls_test()
    # print("generating Doc and PDF")
    # spire_doc_test()
    # print("Done generating Doc")
    # spire_xls_test()
    # print("generating Doc and PDF")
    # spire_doc_test()
    # print("Done generating Doc")
    # spire_xls_test()
    # print("generating Doc and PDF")
    # spire_doc_test()
    # print("Done generating Doc")


if __name__ == "__main__":
    main()
