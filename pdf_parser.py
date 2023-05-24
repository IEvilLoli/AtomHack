from PyPDF2 import PdfFileReader, PdfReader
import pypdf
from tabula import read_pdf


def pdf_parse(filepath):
    pdf_document = filepath
    with open(pdf_document, "rb") as filehandle:
        pdf = PdfReader(filehandle)

        info = pdf.metadata
        pages = len(pdf.pages)
        print("Количество страниц в документе: %i\n\n" % pages)
        print("Мета-описание: ", info)
        for i in range(pages):
            page = pdf.pages[i]
            print("Стр.", i, " мета: ", page, "\n\nСодержание;\n")
            print(page.extract_text())


if __name__ == '__main__':
    pdf_parse("test.pdf")
    print(0)
