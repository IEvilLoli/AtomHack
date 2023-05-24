from PyPDF2 import PdfFileReader, PdfReader

pdf_document = "test.pdf"
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
