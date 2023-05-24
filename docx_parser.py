import re
import docx

def find_type(name_type, text):
    typefile = re.search(name_type, text)
    typefile_cg = typefile[0] if typefile else 'Not found'
    return typefile_cg


def docx_parse(filepath):
    # doc = docx.Document("Чек-лист _5 9 3 10 RUENG.docx")
    doc = docx.Document(filepath)
    print(doc.paragraphs)


    text = []
    for paragraph in doc.paragraphs:
        text.append(paragraph.text)
    print('\n'.join(text))


    print(doc.tables[1].columns[0].cells[1].text)

    full_text_table = []
    for table in doc.tables:
        for column in table.columns:
            for cell in column.cells:
                full_text_table.append(cell.text)
                # print(cell.text)

    text_paragraphs = ' '.join(text)
    text_tables = ' '.join(full_text_table)
    text_all = text_paragraphs+text_tables


    typefile_cg = find_type("Рабочая документация", text_all)
    if typefile_cg == 'Not found':
        typefile_cg = find_type("Чек-лист", text_all)
    if typefile_cg == 'Not found':
        typefile_cg = find_type("Сопроводительное письмо", text_all)

    # новое решение
    print( "----------")

    dict_info_main = {}

    dict_info_main["typefile"] = typefile_cg
    dict_info_main["id_work"] = "12345"
    full_text_table = []
    for j in range(len(doc.tables[1].rows)):
        for i in range(len(doc.tables[1].rows[j].cells)):
            if j != 0:
                print(doc.tables[1].rows[j].cells[i].text)
                text_clear = doc.tables[1].rows[j].cells[i].text
                text_clear = re.sub(r"[/\|\?]", '-', text_clear, count=0)
                if i == 0:
                    dict_info_main["order"] = text_clear
                elif i == 1:
                    dict_info_main["block"] = text_clear
                elif i == 2:
                    dict_info_main["package"] = text_clear
                else:
                    dict_info_main[doc.tables[1].rows[0].cells[i].text] = text_clear
                    # full_text_table.append(cell.text)
    print(dict_info_main)



    return dict_info_main


if __name__ == '__main__':
    docx_parse("data/R23 KK56 50UMA 0 ET WP WD003=r0.docx")
