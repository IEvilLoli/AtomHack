# pip install pypiwin32

import io
import shutil
import docx
import os
import datetime
import re
import docx_parser
import xml.etree.ElementTree as ET
from win32com import client as wc


def indent(elem, level=0):
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


def create_xml(dict_for_xml, filepath):
    # если ведомость

    # если Docs

    # если files_paths

    # если CheckList, IKL, Notes, PDTK
    root = ET.Element('data')
    object = ET.Element('object', id=dict_for_xml["id_work"], createTime="",
                        modifyTime="", status="", createUser="",
                        objectDef="", modifyUser="")
    attributes = ET.Element('attributes')
    attributes.append(ET.Element('attribute', name="A_Creation_Date", datatype="date", value=""))
    attributes.append(ET.Element('attribute', name="A_Name", datatype="string", value=""))
    attributes.append(ET.Element('attribute', name="A_Designation", datatype="string", value=""))
    table = ET.Element('attribute', name="A_Docs_Tbl", datatype="table")
    rows = ET.Element('rows')
    for i in range(1):
        row = ET.Element('row', order="")
        row.append(ET.Element('attribute', name="A_Type_Link", datatype="classifier", value=""))
        row.append(ET.Element('attribute', name="A_Doc_Addition_Ref", datatype="object", value=""))
        row.append(ET.Element('attribute', name="A_Note", datatype="string", value=""))
        rows.append(row)
    table.append(rows)
    attributes.append(table)
    object.append(attributes)
    files = ET.Element('files')
    files.append(ET.Element('file', id="", name="", primary="", bodyId="", modifiedTime="",
                            createdTime="", fileDef="", hash="", size="", path=""))
    object.append(files)
    root.append(object)

    indent(root)
    # xml_str = ET.tostring(root, encoding="utf-8", method="xml")
    etree = ET.ElementTree(root)
    f = io.BytesIO()
    etree.write(f, encoding='utf-8', xml_declaration=True)
    # print(f.getvalue().decode(encoding="utf-8"))
    # Чтобы сразу в файл записать:
    myfile = open(filepath + "/" + dict_for_xml["id_work"] + ".xml", "wb")
    etree.write(myfile, encoding='utf-8', xml_declaration=True)


def build_package(filepath):
    dict_push = docx_parser.docx_parse(filepath)

    if not os.path.isdir(dict_push["order"]):
        os.mkdir(dict_push["order"])

    if not os.path.isdir(dict_push["order"] + "/" + dict_push["block"]):
        os.mkdir(dict_push["order"] + "/" + dict_push["block"])

    if not os.path.isdir(dict_push["order"] + "/" + dict_push["block"] + "/" + dict_push["package"]):
        os.mkdir(dict_push["order"] + "/" + dict_push["block"] + "/" + dict_push["package"])

    curr_path = dict_push["order"] + "/" + dict_push["block"] + "/" + dict_push["package"] + "/AccDocs"

    if not os.path.isdir(curr_path):
        os.mkdir(curr_path)
    if "Check-list" in dict_push["typefile"] or "Чек-лист" in dict_push["typefile"]:
        if not os.path.isdir(curr_path + "/CheckList"):
            os.mkdir(curr_path + "/CheckList")
        if not os.path.isdir(curr_path + "/CheckList" + "/" + dict_push["id_work"]):
            os.mkdir(curr_path + "/CheckList" + "/" + dict_push["id_work"])
        if not os.path.isdir(
                curr_path + "/CheckList" + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files"):
            os.mkdir(curr_path + "/CheckList" + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files")
        shutil.copy2(filepath,
                     curr_path + "/CheckList" + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files")
        create_xml(dict_push, curr_path + "/CheckList" + "/" + dict_push["id_work"])
    elif "IKL" in dict_push["typefile"] or "Пояснительная записка" in dict_push["typefile"]:
        if not os.path.isdir(curr_path + "/IKL"):
            os.mkdir(curr_path + "/IKL")
        if not os.path.isdir(curr_path + "/IKL" + "/" + dict_push["id_work"]):
            os.mkdir(curr_path + "/IKL" + "/" + dict_push["id_work"])
        if not os.path.isdir(curr_path + "/IKL" + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files"):
            os.mkdir(curr_path + "/IKL" + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files")
        shutil.copy2(filepath, curr_path + "/IKL" + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files")
        create_xml(dict_push, curr_path + "/IKL" + "/" + dict_push["id_work"])
    elif "Notes" or "Рабочая документация" in dict_push["typefile"] or "Пояснительная записка" in dict_push["typefile"]:
        if not os.path.isdir(curr_path + "/Notes"):
            os.mkdir(curr_path + "/Notes")
        if not os.path.isdir(curr_path + "/Notes" + "/" + dict_push["id_work"]):
            os.mkdir(curr_path + "/Notes" + "/" + dict_push["id_work"])
        if not os.path.isdir(curr_path + "/Notes" + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files"):
            os.mkdir(curr_path + "/Notes" + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files")
        shutil.copy2(filepath,
                     curr_path + "/Notes" + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files")
        create_xml(dict_push, curr_path + "/Notes" + "/" + dict_push["id_work"])
    elif "PDTK" in dict_push["typefile"] or "ПДТК" in dict_push["typefile"]:
        if not os.path.isdir(curr_path + "/PDTK"):
            os.mkdir(curr_path + "/PDTK")
        if not os.path.isdir(curr_path + "/PDTK" + "/" + dict_push["id_work"]):
            os.mkdir(curr_path + "/PDTK" + "/" + dict_push["id_work"])
        if not os.path.isdir(curr_path + "/PDTK" + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files"):
            os.mkdir(curr_path + "/PDTK" + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files")
        shutil.copy2(filepath, curr_path + "/PDTK" + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files")
        create_xml(dict_push, curr_path + "/PDTK" + "/" + dict_push["id_work"])
    else:
        curr_path = dict_push["order"] + "/" + dict_push["block"] + "/" + dict_push["package"] + "/Docs"
        if not os.path.isdir(curr_path):
            os.mkdir(curr_path)
        if not os.path.isdir(curr_path + "/" + dict_push["id_work"]):
            os.mkdir(curr_path + "/" + dict_push["id_work"])
        if not os.path.isdir(curr_path + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files"):
            os.mkdir(curr_path + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files")
        shutil.copy2(filepath, curr_path + "/" + dict_push["id_work"] + "/" + dict_push["id_work"] + ".files")
        create_xml(dict_push, curr_path + "/" + dict_push["id_work"])


def find_wf(path):
    all_files = []
    for root, dirs, files in os.walk(path):
        for filename in files:
            all_files.append(filename)
            print(filename)

    file_path = ' '.join(all_files)
    # Ищем только файл ведомости
    file_path = re.search(r"[Rr][^.WP]{1,}WP \S{1,}\d\.doc\S*", file_path)
    print("find")
    print(file_path[0] if file_path else 'Not found')
    filepath = "data/" + file_path[0] if file_path else 'Not found'

    # Преобразуем файл doc в docx, т.к. библиотека не работает без этого
    w = wc.Dispatch('word.Application')
    doc_docx = w.Documents.Open(os.path.abspath(filepath))
    doc_docx.SaveAs(os.path.abspath(filepath) + "x", 16)
    doc_docx.Close()
    w.Quit()

    # filepath - финальный относительный путь до нужного документа
    filepath = f"data/{file_path[0]}" + 'x'

    return filepath


if __name__ == '__main__':
    filepath = find_wf("data")
    print(filepath)
    # filepath = "data/Чек-лист _5 9 3 10 RUENG.docx"
    build_package(filepath)
    # print(datetime.datetime.fromtimestamp(os.path.getctime(filepath)).strftime('%Y%m%d%H%M%S'))
