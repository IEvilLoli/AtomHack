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
from pathlib import Path


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
    try:
        root = ET.Element('object', id=dict_for_xml["document_id"], status="", createUser="",
                          objectDef="", modifyUser="User")

        # если ведомость
        if dict_for_xml["typefile"] == "Рабочая документация":
            attributes = ET.Element('attributes')
            attributes.append(ET.Element('attribute', name="A_Create_Time", datatype="date", value=dict_for_xml["Дата "]))
            attributes.append(ET.Element('attribute', name="A_Package_Number", datatype="string", value=dict_for_xml["package"]))
            attributes.append(ET.Element('attribute', name="A_Revision_Number", datatype="string", value=dict_for_xml["Номер ревизии"]))
            attributes.append(ET.Element('attribute', name="A_Inventory_Number", datatype="string", value=""))
            attributes.append(ET.Element('attribute', name="A_Name", datatype="string", value=dict_for_xml["files_list"][0]))
            attributes.append(ET.Element('attribute', name="A_Name_Eng", datatype="string", value=dict_for_xml["files_list"][0]))
            attributes.append(ET.Element('attribute', name="A_Designation", datatype="string", value=""))
            attributes.append(ET.Element('attribute', name="A_Dep", datatype="classifier", value=""))
            attributes.append(ET.Element('attribute', name="A_User", datatype="user", value=""))
            attributes.append(ET.Element('attribute', name="A_Doc_Language", datatype="classifier", value=""))
            root.append(attributes)
            files = ET.Element('files')
            files.append(ET.Element('file', id="", name="", primary="", bodyId="", modifiedTime="",
                                    createdTime="", fileDef="", hash="", size="", path=""))
            root.append(files)

        # если CheckList, IKL, Notes, PDTK
        if dict_for_xml["typefile"] == "Заключение ПДТК" or \
                dict_for_xml["typefile"] == "Additional letter" or \
                dict_for_xml["typefile"] == "Explanatory Note" or \
                dict_for_xml["typefile"] == "Пояснительная записка" or \
                dict_for_xml["typefile"] == "Сопроводительное письмо" or \
                dict_for_xml["typefile"] == "Чек-лист":
            attributes = ET.Element('attributes')
            attributes.append(ET.Element('attribute', name="A_Order", datatype="string", value=dict_for_xml["order"]))
            attributes.append(ET.Element('attribute', name="A_Block", datatype="string", value=dict_for_xml["block"]))
            attributes.append(ET.Element('attribute', name="A_Package", datatype="string", value=dict_for_xml["package"]))
            table = ET.Element('attribute', name="A_Docs_Tbl", datatype="table")
            rows = ET.Element('rows')
            for i in range(len(dict_for_xml["id_work"])):
                t = len(dict_for_xml["id_work"])
                row = ET.Element('row', order="")
                row.append(ET.Element('attribute', name="A_Type_Link", datatype="classifier", value=dict_for_xml["list_other_column"][2*i]))
                row.append(ET.Element('attribute', name="A_Doc_Addition_Ref", datatype="object", value=dict_for_xml["id_element"][i]))
                row.append(ET.Element('attribute', name="A_Note", datatype="string", value=dict_for_xml["list_other_column"][2*i+1]))
                rows.append(row)
            table.append(rows)
            attributes.append(table)
            root.append(attributes)
            files = ET.Element('files')
            files.append(ET.Element('file', id="", name="", primary="", bodyId="", modifiedTime="",
                                    createdTime="", fileDef="", hash="", size="", path=""))
            root.append(files)

        # если files_paths

        # если док в пакете

        indent(root)
        # xml_str = ET.tostring(root, encoding="utf-8", method="xml")
        etree = ET.ElementTree(root)
        f = io.BytesIO()
        etree.write(f, encoding='utf-8', xml_declaration=True)
        # print(f.getvalue().decode(encoding="utf-8"))
        # Чтобы сразу в файл записать:
        myfile = open(filepath + "/" + dict_for_xml["document_id"] + ".xml", "wb")
        etree.write(myfile, encoding='utf-8', xml_declaration=True)
    except:
        root = ET.Element('object', id="", status="", createUser="",
                          objectDef="", modifyUser="User")
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
        root.append(attributes)
        files = ET.Element('files')
        files.append(ET.Element('file', id="", name="", primary="", bodyId="", modifiedTime="",
                                createdTime="", fileDef="", hash="", size="", path=""))
        root.append(files)
        indent(root)
        # xml_str = ET.tostring(root, encoding="utf-8", method="xml")
        etree = ET.ElementTree(root)
        f = io.BytesIO()
        etree.write(f, encoding='utf-8', xml_declaration=True)
        # print(f.getvalue().decode(encoding="utf-8"))
        # Чтобы сразу в файл записать:
        myfile = open(filepath + "/" + dict_for_xml["document_id"] + ".xml", "wb")
        etree.write(myfile, encoding='utf-8', xml_declaration=True)


def build_package(filepath, dict_file_status):
    # получение информации о файле
    dict_push = docx_parser.docx_parse(filepath)
    file_check = filepath.split("\\")[1].split(".")[0]
    if not(file_check in dict_file_status):
        dict_file_status[file_check] = {}
    dict_file_status[file_check]["FE"] = 1
    # print("")

    #формирование основных директорий пакета
    if not os.path.isdir(dict_push["order"]):
        os.mkdir(dict_push["order"])
    if not os.path.isdir(dict_push["order"] + "/" + dict_push["block"]):
        os.mkdir(dict_push["order"] + "/" + dict_push["block"])
    if not os.path.isdir(dict_push["order"] + "/" + dict_push["block"] + "/" + dict_push["package"]):
        os.mkdir(dict_push["order"] + "/" + dict_push["block"] + "/" + dict_push["package"])

    iswf = re.search(r"[Rr][^.WP]{1,}WP \S{1,}\d\.doc\S*", filepath)
    # t = iswf[0]
    iswf_t = iswf[0] if iswf else 'Not found'
    # print(0)
    # iswf.replace(" ", "_")
    if iswf_t != 'Not found':
        for file in dict_push["files_list"]:
            # print(0)
            # if not dict_file_status[file]:
            if not(file in dict_file_status):
                dict_file_status[file] = {}
            dict_file_status[file]["WF"] = 1
            print("")

    # формирование директорий по файлам и копирование файлов, создание xml
    curr_path = dict_push["order"] + "/" + dict_push["block"] + "/" + dict_push["package"] + "/AccDocs"
    if not os.path.isdir(curr_path):
        os.mkdir(curr_path)
    print(dict_push)
    if dict_push["typefile"] == "Check-list" or dict_push["typefile"] == "Чек-лист":
        if not os.path.isdir(curr_path + "/CheckList"):
            os.mkdir(curr_path + "/CheckList")
        if not os.path.isdir(curr_path + "/CheckList" + "/" + dict_push["document_id"]):
            os.mkdir(curr_path + "/CheckList" + "/" + dict_push["document_id"])
        if not os.path.isdir(
                curr_path + "/CheckList" + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files"):
            os.mkdir(curr_path + "/CheckList" + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files")
        shutil.copy2(filepath,
                     curr_path + "/CheckList" + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files")
        create_xml(dict_push, curr_path + "/CheckList" + "/" + dict_push["document_id"])
    elif dict_push["typefile"] == "Additional letter" or dict_push["typefile"] == "Сопроводительное письмо":
        if not os.path.isdir(curr_path + "/IKL"):
            os.mkdir(curr_path + "/IKL")
        if not os.path.isdir(curr_path + "/IKL" + "/" + dict_push["document_id"]):
            os.mkdir(curr_path + "/IKL" + "/" + dict_push["document_id"])
        if not os.path.isdir(curr_path + "/IKL" + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files"):
            os.mkdir(curr_path + "/IKL" + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files")
        shutil.copy2(filepath, curr_path + "/IKL" + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files")
        create_xml(dict_push, curr_path + "/IKL" + "/" + dict_push["document_id"])
    elif dict_push["typefile"] =="Explanatory Note" or dict_push["typefile"] =="Пояснительная записка" or dict_push["typefile"] == "Рабочая документация":
        if not os.path.isdir(curr_path + "/Notes"):
            os.mkdir(curr_path + "/Notes")
        if not os.path.isdir(curr_path + "/Notes" + "/" + dict_push["document_id"]):
            os.mkdir(curr_path + "/Notes" + "/" + dict_push["document_id"])
        if not os.path.isdir(curr_path + "/Notes" + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files"):
            os.mkdir(curr_path + "/Notes" + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files")
        shutil.copy2(filepath,
                     curr_path + "/Notes" + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files")
        create_xml(dict_push, curr_path + "/Notes" + "/" + dict_push["document_id"])
    elif dict_push["typefile"] == "Заключение ПДТК":
        if not os.path.isdir(curr_path + "/PDTK"):
            os.mkdir(curr_path + "/PDTK")
        if not os.path.isdir(curr_path + "/PDTK" + "/" + dict_push["document_id"]):
            os.mkdir(curr_path + "/PDTK" + "/" + dict_push["document_id"])
        if not os.path.isdir(curr_path + "/PDTK" + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files"):
            os.mkdir(curr_path + "/PDTK" + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files")
        shutil.copy2(filepath, curr_path + "/PDTK" + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files")
        create_xml(dict_push, curr_path + "/PDTK" + "/" + dict_push["document_id"])
    else:
        curr_path = dict_push["order"] + "/" + dict_push["block"] + "/" + dict_push["package"] + "/Docs"
        if not os.path.isdir(curr_path):
            os.mkdir(curr_path)
        if not os.path.isdir(curr_path + "/" + dict_push["document_id"]):
            os.mkdir(curr_path + "/" + dict_push["document_id"])
        if not os.path.isdir(curr_path + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files"):
            os.mkdir(curr_path + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files")
        shutil.copy2(filepath, curr_path + "/" + dict_push["document_id"] + "/" + dict_push["document_id"] + ".files")
        create_xml(dict_push, curr_path + "/" + dict_push["document_id"])



def find_wf(path):
    all_files = []
    for root, dirs, files in os.walk(path):
        for filename in files:
            all_files.append(filename)
            # print(filename)

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


def collecting_data(filepath):
    dict_file_status = {}
    # list = []
    p = Path(filepath)
    for x in p.rglob("*"):
        # list.append(x)
        # Преобразование файл doc в docx
        # file = x.split("/")
        new_str = str(x)
        if str(x)[-4:] == ".doc":
            w = wc.Dispatch('word.Application')
            # r = os.path.abspath(file)
            doc_docx = w.Documents.Open(os.path.abspath(x))
            doc_docx.SaveAs(os.path.abspath(x) + "x", 16)
            doc_docx.Close()
            w.Quit()
            new_str = str(x) + 'x'
        build_package(new_str, dict_file_status)
        # dict_file_status

    print("")


if __name__ == '__main__':
    collecting_data("data")
    filepath = find_wf("data")
    # print(filepath)
    # filepath = "data/Чек-лист _5 9 3 10 RUENG.docx"
    # build_package("data/RU_5_9_3_10.docx", dict_file_status = {})
    # build_package("data/R23 KK56 50UMA 0 ET WP WD003=r0.docx", dict_file_status={})
    # print(datetime.datetime.fromtimestamp(os.path.getctime(filepath)).strftime('%Y%m%d%H%M%S'))
