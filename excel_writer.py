import docx_parser
import pandas as pd

# скрипт формирования отчётного excel файла
def create_excel(dict_file_status, path_wfs):
    print(path_wfs)
    for path_wf in path_wfs:
        if path_wf != "":
            dict_push = docx_parser.docx_parse(path_wf)

        file_name = dict_file_status.keys()
        file_name = list(file_name)

        WF = []
        FE = []
        for values in dict_file_status.values():
            if 'WF' in values:
                WF.append(values['WF'])
            else:
                WF.append(0)
            if 'FE' in values:
                FE.append(values["FE"])
            else:
                FE.append(0)

        check_file_WF = []
        check_file_FE = []

        for check in WF:
            if check == 1:
                check_file_WF.append('Да')
            else:
                check_file_WF.append('Нет')
        for check in FE:
            if check == 1:
                check_file_FE.append('Да')
            else:
                check_file_FE.append('Нет')
        # Подсчёт количества файлов физически
        count_ex = 0
        count_wf = 0
        for x in FE:
            count_ex += x
        for x in WF:
            count_wf += x

        salaries2 = pd.DataFrame({'Имя файла': file_name,
                                  'Указан в ведомости': check_file_WF,
                                  'Существует физически': check_file_FE,
                                  })

        if path_wf != "":
            salaries1 = pd.DataFrame({'Контракт': [dict_push['order']],
                                      'Блок': [dict_push["block"]],
                                      'Ведомость': [dict_push["package"]],
                                      'Количество файлов по ведомости': [len(dict_push['files_list'])],
                                      'Количество файлов физически': [count_ex],
                                      })
        else:
            salaries1 = pd.DataFrame({'Контракт': 0,
                                      'Блок': 0,
                                      'Ведомость': 0,
                                      'Количество файлов по ведомости': 0,
                                      'Количество файлов физически': 0,
                                      }, index=[0])

        salary_sheets = {'Общие сведения': salaries1, 'Сведения о ведомости': salaries2}
        writer = pd.ExcelWriter(f'./files{dict_push["order"]}.xlsx', engine='xlsxwriter')
        for sheet_name in salary_sheets.keys():
            salary_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
        writer._save()

