from datetime import datetime
import base64
import os

from fastapi import Request
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment

from custom_webhooks.ExcelList_28 import get_element_value, get_user_folder_id

users_info_collection = list()


async def get_user_name(user_id, fast_bitrix):
    global users_info_collection
    user_info = list(filter(lambda x: x['ID'] == user_id, users_info_collection))
    if not user_info:
        user_info = await fast_bitrix.get_all('user.get', {'ID': user_id})
        users_info_collection += user_info
    try:
        return f"{user_info[0]['LAST_NAME']} {user_info[0]['NAME']} {user_info[0]['SECOND_NAME']}"
    except IndexError:
        return user_id


def get_enumeration_field_value(value, field_info):
    try:
        return list(filter(lambda x: str(x['ID']) == str(value), field_info['items']))[0]['VALUE']
    except IndexError:
        return value


async def excel_smart_185(fast_bitrix, params: Request.query_params):
    items = await fast_bitrix.get_all('crm.item.list', {
        'entityTypeId': '185',
        'select': [
            'title',
            'stageId',
            'ufCrm5_1712046836',
            'ufCrm5_1712047958',
            'ufCrm5_1712049480',
            'ufCrm5_1712049395',
            'ufCrm5_1712049436',
            'ufCrm5_1712648792',
            'ufCrm5_1712301306',
            'ufCrm5_1712049544',
            'ufCrm5_1712049569',
            'ufCrm5_1712049616',
            'ufCrm5_1712049659',
            'ufCrm5_1712049672',
            'ufCrm5_1712051323',
            'ufCrm5_1712049773',
            'ufCrm5_1712049833',
            'ufCrm5_1712049853',
            'ufCrm5_1712049891',
            'ufCrm5_1714650169',
            'ufCrm5_1712049593',
            'ufCrm5_1712049741',
        ]
    })

    users_info = list()

    fields_info = await fast_bitrix.get_all('crm.item.fields', {
        'entityTypeId': '185',
    })
    fields_info = fields_info['fields']

    status_dict = {
        'DT185_9:NEW': 'В работе',
        'DT185_9:SUCCESS': 'Завершено',
        'DT185_9:CLIENT': 'Утверждение УПАиК'
    }

    workbook = openpyxl.Workbook()
    sheet = workbook.active

    sheet.append(
        [
            'Номер',
            'Наименование показателя',
            "Стадия",
            'Направление',
            'Ответственный за показатель',
            'Блок',
            'Приоритет',
            'Соисполнитель',
            'Ответственный в УПАиК',
            'Текущее значение',
            'Результат в %',
            'Дата среза',
            'Целевой результат по итогам встречи (или установленный отдельно ГД ГСП на год)',
            'Основные задачи на текущую и последующую недели',
            'Целевой показатель текущего месяца',
            'Результат исполнения прошлых периодов',
            '',
            'Дополнительные комментарии',
            'Результат исполнения ЯНВАРЯ',
            'Результат исполнения ФЕВРАЛЯ',
            'Результат исполнения МАРТА',
        ]
    )

    for row in items:
        sheet.append(
            [
                row['title'],
                get_enumeration_field_value(row['ufCrm5_1712046836'], fields_info['ufCrm5_1712046836']),
                status_dict[row['stageId']],
                get_enumeration_field_value(row['ufCrm5_1712047958'], fields_info['ufCrm5_1712047958']),
                await get_user_name(row['ufCrm5_1712049480'], fast_bitrix),
                get_enumeration_field_value(row['ufCrm5_1712049395'], fields_info['ufCrm5_1712049395']),
                get_enumeration_field_value(row['ufCrm5_1712049436'], fields_info['ufCrm5_1712049436']),
                await get_user_name(row['ufCrm5_1712648792'], fast_bitrix),
                await get_user_name(row['ufCrm5_1712301306'], fast_bitrix),
                row['ufCrm5_1712049544'],
                row['ufCrm5_1712049569'],
                datetime.fromisoformat(row['ufCrm5_1712049593']).strftime('%d.%m.%Y') if row['ufCrm5_1712049593'] else '',
                row['ufCrm5_1712049616'],
                row['ufCrm5_1712049659'],
                row['ufCrm5_1712049672'],
                row['ufCrm5_1712051323'],
                get_enumeration_field_value(row['ufCrm5_1712049741'], fields_info['ufCrm5_1712049741']),
                row['ufCrm5_1712049773'],
                row['ufCrm5_1712049833'],
                row['ufCrm5_1712049853'],
                row['ufCrm5_1712049891'],
                row['ufCrm5_1714650169'],
            ]
        )

    cell_indicator_colors = {
        'Зеленый': '00b04f',
        'Желтый': 'ffbf00',
        'Красный': 'ff0000'
    }

    row_cell_width = {
        0: 10,
        1: 50,
        2: 15,
        3: 25,
        4: 35,
        5: 25,
        6: 12,
        7: 35,
        8: 35,
        9: 40,
        10: 15,
        11: 12,
        12: 50,
        13: 50,
        14: 50,
        15: 50,
        16: 5,
        17: 50,
        18: 50,
        19: 50,
        20: 50,
        21: 50,
        22: 50,
        23: 50,
        24: 50,
        25: 50,
        26: 50,
        27: 50,
        28: 50,
        29: 50,
    }

    # Изменение ширины ячеек
    for index, row in enumerate(sheet.columns):
        sheet.column_dimensions[get_column_letter(row[0].column)].width = row_cell_width[index]

    # Автоперенос текста, стиль текста и выравнивание
    for row_index, row in enumerate(sheet.rows):
        for cell_index, cell in enumerate(row):
            if row_index < 1:
                cell.font = Font(name='Calibri', size=12, bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:

                # Цвет ячеек
                if cell_index == 16 and row_index > 0:
                    cell_color = cell_indicator_colors[cell.value] if cell.value in cell_indicator_colors else 'ff0000'
                    cell.fill = PatternFill(start_color=cell_color, end_color=cell_color, fill_type='solid')
                    cell.value = ''

                if cell_index in [0, 2, 4, 5, 6, 7, 8, 10, 11]:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

            cell.alignment = cell.alignment.copy(wrapText=True)




    report_name = f'Показатели_функционального_блока{datetime.now().strftime("%d_%m_%Y_%H_%M_%S")}.xlsx'
    workbook.save(report_name)

    bitrix_folder_id = await get_user_folder_id(fast_bitrix, params['user'][5:], 'ПФБ')
    with open(report_name, 'rb') as file:
        report_file = file.read()
    report_file_base64 = str(base64.b64encode(report_file))[2:]
    upload_report = await fast_bitrix.call('disk.folder.uploadfile', {
        'id': bitrix_folder_id,
        'data': {'NAME': report_name},
        'fileContent': report_file_base64
    })

    await fast_bitrix.call('im.notify.system.add', {
        'USER_ID': params['user'][5:],
        'MESSAGE': f'Отчет по показателям функционального блока сформирован. {upload_report["DETAIL_URL"]}'}, raw=True)
    os.remove(report_name)

