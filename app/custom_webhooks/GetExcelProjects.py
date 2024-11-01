import os
import base64
from datetime import datetime

import openpyxl
from fastapi import Request

from custom_webhooks.ExcelList_28 import get_user_folder_id
from custom_webhooks.ExcelSmart_185 import get_user_name


async def get_excel_projects(fast_bitrix, params: Request.query_params):
    projects = await fast_bitrix.get_all('sonet_group.get')
    projects = list(filter(lambda project: params['project_name'].lower() in project['NAME'].lower(), projects))

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(['№ п/п', 'Наименование проекта', 'Руководитель проекта'])
    for index, row in enumerate(projects):
        user_name = await get_user_name(row['OWNER_ID'], fast_bitrix)
        sheet.append([index + 1, row['NAME'], user_name])

    report_name = f'Список_проектов_{datetime.now().strftime("%d_%m_%Y_%H_%M_%S")}.xlsx'
    workbook.save(report_name)

    bitrix_folder_id = await get_user_folder_id(fast_bitrix, params['user'][5:], 'Списки_проектов')
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
        'MESSAGE': f'Список проектов: {upload_report["DETAIL_URL"]}'}, raw=True)
    os.remove(report_name)
