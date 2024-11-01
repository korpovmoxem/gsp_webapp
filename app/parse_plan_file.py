import openpyxl

from tools.tools import fast_bitrix_slow


workbook = openpyxl.load_workbook('Автоматизация ФАКТ на НОЯБРЬ.xlsx')
sheet = workbook.active

group_users = list(map(lambda user: user['USER_ID'], fast_bitrix_slow.get_all('sonet_group.user.get', {'ID': 402})))
group_users_info = fast_bitrix_slow.get_all('user.get', {'ID': group_users})

for row in sheet.rows:
    row_values = list(map(lambda cell: cell.value, row))
    if row_values[0] == '№':
        titles = row_values
    if not any(row_values) or not isinstance(row_values[0], int):
        continue
    row_values = dict(zip(titles, row_values))

    task_description = (f"Зачем:\n{row_values['Зачем'] if row_values['Зачем'] else ''}\n\n"
                        f"Уровень контроля:\n{row_values['Уровень контроля'] if row_values['Уровень контроля'] else ''}\n\n"
                        f"Риски:\n{row_values['Риски'] if row_values['Риски'] else ''}\n\n"
                        f"Необходимая поддержка:\n{row_values['Необходимая поддержка'] if row_values['Необходимая поддержка'] else ''}\n\n")

    if not row_values['Задача']:
        continue

    find_user = list(filter(lambda user: user['LAST_NAME'] == row_values['Ответственный '].split(' ')[0], group_users_info))
    if len(find_user) != 1:
        print(f'Не найден ответственный:\n{row_values}')
        continue
    try:
        fast_bitrix_slow.call('tasks.task.add', {
            'fields': {
                'TITLE': row_values['Задача'],
                'DESCRIPTION': task_description,
                'START_DATE_PLAN': row_values['Дата начала (план))'],
                'END_DATE_PLAN': row_values['Дата окончания (план)'],
                'GROUP_ID': '402',
                'RESPONSIBLE_ID': find_user[0]['ID'],
                'CREATED_BY': '1',
            }
        })
    except:
        print(f'Ошибка: {row_values}')
