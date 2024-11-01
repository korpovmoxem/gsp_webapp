async def on_task_comment_add(fast_bitrix, form_data: dict, request_bitrix=None):
    task_id = form_data['data[FIELDS_AFTER][TASK_ID]']
    task_info = await fast_bitrix.get_all('tasks.task.get', {
        'select': ['UF_AUTO_682875830446', 'STATUS', 'CHANGED_BY', 'TITLE'],
        'taskId': task_id
    })
    result_info = await fast_bitrix.call('tasks.task.result.list', {
        'taskId': task_id
    })
    task_info = task_info['task']
    if result_info:
        if int(task_info['status']) != 3 and ('ufAuto682875830446' not in task_info or result_info['text'] not in task_info['ufAuto682875830446']):
            return await fast_bitrix.call('im.notify.system.add', {
                'USER_ID': task_info['changedBy'],
                'MESSAGE': f'Поле "Отчет/Комментарий" в задаче "{task_info["title"]}" не заполнено, '
                           f'так как задача не находится в статусе "Выполняется"\n'}, raw=True)

        if 'ufAuto682875830446' not in task_info or not task_info['ufAuto682875830446'] or result_info['text'] not in task_info['ufAuto682875830446']:
            return await fast_bitrix.call('tasks.task.update', {
                'taskId': task_id,
                'fields': {
                    'UF_AUTO_682875830446': [result_info['text']]
                }})

