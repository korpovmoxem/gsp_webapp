async def on_task_update(fast_bitrix, form_data: dict, request_bitrix=None):
    task_id = form_data['data[FIELDS_BEFORE][ID]']
    task_info = await fast_bitrix.get_all('tasks.task.get', {
        'select': ['UF_AUTO_682875830446'],
        'taskId': task_id
    })
    task_info = task_info['task']
    if 'ufAuto682875830446' not in task_info or not task_info['ufAuto682875830446']:
        try:
            result_info = await fast_bitrix.call('tasks.task.result.list', {
                'taskId': task_id
            })
        except IndexError:
            return
        if result_info:
            await fast_bitrix.call('tasks.task.update', {
                'taskId': task_id,
                'fields': {
                    'UF_AUTO_682875830446': [result_info['text']]
                }})

