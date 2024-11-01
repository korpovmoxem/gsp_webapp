from time import time

async def on_task_delete(fast_bitrix, form_data: dict,):
    task_id = form_data['data[FIELDS_BEFORE][ID]']
    await fast_bitrix.call('lists.element.add', {
        'IBLOCK_TYPE_ID': 'lists',
        'IBLOCK_ID': '42',
        'ELEMENT_CODE': time(),
        'FIELDS': {
            'NAME': task_id,
        }
    })
