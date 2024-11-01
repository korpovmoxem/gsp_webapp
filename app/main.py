import uvicorn
from fastapi import FastAPI, Request, HTTPException, status

from tools.tools import Authentication, fast_bitrix
from default_webhooks.OnTaskCommentAdd import on_task_comment_add
from default_webhooks.OnTaskDelete import on_task_delete
from custom_webhooks.UpdateList_28 import update_list_28
from custom_webhooks.ExcelList_28 import excel_list_28
from custom_webhooks.ExcelSmart_185 import excel_smart_185
from custom_webhooks.GetExcelProjects import get_excel_projects
from custom_webhooks.KeyTasks_401 import create_key_tasks
from custom_webhooks.ExcelProject_401 import create_key_tasks_report
import custom_webhooks.TaskPlanProjects_GSPI
import custom_webhooks.BlockTasks_GSPI

from test_funcs import test

from time import time

app = FastAPI()
authentication = Authentication()

default_webhooks = {
    'ONTASKCOMMENTADD': on_task_comment_add,
    'ONTASKDELETE': on_task_delete,
}

custom_webhook = {
    'update_list_28': update_list_28,
    'excel_list_28': excel_list_28,
    'excel_smart_185': excel_smart_185,
    'get_excel_projects': get_excel_projects,
    'create_key_tasks': create_key_tasks,
    'create_key_tasks_report': create_key_tasks_report,
    'create_gspi_report': custom_webhooks.TaskPlanProjects_GSPI.create_report,
    'create_gspi_non_completed_report': custom_webhooks.TaskPlanProjects_GSPI.create_non_completed_tasks_report,
    'gspi_notification': custom_webhooks.TaskPlanProjects_GSPI.send_notification,
    'block_tasks_gspi': custom_webhooks.BlockTasks_GSPI.create_report,
}


@app.get('/b24_webhook')
@app.post('/b24_webhook')
async def b24_default(request: Request):
    form_data = await request.form()
    form_data = dict(form_data)

    # Дефолтные вебхуки
    if 'auth[application_token]' in form_data:
        if form_data['auth[application_token]'] != authentication.credentials['outgoing_webhook_token']:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Неверный токен приложение исходящего вебхука",
            )
        await default_webhooks[form_data['event']](fast_bitrix=fast_bitrix, form_data=form_data)
    else:
        if request.query_params['job'] == 'test':
            return time()
        # Кастомные вебхуки
        await custom_webhook[request.query_params['job']](fast_bitrix=fast_bitrix, params=request.query_params)
    return 'ok'


@app.get('/create_key_tasks')
async def create_tasks():
    await create_key_tasks(fast_bitrix)

@app.get('/test')
async def call_test():
    a = await test()
    return len(a)


if __name__ == '__main__':
    uvicorn.run(
        app,
        host='10.64.240.113',
        ssl_keyfile='ssl/gspromWC-dec.key.pem',
        ssl_certfile='ssl/gspromWC.crt.pem',
        ssl_ca_certs='ssl/cachain.crt',
    )
