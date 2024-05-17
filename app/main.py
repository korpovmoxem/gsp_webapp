import uvicorn
from fastapi import FastAPI, Request, HTTPException, status

from tools.tools import Authentication, fast_bitrix
from default_webhooks.OnTaskUpdate import on_task_update
from custom_webhooks.UpdateList_28 import update_list_28
from custom_webhooks.ExcelList_28 import excel_list_28
from custom_webhooks.ExcelSmart_185 import excel_smart_185

from time import time

app = FastAPI()
authentication = Authentication()

default_webhooks = {
    'ONTASKUPDATE': on_task_update,
}

custom_webhook = {
    'update_list_28': update_list_28,
    'excel_list_28': excel_list_28,
    'excel_smart_185': excel_smart_185,
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


if __name__ == '__main__':
    uvicorn.run(
        app,
        host='10.64.240.113',
        ssl_keyfile='ssl/gspromWC-dec.key.pem',
        ssl_certfile='ssl/gspromWC.crt.pem',
        ssl_ca_certs='ssl/cachain.crt',
    )
