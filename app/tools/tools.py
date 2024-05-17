import os

import yaml
from yaml.loader import SafeLoader
from fast_bitrix24 import BitrixAsync, Bitrix
import requests


class Authentication:

    def __init__(self):
        try:
            with open(f'{os.getcwd()}\\authentication.yml', 'r') as file:
                self.credentials = yaml.load(file, SafeLoader)
        except FileNotFoundError:
            with open(f'{os.getcwd()}/authentication.yml', 'r') as file:
                self.credentials = yaml.load(file, SafeLoader)


def send_bitrix_request(method: str, data=None) -> dict | list:
    """
    Отправляет запрос в Б24. Не подходит для выгрузки большого массива данных (больше 50)

    :param method: Метод запроса в Б24
    :param data: Словаь с параметрами запроса
    :return: Ответ от Б24
    """

    bitrix_token = Authentication().credentials['incoming_webhook']
    request_json = requests.post(f"{bitrix_token}{method}", json=data, verify=False).json()

    if 'result' in request_json:
        return request_json['result']
    print(request_json)



fast_bitrix = BitrixAsync(Authentication().credentials['incoming_webhook'])
fast_bitrix_slow = Bitrix(Authentication().credentials['incoming_webhook'])


def test():
    tasks = fast_bitrix_slow.get_all('tasks.task.get', {'taskId': '7258', 'select': ['*', 'UF_*', 'TAGS']})
    print(tasks)


