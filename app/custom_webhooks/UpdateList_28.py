from fastapi import Request

from time import time


def get_element_value(field_name: str, fields_info: dict, element_info: dict) -> str | float:
    try:
        value = list(element_info[fields_info[field_name]['FIELD_ID']].values())[0]
        if '|RUB' in value:
            value = value.replace('|RUB', '')
        return float(value)
    except KeyError:
        return '0'


async def update_list_28(fast_bitrix, params: Request.query_params, request_bitrix=None):
    """
    Создание суммирующих элементов в списке в разрезе месяца и года

    :param fast_bitrix: Асинхронный экземпляр класса Bitrix библиотеки fast_bitrix24
    :param params: Параметры запроса:
                    list_id - ID списка
                    element_id - ID элемента списка
    """
    element_info = await fast_bitrix.get_all('lists.element.get', {
        'IBLOCK_TYPE_ID': 'lists',
        'IBLOCK_ID': params['list_id'],
        'ELEMENT_ID': params['element_id'],
    })
    fields_raw = await fast_bitrix.get_all('lists.field.get', {
        'IBLOCK_TYPE_ID': 'lists',
        'IBLOCK_ID': params['list_id'],
    })
    fields_info = dict()
    for key in fields_raw.keys():
        fields_info[fields_raw[key]['NAME']] = fields_raw[key]

    month_code = get_element_value('Месяц', fields_info, element_info[0])
    year_code = get_element_value('Год', fields_info, element_info[0])
    element_name = 'Итого за месяц'

    # Итого за месяц
    sum_element_info = await fast_bitrix.get_all('lists.element.get', {
        'IBLOCK_TYPE_ID': 'lists',
        'IBLOCK_ID': params['list_id'],
        'FILTER': {
            'NAME': element_name,
            fields_info['Месяц']['FIELD_ID']: month_code,
            fields_info['Год']['FIELD_ID']: year_code,
        }
    })
    if not sum_element_info:
        sum_element_id = await fast_bitrix.call('lists.element.add', {
            'IBLOCK_TYPE_ID': 'lists',
            'IBLOCK_ID': params['list_id'],
            'ELEMENT_CODE': time(),
            'FIELDS': {
                'NAME': element_name,
                fields_info['Месяц']['FIELD_ID']: month_code,
                fields_info['Год']['FIELD_ID']: year_code,
                fields_info['План ремонта ТС, шт']['FIELD_ID']: 0,
                fields_info['ИТОГО план ремонтов ТС на 1 число, шт']['FIELD_ID']: 0,
                fields_info['Дефектовка ТС, шт']['FIELD_ID']: 0,
                fields_info['Заявка на ремонт ТС, шт']['FIELD_ID']: 0,
                fields_info['Закупка ЗЧ для ТС, шт']['FIELD_ID']: 0,
                fields_info['Оплата ЗЧ для ТС, шт']['FIELD_ID']: 0,
                #fields_info['Потребность персонала РММ, чел.']['FIELD_ID']: 0,
                #fields_info['Факт персонала РММ, чел.']['FIELD_ID']: 0,
                fields_info['Поставка ЗЧ для ТС, шт']['FIELD_ID']: 0,
                fields_info['Ремонт ТС, шт']['FIELD_ID']: 0,
                #fields_info['План ремонта, шт. ДОЛГ']['FIELD_ID']: 0,
                #fields_info['Ремонт ТС, шт. ДОЛГ']['FIELD_ID']: 0,
                fields_info['ОБС с учетом КЗ, План']['FIELD_ID']: 0,
                fields_info['ОБС с учетом КЗ, Факт']['FIELD_ID']: 0,
                fields_info['ОперШтаб, План']['FIELD_ID']: 0,
                fields_info['ОперШтаб, Факт']['FIELD_ID']: 0,
                fields_info['Страховой запас, План']['FIELD_ID']: 0,
                fields_info['Страховой запас, Факт']['FIELD_ID']: 0,
                fields_info['Мониторинговый счет, Факт']['FIELD_ID']: 0,
                fields_info['Мониторинговый счет, Остаток на р/сч']['FIELD_ID']: 0,
                fields_info['Перенос с предыдущего месяца']['FIELD_ID']: 0,
            }
        })
    else:
        sum_element_id = sum_element_info[0]['ID']
    date_range_elements_info = await fast_bitrix.get_all('lists.element.get', {
        'IBLOCK_TYPE_ID': 'lists',
        'IBLOCK_ID': params['list_id'],
        'FILTER': {
            '!NAME': element_name,
            fields_info['Месяц']['FIELD_ID']: month_code,
            fields_info['Год']['FIELD_ID']: year_code,
        }
    })
    await fast_bitrix.call('lists.element.update', {
        'IBLOCK_TYPE_ID': 'lists',
        'IBLOCK_ID': params['list_id'],
        'ELEMENT_ID': sum_element_id,
        'FIELDS': {
            'NAME': element_name,
            fields_info['Месяц']['FIELD_ID']: month_code,
            fields_info['Год']['FIELD_ID']: year_code,
            fields_info['План ремонта ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('План ремонта ТС, шт', fields_info, element)), date_range_elements_info)),
            fields_info['ИТОГО план ремонтов ТС на 1 число, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('ИТОГО план ремонтов ТС на 1 число, шт', fields_info, element)), date_range_elements_info)),
            fields_info['Дефектовка ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Дефектовка ТС, шт', fields_info, element)), date_range_elements_info)),
            fields_info['Заявка на ремонт ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Заявка на ремонт ТС, шт', fields_info, element)), date_range_elements_info)),
            fields_info['Закупка ЗЧ для ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Закупка ЗЧ для ТС, шт', fields_info, element)), date_range_elements_info)),
            fields_info['Оплата ЗЧ для ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Оплата ЗЧ для ТС, шт', fields_info, element)), date_range_elements_info)),
            #fields_info['Потребность персонала РММ, чел.']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Потребность персонала РММ, чел.', fields_info, element)), date_range_elements_info)),
            #fields_info['Факт персонала РММ, чел.']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Факт персонала РММ, чел.', fields_info, element)), date_range_elements_info)),
            fields_info['Поставка ЗЧ для ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Поставка ЗЧ для ТС, шт', fields_info, element)), date_range_elements_info)),
            fields_info['Ремонт ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Ремонт ТС, шт', fields_info, element)), date_range_elements_info)),
            #fields_info['План ремонта, шт. ДОЛГ']['FIELD_ID']: sum(map(lambda element: int(get_element_value('План ремонта, шт. ДОЛГ', fields_info, element)), date_range_elements_info)),
            #fields_info['Ремонт ТС, шт. ДОЛГ']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Ремонт ТС, шт. ДОЛГ', fields_info, element)), date_range_elements_info)),
            fields_info['ОБС с учетом КЗ, План']['FIELD_ID']: sum(map(lambda element: int(get_element_value('ОБС с учетом КЗ, План', fields_info, element)), date_range_elements_info)),
            fields_info['ОБС с учетом КЗ, Факт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('ОБС с учетом КЗ, Факт', fields_info, element)), date_range_elements_info)),
            fields_info['ОперШтаб, План']['FIELD_ID']: sum(map(lambda element: int(get_element_value('ОперШтаб, План', fields_info, element)), date_range_elements_info)),
            fields_info['ОперШтаб, Факт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('ОперШтаб, Факт', fields_info, element)), date_range_elements_info)),
            fields_info['Страховой запас, План']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Страховой запас, План', fields_info, element)), date_range_elements_info)),
            fields_info['Страховой запас, Факт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Страховой запас, Факт', fields_info, element)), date_range_elements_info)),
            fields_info['Мониторинговый счет, Факт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Мониторинговый счет, Факт', fields_info, element)), date_range_elements_info)),
            fields_info['Мониторинговый счет, Остаток на р/сч']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Мониторинговый счет, Остаток на р/сч', fields_info, element)), date_range_elements_info)),
            fields_info['Перенос с предыдущего месяца']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Перенос с предыдущего месяца', fields_info, element)), date_range_elements_info)),

        }})

    # Итого общее
    element_name = 'Итого за год'
    sum_element_info = await fast_bitrix.get_all('lists.element.get', {
        'IBLOCK_TYPE_ID': 'lists',
        'IBLOCK_ID': params['list_id'],
        'FILTER': {
            'NAME': element_name,
            fields_info['Год']['FIELD_ID']: year_code,
        }
    })
    if not sum_element_info:
        sum_element_id = await fast_bitrix.call('lists.element.add', {
            'IBLOCK_TYPE_ID': 'lists',
            'IBLOCK_ID': params['list_id'],
            'ELEMENT_CODE': time(),
            'FIELDS': {
                'NAME': element_name,
                fields_info['Год']['FIELD_ID']: year_code,
                #fields_info['План ремонта ТС, шт']['FIELD_ID']: 0,
                fields_info['ИТОГО план ремонтов ТС на 1 число, шт']['FIELD_ID']: 0,
                fields_info['Дефектовка ТС, шт']['FIELD_ID']: 0,
                fields_info['Заявка на ремонт ТС, шт']['FIELD_ID']: 0,
                fields_info['Закупка ЗЧ для ТС, шт']['FIELD_ID']: 0,
                fields_info['Оплата ЗЧ для ТС, шт']['FIELD_ID']: 0,
                #fields_info['Потребность персонала РММ, чел.']['FIELD_ID']: 0,
                #fields_info['Факт персонала РММ, чел.']['FIELD_ID']: 0,
                fields_info['Поставка ЗЧ для ТС, шт']['FIELD_ID']: 0,
                fields_info['Ремонт ТС, шт']['FIELD_ID']: 0,
                #fields_info['План ремонта, шт. ДОЛГ']['FIELD_ID']: 0,
                #fields_info['Ремонт ТС, шт. ДОЛГ']['FIELD_ID']: 0,
                fields_info['ОБС с учетом КЗ, План']['FIELD_ID']: 0,
                fields_info['ОБС с учетом КЗ, Факт']['FIELD_ID']: 0,
                fields_info['ОперШтаб, План']['FIELD_ID']: 0,
                fields_info['ОперШтаб, Факт']['FIELD_ID']: 0,
                fields_info['Мониторинговый счет, Факт']['FIELD_ID']: 0,
                fields_info['Мониторинговый счет, Остаток на р/сч']['FIELD_ID']: 0,
            }
        })
    else:
        sum_element_id = sum_element_info[0]['ID']
    date_range_elements_info = await fast_bitrix.get_all('lists.element.get', {
        'IBLOCK_TYPE_ID': 'lists',
        'IBLOCK_ID': params['list_id'],
        'FILTER': {
            '!NAME': element_name,
            fields_info['Год']['FIELD_ID']: year_code,
        }
    })
    await fast_bitrix.call('lists.element.update', {
        'IBLOCK_TYPE_ID': 'lists',
        'IBLOCK_ID': params['list_id'],
        'ELEMENT_ID': sum_element_id,
        'FIELDS': {
            'NAME': element_name,
            fields_info['Год']['FIELD_ID']: year_code,
            #fields_info['План ремонта ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('План ремонта ТС, шт', fields_info, element)), date_range_elements_info)),
            fields_info['ИТОГО план ремонтов ТС на 1 число, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('ИТОГО план ремонтов ТС на 1 число, шт', fields_info, element)), date_range_elements_info)),
            fields_info['Дефектовка ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Дефектовка ТС, шт', fields_info, element)), date_range_elements_info)),
            fields_info['Заявка на ремонт ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Заявка на ремонт ТС, шт', fields_info, element)), date_range_elements_info)),
            fields_info['Закупка ЗЧ для ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Закупка ЗЧ для ТС, шт', fields_info, element)), date_range_elements_info)),
            fields_info['Оплата ЗЧ для ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Оплата ЗЧ для ТС, шт', fields_info, element)), date_range_elements_info)),
            #fields_info['Потребность персонала РММ, чел.']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Потребность персонала РММ, чел.', fields_info, element)), date_range_elements_info)),
            #fields_info['Факт персонала РММ, чел.']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Факт персонала РММ, чел.', fields_info, element)), date_range_elements_info)),
            fields_info['Поставка ЗЧ для ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Поставка ЗЧ для ТС, шт', fields_info, element)), date_range_elements_info)),
            fields_info['Ремонт ТС, шт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Ремонт ТС, шт', fields_info, element)), date_range_elements_info)),
            #fields_info['План ремонта, шт. ДОЛГ']['FIELD_ID']: sum(map(lambda element: int(get_element_value('План ремонта, шт. ДОЛГ', fields_info, element)), date_range_elements_info)),
            #fields_info['Ремонт ТС, шт. ДОЛГ']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Ремонт ТС, шт. ДОЛГ', fields_info, element)), date_range_elements_info)),
            fields_info['ОБС с учетом КЗ, План']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Новое ОБС с учетом КЗ, План', fields_info, element)), date_range_elements_info)),
            fields_info['ОБС с учетом КЗ, Факт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Новое ОБС с учетом КЗ, Факт', fields_info, element)), date_range_elements_info)),
            fields_info['ОперШтаб, План']['FIELD_ID']: sum(map(lambda element: int(get_element_value('ОперШтаб, План', fields_info, element)), date_range_elements_info)),
            fields_info['ОперШтаб, Факт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('ОперШтаб, Факт', fields_info, element)), date_range_elements_info)),
            fields_info['Мониторинговый счет, Факт']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Мониторинговый счет, Факт', fields_info, element)), date_range_elements_info)),
            fields_info['Мониторинговый счет, Остаток на р/сч']['FIELD_ID']: sum(map(lambda element: int(get_element_value('Мониторинговый счет, Остаток на р/сч', fields_info, element)), date_range_elements_info)),
        }
    })





