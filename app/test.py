from tools.tools import fast_bitrix_slow


def test():
    a = fast_bitrix_slow.get_all('crm.item.fields', {
        'entityTypeId': '185',
    })
    return a

fields = test()['fields']
for i in fields:
    print(i, fields[i])




