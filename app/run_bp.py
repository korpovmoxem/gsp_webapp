from tools.tools import fast_bitrix_slow, fast_bitrix

elements = fast_bitrix_slow.get_all('lists.element.get', {
        'IBLOCK_TYPE_ID': 'lists',
        'IBLOCK_ID': '28',
    })

for index, elem in enumerate(elements, 1):
    print(index)
    fast_bitrix_slow.call('bizproc.workflow.start', {'TEMPLATE_ID': '251', 'DOCUMENT_ID': ['lists', 'Bitrix\Lists\BizprocDocumentLists', elem['ID']]})