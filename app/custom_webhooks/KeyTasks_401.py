async def create_key_tasks(fast_bitrix):
    project_users = await fast_bitrix.get_all(
        'sonet_group.user.get',
        {
            'ID': '401'
        }
    )
    project_users = list(map(lambda user: 'user_' + user['USER_ID'], filter(lambda elem: elem['ROLE'] == 'K', project_users)))
    for user in project_users:
        await fast_bitrix.call(
            'bizproc.workflow.start',
            {
                'TEMPLATE_ID': '533',
                'DOCUMENT_ID': ['lists', 'BizprocDocument', '5586'],
                'PARAMETERS': {
                    'responsible': user,
                }
            }
        )