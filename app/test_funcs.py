#from tools.tools import fast_bitrix_slow, fast_bitrix
import sys
from email.utils import formataddr

import re

'''
def test():
    a = fast_bitrix_slow.get_all('tasks.task.list', {
        'select': ['*', 'UF_*']})
    return a

fields = test()
with open('tasks.txt', 'a') as file:
    for i in fields:
        a = i['id']
        try:
            b = i['ufAuto851221265619'] if 'ufAuto851221265619' in i else ''
        except:
            b = ''
        file.write(f"{a},{b},{i['groupId']},{i['createdDate']};\n")

    a = fast_bitrix_slow.get_all('tasks.task.list', {
        'filter': {
            'GROUP_ID': ['225', '228', '323', '231', '232', '234', '224', '222', '226', '223', '227', '255', '256', '257', '317', '122', '278', '263', '275', '267', '279', '335', '266', '268', '277', '271', '276', '264', '320', '336', '274', '284', '270', '273', '251', '378', '218', '366', '379', '369', '374', '362', '363', '357', '358', '364', '350', '360', '359', '361', '384', '385', '386', '387', '344', '345', '389', '390', '391', '392', '393', '394', '395', '396', '397', '398'],
            '>CREATED_DATE': '2023-12-31'

        }})
    return '\n'.join(list(map(lambda x: x['id'], a)))

with open('tasks.txt', 'w') as file:
    file.write(test())


l = list()
for i in sys.stdin:
    l.append(str(i).strip().strip(','))
print(l)
'''
'''
async def test():
    a = await fast_bitrix.get_all('tasks.task.list', {'filter': {'STATUS': '3'}})
    return a


def a():
    tasks = fast_bitrix_slow.get_all('tasks.task.list', {
        'filter': {
            'GROUP_ID': '255'
        }
    })
    for index, task in enumerate(tasks, 1):
        print(index, len(tasks))
        history = fast_bitrix_slow.get_all('tasks.task.history.list', {
            'taskId': task['id']
        })
        completed_status = list(filter(lambda x: x['value']['from'] == '5' and x['value']['to'] == '2' and x['user']['id'] == '1', history))
        if completed_status:
            fast_bitrix_slow.call('tasks.task.update', {
                'taskId': task['id'],
                'fields': {
                    'STATUS': '5'
                }
            })

from datetime import date, timedelta
def b():
    start_week = date.today()
    while start_week.isoweekday() != 1:
        start_week = start_week - timedelta(days=1)
    end_week = start_week + timedelta(days=6)

import smtplib
from email.mime.text import MIMEText
from email.utils import make_msgid

def asc():
    user_info = fast_bitrix_slow.get_all('user.get', {'ID': '296'})
    message = MIMEText('test')
    message['From'] = formataddr(('План мероприятий', '1c_b24_task_plan@gsprom.ru'))
    message['Subject'] = f'Контроль выполнения плана мероприятий за'
    message['Message-ID'] = make_msgid()
    server = smtplib.SMTP('email.gsprom.ru:587')
    server.ehlo()
    server.starttls()
    server.ehlo
    server.login('1c_b24_task_plan@gsprom.ru', 'bk=/PN1F3Q%H')
    message['To'] = user_info[0]['EMAIL']
    server.sendmail(message['From'], message['To'], message.as_string())
    server.quit()

asc()
'''

for row in sys.stdin:
    email = re.search('(\w*)@.*', row)
    print(email.groups(1)[0])