import imapclient
import gspread
from google.oauth2.service_account import Credentials
from email.header import decode_header
import re
from datetime import datetime, timedelta
import datetime as dt
from gspread_formatting import *
import time
import os

MAIL_USERNAME = "incomingTD@tmpk.net"
MAIL_PASSWORD = "isqMJeZ8RYwybDgeDQ37"

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'credentials.json'

spreadsheet_id = '1XARhrr6YIIxxaWsafe_aoNB7j5Z6Nk7PYFv53eOA1pU'

last_day_file = open(r'last_day', 'r+')
curr_date = dt.date.today().strftime('%d-%b-%Y')
last_day = last_day_file.read()
last_day_file.seek(0)
last_day_file.write(curr_date)
last_day_file.close()

ignored_numbers_file = open(r'ignored numbers', 'r+')
ignored_numbers = []
for line in ignored_numbers_file:
    ignored_numbers.append(line.strip())
ignored_numbers_file.seek(0)
ignored_numbers_file.truncate()
ignored_numbers_file.write('000000')
ignored_numbers_file.close

start_date = last_day

task_type_dict = {
    'узел': 'Реорганизация узла',
    'выделение': 'Анализ тех. возможности',
    'обрывы': 'Изменение док-ции',
    'юр.лицо': 'Юр. лицо',
    'юр. лицо': 'Юр. лицо',
    'физ.лицо': 'Физ. лицо',
    'физ. лицо': 'Физ. лицо',
    'Физ лицо': 'физ. лицо',
    'Безопасный регион': 'Юр. лицо',
    'failure': 'Failure'
}

def connect_to_mail():
    mail = imapclient.IMAPClient('imap.mail.ru', ssl=True)
    mail.login(MAIL_USERNAME, MAIL_PASSWORD)
    mail.select_folder('INBOX')
    return mail

def connect_to_sheets():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(spreadsheet_id).sheet1

def adjust_date_for_time_and_weekend(received_date):
    if received_date.hour >= 16:
        received_date += timedelta(days=1)
    while received_date.weekday() >= 5:
        received_date += timedelta(days=1)
    return received_date

def convert_date_to_number(date):
    epoch = datetime(1899, 12, 30)
    delta = date - epoch
    return delta.days

last_msg_id_file = open(r'last_msg_id', 'r')
last_msg_id_file_content = last_msg_id_file.read()
last_msg_id_file.close()

if last_msg_id_file_content == '':
    last_processed_id = 0
else:
    last_processed_id = int(last_msg_id_file_content)

while True:
    try:
        with connect_to_mail() as mail:
            worksheet = connect_to_sheets()
            messages = mail.search(['SINCE', start_date])
            
            if len(messages) > 0:
                print(f'Found {len(messages)} messages since {start_date}')

                for msg_id in messages:
                    if msg_id <= last_processed_id:
                        continue
                    
                    msg_data = mail.fetch([msg_id], ['ENVELOPE'])
                    msg = msg_data[msg_id]
                    envelope = msg[b'ENVELOPE']

                    if envelope.date:
                        adjusted_date = adjust_date_for_time_and_weekend(envelope.date)
                        date_number = convert_date_to_number(adjusted_date)  # Конвертация даты в числовой формат
                    else:
                        date_number = None
                    subject = envelope.subject.decode() if isinstance(envelope.subject, bytes) else envelope.subject
                    #Ошибка ниже
                    if subject == None:
                        subject = "Неопределенный"
                    decoded_subject = decode_header(subject)[0]
                    if isinstance(decoded_subject[0], bytes):
                        if isinstance(decoded_subject[0], bytes):
                            subject = decoded_subject[0].decode(decoded_subject[1] or 'utf-8')
                    else:
                        subject = decoded_subject[0]

                    if subject.startswith(('Re:', 'Fwd:')):
                        continue 

                    task_type = 'Неопределенный'
                    for key, value in task_type_dict.items():
                        if key.lower() in subject.lower():
                            task_type = value
                            break
                    
                    # Special cases
                    if 'Изменение ИД' in subject or 'изменение ИД' in subject or 'Проектирование Сектора' in subject or 'Подготовка сметы' in subject:
                        task_type = 'Изменение док-ции'
                
                    number_match = re.search(r'\b\d{5,6}\b', subject)
                    number = number_match.group(0) if number_match else '-'
                    if not number in ignored_numbers:

                        last_row = len(worksheet.col_values(1))
                        id_formula = f"=A{last_row}+1"

                        # Записываем числовое значение даты
                        worksheet.append_row(['', date_number, '', task_type, number])

                        worksheet.update_acell(f"A{last_row + 1}", id_formula)

                        cell_range = f'B{last_row + 1}'
                        format_cell_range(worksheet, cell_range, cellFormat(numberFormat=NumberFormat(type='DATE', pattern='dd.mm.yy')))  # Форматируем ячейку как дату

                    last_processed_id = msg_id
                    
                    last_msg_id_file = open(r'last_msg_id', 'w+')
                    last_msg_id_file.seek(0)
                    last_msg_id_file.write(str(last_processed_id))
                    last_msg_id_file.close()
                    
                    
                print("Данные успешно записаны в Google Таблицу.")
            else:
                print("Новых сообщений нет.")
        time.sleep(60)
    except Exception as e:
        print(f"Ошибка!: {e}")
        time.sleep(600)
