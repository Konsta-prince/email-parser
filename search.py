from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import imaplib
import json
import email
from email.header import decode_header
from datetime import datetime, timedelta
import re
import os
import sys

def load_config():
    #Определяем папку, где лежит программа (.exe или .py)
    if getattr(sys, 'frozen', False):
        # Если запущен как .exe
        base_path = os.path.dirname(sys.executable)
    else:
        # Если запущен как .py скрипт
        base_path = os.path.dirname(os.path.abspath(__file__))

    config_path = os.path.join(base_path, 'config.json')

    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
        return config
config = load_config()

email_settings = config['email_settings']
search_settings = config['search_settings']

imap_server = email_settings['imap_server']
email_address = email_settings['email_address']
password = email_settings['password']
days_to_search = search_settings['days_to_search']
keywords = search_settings['keywords']

# Вычисляем дату 30 дней назад
date_from = (datetime.now() - timedelta(days=days_to_search)).strftime("%d-%b-%Y")

print(f"Ищем письма с {date_from}...")
print(f"Ключевые слова: {keywords}")
print("-" * 50)

def connect_to_mail(server, email, password):
    # Подключаемся
    mail = imaplib.IMAP4_SSL(server)
    mail.login(email, password)
    mail.select('INBOX')

    return mail

mail = connect_to_mail(imap_server, email_address, password)

# Ищем все письма с указанной даты
status, messages = mail.search(None, f'SINCE {date_from}')

if status != 'OK':
    print("Не удалось найти письма")
    mail.logout()
    exit()

email_ids = messages[0].split()
print(f"Найдено писем за период: {len(email_ids)}")
print("-" * 50)

# Функция для декодирования заголовков
def decode_header_value(value):
    if value is None:
        return ""
    decoded_parts = decode_header(value)
    decoded_value = ""
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            decoded_value += part.decode(encoding or 'utf-8', errors='ignore')
        else:
            decoded_value += part
    return decoded_value

# Функция для очистки HTML тегов
def clean_html(text):
    # Убираем все теги <...>
    clean_text = re.sub(r'<[^>]+>', '', text)
    # Убираем лишние пробелы и переносы строк
    clean_text = re.sub(r'\s+', ' ', clean_text).strip()
    return clean_text

# Функция для получения тела письма
def get_email_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))
            
            # Пропускаем вложения
            if "attachment" in content_disposition:
                continue
            
            # Сначала ищем текстовую версию (text/plain)
            if content_type == "text/plain":
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or 'utf-8'
                    body = payload.decode(charset, errors='ignore')
                    break
            # Если текстовой нет, берем HTML
            elif content_type == "text/html" and not body:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or 'utf-8'
                    body = payload.decode(charset, errors='ignore')
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or 'utf-8'
            body = payload.decode(charset, errors='ignore')
    
    # Очищаем HTML теги
    body = clean_html(body)
    return body

# Функция для проверки наличия вложений
def has_attachments(msg):
    if msg.is_multipart():
        for part in msg.walk():
            content_disposition = str(part.get("Content-Disposition", ""))
            if "attachment" in content_disposition:
                return True
    return False

# Функция для извлечения чистого email из строки "Name <email@domain.com>"
def extract_email(from_string):
    match = re.search(r'<([^>]+)>', from_string)
    if match:
        return match.group(1)
    # Если email без скобок
    if '@' in from_string:
        return from_string.strip()
    return from_string

# Функция для форматирования даты
def format_date(date_string):
    try:
        # Пытаемся распарсить дату в формате IMAP
        # Пример: "Mon, 21 Sep 2026 10:30:00 +0300"
        date_obj = datetime.strptime(date_string, "%a, %d %b %Y %H:%M:%S %z")
        return date_obj.strftime("%Y-%m-%d %H:%M:%S")
    except:
        try:
            # Альтернативный формат
            date_obj = datetime.strptime(date_string, "%d %b %Y %H:%M:%S %z")
            return date_obj.strftime("%Y-%m-%d %H:%M:%S")
        except:
            return date_string

# Проверяем каждое письмо
found_emails = []

for email_id in email_ids:
    status, msg_data = mail.fetch(email_id, '(RFC822)')
    
    if status != 'OK':
        continue
    
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)
    
    # Получаем заголовки
    subject = decode_header_value(msg.get('Subject', ''))
    from_raw = decode_header_value(msg.get('From', ''))
    from_email = extract_email(from_raw)
    date_raw = msg.get('Date', '')
    date_formatted = format_date(date_raw)
    
    # Получаем тело письма
    body = get_email_body(msg)
    
    # Проверяем вложения
    has_attach = "Да" if has_attachments(msg) else "Нет"
    
    # Проверяем наличие ключевых слов
    found_keywords = []
    for keyword in keywords:
        if keyword.lower() in body.lower() or keyword.lower() in subject.lower():
            found_keywords.append(keyword)
    
    if found_keywords:
        found_emails.append({
            'date': date_formatted,
            'from_email': from_email,
            'subject': subject,
            'body': body[:500],  # Первые 500 символов
            'has_attachments': has_attach,
            'keywords': ', '.join(found_keywords)
        })
        
        print(f"✓ Найдено совпадение!")
        print(f"  Дата: {date_formatted}")
        print(f"  От: {from_email}")
        print(f"  Тема: {subject}")
        print(f"  Вложения: {has_attach}")
        print(f"  Ключевые слова: {', '.join(found_keywords)}")
        print(f"  Тело (первые 100 символов): {body[:100]}...")
        print("-" * 50)

print(f"\nИтого найдено писем с ключевыми словами: {len(found_emails)}")

# Создаем новый Excel файл
wb = Workbook()
ws = wb.active
ws.title = "Найденные письма"

# Заголовки столбцов
headers = ["Дата", "Email отправителя", "Тема", "Тело письма", "Вложения", "Ключевые слова"]
ws.append(headers)

# Форматируем заголовки (жирный шрифт)
for cell in ws[1]:
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')

# Добавляем данные
for email_data in found_emails:
    ws.append([
        email_data['date'],
        email_data['from_email'],
        email_data['subject'],
        email_data['body'],
        email_data['has_attachments'],
        email_data['keywords']
    ])

# Автоматически подстраиваем ширину столбцов
for column in ws.columns:
    max_length = 0
    column_letter = column[0].column_letter
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = min(max_length + 2, 50)  # Максимум 50 символов
    ws.column_dimensions[column_letter].width = adjusted_width

# Сохраняем файл
filename = f"parsed_emails_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
wb.save(filename)
print(f"\n✓ Данные сохранены в файл: {filename}")
print(f"✓ Всего писем: {len(found_emails)}")

# Отключаемся
mail.close()
mail.logout()