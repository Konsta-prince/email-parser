import imaplib
import json
import email
from email.header import decode_header
from datetime import datetime, timedelta
import re
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# ==========================================
# БЛОК 1: ВСЕ ФУНКЦИИ (ИНСТРУМЕНТЫ)
# ==========================================

#Читает настройки из config.json
def load_config():
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    config_path = os.path.join(base_path, 'config.json')
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

#Функция для подключения к почтовому серверу 
def connect_to_mail(server, email_addr, password):
    mail = imaplib.IMAP4_SSL(server)
    mail.login(email_addr, password)
    mail.select('INBOX')
    return mail

# ==========================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (ОЧИСТКА ДАННЫХ)
# Эти функции превращают "сырые" данные письма 
# в чистый текст, правильные даты и email-адреса.
# ==========================================
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

def clean_html(text):
    clean_text = re.sub(r'<[^>]+>', '', text)
    clean_text = re.sub(r'\s+', ' ', clean_text).strip()
    return clean_text

def get_email_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))
            if "attachment" in content_disposition:
                continue
            if content_type == "text/plain":
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or 'utf-8'
                    body = payload.decode(charset, errors='ignore')
                    break
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
    return clean_html(body)

def has_attachments(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if "attachment" in str(part.get("Content-Disposition", "")):
                return True
    return False

def extract_email(from_string):
    match = re.search(r'<([^>]+)>', from_string)
    if match:
        return match.group(1)
    if '@' in from_string:
        return from_string.strip()
    return from_string

def format_date(date_string):
    try:
        date_obj = datetime.strptime(date_string, "%a, %d %b %Y %H:%M:%S %z")
        return date_obj.strftime("%Y-%m-%d %H:%M:%S")
    except:
        try:
            date_obj = datetime.strptime(date_string, "%d %b %Y %H:%M:%S %z")
            return date_obj.strftime("%Y-%m-%d %H:%M:%S")
        except:
            return date_string
#===========================================

#Подключение к почте, сбор темы,текста и даты, сохранение в список 
def search_and_process_emails(mail, date_from, keywords):
    print(f"Ищем письма с {date_from}...")
    print(f"Ключевые слова: {keywords}")
    print("-" * 50)

    status, messages = mail.search(None, f'SINCE {date_from}')
    if status != 'OK':
        print("Не удалось найти письма")
        return []

    email_ids = messages[0].split()
    print(f"Найдено писем за период: {len(email_ids)}")
    print("-" * 50)

    found_emails = []
    for email_id in email_ids:
        status, msg_data = mail.fetch(email_id, '(RFC822)')
        if status != 'OK':
            continue
        
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)
        
        subject = decode_header_value(msg.get('Subject', ''))
        from_raw = decode_header_value(msg.get('From', ''))
        from_email = extract_email(from_raw)
        date_formatted = format_date(msg.get('Date', ''))
        body = get_email_body(msg)
        has_attach = "Да" if has_attachments(msg) else "Нет"
        
        found_keywords = [kw for kw in keywords if kw.lower() in body.lower() or kw.lower() in subject.lower()]
        
        if found_keywords:
            found_emails.append({
                'date': date_formatted,
                'from_email': from_email,
                'subject': subject,
                'body': body[:500],
                'has_attachments': has_attach,
                'keywords': ', '.join(found_keywords)
            })
            print(f"✓ Найдено совпадение: {subject}")

    print(f"\nИтого найдено писем с ключевыми словами: {len(found_emails)}")
    return found_emails

# Функция для создания и сохранения отчета в Excel с форматированием
def save_to_excel(found_emails):
    if not found_emails:
        print("Нет данных для сохранения.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Найденные письма"

    headers = ["Дата", "Email отправителя", "Тема", "Тело письма", "Вложения", "Ключевые слова"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for email_data in found_emails:
        ws.append([
            email_data['date'],
            email_data['from_email'],
            email_data['subject'],
            email_data['body'],
            email_data['has_attachments'],
            email_data['keywords']
        ])

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

    filename = f"parsed_emails_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(filename)
    print(f"\n✓ Данные сохранены в файл: {filename}")


# ==========================================
# БЛОК 2: ОСНОВНОЙ КОД (ВЫЗОВ ИНСТРУМЕНТОВ)
# ==========================================

if __name__ == "__main__":
    # 1. Загружаем настройки
    config = load_config()
    email_settings = config['email_settings']
    search_settings = config['search_settings']

    # 2. Готовим переменные
    imap_server = email_settings['imap_server']
    email_address = email_settings['email_address']
    password = email_settings['password']
    days_to_search = search_settings['days_to_search']
    keywords = search_settings['keywords']
    
    date_from = (datetime.now() - timedelta(days=days_to_search)).strftime("%d-%b-%Y")

    # 3. Подключаемся к почте
    mail = connect_to_mail(imap_server, email_address, password)

    # 4. Ищем и обрабатываем письма
    found_emails = search_and_process_emails(mail, date_from, keywords)

    # 5. Сохраняем в Excel
    save_to_excel(found_emails)

    # 6. Отключаемся
    mail.close()
    mail.logout()
    print("✓ Программа завершена.")