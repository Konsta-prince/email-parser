import imaplib
import json
import email
from email.header import decode_header
from datetime import datetime, timedelta
import re
import os
import sys
import logging  # НОВОЕ: модуль для логирования
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# ==========================================
# НАСТРОЙКА ЛОГИРОВАНИЯ (НОВОЕ)
# ==========================================
def setup_logging():
    """Настраивает запись логов в файл log.txt"""
    # Определяем папку, где лежит программа (работает и для .exe, и для .py)
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    log_file = os.path.join(base_path, 'log.txt')
    
    # Настраиваем логгер
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

# ==========================================
# БЛОК 1: ВСЕ ФУНКЦИИ (ИНСТРУМЕНТЫ)
# ==========================================

def load_config():
    """Читает настройки из config.json с проверкой ошибок"""
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    config_path = os.path.join(base_path, 'config.json')
    
    # НОВОЕ: проверяем, существует ли файл
    if not os.path.exists(config_path):
        raise FileNotFoundError(
            f"Файл настроек 'config.json' не найден в папке: {base_path}\n"
            "Убедитесь, что config.json лежит в той же папке, что и программа."
        )
    
    # НОВОЕ: обрабатываем ошибки чтения и парсинга JSON
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
    except json.JSONDecodeError as e:
        raise ValueError(
            f"Ошибка в файле config.json: {e}\n"
            "Проверьте, что вы не удалили кавычки, запятые или скобки. "
            "Формат должен быть точно как в примере."
        )
    except Exception as e:
        raise Exception(f"Не удалось прочитать файл настроек: {e}")
    
    return config

def connect_to_mail(server, email_addr, password):
    """Подключается к почтовому серверу с обработкой ошибок"""
    try:
        mail = imaplib.IMAP4_SSL(server)
        mail.login(email_addr, password)
        mail.select('INBOX')
        return mail
    except imaplib.IMAP4.error as e:
        error_msg = str(e)
        if 'AUTHENTICATIONFAILED' in error_msg or 'LOGIN' in error_msg:
            raise ConnectionError(
                "Не удалось войти в почту. Проверьте:\n"
                "1. Правильно ли указан email в config.json\n"
                "2. Используете ли вы 'Пароль приложения' (обычный пароль от почты не подойдет)\n"
                "3. Включен ли IMAP в настройках вашей почты"
            )
        else:
            raise ConnectionError(f"Ошибка подключения к серверу {server}: {e}")
    except imaplib.IMAP4.abort as e:
        raise ConnectionError(f"Соединение с сервером прервано: {e}")
    except Exception as e:
        raise ConnectionError(f"Неизвестная ошибка при подключении: {e}")

# ==========================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (ОЧИСТКА ДАННЫХ)
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

# ==========================================
# ОСНОВНЫЕ ФУНКЦИИ РАБОТЫ
# ==========================================

def search_and_process_emails(mail, date_from, keywords):
    """Подключается к почте, ищет письма по ключевым словам"""
    logging.info(f"Начинаем поиск писем с {date_from}")
    logging.info(f"Ключевые слова: {keywords}")
    
    print(f"Ищем письма с {date_from}...")
    print(f"Ключевые слова: {keywords}")
    print("-" * 50)

    try:
        status, messages = mail.search(None, f'SINCE {date_from}')
    except Exception as e:
        logging.error(f"Ошибка при поиске писем: {e}")
        print(f"❌ Ошибка при поиске писем: {e}")
        return []

    if status != 'OK':
        logging.warning("Сервер вернул статус не OK при поиске")
        print("Не удалось найти письма")
        return []

    email_ids = messages[0].split()
    logging.info(f"Найдено писем за период: {len(email_ids)}")
    print(f"Найдено писем за период: {len(email_ids)}")
    print("-" * 50)

    found_emails = []
    for email_id in email_ids:
        try:
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
                logging.info(f"Совпадение: {subject} (от: {from_email})")
        except Exception as e:
            logging.error(f"Ошибка при обработке письма {email_id}: {e}")
            continue

    logging.info(f"Итого найдено писем с ключевыми словами: {len(found_emails)}")
    print(f"\nИтого найдено писем с ключевыми словами: {len(found_emails)}")
    return found_emails

def save_to_excel(found_emails):
    """Сохраняет найденные письма в Excel-файл"""
    if not found_emails:
        logging.info("Нет данных для сохранения в Excel")
        print("Нет данных для сохранения.")
        return None

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
    logging.info(f"Файл сохранен: {filename}")
    print(f"\n✓ Данные сохранены в файл: {filename}")
    return filename

# ==========================================
# БЛОК 2: ОСНОВНОЙ КОД (ТОЧКА ВХОДА)
# ==========================================

def main():
    """Главная функция программы"""
    print("=" * 50)
    print("  ПАРСЕР ПОЧТЫ - Версия 2.0")
    print("=" * 50)
    
    # Настраиваем логирование
    setup_logging()
    logging.info("=" * 30)
    logging.info("Программа запущена")
    
    try:
        # 1. Загружаем настройки
        logging.info("Загрузка настроек из config.json...")
        config = load_config()
        
        email_settings = config['email_settings']
        search_settings = config['search_settings']
        
        # 2. Готовим переменные
        imap_server = email_settings['imap_server'].strip()
        email_address = email_settings['email_address'].strip()
        password = email_settings['password'].strip()
        days_to_search = search_settings['days_to_search']
        keywords = search_settings['keywords']
        
        # НОВОЕ: проверяем, что ключевые слова не пустые
        if not keywords:
            logging.warning("Список ключевых слов пуст")
            print("⚠️ ВНИМАНИЕ: В config.json не указаны ключевые слова!")
            print("   Программа найдет все письма за период, но не будет их фильтровать.")
            print("   Рекомендуется добавить ключевые слова в config.json.\n")
        
        logging.info(f"Почта: {email_address}")
        logging.info(f"Сервер: {imap_server}")
        logging.info(f"Период поиска: {days_to_search} дней")
        
        date_from = (datetime.now() - timedelta(days=days_to_search)).strftime("%d-%b-%Y")

        # 3. Подключаемся к почте
        print(f"\n📧 Подключаемся к {imap_server}...")
        logging.info(f"Подключение к {imap_server}...")
        mail = connect_to_mail(imap_server, email_address, password)
        print("✓ Успешно подключено к почте")
        logging.info("Успешное подключение к почте")

        # 4. Ищем и обрабатываем письма
        found_emails = search_and_process_emails(mail, date_from, keywords)

        # 5. Сохраняем в Excel
        save_to_excel(found_emails)

        # 6. Отключаемся
        mail.close()
        mail.logout()
        logging.info("Отключение от почты")
        print("\n✓ Программа завершена успешно.")
        
    except FileNotFoundError as e:
        logging.error(f"Файл не найден: {e}")
        print(f"\n❌ ОШИБКА: {e}")
    except ValueError as e:
        logging.error(f"Ошибка в настройках: {e}")
        print(f"\n❌ ОШИБКА В НАСТРОЙКАХ: {e}")
    except ConnectionError as e:
        logging.error(f"Ошибка подключения: {e}")
        print(f"\n❌ ОШИБКА ПОДКЛЮЧЕНИЯ: {e}")
    except Exception as e:
        logging.error(f"Неизвестная ошибка: {e}", exc_info=True)
        print(f"\n❌ ПРОИЗОШЛА ОШИБКА: {e}")
        print("Если ошибка повторяется, отправьте файл log.txt разработчику.")
    
    finally:
        logging.info("Программа завершена")
        logging.info("=" * 30)
        # НОВОЕ: блокировка закрытия окна
        input("\nНажмите Enter, чтобы закрыть программу...")

if __name__ == "__main__":
    main()