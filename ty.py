import imaplib
import json

# Загружаем настройки из файла config.json
with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

# Достаем настройки почты
email_settings = config['email_settings']
imap_server = email_settings['imap_server']
email_address = email_settings['email_address']
password = email_settings['password']

print(f"Подключаемся к {imap_server}...")
print(f"Почта: {email_address}")

# Пытаемся подключиться
try:
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(email_address, password)
    print("✓ Успешно подключились к почте!")
    
    # Выбираем папку "Входящие"
    mail.select('INBOX')
    print("✓ Открыли папку 'Входящие'")
    
    # Отключаемся
    mail.close()
    mail.logout()
    print("✓ Отключились от почты")
    
except Exception as e:
    print(f"✗ Ошибка подключения: {e}")