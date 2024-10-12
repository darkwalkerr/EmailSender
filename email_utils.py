# email_utils.py

import re
import json
import os

def is_valid_email(email):
    """
    Проверяет корректность email адреса.
    Возвращает True, если email корректен, иначе False.
    """
    regex = r"[^@]+@[^@]+\.[^@]+"
    return re.match(regex, email) is not None

def save_sender_credentials(email, login, password):
    """
    Сохраняет email отправителя, логин и пароль в файл config.json
    """
    config = {
        'sender_email': email,
        'login': login,
        'password': password
    }
    with open('config.json', 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

def load_sender_credentials():
    """
    Загружает email отправителя, логин и пароль из файла config.json
    Если файл не существует, возвращает пустые строки
    """
    if os.path.exists('config.json'):
        with open('config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)
            email = config.get('sender_email', '')
            login = config.get('login', '')
            password = config.get('password', '')
            return email, login, password
    else:
        return '', '', ''

def save_recipients(recipients):
    """
    Сохраняет список получателей в файл recipients.json
    """
    with open('recipients.json', 'w', encoding='utf-8') as f:
        json.dump(recipients, f, ensure_ascii=False, indent=4)

def load_recipients():
    """
    Загружает список получателей из файла recipients.json
    Если файл не существует, возвращает пустой список
    """
    if os.path.exists('recipients.json'):
        with open('recipients.json', 'r', encoding='utf-8') as f:
            recipients = json.load(f)
            return recipients
    else:
        return []
