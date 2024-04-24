from google.oauth2 import service_account
from googleapiclient.discovery import build
import re

SERVICE_ACCOUNT_FILE = '321.json'
DOCUMENT_ID = '1TdK-HwMSZ7BE8qYcqWM6CdmQ932EOg88lCllE0h28Ak'

def create_docs_service():
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=['https://www.googleapis.com/auth/documents']
    )
    service = build('docs', 'v1', credentials=credentials)
    return service

def get_document_content(service, document_id):
    document = service.documents().get(documentId=document_id).execute()
    content = document.get('body').get('content')
    if isinstance(content, list):
        content = '\n'.join([c.get('paragraph').get('elements')[0].get('textRun').get('content') for c in content if 'paragraph' in c])
    return content

def replace_words(text, replacements):
    for old_word, new_word in replacements.items():
        pattern = r'\b' + re.escape(old_word) + r'\b'
        text = re.sub(pattern, new_word, text)
    return text
 
def main():
    docs_service = create_docs_service()
    document_content = get_document_content(docs_service, DOCUMENT_ID)

    replacements = {
        "ЦК": "МК",
        "Центральная компания": "Маркетинговая компания",
        "ВЫСОЦКИЙ КОНСАЛТИНГ": "BUSINESS BOOSTER",
        "Visotsky Inc": "BUSINESS BOOSTER",
        "Visotsky Consulting": "BUSINESS BOOSTER",
        "VC Int": "BUSINESS BOOSTER",
        "ВК": "BUSINESS BOOSTER",
        "УК": "ОО",
        "Цель": "Видение ",
        "Замысел": "Миссия",
        "Инспекции": "Мониторинг",
        "Инспекции основ": "Мониторинг",
        "Tonnus": "BB Platform",
        "Статистика": "Метрика",
        "ИП": "Регламент",
        "ДУК": "Директива",
        "Проверочный список": "Чек-лист",
        "Оргсхема": "Оргструктура"
    }

    modified_text = replace_words(document_content, replacements)
    print(modified_text)

if __name__ == "__main__":
    print('Updated document')
    main()