"""
Модуль генерации юридических документов (претензий, жалоб, заявлений)
"""
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pathlib import Path
import re
from gigachat import GigaChat
from gigachat.models import Chat, Messages, MessagesRole


def safe_decode(content):
    """Безопасно декодирует байты в строку"""
    if isinstance(content, bytes):
        return content.decode('utf-8', errors='ignore')
    return content


def analyze_chat_for_template(chat_history: list, category: str, auth_key: str) -> dict:
    """
    Анализирует историю чата и извлекает данные для документа
    """
    try:
        # Формируем промт
        prompt = f"""Проанализируй юридическую консультацию и извлеки данные для документа.

КАТЕГОРИЯ: {category}

ИСТОРИЯ ДИАЛОГА:
"""
        for msg in chat_history[-10:]:
            if msg["role"] == "user":
                prompt += f"Вопрос: {msg['content']}\n"
            else:
                prompt += f"Ответ: {msg['content']}\n"
        
        prompt += """
ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ (каждое поле с новой строки):

ТИП: [претензия/жалоба/заявление]
КОМУ: [наименование организации, должность]
ОТ: [ФИО, адрес, телефон - если не указано, оставь пустым]
СИТУАЦИЯ: [краткое описание 2-3 предложения]
ЗАКОНЫ: [статьи законов]
ТРЕБОВАНИЯ: [нумерованный список требований]
ПРИЛОЖЕНИЯ: [список документов]

ВАЖНО:
- Если ФИО/адрес не указаны — оставь поле пустым
- ТИП определи по контексту
- СИТУАЦИЯ кратко на основе вопроса
- ЗАКОНЫ из ответа ИИ
- ТРЕБОВАНИЯ по ситуации"""

        with GigaChat(
            credentials=auth_key,
            scope="GIGACHAT_API_PERS",
            verify_ssl_certs=False
        ) as giga:
            messages = [Messages(role=MessagesRole.USER, content=prompt)]
            response = giga.chat(Chat(messages=messages))
            result_text = safe_decode(response.choices[0].message.content)
        
        # Парсим ответ
        result = parse_template_text(result_text)
        return result
        
    except Exception as e:
        print(f"Ошибка анализа: {e}")
        return get_default_template()


def parse_template_text(text: str) -> dict:
    """Парсит текстовый ответ в словарь"""
    result = get_default_template()
    
    try:
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if line.startswith('ТИП:'):
                result['document_type'] = line.replace('ТИП:', '').strip()
            elif line.startswith('КОМУ:'):
                result['recipient'] = line.replace('КОМУ:', '').strip()
            elif line.startswith('ОТ:'):
                result['sender'] = line.replace('ОТ:', '').strip()
            elif line.startswith('СИТУАЦИЯ:'):
                result['situation'] = line.replace('СИТУАЦИЯ:', '').strip()
            elif line.startswith('ЗАКОНЫ:'):
                result['legal_basis'] = line.replace('ЗАКОНЫ:', '').strip()
            elif line.startswith('ТРЕБОВАНИЯ:'):
                result['requirements'] = line.replace('ТРЕБОВАНИЯ:', '').strip()
            elif line.startswith('ПРИЛОЖЕНИЯ:'):
                result['attachments'] = line.replace('ПРИЛОЖЕНИЯ:', '').strip()
    except Exception as e:
        print(f"Ошибка парсинга: {e}")
    
    return result


def get_default_template() -> dict:
    """Возвращает шаблон по умолчанию"""
    return {
        "document_type": "ПРЕТЕНЗИЯ",
        "recipient": "[Наименование организации]",
        "sender": "[ФИО, адрес, телефон]",
        "situation": "[Описание ситуации]",
        "legal_basis": "[Статьи законов]",
        "requirements": "[Требования]",
        "attachments": "[Приложения]"
    }


def generate_legal_document(template_data: dict, output_dir: str) -> str:
    """
    Создаёт Word-документ на основе данных шаблона
    """
    doc = Document()
    
    # === СТИЛИ ===
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    # === ШАПКА ===
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(template_data.get('recipient', '[Наименование организации]'))
    run.font.size = Pt(11)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sender = template_data.get('sender', '')
    if sender:
        run = p.add_run(f"от {sender}")
    else:
        run = p.add_run("от ___________________________________\n(ФИО, адрес, телефон)")
    run.font.size = Pt(11)
    
    doc.add_paragraph()
    
    # === ЗАГОЛОВОК ===
    doc_type = template_data.get('document_type', 'ПРЕТЕНЗИЯ').upper()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(doc_type)
    run.bold = True
    run.font.size = Pt(14)
    
    doc.add_paragraph()
    
    # === СИТУАЦИЯ ===
    situation = template_data.get('situation', '[Описание ситуации]')
    doc.add_paragraph(situation)
    doc.add_paragraph()
    
    # === ПРАВОВОЕ ОБОСНОВАНИЕ ===
    legal_basis = template_data.get('legal_basis', '[Статьи законов]')
    p = doc.add_paragraph()
    run = p.add_run("Правовое обоснование: ")
    run.bold = True
    p.add_run(legal_basis)
    doc.add_paragraph()
    
    # === ТРЕБОВАНИЯ ===
    p = doc.add_paragraph()
    run = p.add_run("На основании вышеизложенного ТРЕБУЮ:")
    run.bold = True
    
    requirements = template_data.get('requirements', '[Требования]')
    req_lines = [line.strip() for line in requirements.split('\n') if line.strip()]
    for i, line in enumerate(req_lines, 1):
        # Убираем номера, если они есть
        line = re.sub(r'^\d+[\.\)]\s*', '', line)
        if line:
            doc.add_paragraph(f"{i}. {line}")
    
    doc.add_paragraph()
    
    # === ПРИЛОЖЕНИЯ ===
    p = doc.add_paragraph()
    run = p.add_run("Приложения:")
    run.bold = True
    
    attachments = template_data.get('attachments', '[Приложения]')
    att_lines = [line.strip() for line in attachments.split('\n') if line.strip()]
    for i, line in enumerate(att_lines, 1):
        line = re.sub(r'^\d+[\.\)]\s*', '', line)
        if line:
            doc.add_paragraph(f"{i}. {line}")
    
    doc.add_paragraph()
    doc.add_paragraph()
    
    # === ДАТА И ПОДПИСЬ ===
    p = doc.add_paragraph()
    p.add_run("Дата: ___________                    Подпись: ___________")
    
    # === СОХРАНЕНИЕ ===
    safe_name = re.sub(r'[^\w\-.]', '_', doc_type)
    output_path = Path(output_dir) / f"{safe_name}.docx"
    doc.save(str(output_path))
    
    return str(output_path)
