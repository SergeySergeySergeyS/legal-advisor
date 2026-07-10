"""
Модуль генерации юридических документов (претензий, жалоб, заявлений)
"""
from docx import Document
from docx.shared import Pt, RGBColor
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
    Анализирует историю чата и определяет тип документа + извлекает данные
    
    Returns:
        dict с ключами:
        - document_type: тип документа (претензия, жалоба, иск)
        - recipient: кому адресован
        - sender: от кого
        - situation: описание ситуации
        - legal_basis: юридическое обоснование
        - requirements: требования
        - attachments: приложения
    """
    try:
        # Формируем промт для анализа
        prompt = f"""Проанализируй следующую юридическую консультацию и извлеки данные для создания документа.

КАТЕГОРИЯ: {category}

ИСТОРИЯ ДИАЛОГА:
"""
        
        for msg in chat_history[-10:]:
            if msg["role"] == "user":
                prompt += f"Вопрос: {msg['content']}\n"
            else:
                prompt += f"Ответ: {msg['content']}\n"
        
        prompt += """
ВЕРНИ РЕЗУЛЬТАТ В ФОРМАТЕ JSON (без markdown, просто текст):
{
    "document_type": "тип документа (претензия/жалоба/заявление/иск)",
    "recipient": "кому адресован (организация, должность)",
    "sender": "от кого (ФИО, адрес, телефон)",
    "situation": "краткое описание ситуации (2-3 предложения)",
    "legal_basis": "юридическое обоснование (статьи законов)",
    "requirements": "требования (нумерованный список)",
    "attachments": "приложения (список документов)"
}

ВАЖНО:
- Если данные не указаны явно (ФИО, адрес) — оставь пустые поля для заполнения
- document_type определи по контексту (претензия продавцу, жалоба в госорган, иск в суд)
- situation сформулируй кратко на основе вопроса пользователя
- legal_basis возьми из ответа ИИ (статьи законов)
- requirements сформулируй на основе ситуации
- attachments определи из контекста"""

        with GigaChat(
            credentials=auth_key,
            scope="GIGACHAT_API_PERS",
            verify_ssl_certs=False
        ) as giga:
            messages = [Messages(role=MessagesRole.USER, content=prompt)]
            response = giga.chat(Chat(messages=messages))
            result_text = safe_decode(response.choices[0].message.content)
        
        # Парсим JSON из ответа (простой способ)
        result = parse_json_from_text(result_text)
        
        return result
        
    except Exception as e:
        print(f"Ошибка анализа чата: {e}")
        return {
            "document_type": "претензия",
            "recipient": "[Наименование организации]",
            "sender": "[ФИО, адрес, телефон]",
            "situation": "[Описание ситуации]",
            "legal_basis": "[Статьи законов]",
            "requirements": "[Требования]",
            "attachments": "[Приложения]"
        }


def parse_json_from_text(text: str) -> dict:
    """Извлекает JSON из текста ответа ИИ"""
    try:
        # Ищем JSON в тексте
        json_match = re.search(r'\{[^{}]*\}', text, re.DOTALL)
        if json_match:
            json_str = json_match.group(0)
            # Простая очистка
            json_str = json_str.replace('\n', ' ').replace('\r', '')
            # Парсим вручную (без import json для простоты)
            result = {}
            pairs = re.findall(r'"([^"]+)"\s*:\s*"([^"]*)"', json_str)
            for key, value in pairs:
                result[key] = value
            return result
    except Exception as e:
        print(f"Ошибка парсинга JSON: {e}")
    
    # Если не получилось — возвращаем шаблон
    return {
        "document_type": "претензия",
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
    
    Args:
        template_data: словарь с данными документа
        output_dir: директория для сохранения
    
    Returns:
        str: путь к созданному файлу
    """
    doc = Document()
    
    # === СТИЛИ ===
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    # === ШАПКА ДОКУМЕНТА ===
    # Кому
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(f"{template_data.get('recipient', '[Наименование организации]')}")
    run.font.size = Pt(11)
    
    # От кого
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(f"от {template_data.get('sender', '[ФИО, адрес, телефон]')}")
    run.font.size = Pt(11)
    
    doc.add_paragraph()  # Пустая строка
    
    # === ЗАГОЛОВОК ===
    doc_type = template_data.get('document_type', 'ПРЕТЕНЗИЯ').upper()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(doc_type)
    run.bold = True
    run.font.size = Pt(14)
    
    doc.add_paragraph()  # Пустая строка
    
    # === ОПИСАНИЕ СИТУАЦИИ ===
    situation = template_data.get('situation', '[Описание ситуации]')
    doc.add_paragraph(situation)
    
    doc.add_paragraph()  # Пустая строка
    
    # === ЮРИДИЧЕСКОЕ ОБОСНОВАНИЕ ===
    legal_basis = template_data.get('legal_basis', '[Статьи законов]')
    p = doc.add_paragraph()
    run = p.add_run("Правовое обоснование: ")
    run.bold = True
    p.add_run(legal_basis)
    
    doc.add_paragraph()  # Пустая строка
    
    # === ТРЕБОВАНИЯ ===
    p = doc.add_paragraph()
    run = p.add_run("На основании вышеизложенного ТРЕБУЮ:")
    run.bold = True
    
    requirements = template_data.get('requirements', '[Требования]')
    # Разбиваем требования по пунктам
    req_lines = requirements.split('\n')
    for i, line in enumerate(req_lines, 1):
        line = line.strip()
        if line and not line.startswith(tuple('0123456789')):
            doc.add_paragraph(f"{i}. {line}", style='List Number')
        elif line:
            doc.add_paragraph(line)
    
    doc.add_paragraph()  # Пустая строка
    
    # === ПРИЛОЖЕНИЯ ===
    p = doc.add_paragraph()
    run = p.add_run("Приложения:")
    run.bold = True
    
    attachments = template_data.get('attachments', '[Приложения]')
    att_lines = attachments.split('\n')
    for i, line in enumerate(att_lines, 1):
        line = line.strip()
        if line and not line.startswith(tuple('0123456789')):
            doc.add_paragraph(f"{i}. {line}")
        elif line:
            doc.add_paragraph(line)
    
    doc.add_paragraph()  # Пустая строка
    doc.add_paragraph()  # Пустая строка
    
    # === ДАТА И ПОДПИСЬ ===
    p = doc.add_paragraph()
    p.add_run("Дата: ___________                    Подпись: ___________")
    
    # === СОХРАНЕНИЕ ===
    output_path = Path(output_dir) / f"{doc_type}_документ.docx"
    doc.save(str(output_path))
    
    return str(output_path)
