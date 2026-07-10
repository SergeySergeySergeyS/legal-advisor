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
        # Формируем промт с более детальными инструкциями
        prompt = f"""Ты — юридический помощник. Проанализируй консультацию и создай детальный документ.

КАТЕГОРИЯ: {category}

ИСТОРИЯ ДИАЛОГА:
"""
        for msg in chat_history[-10:]:
            if msg["role"] == "user":
                prompt += f"ВОПРОС КЛИЕНТА: {msg['content']}\n"
            else:
                prompt += f"ОТВЕТ ЮРИСТА: {msg['content']}\n"
        
        prompt += """
ЗАДАЧА: Извлеки МАКСИМУМ деталей из ответа юриста и создай структуру документа.

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ (каждое поле с новой строки, БЕЗ markdown):

ТИП: [претензия/жалоба/заявление/иск]
КОМУ: [конкретное наименование организации ИЛИ должность, например: "Директору ООО 'ТелефонМаркет'" или "В Государственную инспекцию труда"]
ОТ: [если в вопросе есть ФИО/адрес/телефон — укажи, иначе оставь ПУСТЫМ]
СИТУАЦИЯ: [ДЕТАЛЬНОЕ описание ситуации из вопроса клиента + факты из ответа юриста, 4-6 предложений]
ЗАКОНЫ: [ВСЕ статьи законов, упомянутые в ответе юриста, с названиями, например: "Статья 18 Закона РФ 'О защите прав потребителей' от 07.02.1992 № 2300-1; Статья 503 ГК РФ"]
ТРЕБОВАНИЯ: [КОНКРЕТНЫЕ требования на основе ситуации, нумерованный список, например: "1. Вернуть денежные средства в размере 15000 рублей; 2. Выплатить неустойку в размере 1% от стоимости товара за каждый день просрочки"]
ПРИЛОЖЕНИЯ: [КОНКРЕТНЫЙ список документов, упомянутых в ответе юриста, нумерованный список]

КРИТИЧЕСКИ ВАЖНО:
- ЗАКОНЫ: Выпиши ВСЕ статьи законов из ответа юриста (если упомянуты ст. 18, 22, 13 ЗоЗПП — укажи ВСЕ)
- СИТУАЦИЯ: Используй конкретные факты из вопроса (марка товара, сумма, сроки)
- ТРЕБОВАНИЯ: Формулируй КОНКРЕТНЫЕ требования (суммы, сроки, проценты) на основе законов из ответа
- ПРИЛОЖЕНИЯ: Перечисли ВСЕ документы, упомянутые юристом (чек, гарантия, акт осмотра и т.д.)
- Если в вопросе НЕТ конкретных данных (ФИО, адрес) — оставь поле пустым для заполнения клиентом"""

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
