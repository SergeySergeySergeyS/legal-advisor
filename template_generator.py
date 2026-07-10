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
    """Парсит текстовый ответ в словарь — улучшенная версия"""
    result = get_default_template()
    
    try:
        lines = text.split('\n')
        current_section = None
        current_content = []
        
        # Функция для сохранения накопленного содержимого
        def save_section():
            nonlocal current_section, current_content
            if current_section and current_content:
                content = '\n'.join(current_content).strip()
                if content:
                    result[current_section] = content
            current_content = []
        
        for line in lines:
            line_stripped = line.strip()
            if not line_stripped:
                continue
            
            # Определяем начало новой секции по ключевым словам
            line_upper = line_stripped.upper()
            
            # ТИП документа
            if any(x in line_upper for x in ['ТИП:', 'ТИП ДОКУМЕНТА:', 'DOCTYPE:']):
                save_section()
                current_section = 'document_type'
                content = re.sub(r'^ТИП( ДОКУМЕНТА)?:\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            # КОМУ
            elif any(x in line_upper for x in ['КОМУ:', 'АДРЕСАТ:', 'В:', 'АДРЕСАТУ:']):
                save_section()
                current_section = 'recipient'
                content = re.sub(r'^(КОМУ|АДРЕСАТ|В|АДРЕСАТУ):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            # ОТ КОГО
            elif any(x in line_upper for x in ['ОТ:', 'ОТ КОГО:', 'ЗАЯВИТЕЛЬ:', 'ИСТЕЦ:']):
                save_section()
                current_section = 'sender'
                content = re.sub(r'^(ОТ( КОГО)?|ЗАЯВИТЕЛЬ|ИСТЕЦ):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            # СИТУАЦИЯ
            elif any(x in line_upper for x in ['СИТУАЦИЯ:', 'ОПИСАНИЕ:', 'ФАКТЫ:', 'ОБСТОЯТЕЛЬСТВА:']):
                save_section()
                current_section = 'situation'
                content = re.sub(r'^(СИТУАЦИЯ|ОПИСАНИЕ|ФАКТЫ|ОБСТОЯТЕЛЬСТВА):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            # ЗАКОНЫ / ПРАВОВОЕ ОБОСНОВАНИЕ
            elif any(x in line_upper for x in ['ЗАКОНЫ:', 'ПРАВОВОЕ ОБОСНОВАНИЕ', 'ПРАВОВАЯ БАЗА', 
                                                'СТАТЬИ:', 'НОРМАТИВНАЯ БАЗА', 'ОБОСНОВАНИЕ:']):
                save_section()
                current_section = 'legal_basis'
                content = re.sub(r'^(ЗАКОНЫ|ПРАВОВОЕ ОБОСНОВАНИЕ|ПРАВОВАЯ БАЗА|СТАТЬИ|НОРМАТИВНАЯ БАЗА|ОБОСНОВАНИЕ):\s*', 
                                '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            # ТРЕБОВАНИЯ
            elif any(x in line_upper for x in ['ТРЕБОВАНИЯ:', 'ТРЕБУЮ:', 'ПРОШУ:', 'ПРОШУ ПРИНЯТЬ']):
                save_section()
                current_section = 'requirements'
                content = re.sub(r'^(ТРЕБОВАНИЯ|ТРЕБУЮ|ПРОШУ|ПРОШУ ПРИНЯТЬ):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            # ПРИЛОЖЕНИЯ
            elif any(x in line_upper for x in ['ПРИЛОЖЕНИЯ:', 'ПРИЛАГАЮ:', 'ПРИЛОЖЕННЫЕ ДОКУМЕНТЫ']):
                save_section()
                current_section = 'attachments'
                content = re.sub(r'^(ПРИЛОЖЕНИЯ|ПРИЛАГАЮ|ПРИЛОЖЕННЫЕ ДОКУМЕНТЫ):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            # Если мы внутри секции — добавляем строку
            elif current_section:
                current_content.append(line_stripped)
        
        # Сохраняем последнюю секцию
        save_section()
        
        # Постобработка: очищаем от markdown и лишних символов
        for key in result:
            if result[key]:
                # Убираем markdown
                result[key] = result[key].replace('**', '').replace('*', '')
                result[key] = result[key].replace('__', '').replace('_', '')
                # Убираем лишние пробелы
                result[key] = re.sub(r'\s+', ' ', result[key]).strip()
        
        print(f"✅ Распознанные поля: {list(result.keys())}")
        print(f"📄 Результат парсинга: {result}")
        
    except Exception as e:
        print(f"❌ Ошибка парсинга: {e}")
    
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
