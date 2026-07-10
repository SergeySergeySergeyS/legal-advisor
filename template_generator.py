"""
Модуль генерации юридических документов (претензий, жалоб, исков, ходатайств)
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


# === ДОПУСТИМЫЕ ЗАКОНЫ ПО КАТЕГОРИЯМ ===
VALID_LAWS_BY_CATEGORY = {
    "🛒 Защита прав потребителей": [
        "о защите прав потребителей", "зозпп", "2300-1",
        "гражданский кодекс", "гк рф",
        "потребител"
    ],
    "🏠 ЖКХ": [
        "жилищный кодекс", "жк рф",
        "постановление правительства", "№ 354", "№ 491", "№354", "№491",
        "гражданский кодекс", "гк рф",
        "управляющ", "жиль"
    ],
    "💼 Трудовые споры": [
        "трудовой кодекс", "тк рф",
        "коллективн", "заработн", "увольнен",
        "дисциплинарн"
    ],
    "🚗 Споры с ГИБДД": [
        "коап", "коап рф",
        "правила дорожного", "пдд",
        "безопасност.*дорож",
        "осаго", "каско",
        "транспортн",
        "административн"
    ]
}

# === ЗАПРЕЩЁННЫЕ ЗАКОНЫ ПО КАТЕГОРИЯМ ===
INVALID_LAWS_BY_CATEGORY = {
    "🛒 Защита прав потребителей": ["тк рф", "коап", "жилищн"],
    "🏠 ЖКХ": ["о защите прав потребителей", "тк рф", "коап"],
    "💼 Трудовые споры": ["о защите прав потребителей", "коап", "жилищн"],
    "🚗 Споры с ГИБДД": [
        "о защите прав потребителей", "зозпп", "2300-1",
        "тк рф", "трудов",
        "жилищн", "жк рф",
        "гражданск"
    ]
}


# === ПРОМТЫ ДЛЯ РАЗНЫХ ТИПОВ ДОКУМЕНТОВ ===
DOCUMENT_PROMPTS = {
    "претензия": """Ты — юридический помощник. Создай ПРЕТЕНЗИЮ к контрагенту.

КРИТИЧЕСКИ ВАЖНО:
- Используй ТОЛЬКО информацию из ПОСЛЕДНЕГО вопроса клиента
- НЕ выдумывай факты, которых нет в вопросе
- Если данные не указаны (ФИО, адрес) — оставь поле ПУСТЫМ

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ:

ТИП: претензия
КОМУ: [должность и наименование организации]
ОТ: [если есть ФИО/адрес/телефон — укажи, иначе оставь ПУСТЫМ]
СИТУАЦИЯ: [ДЕТАЛЬНОЕ описание 4-6 предложений от ПЕРВОГО ЛИЦА]
ЗАКОНЫ: [ВСЕ статьи законов из ответа юриста, через точку с запятой]
ТРЕБОВАНИЯ: [КОНКРЕТНЫЕ требования с суммами и сроками, нумерованный список]
ПРИЛОЖЕНИЯ: [КОНКРЕТНЫЙ список документов, нумерованный список]

ВАЖНО:
- Претензия адресуется КОНТРАГЕНТУ
- СИТУАЦИЯ: пиши ОТ ПЕРВОГО ЛИЦА
- ЗАКОНЫ: используй ТОЛЬКО релевантные статьи
- Не используй markdown""",

    "жалоба": """Ты — юридический помощник. Создай ЖАЛОБУ в государственный орган.

КРИТИЧЕСКИ ВАЖНО:
- Используй ТОЛЬКО информацию из ПОСЛЕДНЕГО вопроса клиента
- НЕ выдумывай факты, которых нет в вопросе
- Если данные не указаны — оставь поле ПУСТЫМ

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ:

ТИП: жалоба
КОМУ: [наименование госоргана и должность руководителя]
ОТ: [если есть ФИО/адрес/телефон — укажи, иначе оставь ПУСТЫМ]
СИТУАЦИЯ: [ДЕТАЛЬНОЕ описание 4-6 предложений от ПЕРВОГО ЛИЦА]
ЗАКОНЫ: [ВСЕ статьи законов из ответа юриста, через точку с запятой]
ТРЕБОВАНИЯ: [КОНКРЕТНЫЕ требования — ПРОВЕСТИ ПРОВЕРКУ, ПРИНЯТЬ МЕРЫ, нумерованный список]
ПРИЛОЖЕНИЯ: [КОНКРЕТНЫЙ список документов, нумерованный список]

ВАЖНО:
- Жалоба адресуется В ГОСОРГАН
- ТРЕБОВАНИЯ: формулируй как "ПРОВЕСТИ ПРОВЕРКУ", "ПРИНЯТЬ МЕРЫ"
- ЗАКОНЫ: используй ТОЛЬКО релевантные статьи
- Не используй markdown""",

    "иск": """Ты — юридический помощник. Создай ИСКОВОЕ ЗАЯВЛЕНИЕ в суд.

КРИТИЧЕСКИ ВАЖНО:
- Используй ТОЛЬКО информацию из ПОСЛЕДНЕГО вопроса клиента
- НЕ выдумывай факты, которых нет в вопросе
- Если данные не указаны (ФИО истца, адрес ответчика) — оставь поле ПУСТЫМ
- НЕ пиши шаблоны вроде "если известен — укажи конкретнее"

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ:

ТИП: исковое заявление
КОМУ: [наименование суда — если не указано, напиши "В районный суд по месту жительства истца"]
ИСТЕЦ: [ФИО, адрес, телефон — если не указано, оставь ПУСТЫМ]
ОТВЕТЧИК: [ФИО или наименование организации, адрес, ИНН — если не указано, оставь ПУСТЫМ]
СИТУАЦИЯ: [ДЕТАЛЬНОЕ описание 5-7 предложений от ПЕРВОГО ЛИЦА]
ЗАКОНЫ: [ВСЕ статьи законов из ответа юриста, через точку с запятой]
ЦЕНА_ИСКА: [сумма иска в рублях с расчётом]
ТРЕБОВАНИЯ: [КОНКРЕТНЫЕ требования к суду, нумерованный список]
ПРИЛОЖЕНИЯ: [КОНКРЕТНЫЙ список документов, нумерованный список]

ВАЖНО:
- Иск адресуется В СУД
- Обязательно укажи ИСТЦА и ОТВЕТЧИКА (если есть данные)
- ЦЕНА_ИСКА: рассчитай общую сумму
- ТРЕБОВАНИЯ: формулируй как "ПРОШУ СУД: 1. Взыскать..."
- В ПРИЛОЖЕНИЯХ обязательно: "Копия искового заявления для ответчика", "Квитанция об уплате госпошлины"
- ЗАКОНЫ: используй ТОЛЬКО релевантные статьи
- Не используй markdown""",

    "ходатайство": """Ты — юридический помощник. Создай ХОДАТАЙСТВО.

КРИТИЧЕСКИ ВАЖНО:
- Используй ТОЛЬКО информацию из ПОСЛЕДНЕГО вопроса клиента
- НЕ выдумывай факты
- Если данные не указаны — оставь поле ПУСТЫМ

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ:

ТИП: ходатайство
КОМУ: [наименование органа/организации]
ОТ: [если есть ФИО/адрес/телефон — укажи, иначе оставь ПУСТЫМ]
СИТУАЦИЯ: [ДЕТАЛЬНОЕ описание 3-5 предложений от ПЕРВОГО ЛИЦА]
ЗАКОНЫ: [ВСЕ статьи законов из ответа юриста, через точку с запятой]
ТРЕБОВАНИЯ: [КОНКРЕТНАЯ просьба, нумерованный список]
ПРИЛОЖЕНИЯ: [КОНКРЕТНЫЙ список документов, нумерованный список]

ВАЖНО:
- ТРЕБОВАНИЯ: формулируй как "ПРОШУ: 1. Предоставить..."
- ЗАКОНЫ: используй ТОЛЬКО релевантные статьи
- Не используй markdown"""
}


def analyze_chat_for_template(chat_history: list, category: str, auth_key: str, doc_type: str = "претензия") -> dict:
    """Анализирует историю чата и извлекает данные для документа"""
    try:
        prompt_template = DOCUMENT_PROMPTS.get(doc_type, DOCUMENT_PROMPTS["претензия"])
        
        # Берём ТОЛЬКО последний вопрос и последний ответ
        last_user_question = ""
        last_assistant_answer = ""
        
        for msg in reversed(chat_history):
            if msg["role"] == "user" and not last_user_question:
                last_user_question = msg['content']
            elif msg["role"] == "assistant" and not last_assistant_answer:
                last_assistant_answer = msg['content']
            if last_user_question and last_assistant_answer:
                break
        
        prompt = f"""КАТЕГОРИЯ: {category}

ПОСЛЕДНИЙ ВОПРОС КЛИЕНТА (используй ТОЛЬКО эту информацию):
{last_user_question}

ОТВЕТ ЮРИСТА (используй статьи законов из этого ответа):
{last_assistant_answer}

{prompt_template}

ВАЖНО: Создавай документ ТОЛЬКО на основе ПОСЛЕДНЕГО вопроса."""
        
        with GigaChat(
            credentials=auth_key,
            scope="GIGACHAT_API_PERS",
            verify_ssl_certs=False
        ) as giga:
            messages = [Messages(role=MessagesRole.USER, content=prompt)]
            response = giga.chat(Chat(messages=messages))
            result_text = safe_decode(response.choices[0].message.content)
        
        print(f"🔍 Сырой ответ ИИ для типа '{doc_type}':\n{result_text}\n")
        
        result = parse_template_text(result_text, doc_type)
        result['legal_basis'] = validate_legal_basis(result.get('legal_basis', ''), category)
        
        return result
        
    except Exception as e:
        print(f"❌ Ошибка анализа: {e}")
        return get_default_template(doc_type)


def validate_legal_basis(legal_basis: str, category: str) -> str:
    """Проверяет правовое обоснование на соответствие категории"""
    if not legal_basis or legal_basis == '[Статьи законов]':
        return legal_basis
    
    valid_keywords = VALID_LAWS_BY_CATEGORY.get(category, [])
    invalid_keywords = INVALID_LAWS_BY_CATEGORY.get(category, [])
    
    articles = [a.strip() for a in legal_basis.split(';') if a.strip()]
    
    filtered_articles = []
    for article in articles:
        article_lower = article.lower()
        is_invalid = any(kw in article_lower for kw in invalid_keywords)
        is_valid = any(re.search(kw, article_lower) for kw in valid_keywords) if valid_keywords else True
        
        if not is_invalid:
            filtered_articles.append(article)
        else:
            print(f"🚫 Удалена нерелевантная статья: {article}")
    
    if not filtered_articles:
        print(f"⚠️ Все статьи отфильтрованы, возвращаем оригинал")
        return legal_basis
    
    return '; '.join(filtered_articles)


def parse_template_text(text: str, doc_type: str = "претензия") -> dict:
    """Парсит текстовый ответ в словарь"""
    result = get_default_template(doc_type)
    
    try:
        lines = text.split('\n')
        current_section = None
        current_content = []
        
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
            
            line_upper = line_stripped.upper()
            
            if any(x in line_upper for x in ['ТИП:', 'ТИП ДОКУМЕНТА:']):
                save_section()
                current_section = 'document_type'
                content = re.sub(r'^ТИП( ДОКУМЕНТА)?:\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            elif any(x in line_upper for x in ['КОМУ:', 'АДРЕСАТ:', 'АДРЕСАТУ:', 'В СУД:', 'В:']):
                save_section()
                current_section = 'recipient'
                content = re.sub(r'^(КОМУ|АДРЕСАТ|АДРЕСАТУ|В СУД|В):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            elif any(x in line_upper for x in ['ИСТЕЦ:', 'ЗАЯВИТЕЛЬ:']):
                save_section()
                current_section = 'sender'
                content = re.sub(r'^(ИСТЕЦ|ЗАЯВИТЕЛЬ):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            elif any(x in line_upper for x in ['ОТ:', 'ОТ КОГО:']):
                save_section()
                current_section = 'sender'
                content = re.sub(r'^ОТ( КОГО)?:\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            elif any(x in line_upper for x in ['ОТВЕТЧИК:']):
                save_section()
                current_section = 'defendant'
                content = re.sub(r'^ОТВЕТЧИК:\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            elif any(x in line_upper for x in ['СИТУАЦИЯ:', 'ОПИСАНИЕ:', 'ФАКТЫ:', 'ОБСТОЯТЕЛЬСТВА:']):
                save_section()
                current_section = 'situation'
                content = re.sub(r'^(СИТУАЦИЯ|ОПИСАНИЕ|ФАКТЫ|ОБСТОЯТЕЛЬСТВА):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            elif any(x in line_upper for x in ['ЗАКОНЫ:', 'ПРАВОВОЕ ОБОСНОВАНИЕ', 'ПРАВОВАЯ БАЗА', 
                                                'СТАТЬИ:', 'НОРМАТИВНАЯ БАЗА', 'ОБОСНОВАНИЕ:']):
                save_section()
                current_section = 'legal_basis'
                content = re.sub(r'^(ЗАКОНЫ|ПРАВОВОЕ ОБОСНОВАНИЕ|ПРАВОВАЯ БАЗА|СТАТЬИ|НОРМАТИВНАЯ БАЗА|ОБОСНОВАНИЕ):\s*', 
                                '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            elif any(x in line_upper for x in ['ЦЕНА ИСКА:', 'ЦЕНА_ИСКА:', 'СУММА ИСКА:']):
                save_section()
                current_section = 'claim_amount'
                content = re.sub(r'^(ЦЕНА ИСКА|ЦЕНА_ИСКА|СУММА ИСКА):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            elif any(x in line_upper for x in ['ТРЕБОВАНИЯ:', 'ТРЕБУЮ:', 'ПРОШУ:', 'ПРОШУ СУД:', 'ХОДАТАЙСТВУЮ:']):
                save_section()
                current_section = 'requirements'
                content = re.sub(r'^(ТРЕБОВАНИЯ|ТРЕБУЮ|ПРОШУ СУД|ПРОШУ|ХОДАТАЙСТВУЮ):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            elif any(x in line_upper for x in ['ПРИЛОЖЕНИЯ:', 'ПРИЛАГАЮ:', 'ПРИЛОЖЕННЫЕ ДОКУМЕНТЫ']):
                save_section()
                current_section = 'attachments'
                content = re.sub(r'^(ПРИЛОЖЕНИЯ|ПРИЛАГАЮ|ПРИЛОЖЕННЫЕ ДОКУМЕНТЫ):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            elif current_section:
                current_content.append(line_stripped)
        
        save_section()
        
        for key in result:
            if result[key]:
                result[key] = result[key].replace('**', '').replace('*', '')
                result[key] = result[key].replace('__', '').replace('_', '')
                result[key] = re.sub(r'\s+', ' ', result[key]).strip()
        
        print(f"✅ Распознанные поля для '{doc_type}': {list(result.keys())}")
        
    except Exception as e:
        print(f"❌ Ошибка парсинга: {e}")
    
    return result


def get_default_template(doc_type: str = "претензия") -> dict:
    """Возвращает шаблон по умолчанию"""
    return {
        "document_type": doc_type.upper(),
        "recipient": "[Наименование организации/органа]",
        "sender": "[ФИО, адрес, телефон]",
        "situation": "[Описание ситуации]",
        "legal_basis": "[Статьи законов]",
        "requirements": "[Требования]",
        "attachments": "[Приложения]",
        "defendant": "",
        "claim_amount": ""
    }


def split_into_items(text: str) -> list:
    """Разбивает текст на пункты. УЛУЧШЕННАЯ ВЕРСИЯ: объединяет короткие строки и удаляет дубликаты."""
    if not text:
        return []
    
    lines = []
    lines_by_newline = [line.strip() for line in text.split('\n') if line.strip()]
    
    if len(lines_by_newline) > 1:
        lines = lines_by_newline
    else:
        lines = re.split(r'\s+(?=\d+[\.\)])\s*', text)
        lines = [line.strip() for line in lines if line.strip()]
    
    if len(lines) <= 1 and ';' in text:
        lines = [line.strip() for line in text.split(';') if line.strip()]
    
    if len(lines) == 0:
        lines = [text]
    
    # Объединяем короткие строки с предыдущими
    merged_lines = []
    for line in lines:
        starts_with_number = re.match(r'^\d+[\.\)]\s*', line)
        
        if starts_with_number:
            merged_lines.append(line)
        elif len(line) < 15 and merged_lines:
            merged_lines[-1] = merged_lines[-1] + ' ' + line
        else:
            merged_lines.append(line)
    
    # === НОВОЕ: Удаляем дубликаты ===
    unique_lines = []
    seen = set()
    for line in merged_lines:
        # Убираем номера для сравнения
        line_without_number = re.sub(r'^\d+[\.\)]\s*', '', line).strip().lower()
        if line_without_number not in seen:
            seen.add(line_without_number)
            unique_lines.append(line)
    
    return unique_lines


def generate_legal_document(template_data: dict, output_dir: str, doc_type: str = "претензия") -> str:
    """Создаёт Word-документ на основе данных шаблона"""
    doc = Document()
    
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    # === ШАПКА ===
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    recipient = template_data.get('recipient', '[Наименование]')
    # Убираем шаблонные подсказки
    if 'если известен' in recipient.lower() or 'укажи конкретнее' in recipient.lower():
        recipient = '[Наименование организации]'
    run = p.add_run(recipient)
    run.font.size = Pt(11)
    
    if doc_type == "иск":
        defendant = template_data.get('defendant', '').strip()
        # Убираем шаблонные подсказки
        if defendant and ('если известен' in defendant.lower() or 'оставь пустым' in defendant.lower()):
            defendant = ''
        if defendant:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(f"Ответчик: {defendant}")
            run.font.size = Pt(11)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sender = template_data.get('sender', '')
    # Убираем шаблонные подсказки
    if sender and ('оставь пустым' in sender.lower() or 'если неизвестны' in sender.lower()):
        sender = ''
    
    if doc_type == "иск":
        if sender:
            run = p.add_run(f"Истец: {sender}")
        else:
            run = p.add_run("Истец: ___________________________________\n(ФИО, адрес, телефон, паспорт)")
    else:
        if sender:
            run = p.add_run(f"от {sender}")
        else:
            run = p.add_run("от ___________________________________\n(ФИО, адрес, телефон)")
    run.font.size = Pt(11)
    
    doc.add_paragraph()
    
    # === ЗАГОЛОВОК ===
    doc_titles = {
        "претензия": "ПРЕТЕНЗИЯ",
        "жалоба": "ЖАЛОБА",
        "иск": "ИСКОВОЕ ЗАЯВЛЕНИЕ",
        "ходатайство": "ХОДАТАЙСТВО"
    }
    doc_title = doc_titles.get(doc_type, doc_type.upper())
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(doc_title)
    run.bold = True
    run.font.size = Pt(14)
    
    doc.add_paragraph()
    
    # === СИТУАЦИЯ ===
    situation = template_data.get('situation', '[Описание ситуации]')
    doc.add_paragraph(situation)
    doc.add_paragraph()
    
    # === ПРАВОВОЕ ОБОСНОВАНИЕ ===
    legal_basis = template_data.get('legal_basis', '').strip()
    if not legal_basis or legal_basis == '[Статьи законов]':
        legal_basis = ("В соответствии с действующим законодательством Российской Федерации, "
                      "а также на основании прав, предоставленных мне нормативными правовыми актами, "
                      "регулирующими данную ситуацию.")
    
    p = doc.add_paragraph()
    run = p.add_run("Правовое обоснование: ")
    run.bold = True
    p.add_run(legal_basis)
    doc.add_paragraph()
    
    # === ЦЕНА ИСКА (только для иска) ===
    if doc_type == "иск":
        claim_amount = template_data.get('claim_amount', '').strip()
        if claim_amount:
            p = doc.add_paragraph()
            run = p.add_run("Цена иска: ")
            run.bold = True
            p.add_run(claim_amount)
            doc.add_paragraph()
    
    # === ТРЕБОВАНИЯ ===
    req_titles = {
        "претензия": "На основании вышеизложенного ТРЕБУЮ:",
        "жалоба": "На основании вышеизложенного ПРОШУ:",
        "иск": "ПРОШУ СУД:",
        "ходатайство": "На основании вышеизложенного ХОДАТАЙСТВУЮ:"
    }
    req_title = req_titles.get(doc_type, "ТРЕБУЮ:")
    
    p = doc.add_paragraph()
    run = p.add_run(req_title)
    run.bold = True
    
    requirements = template_data.get('requirements', '').strip()
    if not requirements or requirements == '[Требования]':
        requirements = "Удовлетворить мои законные требования в соответствии с действующим законодательством РФ."
    
    req_lines = split_into_items(requirements)
    
    for i, line in enumerate(req_lines, 1):
        line = re.sub(r'^\d+[\.\)]\s*', '', line)
        if line:
            p = doc.add_paragraph(f"{i}. {line}")
            p.paragraph_format.left_indent = Pt(20)
    
    doc.add_paragraph()
    
    # === ПРИЛОЖЕНИЯ ===
    p = doc.add_paragraph()
    run = p.add_run("Приложения:")
    run.bold = True
    
    attachments = template_data.get('attachments', '').strip()
    if not attachments or attachments == '[Приложения]':
        attachments = "1. Копия данного документа\n2. Копии документов, подтверждающих требования"
    
    att_lines = split_into_items(attachments)
    
    if doc_type == "иск":
        att_text = ' '.join(att_lines).lower()
        if 'госпошлин' not in att_text:
            att_lines.append("Квитанция об уплате государственной пошлины")
        if 'копия искового' not in att_text and 'для ответчика' not in att_text:
            att_lines.append("Копия искового заявления для ответчика")
    
    for i, line in enumerate(att_lines, 1):
        line = re.sub(r'^\d+[\.\)]\s*', '', line)
        if line:
            p = doc.add_paragraph(f"{i}. {line}")
            p.paragraph_format.left_indent = Pt(20)
    
    doc.add_paragraph()
    doc.add_paragraph()
    
    # === ДАТА И ПОДПИСЬ ===
    p = doc.add_paragraph()
    p.add_run("Дата: ___________                    Подпись: ___________")
    
    # === СОХРАНЕНИЕ ===
    safe_name = re.sub(r'[^\w\-.а-яА-Я]', '_', doc_title)
    output_path = Path(output_dir) / f"{safe_name}.docx"
    doc.save(str(output_path))
    
    return str(output_path)
