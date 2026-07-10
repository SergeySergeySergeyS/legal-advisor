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


# === ДОПУСТИМЫЕ ЗАКОНЫ ПО КАТЕГОРИЯМ (для валидации) ===
VALID_LAWS_BY_CATEGORY = {
    "🛒 Защита прав потребителей": [
        "о защите прав потребителей", "зозпп", "2300-1",
        "гражданский кодекс", "гк рф", "гк рф",
        "потребитель"
    ],
    "🏠 ЖКХ": [
        "жилищный кодекс", "жк рф", "жк рф",
        "постановление правительства", "№ 354", "№ 491", "№354", "№491",
        "гражданский кодекс", "гк рф",
        "управляющ", "жиль"
    ],
    "💼 Трудовые споры": [
        "трудовой кодекс", "тк рф", "тк рф",
        "коллективн", "заработн", "увольнен",
        "дисциплинарн"
    ],
    "🚗 Споры с ГИБДД": [
        "коап", "коап рф",
        "правила дорожного", "пдд",
        "безопасност.*дорож",
        "осаго", "каско",
        "транспортн"
    ]
}

# === ЗАПРЕЩЁННЫЕ ЗАКОНЫ ПО КАТЕГОРИЯМ (явные ошибки) ===
INVALID_LAWS_BY_CATEGORY = {
    "🛒 Защита прав потребителей": ["тк рф", "коап", "жилищн"],
    "🏠 ЖКХ": ["о защите прав потребителей", "тк рф", "коап"],
    "💼 Трудовые споры": ["о защите прав потребителей", "коап", "жилищн"],
    "🚗 Споры с ГИБДД": ["о защите прав потребителей", "тк рф", "жилищн"]
}


# === ПРОМТЫ ДЛЯ РАЗНЫХ ТИПОВ ДОКУМЕНТОВ ===
DOCUMENT_PROMPTS = {
    "претензия": """Ты — юридический помощник. Создай ПРЕТЕНЗИЮ к контрагенту (продавцу, исполнителю, арендодателю).

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ:

ТИП: претензия
КОМУ: [должность и наименование организации]
ОТ: [если есть ФИО/адрес/телефон — укажи, иначе оставь ПУСТЫМ]
СИТУАЦИЯ: [ДЕТАЛЬНОЕ описание 4-6 предложений от ПЕРВОГО ЛИЦА "Я приобрёл..."]
ЗАКОНЫ: [ВСЕ статьи законов из ответа юриста с названиями, через точку с запятой]
ТРЕБОВАНИЯ: [КОНКРЕТНЫЕ требования с суммами и сроками, нумерованный список]
ПРИЛОЖЕНИЯ: [КОНКРЕТНЫЙ список документов, нумерованный список]

КРИТИЧЕСКИ ВАЖНО:
- Претензия адресуется КОНТРАГЕНТУ (продавцу, исполнителю, арендодателю)
- СИТУАЦИЯ: пиши ОТ ПЕРВОГО ЛИЦА ("Я приобрёл")
- ТРЕБОВАНИЯ: указывай КОНКРЕТНЫЕ суммы и сроки
- ЗАКОНЫ: используй ТОЛЬКО те статьи, которые реально относятся к данной ситуации. НЕ добавляй нерелевантные законы!
- Не используй markdown""",

    "жалоба": """Ты — юридический помощник. Создай ЖАЛОБУ в государственный орган.

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ:

ТИП: жалоба
КОМУ: [наименование госоргана и должность руководителя]
ОТ: [если есть ФИО/адрес/телефон — укажи, иначе оставь ПУСТЫМ]
СИТУАЦИЯ: [ДЕТАЛЬНОЕ описание 4-6 предложений от ПЕРВОГО ЛИЦА]
ЗАКОНЫ: [ВСЕ статьи законов из ответа юриста с названиями, через точку с запятой]
ТРЕБОВАНИЯ: [КОНКРЕТНЫЕ требования к госоргану — ПРОВЕСТИ ПРОВЕРКУ, ПРИНЯТЬ МЕРЫ и т.д., нумерованный список]
ПРИЛОЖЕНИЯ: [КОНКРЕТНЫЙ список документов, нумерованный список]

КРИТИЧЕСКИ ВАЖНО:
- Жалоба адресуется В ГОСОРГАН (Роспотребнадзор, ГИТ, Жилинспекция, ГИБДД, Прокуратура)
- ТРЕБОВАНИЯ: формулируй как "ПРОВЕСТИ ПРОВЕРКУ", "ПРИНЯТЬ МЕРЫ", "ВЫНЕСТИ ПРЕДПИСАНИЕ"
- СИТУАЦИЯ: пиши ОТ ПЕРВОГО ЛИЦА
- ЗАКОНЫ: используй ТОЛЬКО те статьи, которые реально относятся к данной ситуации. НЕ добавляй нерелевантные законы!
- Не используй markdown""",

    "иск": """Ты — юридический помощник. Создай ИСКОВОЕ ЗАЯВЛЕНИЕ в суд.

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ:

ТИП: исковое заявление
КОМУ: [наименование суда]
ИСТЕЦ: [ФИО, адрес, телефон, паспортные данные — если есть, иначе оставь ПУСТЫМ]
ОТВЕТЧИК: [ФИО или наименование организации, адрес — если есть, иначе оставь ПУСТЫМ]
СИТУАЦИЯ: [ДЕТАЛЬНОЕ описание 5-7 предложений от ПЕРВОГО ЛИЦА]
ЗАКОНЫ: [ВСЕ статьи законов из ответа юриста с названиями, через точку с запятой]
ЦЕНА_ИСКА: [сумма иска в рублях с расчётом]
ТРЕБОВАНИЯ: [КОНКРЕТНЫЕ требования к суду, нумерованный список]
ПРИЛОЖЕНИЯ: [КОНКРЕТНЫЙ список документов, нумерованный список]

КРИТИЧЕСКИ ВАЖНО:
- Иск адресуется В СУД
- Обязательно укажи ИСТЦА и ОТВЕТЧИКА
- ЦЕНА_ИСКА: рассчитай общую сумму с обоснованием
- ТРЕБОВАНИЯ: формулируй как "ПРОШУ СУД: 1. Взыскать... 2. Признать..."
- В ПРИЛОЖЕНИЯХ обязательно: "Копия искового заявления для ответчика", "Квитанция об уплате госпошлины"
- ЗАКОНЫ: используй ТОЛЬКО те статьи, которые реально относятся к данной ситуации. НЕ добавляй нерелевантные законы!
- Не используй markdown""",

    "ходатайство": """Ты — юридический помощник. Создай ХОДАТАЙСТВО в государственный орган или организацию.

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ:

ТИП: ходатайство
КОМУ: [наименование органа/организации и должность]
ОТ: [если есть ФИО/адрес/телефон — укажи, иначе оставь ПУСТЫМ]
СИТУАЦИЯ: [ДЕТАЛЬНОЕ описание 3-5 предложений от ПЕРВОГО ЛИЦА]
ЗАКОНЫ: [ВСЕ статьи законов из ответа юриста с названиями, через точку с запятой]
ТРЕБОВАНИЯ: [КОНКРЕТНАЯ просьба, нумерованный список]
ПРИЛОЖЕНИЯ: [КОНКРЕТНЫЙ список документов, нумерованный список]

КРИТИЧЕСКИ ВАЖНО:
- Ходатайство — это просьба о совершении действия
- ТРЕБОВАНИЯ: формулируй как "ПРОШУ: 1. Предоставить... 2. Перенести... 3. Выдать..."
- СИТУАЦИЯ: пиши ОТ ПЕРВОГО ЛИЦА
- ЗАКОНЫ: используй ТОЛЬКО те статьи, которые реально относятся к данной ситуации. НЕ добавляй нерелевантные законы!
- Не используй markdown
- ВАЖНО: даты и номера НЕ переноси на новую строку. Пиши "от 04.07.2026" целиком в одной строке."""
}


def analyze_chat_for_template(chat_history: list, category: str, auth_key: str, doc_type: str = "претензия") -> dict:
    """
    Анализирует историю чата и извлекает данные для документа
    """
    try:
        prompt_template = DOCUMENT_PROMPTS.get(doc_type, DOCUMENT_PROMPTS["претензия"])
        
        prompt = f"""КАТЕГОРИЯ: {category}

ИСТОРИЯ ДИАЛОГА:
"""
        for msg in chat_history[-10:]:
            if msg["role"] == "user":
                prompt += f"ВОПРОС КЛИЕНТА: {msg['content']}\n"
            else:
                prompt += f"ОТВЕТ ЮРИСТА: {msg['content']}\n"
        
        prompt += f"\n{prompt_template}"
        
        with GigaChat(
            credentials=auth_key,
            scope="GIGACHAT_API_PERS",
            verify_ssl_certs=False
        ) as giga:
            messages = [Messages(role=MessagesRole.USER, content=prompt)]
            response = giga.chat(Chat(messages=messages))
            result_text = safe_decode(response.choices[0].message.content)
        
        print(f"🔍 Сырой ответ ИИ для типа '{doc_type}':\n{result_text}\n")
        
        # Парсим ответ
        result = parse_template_text(result_text, doc_type)
        
        # ВАЛИДАЦИЯ: удаляем нерелевантные законы
        result['legal_basis'] = validate_legal_basis(result.get('legal_basis', ''), category)
        
        return result
        
    except Exception as e:
        print(f"❌ Ошибка анализа: {e}")
        return get_default_template(doc_type)


def validate_legal_basis(legal_basis: str, category: str) -> str:
    """
    Проверяет правовое обоснование на соответствие категории.
    Удаляет нерелевантные законы.
    """
    if not legal_basis or legal_basis == '[Статьи законов]':
        return legal_basis
    
    # Получаем списки допустимых и запрещённых законов
    valid_keywords = VALID_LAWS_BY_CATEGORY.get(category, [])
    invalid_keywords = INVALID_LAWS_BY_CATEGORY.get(category, [])
    
    # Разбиваем на отдельные статьи (по точке с запятой)
    articles = [a.strip() for a in legal_basis.split(';') if a.strip()]
    
    # Фильтруем статьи
    filtered_articles = []
    for article in articles:
        article_lower = article.lower()
        
        # Проверяем, есть ли запрещённые ключевые слова
        is_invalid = any(kw in article_lower for kw in invalid_keywords)
        
        # Проверяем, есть ли допустимые ключевые слова
        is_valid = any(kw in article_lower for kw in valid_keywords) if valid_keywords else True
        
        # Если статья не запрещена И (допустима ИЛИ нет списка допустимых)
        if not is_invalid:
            filtered_articles.append(article)
        else:
            print(f"🚫 Удалена нерелевантная статья: {article}")
    
    # Если всё удалили — возвращаем оригинал (лучше что-то, чем ничего)
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
            
            # ТИП документа
            if any(x in line_upper for x in ['ТИП:', 'ТИП ДОКУМЕНТА:']):
                save_section()
                current_section = 'document_type'
                content = re.sub(r'^ТИП( ДОКУМЕНТА)?:\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            # КОМУ
            elif any(x in line_upper for x in ['КОМУ:', 'АДРЕСАТ:', 'АДРЕСАТУ:', 'В СУД:', 'В:']):
                save_section()
                current_section = 'recipient'
                content = re.sub(r'^(КОМУ|АДРЕСАТ|АДРЕСАТУ|В СУД|В):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            # ИСТЕЦ (для иска)
            elif any(x in line_upper for x in ['ИСТЕЦ:', 'ЗАЯВИТЕЛЬ:']):
                save_section()
                current_section = 'sender'
                content = re.sub(r'^(ИСТЕЦ|ЗАЯВИТЕЛЬ):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            # ОТ КОГО (для других документов)
            elif any(x in line_upper for x in ['ОТ:', 'ОТ КОГО:']):
                save_section()
                current_section = 'sender'
                content = re.sub(r'^ОТ( КОГО)?:\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            # ОТВЕТЧИК (для иска)
            elif any(x in line_upper for x in ['ОТВЕТЧИК:']):
                save_section()
                current_section = 'defendant'
                content = re.sub(r'^ОТВЕТЧИК:\s*', '', line_stripped, flags=re.IGNORECASE)
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
            
            # ЦЕНА ИСКА (для иска)
            elif any(x in line_upper for x in ['ЦЕНА ИСКА:', 'ЦЕНА_ИСКА:', 'СУММА ИСКА:']):
                save_section()
                current_section = 'claim_amount'
                content = re.sub(r'^(ЦЕНА ИСКА|ЦЕНА_ИСКА|СУММА ИСКА):\s*', '', line_stripped, flags=re.IGNORECASE)
                current_content = [content]
            
            # ТРЕБОВАНИЯ
            elif any(x in line_upper for x in ['ТРЕБОВАНИЯ:', 'ТРЕБУЮ:', 'ПРОШУ:', 'ПРОШУ СУД:', 'ХОДАТАЙСТВУЮ:']):
                save_section()
                current_section = 'requirements'
                content = re.sub(r'^(ТРЕБОВАНИЯ|ТРЕБУЮ|ПРОШУ СУД|ПРОШУ|ХОДАТАЙСТВУЮ):\s*', '', line_stripped, flags=re.IGNORECASE)
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
        
        # Постобработка
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
    """
    Разбивает текст на пункты по номерам или переводам строк.
    УЛУЧШЕННАЯ ВЕРСИЯ: объединяет короткие строки (< 15 символов) с предыдущими.
    """
    if not text:
        return []
    
    lines = []
    
    # Сначала пробуем разбить по переводам строк
    lines_by_newline = [line.strip() for line in text.split('\n') if line.strip()]
    
    if len(lines_by_newline) > 1:
        lines = lines_by_newline
    else:
        # Если всё в одной строке — разбиваем по номерам (1., 2., 3.)
        lines = re.split(r'\s+(?=\d+[\.\)])\s*', text)
        lines = [line.strip() for line in lines if line.strip()]
    
    # Если не получилось — пробуем по точке с запятой
    if len(lines) <= 1 and ';' in text:
        lines = [line.strip() for line in text.split(';') if line.strip()]
    
    if len(lines) == 0:
        lines = [text]
    
    # === НОВОЕ: Объединяем короткие строки с предыдущими ===
    # Если строка короче 15 символов и не начинается с цифры — это продолжение предыдущей
    merged_lines = []
    for line in lines:
        # Проверяем, начинается ли строка с цифры (новый пункт)
        starts_with_number = re.match(r'^\d+[\.\)]\s*', line)
        
        if starts_with_number:
            # Это новый пункт — добавляем отдельно
            merged_lines.append(line)
        elif len(line) < 15 and merged_lines:
            # Короткая строка без номера — объединяем с предыдущей
            merged_lines[-1] = merged_lines[-1] + ' ' + line
        else:
            # Обычная строка — добавляем отдельно
            merged_lines.append(line)
    
    return merged_lines


def generate_legal_document(template_data: dict, output_dir: str, doc_type: str = "претензия") -> str:
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
    run = p.add_run(template_data.get('recipient', '[Наименование]'))
    run.font.size = Pt(11)
    
    # Для иска — добавляем ОТВЕТЧИКА
    if doc_type == "иск":
        defendant = template_data.get('defendant', '').strip()
        if defendant:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(f"Ответчик: {defendant}")
            run.font.size = Pt(11)
    
    # От кого / Истец
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sender = template_data.get('sender', '')
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
    
    # === ТРЕБОВАНИЯ / ПРОШУ ===
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
    
    # Для иска — добавляем обязательные приложения, если их нет
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
