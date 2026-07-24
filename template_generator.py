"""
Модуль генерации юридических документов (претензий, жалоб, исков, ходатайств)
ФИНАЛЬНАЯ версия с учётом:
- Технически сложных товаров (ТСТ)
- Альтернативной подсудности по ЗоЗПП
- Освобождения от госпошлины (ст. 333.36 НК РФ)
- Очистки от глюков и мусора
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
    "🛒 Защита прав потребителей": ["о защите прав потребителей", "зозпп", "2300-1", "гражданский кодекс", "гк рф", "потребител"],
    "💼 Трудовые споры": ["трудовой кодекс", "тк рф", "коллективн", "заработн", "увольнен", "дисциплинарн"],
    "🚗 Споры с ГИБДД": ["коап", "коап рф", "правила дорожного", "пдд", "безопасност.*дорож", "осаго", "транспортн"],
    "👨‍👩‍👧 Семейное право": ["семейный кодекс", "ск рф", "алимент", "развод", "брачн", "раздел имуществ", "усыновлен"],
    "📜 Наследство": ["гражданский кодекс", "гк рф", "наследств", "завещан", "наследодател", "обязательн.*доля"],
    "🏠 Жилищные вопросы (вкл. ЖКХ)": ["жилищный кодекс", "жк рф", "постановление правительства", "№ 354", "№ 491", "управляющ", "жиль", "приватизац"]
}

# === ЗАПРЕЩЁННЫЕ ЗАКОНЫ ПО КАТЕГОРИЯМ ===
INVALID_LAWS_BY_CATEGORY = {
    "🛒 Защита прав потребителей": ["тк рф", "коап", "жилищн", "ск рф"],
    "💼 Трудовые споры": ["о защите прав потребителей", "коап", "жилищн", "ск рф"],
    "🚗 Споры с ГИБДД": ["о защите прав потребителей", "тк рф", "жилищн", "ск рф", "гк рф"],
    "👨‍👩‍👧 Семейное право": ["тк рф", "коап", "жилищн", "зозпп"],
    "📜 Наследство": ["тк рф", "коап", "жилищн", "зозпп", "ск рф"],
    "🏠 Жилищные вопросы (вкл. ЖКХ)": ["о защите прав потребителей", "тк рф", "коап", "ск рф"]
}

# === ТЕХНИЧЕСКИ СЛОЖНЫЕ ТОВАРЫ (для промтов) ===
TST_EXAMPLES = """
К технически сложным товарам относятся (Перечень, утв. Постановлением Правительства РФ от 10.11.2011 № 924):
- смартфоны, мобильные телефоны
- компьютеры (стационарные, портативные, планшеты)
- телевизоры, мониторы
- автомобили, мотоциклы
- бытовая техника (стиральные машины, холодильники, посудомойки)
- фото- и видеокамеры
- часы (наручные, карманные — электронные/механические с 2+ функциями)
- садовая техника (газонокосилки, бензопилы)
- навигационное оборудование
"""


# === ПРОМТЫ ДЛЯ РАЗНЫХ ТИПОВ ДОКУМЕНТОВ ===
DOCUMENT_PROMPTS = {
    "претензия": """Ты — юридический помощник. Создай ПРЕТЕНЗИЮ к контрагенту.

КРИТИЧЕСКИ ВАЖНО:
- Используй ТОЛЬКО информацию из ПОСЛЕДНЕГО вопроса клиента
- НЕ выдумывай факты
- Если данные не указаны — оставь поле ПУСТЫМ
- НЕ используй никакие скобки: ни круглые (), ни квадратные [], ни угловые <>
- НЕ пиши placeholder'ы вроде [ФИО], <ФИО>, (ФИО)

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ:

ТИП: претензия
КОМУ: [должность и наименование организации]
ОТ: [ФИО, адрес, телефон — если нет, оставь пустым]
СИТУАЦИЯ: [4-6 предложений от ПЕРВОГО ЛИЦА]
ЗАКОНЫ: [статьи законов через точку с запятой]
ТРЕБОВАНИЯ: [нумерованный список]
ПРИЛОЖЕНИЯ: [нумерованный список]

ВАЖНО:
- НЕ используй markdown
- НЕ используй скобки любого типа""",

    "жалоба": """Ты — юридический помощник. Создай ЖАЛОБУ в государственный орган.

КРИТИЧЕСКИ ВАЖНО:
- Используй ТОЛЬКО информацию из ПОСЛЕДНЕГО вопроса клиента
- НЕ выдумывай факты
- Если данные не указаны — оставь поле ПУСТЫМ
- НЕ используй никакие скобки: ни круглые (), ни квадратные [], ни угловые <>

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ:

ТИП: жалоба
КОМУ: [наименование госоргана]
ОТ: [ФИО, адрес, телефон — если нет, оставь пустым]
СИТУАЦИЯ: [4-6 предложений от ПЕРВОГО ЛИЦА]
ЗАКОНЫ: [статьи законов через точку с запятой]
ТРЕБОВАНИЯ: [нумерованный список — ПРОВЕСТИ ПРОВЕРКУ, ПРИНЯТЬ МЕРЫ]
ПРИЛОЖЕНИЯ: [нумерованный список]

ВАЖНО:
- НЕ используй markdown
- НЕ используй скобки любого типа""",

    "иск": """Ты — юридический помощник. Создай ИСКОВОЕ ЗАЯВЛЕНИЕ в суд.

КРИТИЧЕСКИ ВАЖНО:
- Используй ТОЛЬКО информацию из ПОСЛЕДНЕГО вопроса клиента
- НЕ выдумывай факты
- Если данные не указаны — оставь поле ПУСТЫМ
- НЕ используй никакие скобки: ни круглые (), ни квадратные [], ни угловые <>
- НЕ пиши placeholder'ы вроде [ФИО], <ФИО>, (ФИО)

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ:

ТИП: исковое заявление
КОМУ: [наименование суда С УКАЗАНИЕМ ПОДСУДНОСТИ — "В [название] районный суд г. [город] по месту жительства истца"]
ИСТЕЦ: [ФИО, адрес, телефон — если нет, оставь пустым]
ОТВЕТЧИК: [наименование организации, адрес, ИНН — если нет, оставь пустым]
СИТУАЦИЯ: [5-7 предложений от ПЕРВОГО ЛИЦА БЕЗ placeholder'ов]
ЗАКОНЫ: [статьи законов через точку с запятой]
ЦЕНА_ИСКА: [расчёт цены иска с разбивкой по строкам]
ТРЕБОВАНИЯ: [нумерованный список с СУММАМИ]
ПРИЛОЖЕНИЯ: [нумерованный список]

=== ПОДСУДНОСТЬ (КРИТИЧЕСКИ ВАЖНО!) ===

Для исков по ЗоЗПП действует АЛЬТЕРНАТИВНАЯ ПОДСУДНОСТЬ:
- Ст. 17 ЗоЗПП — иски о защите прав потребителей могут подаваться:
  1. По месту нахождения организации-ответчика
  2. По месту жительства или пребывания истца (ВЫГОДНЕЕ ДЛЯ ИСТЦА!)
  3. По месту заключения или исполнения договора
- Ст. 29 ГПК РФ — альтернативная подсудность

Поэтому в поле КОМУ ОБЯЗАТЕЛЬНО пиши:
"В [название] районный суд по месту жительства истца"

Пример: "В районный суд по месту жительства истца"

=== ТЕХНИЧЕСКИ СЛОЖНЫЕ ТОВАРЫ (ТСТ) ===

""" + TST_EXAMPLES + """

ПРАВИЛА ДЛЯ ТСТ (ст. 18 п. 1 ЗоЗПП):
- В течение 15 дней со дня передачи товара — потребитель вправе отказаться от товара и потребовать возврата денег ИЛИ замены
- ПОСЛЕ 15 дней — требования удовлетворяются ТОЛЬКО в случаях:
  1. Обнаружение существенного недостатка товара
  2. Нарушение продавцом срока ремонта (45 дней)
  3. Невозможность использования товара >30 дней в году из-за неоднократного ремонта

ПРАВИЛА ПРИМЕНЕНИЯ:
- Если в вопросе УКАЗАНО, что товар куплен менее 15 дней назад → пиши: "В соответствии с п. 1 ст. 18 ЗоЗПП, в течение 15 дней со дня передачи технически сложного товара потребитель вправе отказаться от исполнения договора купли-продажи"
- Если прошло БОЛЬШЕ 15 дней → пиши: "В соответствии с п. 1 ст. 18 ЗоЗПП, в отношении технически сложного товара требование о возврате денег подлежит удовлетворению в случае обнаружения существенного недостатка товара"
- Если упоминается экспертиза → добавь: "Готов предоставить товар для проведения проверки качества и/или экспертизы в соответствии с п. 5 ст. 18 ЗоЗПП"
- В правовое обоснование добавь: "Постановление Правительства РФ от 10.11.2011 № 924 (Перечень технически сложных товаров)"

=== РАСЧЁТ ЦЕНЫ ИСКА ===

Для исков по ЗоЗПП ОБЯЗАТЕЛЬНО включи:
1. Стоимость товара/услуги
2. НЕУСТОЙКА по ст. 23 ЗоЗПП — 1% от стоимости за КАЖДЫЙ день просрочки
3. МОРАЛЬНЫЙ ВРЕД по ст. 15 ЗоЗПП — 5 000 - 20 000 рублей
4. СУДЕБНЫЕ РАСХОДЫ — если указаны

ПРИМЕР (каждый пункт с новой строки):
"Стоимость товара: 80 000 рублей
Неустойка по ст. 23 ЗоЗПП (1% × 80 000 руб. × 29 дней): 23 200 рублей
Компенсация морального вреда: 10 000 рублей
ИТОГО: 113 200 рублей"

=== ОСВОБОЖДЕНИЕ ОТ ГОСПОШЛИНЫ ===

Согласно ст. 333.36 НК РФ, потребители ОСВОБОЖДАЮТСЯ от госпошлины по искам до 1 000 000 руб.
Поэтому:
- НЕ включай "Квитанция об уплате госпошлины" в приложения
- Добавь ст. 333.36 НК РФ в правовое обоснование

=== ТРЕБОВАНИЯ ===

Включи:
1. Стоимость товара
2. Неустойку по ст. 23 ЗоЗПП
3. Моральный вред по ст. 15 ЗоЗПП
4. Штраф 50% по ст. 13 п. 6 ЗоЗПП

НЕ включай пояснения типа "не заявлено", "отсутствует".

=== ПРИЛОЖЕНИЯ ===

Включи:
- Копия искового заявления для ответчика
- Копия договора купли-продажи (или чека)
- Копия чека
- Копия гарантии
- Копия претензии с отметкой
- Расчёт цены иска
- Если есть — акт экспертизы, заключение эксперта

НЕ включай:
- Квитанция о госпошлине
- Копии статей законодательства
- Любые куски статей законов (например, "36 Налогового кодекса")
- Устав ООО, свидетельство о регистрации (это можно истребовать через суд)

ВАЖНО:
- НЕ используй скобки любого типа
- НЕ используй placeholder'ы
- НЕ используй markdown""",

    "ходатайство": """Ты — юридический помощник. Создай ХОДАТАЙСТВО.

КРИТИЧЕСКИ ВАЖНО:
- Используй ТОЛЬКО информацию из ПОСЛЕДНЕГО вопроса клиента
- НЕ выдумывай факты
- Если данные не указаны — оставь поле ПУСТЫМ
- НЕ используй никакие скобки: ни круглые (), ни квадратные [], ни угловые <>

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ:

ТИП: ходатайство
КОМУ: [наименование органа]
ОТ: [ФИО, адрес, телефон — если нет, оставь пустым]
СИТУАЦИЯ: [3-5 предложений от ПЕРВОГО ЛИЦА]
ЗАКОНЫ: [статьи законов через точку с запятой]
ТРЕБОВАНИЯ: [нумерованный список — ПРОШУ: 1. Предоставить...]
ПРИЛОЖЕНИЯ: [нумерованный список]

ВАЖНО:
- НЕ используй markdown
- НЕ используй скобки любого типа"""
}


def analyze_chat_for_template(chat_history: list, category: str, auth_key: str, doc_type: str = "претензия") -> dict:
    """Анализирует историю чата и извлекает данные для документа"""
    try:
        prompt_template = DOCUMENT_PROMPTS.get(doc_type, DOCUMENT_PROMPTS["претензия"])
        
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

ПОСЛЕДНИЙ ВОПРОС КЛИЕНТА:
{last_user_question}

ОТВЕТ ЮРИСТА:
{last_assistant_answer}

{prompt_template}

Создавай документ ТОЛЬКО на основе ПОСЛЕДНЕГО вопроса."""
        
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
        "recipient": "",
        "sender": "",
        "situation": "",
        "legal_basis": "",
        "requirements": "",
        "attachments": "",
        "defendant": "",
        "claim_amount": ""
    }


def clean_template_hints(text: str) -> str:
    """Удаляет шаблонные подсказки, скобки и лишние символы"""
    if not text:
        return text
    
    patterns = [
        r'\(укажите[^)]*\)',
        r'\(если[^)]*\)',
        r'\(оставь[^)]*\)',
        r'\(если известен[^)]*\)',
        r'\(если имеются[^)]*\)',
        r'\(если понесены[^)]*\)',
        r'\(если такие[^)]*\)',
        r'\(при наличии[^)]*\)',
        r'\(при их наличии[^)]*\)',
        r'\(если таковые[^)]*\)',
        r'\(укажи[^)]*\)',
        r'\(если известно[^)]*\)',
        r'\(если они[^)]*\)',
        r'\(если есть[^)]*\)',
    ]
    
    for pattern in patterns:
        text = re.sub(pattern, '', text, flags=re.IGNORECASE)
    
    # Убираем квадратные скобки [...]
    text = re.sub(r'\[([^\]]*)\]', r'\1', text)
    
    # Убираем угловые скобки <...>
    text = re.sub(r'<([^>]*)>', r'\1', text)
    
    # Убираем круглые скобки с placeholder'ами
    text = re.sub(r'\([А-ЯA-Z][А-Яа-яA-Za-z\s,]+\)', '', text)
    
    # Убираем кавычки вокруг значения
    text = text.strip()
    text = text.strip('"').strip("'").strip('«').strip('»')
    text = text.strip()
    
    text = re.sub(r'\s+', ' ', text).strip()
    text = text.rstrip(',;')
    
    return text


def format_claim_amount(text: str) -> str:
    """Форматирует цену иска — разбивает на строки"""
    if not text:
        return text
    
    if ' - ' in text and '\n' not in text:
        parts = [p.strip() for p in text.split(' - ') if p.strip()]
        if len(parts) > 1:
            return '\n'.join(parts)
    
    return text


def remove_duplicate_headers(text: str) -> str:
    """Удаляет дублирующиеся заголовки"""
    if not text:
        return text
    
    text = re.sub(r'Цена иска:\s*Цена иска:', 'Цена иска:', text, flags=re.IGNORECASE)
    text = re.sub(r'^\s*ПРОШУ СУД:\s*', '', text, flags=re.IGNORECASE)
    text = re.sub(r'^\s*1\.\s*ПРОШУ СУД:\s*', '', text, flags=re.IGNORECASE)
    
    return text


def is_garbage_attachment(line: str) -> bool:
    """
    Проверяет, является ли приложение мусором (кусок статьи закона и т.п.)
    """
    line_lower = line.lower().strip()
    
    # Мусорные приложения
    garbage_patterns = [
        r'^\d+\s*(налогов|нк\s*рф)',  # "36 Налогового кодекса"
        r'стать(и|я|ей)\s+законодательств',  # "статьи законодательства"
        r'копи(я|и|ю)\s+статей',  # "копия статей"
        r'копи(я|и|ю)\s+законодательств',  # "копия законодательства"
        r'^\d+\s+(гк|тк|коап|жк|закон|кодекс)',  # "475 ГК РФ"
        r'устав\s+(ооо|зао|ао|пао)',  # "Устав ООО"
        r'свидетельств[ао]\s+о\s+регистрац',  # "свидетельство о регистрации"
        r'егрюл|егрип',  # выписки из ЕГРЮЛ
    ]
    
    for pattern in garbage_patterns:
        if re.search(pattern, line_lower):
            return True
    
    return False


def split_into_items(text: str) -> list:
    """Разбивает текст на пункты с удалением дубликатов и мусора"""
    if not text:
        return []
    
    text = remove_duplicate_headers(text)
    
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
    
    merged_lines = []
    for line in lines:
        starts_with_number = re.match(r'^\d+[\.\)]\s*', line)
        
        if starts_with_number:
            merged_lines.append(line)
        elif len(line) < 15 and merged_lines:
            merged_lines[-1] = merged_lines[-1] + ' ' + line
        else:
            merged_lines.append(line)
    
    unique_lines = []
    seen = set()
    for line in merged_lines:
        line_without_number = re.sub(r'^\d+[\.\)]\s*', '', line).strip()
        line_cleaned = clean_template_hints(line_without_number).lower()
        
        if any(x in line_cleaned for x in ['прошу суд', 'требую', 'прошу', 'ходатайствую']):
            continue
        
        if any(x in line_cleaned for x in ['не заявлены', 'отсутствуют', 'не подлежат', 'ввиду отсутствия']):
            continue
        
        if line_cleaned not in seen and line_cleaned:
            seen.add(line_cleaned)
            cleaned_line = clean_template_hints(line)
            if cleaned_line:
                unique_lines.append(cleaned_line)
    
    return unique_lines


def fix_unclosed_quotes(text: str) -> str:
    """Исправляет незакрытые кавычки"""
    if not text:
        return text
    
    open_quotes = text.count('«')
    close_quotes = text.count('»')
    
    if open_quotes > close_quotes:
        text = text + '»' * (open_quotes - close_quotes)
    
    return text


def ensure_jurisdiction(recipient: str) -> str:
    """
    Для исков в суд добавляет "по месту жительства истца", если это не указано.
    Согласно ст. 17 ЗоЗПП и ст. 29 ГПК РФ.
    """
    if not recipient:
        return "В районный суд по месту жительства истца"
    
    recipient_lower = recipient.lower()
    
    # Если это суд, но нет указания подсудности
    if 'суд' in recipient_lower:
        if 'по месту жительства' not in recipient_lower and \
           'по месту нахождения' not in recipient_lower and \
           'мировой' not in recipient_lower and \
           'арбитраж' not in recipient_lower:
            # Добавляем подсудность
            if recipient_lower.startswith('в '):
                recipient = recipient + ' по месту жительства истца'
            else:
                recipient = 'В ' + recipient + ' по месту жительства истца'
    
    return recipient


def generate_legal_document(template_data: dict, output_dir: str, doc_type: str = "претензия") -> str:
    """Создаёт Word-документ на основе данных шаблона"""
    doc = Document()
    
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    # === ШАПКА ===
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    recipient = template_data.get('recipient', '')
    recipient = clean_template_hints(recipient)
    
    # Для иска — обеспечиваем правильную подсудность
    if doc_type == "иск":
        recipient = ensure_jurisdiction(recipient)
    
    if not recipient:
        recipient = '[Наименование]'
    run = p.add_run(recipient)
    run.font.size = Pt(11)
    
    if doc_type == "иск":
        defendant = template_data.get('defendant', '').strip()
        defendant = clean_template_hints(defendant)
        if not defendant and recipient and recipient != '[Наименование]' and 'суд' not in recipient.lower():
            defendant = recipient
        if defendant:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(f"Ответчик: {defendant}")
            run.font.size = Pt(11)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sender = template_data.get('sender', '')
    sender = clean_template_hints(sender)
    
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
    situation = template_data.get('situation', '')
    situation = clean_template_hints(situation)
    situation = fix_unclosed_quotes(situation)
    if not situation:
        situation = '[Описание ситуации]'
    doc.add_paragraph(situation)
    doc.add_paragraph()
    
    # === ПРАВОВОЕ ОБОСНОВАНИЕ ===
    legal_basis = template_data.get('legal_basis', '').strip()
    legal_basis = clean_template_hints(legal_basis)
    legal_basis = fix_unclosed_quotes(legal_basis)
    if not legal_basis:
        legal_basis = ("В соответствии с действующим законодательством Российской Федерации, "
                      "а также на основании прав, предоставленных мне нормативными правовыми актами, "
                      "регулирующими данную ситуацию.")
    
    p = doc.add_paragraph()
    run = p.add_run("Правовое обоснование: ")
    run.bold = True
    p.add_run(legal_basis)
    doc.add_paragraph()
    
    # === ЦЕНА ИСКА ===
    if doc_type == "иск":
        claim_amount = template_data.get('claim_amount', '').strip()
        claim_amount = clean_template_hints(claim_amount)
        claim_amount = remove_duplicate_headers(claim_amount)
        claim_amount = format_claim_amount(claim_amount)
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
    requirements = clean_template_hints(requirements)
    requirements = remove_duplicate_headers(requirements)
    requirements = fix_unclosed_quotes(requirements)
    if not requirements:
        requirements = "Удовлетворить мои законные требования в соответствии с действующим законодательством РФ."
    
    req_lines = split_into_items(requirements)
    
    for i, line in enumerate(req_lines, 1):
        line = re.sub(r'^\d+[\.\)]\s*', '', line)
        line = clean_template_hints(line)
        line = fix_unclosed_quotes(line)
        if line:
            p = doc.add_paragraph(f"{i}. {line}")
            p.paragraph_format.left_indent = Pt(20)
    
    doc.add_paragraph()
    
    # === ПРИЛОЖЕНИЯ ===
    p = doc.add_paragraph()
    run = p.add_run("Приложения:")
    run.bold = True
    
    attachments = template_data.get('attachments', '').strip()
    attachments = clean_template_hints(attachments)
    if not attachments:
        attachments = "1. Копия данного документа\n2. Копии документов, подтверждающих требования"
    
    att_lines = split_into_items(attachments)
    
    if doc_type == "иск":
        att_text = ' '.join(att_lines).lower()
        
        # Удаляем квитанцию о госпошлине (ст. 333.36 НК РФ)
        att_lines = [line for line in att_lines 
                     if 'госпошлин' not in line.lower() and 'государственной пошлин' not in line.lower()]
        
        # Удаляем мусорные приложения (куски статей, уставы и т.д.)
        att_lines = [line for line in att_lines if not is_garbage_attachment(line)]
        
        if 'копия искового' not in att_text and 'для ответчика' not in att_text:
            att_lines.append("Копия искового заявления для ответчика")
        if 'расчёт' not in att_text and 'расчет' not in att_text:
            att_lines.append("Расчёт цены иска")
    
    for i, line in enumerate(att_lines, 1):
        line = re.sub(r'^\d+[\.\)]\s*', '', line)
        line = clean_template_hints(line)
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
