"""
Модуль генерации юридических документов (претензий, жалоб, исков, ходатайств)
Финальная версия с учётом всех исправлений
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
        "гражданский кодекс", "гк рф",
        "потребител",
        "налоговый кодекс", "нк рф", "333.36"
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

# === ЗАПРЕЩЁННЫЕ ЗАКОНЫ ПО КАТЕГОРИЯМ (явные ошибки) ===
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
- НИКОГДА не пиши инструкции в скобках типа "(укажи название)", "(если известно)", "(при наличии)"
- Если данных нет — просто ОСТАВЬ поле пустым, без подсказок

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
- Не используй markdown
- Не используй кавычки вокруг значений""",

    "жалоба": """Ты — юридический помощник. Создай ЖАЛОБУ в государственный орган.

КРИТИЧЕСКИ ВАЖНО:
- Используй ТОЛЬКО информацию из ПОСЛЕДНЕГО вопроса клиента
- НЕ выдумывай факты, которых нет в вопросе
- Если данные не указаны — оставь поле ПУСТЫМ
- НИКОГДА не пиши инструкции в скобках типа "(укажи название)", "(если известно)", "(при наличии)"
- Если данных нет — просто ОСТАВЬ поле пустым, без подсказок

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
- Не используй markdown
- Не используй кавычки вокруг значений""",

    "иск": """Ты — юридический помощник. Создай ИСКОВОЕ ЗАЯВЛЕНИЕ в суд.

КРИТИЧЕСКИ ВАЖНО:
- Используй ТОЛЬКО информацию из ПОСЛЕДНЕГО вопроса клиента
- НЕ выдумывай факты, которых нет в вопросе
- Если данные не указаны (ФИО истца, адрес ответчика) — оставь поле ПУСТЫМ
- НИКОГДА не пиши инструкции в скобках типа "(укажи название)", "(если известно)", "(при наличии)"
- Если данных нет — просто ОСТАВЬ поле пустым, без подсказок

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ:

ТИП: исковое заявление
КОМУ: [наименование суда — если не указано, напиши "В районный суд по месту жительства истца"]
ИСТЕЦ: [ФИО, адрес, телефон — если не указано, оставь ПУСТЫМ]
ОТВЕТЧИК: [ФИО или наименование организации, адрес, ИНН — если не указано, оставь ПУСТЫМ]
СИТУАЦИЯ: [ДЕТАЛЬНОЕ описание 5-7 предложений от ПЕРВОГО ЛИЦА]
ЗАКОНЫ: [ВСЕ статьи законов из ответа юриста, через точку с запятой]
ЦЕНА_ИСКА: [ДЕТАЛЬНЫЙ расчёт цены иска с разбивкой по пунктам, БЕЗ заголовка "Цена иска:" и БЕЗ кавычек]
ТРЕБОВАНИЯ: [КОНКРЕТНЫЕ требования к суду с СУММАМИ, нумерованный список, БЕЗ заголовка "ПРОШУ СУД:"]
ПРИЛОЖЕНИЯ: [КОНКРЕТНЫЙ список документов, нумерованный список]

=== ПРАВИЛА РАСЧЁТА ЦЕНЫ ИСКА ===

Для исков по защите прав потребителей ОБЯЗАТЕЛЬНО включи в цену иска:

1. ОСНОВНАЯ СУММА — стоимость товара/услуги
2. НЕУСТОЙКА по ст. 23 ЗоЗПП — 1% от стоимости товара за КАЖДЫЙ день просрочки
   ФОРМУЛА: Стоимость × 1% × Количество дней просрочки
   Дни просрочки считай от даты, когда продавец должен был удовлетворить требование 
   (10 дней с момента обращения потребителя — ст. 22 ЗоЗПП) до текущей даты 10 июля 2026 г.
   Если дата обращения не указана — считай от даты обнаружения недостатка + 10 дней.
3. МОРАЛЬНЫЙ ВРЕД по ст. 15 ЗоЗПП — обычно 5 000 - 20 000 рублей
4. СУДЕБНЫЕ РАСХОДЫ — если указаны в вопросе

ПРИМЕР РАСЧЁТА (в поле ЦЕНА_ИСКА пиши ТОЛЬКО расчёт, БЕЗ заголовка и БЕЗ кавычек):
"- Стоимость товара: 80 000 рублей
- Неустойка по ст. 23 ЗоЗПП (1% × 80 000 руб. × 29 дней просрочки): 23 200 рублей
- Компенсация морального вреда: 10 000 рублей
ИТОГО: 113 200 рублей"

=== ПРАВИЛА ФОРМУЛИРОВКИ ТРЕБОВАНИЙ ===

Для исков по ЗоЗПП ОБЯЗАТЕЛЬНО включи в требования:
1. Взыскать стоимость товара (основная сумма)
2. Взыскать НЕУСТОЙКУ по ст. 23 ЗоЗПП (с указанием расчёта)
3. Взыскать компенсацию морального вреда по ст. 15 ЗоЗПП
4. Взыскать ШТРАФ 50% от присуждённой суммы по ст. 13 п. 6 ЗоЗПП
5. Взыскать судебные расходы (если есть, НО НЕ госпошлину!)

=== ОСВОБОЖДЕНИЕ ОТ ГОСПОШЛИНЫ ===

ВАЖНО: Согласно ст. 333.36 НК РФ, потребители по искам о защите прав потребителей 
ОСВОБОЖДАЮТСЯ от уплаты государственной пошлины, если цена иска не превышает 1 000 000 рублей.

Поэтому:
- НЕ включай в приложения "Квитанция об уплате государственной пошлины"
- НЕ включай в требования "Взыскать расходы на уплату госпошлины"
- В правовое обоснование добавь ст. 333.36 НК РФ

ПРИМЕР ТРЕБОВАНИЙ (в поле ТРЕБОВАНИЯ пиши ТОЛЬКО список, БЕЗ заголовка "ПРОШУ СУД:"):
"1. Взыскать с ответчика стоимость товара в размере 80 000 рублей
2. Взыскать неустойку по ст. 23 Закона РФ «О защите прав потребителей» в размере 23 200 рублей (1% × 80 000 руб. × 29 дней)
3. Взыскать компенсацию морального вреда в размере 10 000 рублей
4. Взыскать штраф в размере 50% от присуждённой суммы на основании п. 6 ст. 13 Закона РФ «О защите прав потребителей»"

ПРИМЕР ПРИЛОЖЕНИЙ (БЕЗ квитанции о госпошлине!):
"1. Копия искового заявления для ответчика
2. Копия чека
3. Копия гарантийного талона
4. Копия претензии с отметкой о вручении
5. Расчёт цены иска"

ВАЖНО:
- Иск адресуется В СУД
- Обязательно укажи ИСТЦА и ОТВЕТЧИКА (если есть данные)
- ЦЕНА_ИСКА: рассчитай ВСЕ суммы с разбивкой, БЕЗ заголовка "Цена иска:" и БЕЗ кавычек
- ТРЕБОВАНИЯ: нумерованный список с КОНКРЕТНЫМИ СУММАМИ, БЕЗ заголовка "ПРОШУ СУД:"
- В ПРИЛОЖЕНИЯХ НЕ ДОЛЖНО БЫТЬ "Квитанция об уплате государственной пошлины"
- ЗАКОНЫ: используй ТОЛЬКО релевантные статьи, включая ст. 333.36 НК РФ
- Не используй markdown
- Не используй кавычки вокруг значений""",

    "ходатайство": """Ты — юридический помощник. Создай ХОДАТАЙСТВО.

КРИТИЧЕСКИ ВАЖНО:
- Используй ТОЛЬКО информацию из ПОСЛЕДНЕГО вопроса клиента
- НЕ выдумывай факты
- Если данные не указаны — оставь поле ПУСТЫМ
- НИКОГДА не пиши инструкции в скобках типа "(укажи название)", "(если известно)", "(при наличии)"
- Если данных нет — просто ОСТАВЬ поле пустым, без подсказок

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
- Не используй markdown
- Не используй кавычки вокруг значений"""
}


def analyze_chat_for_template(chat_history: list, category: str, auth_key: str, doc_type: str = "претензия") -> dict:
    """
    Анализирует историю чата и извлекает данные для документа.
    Использует ТОЛЬКО последний вопрос и последний ответ.
    """
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
    
    valid_keywords = VALID_LAWS_BY_CATEGORY.get(category, [])
    invalid_keywords = INVALID_LAWS_BY_CATEGORY.get(category, [])
    
    # Разбиваем на отдельные статьи (по точке с запятой)
    articles = [a.strip() for a in legal_basis.split(';') if a.strip()]
    
    filtered_articles = []
    for article in articles:
        article_lower = article.lower()
        
        # Проверяем, есть ли запрещённые ключевые слова
        is_invalid = any(kw in article_lower for kw in invalid_keywords)
        
        # Проверяем, есть ли допустимые ключевые слова
        is_valid = any(re.search(kw, article_lower) for kw in valid_keywords) if valid_keywords else True
        
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


def clean_template_hints(text: str) -> str:
    """Удаляет шаблонные подсказки из текста, кавычки и лишние символы"""
    if not text:
        return text
    
    # Убираем все фразы в скобках с инструкциями
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
    
    # Убираем кавычки вокруг значения (в начале и в конце)
    text = text.strip()
    text = text.strip('"').strip("'").strip('«').strip('»')
    text = text.strip()
    
    # Убираем лишние пробелы
    text = re.sub(r'\s+', ' ', text).strip()
    
    # Убираем знаки препинания в конце
    text = text.rstrip(',;')
    
    return text


def remove_duplicate_headers(text: str) -> str:
    """Удаляет дублирующиеся заголовки из текста"""
    if not text:
        return text
    
    # Удаляем дублирование "Цена иска: Цена иска:"
    text = re.sub(r'Цена иска:\s*Цена иска:', 'Цена иска:', text, flags=re.IGNORECASE)
    
    # Удаляем "ПРОШУ СУД:" из начала текста требований
    text = re.sub(r'^\s*ПРОШУ СУД:\s*', '', text, flags=re.IGNORECASE)
    text = re.sub(r'^\s*1\.\s*ПРОШУ СУД:\s*', '', text, flags=re.IGNORECASE)
    
    return text


def split_into_items(text: str) -> list:
    """
    Разбивает текст на пункты по номерам или переводам строк.
    УЛУЧШЕННАЯ ВЕРСИЯ: объединяет короткие строки и удаляет дубликаты.
    """
    if not text:
        return []
    
    # Удаляем дублирующиеся заголовки
    text = remove_duplicate_headers(text)
    
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
    
    # === Объединяем короткие строки с предыдущими ===
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
    
    # === Удаляем дубликаты ===
    unique_lines = []
    seen = set()
    for line in merged_lines:
        # Убираем номера и очищаем от подсказок для сравнения
        line_without_number = re.sub(r'^\d+[\.\)]\s*', '', line).strip()
        line_cleaned = clean_template_hints(line_without_number).lower()
        
        # === Пропускаем строки с заголовками ===
        if any(x in line_cleaned for x in ['прошу суд', 'требую', 'прошу', 'ходатайствую']):
            continue
        
        # === Пропускаем квитанции о госпошлине для исков по ЗоЗПП ===
        # (Проверка делается в generate_legal_document, но на всякий случай и здесь)
        
        if line_cleaned not in seen and line_cleaned:
            seen.add(line_cleaned)
            # Очищаем от подсказок
            cleaned_line = clean_template_hints(line)
            if cleaned_line:
                unique_lines.append(cleaned_line)
    
    return unique_lines


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
    recipient = template_data.get('recipient', '[Наименование]')
    recipient = clean_template_hints(recipient)
    if 'если известен' in recipient.lower() or 'укажи конкретнее' in recipient.lower():
        recipient = '[Наименование организации]'
    run = p.add_run(recipient)
    run.font.size = Pt(11)
    
    # Для иска — добавляем ОТВЕТЧИКА
    if doc_type == "иск":
        defendant = template_data.get('defendant', '').strip()
        defendant = clean_template_hints(defendant)
        if defendant and ('если известен' in defendant.lower() or 'оставь пустым' in defendant.lower()):
            defendant = ''
        # Если ответчик пустой, но есть получатель — используем его
        if not defendant and recipient and recipient != '[Наименование организации]' and recipient != '[Наименование]':
            defendant = recipient
        if defendant:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(f"Ответчик: {defendant}")
            run.font.size = Pt(11)
    
    # От кого / Истец
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sender = template_data.get('sender', '')
    sender = clean_template_hints(sender)
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
    situation = clean_template_hints(situation)
    doc.add_paragraph(situation)
    doc.add_paragraph()
    
    # === ПРАВОВОЕ ОБОСНОВАНИЕ ===
    legal_basis = template_data.get('legal_basis', '').strip()
    legal_basis = clean_template_hints(legal_basis)
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
        claim_amount = clean_template_hints(claim_amount)
        claim_amount = remove_duplicate_headers(claim_amount)
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
    requirements = clean_template_hints(requirements)
    requirements = remove_duplicate_headers(requirements)
    if not requirements or requirements == '[Требования]':
        requirements = "Удовлетворить мои законные требования в соответствии с действующим законодательством РФ."
    
    req_lines = split_into_items(requirements)
    
    for i, line in enumerate(req_lines, 1):
        line = re.sub(r'^\d+[\.\)]\s*', '', line)
        line = clean_template_hints(line)
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
    if not attachments or attachments == '[Приложения]':
        attachments = "1. Копия данного документа\n2. Копии документов, подтверждающих требования"
    
    att_lines = split_into_items(attachments)
    
    # === Специальная обработка для исков ===
    if doc_type == "иск":
        att_text = ' '.join(att_lines).lower()
        
        # === ИСПРАВЛЕНИЕ: НЕ добавляем квитанцию о госпошлине для исков по ЗоЗПП ===
        # Согласно ст. 333.36 НК РФ, потребители освобождены от госпошлины
        
        # === Удаляем квитанцию о госпошлине, если ИИ её добавил ===
        att_lines = [line for line in att_lines 
                     if 'госпошлин' not in line.lower() and 'государственной пошлин' not in line.lower()]
        
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
