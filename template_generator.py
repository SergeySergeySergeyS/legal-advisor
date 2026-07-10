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
        # Формируем промт с ПРИМЕРАМИ
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
ЗАДАЧА: Извлеки МАКСИМУМ деталей и создай структуру документа.

ВЕРНИ РЕЗУЛЬТАТ СТРОГО В ТАКОМ ФОРМАТЕ (каждое поле с новой строки):

ТИП: [претензия/жалоба/заявление/иск]
КОМУ: [конкретное наименование организации]
ОТ: [если в вопросе есть ФИО/адрес/телефон — укажи, иначе оставь ПУСТЫМ]
СИТУАЦИЯ: [ДЕТАЛЬНОЕ описание 4-6 предложений от ПЕРВОГО ЛИЦА "Я купил...", используй факты из вопроса]
ЗАКОНЫ: [ВСЕ статьи законов из ответа юриста с названиями, через точку с запятой]
ТРЕБОВАНИЯ: [КОНКРЕТНЫЕ требования с суммами и сроками, нумерованный список]
ПРИЛОЖЕНИЯ: [КОНКРЕТНЫЙ список документов, нумерованный список]

=== ПРИМЕР ПРАВИЛЬНОГО ОТВЕТА ===

ТИП: претензия
КОМУ: Директору ООО "ТехноМаркет"
ОТ: 
СИТУАЦИЯ: 1 июля 2026 года я приобрёл в вашем магазине мобильный телефон iPhone 15 стоимостью 80 000 рублей, что подтверждается кассовым чеком. 8 июля 2026 года телефон перестал включаться. Я обратился в магазин с требованием возврата денежных средств, однако продавец отказался удовлетворить моё требование, предложив обратиться в сервисный центр.
ЗАКОНЫ: п. 1 ст. 18 Закона РФ "О защите прав потребителей" от 07.02.1992 № 2300-1; ст. 22 того же закона (10 дней на удовлетворение требований); ст. 23 того же закона (неустойка 1% в день); ст. 15 того же закона (компенсация морального вреда)
ТРЕБОВАНИЯ: 
1. Вернуть уплаченные денежные средства в размере 80 000 рублей в течение 10 дней
2. Выплатить неустойку в размере 1% от стоимости товара за каждый день просрочки
3. Компенсировать моральный вред
ПРИЛОЖЕНИЯ: 
1. Копия кассового чека
2. Копия гарантийного талона
3. Копия данной претензии с отметкой о вручении

=== КОНЕЦ ПРИМЕРА ===

КРИТИЧЕСКИ ВАЖНО:
- СИТУАЦИЯ: пиши ОТ ПЕРВОГО ЛИЦА ("Я приобрёл", а НЕ "Клиент приобрёл")
- ЗАКОНЫ: выпиши ВСЕ статьи из ответа юриста (если упомянуты ст. 18, 22, 13 — укажи ВСЕ)
- ТРЕБОВАНИЯ: указывай КОНКРЕТНЫЕ суммы и сроки
- Не используй markdown (**, *, __)
- Не добавляй пояснений — только поля в указанном формате"""

        with GigaChat(
            credentials=auth_key,
            scope="GIGACHAT_API_PERS",
            verify_ssl_certs=False
        ) as giga:
            messages = [Messages(role=MessagesRole.USER, content=prompt)]
            response = giga.chat(Chat(messages=messages))
            result_text = safe_decode(response.choices[0].message.content)
        
        print(f"🔍 Сырой ответ ИИ:\n{result_text}\n")
        
        # Парсим ответ
        result = parse_template_text(result_text)
        return result
        
    except Exception as e:
        print(f"❌ Ошибка анализа: {e}")
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
            elif any(x in line_upper for x in ['КОМУ:', 'АДРЕСАТ:', 'АДРЕСАТУ:']):
                save_section()
                current_section = 'recipient'
                content = re.sub(r'^(КОМУ|АДРЕСАТ|АДРЕСАТУ):\s*', '', line_stripped, flags=re.IGNORECASE)
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
    legal_basis = template_data.get('legal_basis', '').strip()
    if not legal_basis or legal_basis == '[Статьи законов]':
        # Если ИИ не заполнил — добавляем универсальный текст
        legal_basis = ("В соответствии с действующим законодательством Российской Федерации, "
                      "а также на основании прав, предоставленных мне нормативными правовыми актами, "
                      "регулирующими данную ситуацию.")
    
    p = doc.add_paragraph()
    run = p.add_run("Правовое обоснование: ")
    run.bold = True
    p.add_run(legal_basis)
    doc.add_paragraph()
    
    # === ТРЕБОВАНИЯ ===
    p = doc.add_paragraph()
    run = p.add_run("На основании вышеизложенного ТРЕБУЮ:")
    run.bold = True
    
    requirements = template_data.get('requirements', '').strip()
    if not requirements or requirements == '[Требования]':
        requirements = "Удовлетворить мои законные требования в соответствии с действующим законодательством РФ."
    
    # Разбиваем требования по пунктам (ищем цифры в начале строки)
    req_lines = []
    current_line = ""
    for line in requirements.split('\n'):
        line = line.strip()
        if not line:
            continue
        # Если строка начинается с цифры — это новый пункт
        if re.match(r'^\d+[\.\)]\s*', line):
            if current_line:
                req_lines.append(current_line.strip())
            current_line = re.sub(r'^\d+[\.\)]\s*', '', line)
        else:
            current_line += " " + line
    if current_line:
        req_lines.append(current_line.strip())
    
    # Если не получилось разбить — пробуем разбить по точке с запятой
    if len(req_lines) <= 1 and ';' in requirements:
        req_lines = [line.strip() for line in requirements.split(';') if line.strip()]
    
    # Если всё ещё один пункт — добавляем как есть
    if len(req_lines) == 0:
        req_lines = [requirements]
    
    # Выводим каждый пункт отдельным абзацем
    for i, line in enumerate(req_lines, 1):
        # Убираем номера, если они есть
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
        attachments = "1. Копия данной претензии с отметкой о вручении\n2. Копии документов, подтверждающих мои требования"
    
    # Разбиваем приложения по пунктам
    att_lines = []
    current_line = ""
    for line in attachments.split('\n'):
        line = line.strip()
        if not line:
            continue
        # Если строка начинается с цифры — это новый пункт
        if re.match(r'^\d+[\.\)]\s*', line):
            if current_line:
                att_lines.append(current_line.strip())
            current_line = re.sub(r'^\d+[\.\)]\s*', '', line)
        else:
            current_line += " " + line
    if current_line:
        att_lines.append(current_line.strip())
    
    # Если не получилось разбить — пробуем разбить по точке с запятой
    if len(att_lines) <= 1 and ';' in attachments:
        att_lines = [line.strip() for line in attachments.split(';') if line.strip()]
    
    # Если всё ещё один пункт — добавляем как есть
    if len(att_lines) == 0:
        att_lines = [attachments]
    
    # Выводим каждый пункт отдельным абзацем
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
    safe_name = re.sub(r'[^\w\-.]', '_', doc_type)
    output_path = Path(output_dir) / f"{safe_name}.docx"
    doc.save(str(output_path))
    
    return str(output_path)
