"""
🎓 Юридический консультант для граждан на базе ИИ
Главный файл приложения
"""
import streamlit as st
from datetime import datetime
import tempfile
from pathlib import Path
import legal_advisor
import template_generator


# === НАСТРОЙКИ СТРАНИЦЫ ===
st.set_page_config(
    page_title="🎓 Юридический консультант",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# === КАТЕГОРИИ ВОПРОСОВ ===
CATEGORIES = [
    "🛒 Защита прав потребителей",
    "🏠 ЖКХ",
    "💼 Трудовые споры",
    "🚗 Споры с ГИБДД",
]

# === ЛИМИТЫ ТАРИФОВ ===
FREE_LIMIT = 999


# === ИНИЦИАЛИЗАЦИЯ SESSION STATE ===
if "messages" not in st.session_state:
    st.session_state.messages = []

if "questions_today" not in st.session_state:
    st.session_state.questions_today = 0

if "last_reset_date" not in st.session_state:
    st.session_state.last_reset_date = datetime.now().date()

if "selected_category" not in st.session_state:
    st.session_state.selected_category = CATEGORIES[0]


# === СБРОС СЧЁТЧИКА ===
today = datetime.now().date()
if st.session_state.last_reset_date != today:
    st.session_state.questions_today = 0
    st.session_state.last_reset_date = today


# === ЗАГОЛОВОК ===
st.title("🎓 Юридический консультант")
st.markdown("""
**Ваш персональный ИИ-юрист — отвечает на вопросы простым языком, со ссылками на законы и пошаговыми инструкциями**

💬 Задайте вопрос → 🤖 Получите консультацию → 📄 Скачайте шаблон документа
""")


# === БОКОВАЯ ПАНЕЛЬ ===
with st.sidebar:
    st.header("⚙️ Настройки")
    
    # === КЛЮЧ GIGACHAT ===
    try:
        auth_key = st.secrets.get("AUTH_KEY", "")
    except Exception:
        auth_key = ""
    
    if not auth_key:
        import os
        auth_key = os.environ.get("AUTH_KEY", "")
    
    if not auth_key:
        auth_key = st.text_input(
            "🔑 Ключ GigaChat API",
            type="password",
            help="Получите ключ на developers.sber.ru"
        )
        if not auth_key:
            st.warning("⚠️ Введите ключ GigaChat API")
            st.stop()
    else:
        st.success("✅ Ключ GigaChat получен")
    
    st.markdown("---")
    
    # === КАТЕГОРИЯ ===
    st.markdown("### 📋 Категория вопроса")
    selected_category = st.selectbox(
        "Выберите тему:",
        CATEGORIES,
        index=CATEGORIES.index(st.session_state.selected_category)
    )
    st.session_state.selected_category = selected_category
    
    st.markdown("---")
    
    # === СЧЁТЧИК ===
    st.markdown("### 📊 Ваш тариф")
    remaining = max(0, FREE_LIMIT - st.session_state.questions_today)
    
    if remaining > 0:
        st.info(f"🆓 **Free-тариф**\n\nОсталось вопросов сегодня: **{remaining}** из {FREE_LIMIT}")
    else:
        st.warning(f"⚠️ **Лимит исчерпан**\n\nВозвращайтесь завтра")
    
    st.markdown("---")
    
    # === УПРАВЛЕНИЕ ===
    st.markdown("### 🗑️ Управление чатом")
    if st.button("🗑️ Очистить историю", use_container_width=True):
        st.session_state.messages = []
        st.rerun()
    
    st.markdown("---")
    
    st.markdown("### ℹ️ О сервисе")
    st.markdown("""
    🤖 GigaChat (Сбер)
    ⚖️ Ссылки на статьи законов
    📋 Пошаговые инструкции
    📄 Генератор документов
    """)
    
    st.markdown("---")
    st.markdown(f"📅 **Сегодня:** {datetime.now().strftime('%d.%m.%Y')}")


# === ОСНОВНАЯ ОБЛАСТЬ ===

# Приветствие
if not st.session_state.messages:
    st.markdown(f"""
    <div style='background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin-bottom: 20px;'>
        <h3>👋 Здравствуйте! Я ваш ИИ-юрист.</h3>
        <p>Сейчас выбрана категория: <b>{selected_category}</b></p>
        <p>Опишите вашу ситуацию, и я помогу разобраться:</p>
        <ul>
            <li>📝 Краткий ответ</li>
            <li>⚖️ Применимые законы</li>
            <li>📋 Пошаговая инструкция</li>
            <li>📄 Необходимые документы</li>
            <li>🏛️ Куда обращаться</li>
            <li>⏰ Важные сроки</li>
            <li>💰 Возможные требования</li>
        </ul>
        <p><b>После консультации вы сможете скачать готовый документ!</b></p>
        <p><b>Напишите ваш вопрос ниже 👇</b></p>
    </div>
    """, unsafe_allow_html=True)

# История чата
for i, message in enumerate(st.session_state.messages):
    with st.chat_message(message["role"]):
        st.markdown(message["content"])
        
        # Кнопка скачивания документа ПОСЛЕ каждого ответа ИИ
        if message["role"] == "assistant":
            button_key = f"doc_btn_{i}"
            download_key = f"doc_download_{i}"
            
            if st.button("📄 Скачать готовую претензию", key=button_key):
                with st.spinner("📝 Формирую документ..."):
                    try:
                        # Берём историю до этого ответа
                        history_for_doc = st.session_state.messages[:i+1]
                        
                        # Анализируем чат
                        template_data = template_generator.analyze_chat_for_template(
                            chat_history=history_for_doc,
                            category=st.session_state.selected_category,
                            auth_key=auth_key
                        )
                        
                        # Создаём документ во временной папке
                        with tempfile.TemporaryDirectory() as temp_dir:
                            doc_path = template_generator.generate_legal_document(
                                template_data=template_data,
                                output_dir=temp_dir
                            )
                            
                            # Читаем файл
                            with open(doc_path, 'rb') as f:
                                doc_bytes = f.read()
                            
                            # Сохраняем в session_state для кнопки скачивания
                            st.session_state[download_key] = {
                                'data': doc_bytes,
                                'filename': f"{template_data.get('document_type', 'документ')}.docx"
                            }
                            st.rerun()
                            
                    except Exception as e:
                        st.error(f"❌ Ошибка создания документа: {e}")
            
            # Кнопка скачивания (появляется после генерации)
            if download_key in st.session_state:
                st.download_button(
                    label="💾 Скачать документ",
                    data=st.session_state[download_key]['data'],
                    file_name=st.session_state[download_key]['filename'],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_{i}",
                    use_container_width=True
                )

# === ПОЛЕ ВВОДА ===
if prompt := st.chat_input("💬 Напишите ваш юридический вопрос..."):
    # Проверка лимита
    if st.session_state.questions_today >= FREE_LIMIT:
        st.error("⚠️ Лимит исчерпан! Возвращайтесь завтра.")
        st.stop()
    
    # Сообщение пользователя
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)
    
    # Ответ ИИ
    with st.chat_message("assistant"):
        with st.spinner("🤖 Консультируюсь с законами..."):
            chat_history = [
                {"role": msg["role"], "content": msg["content"]}
                for msg in st.session_state.messages[:-1]
            ]
            
            answer = legal_advisor.ask_legal_advisor(
                question=prompt,
                category=st.session_state.selected_category,
                chat_history=chat_history,
                auth_key=auth_key
            )
            
            formatted_answer = legal_advisor.format_answer_for_display(answer)
            st.markdown(formatted_answer)
            
            st.session_state.messages.append({"role": "assistant", "content": formatted_answer})
            st.session_state.questions_today += 1


# === ПОДВАЛ ===
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #888; font-size: 12px;'>
⚖️ <b>Важно:</b> Консультации носят информационный характер и не заменяют очную консультацию юриста.<br>
Для сложных случаев рекомендуется обратиться к профессиональному юристу.
</div>
""", unsafe_allow_html=True)
