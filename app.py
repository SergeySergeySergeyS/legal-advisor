"""
🎓 Юридический консультант для граждан на базе ИИ
Главный файл приложения
"""
import streamlit as st
from datetime import datetime
import legal_advisor


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
FREE_LIMIT = 999  # вопросов в день (временно для тестирования)


# === ИНИЦИАЛИЗАЦИЯ SESSION STATE ===
if "messages" not in st.session_state:
    st.session_state.messages = []

if "questions_today" not in st.session_state:
    st.session_state.questions_today = 0

if "last_reset_date" not in st.session_state:
    st.session_state.last_reset_date = datetime.now().date()

if "selected_category" not in st.session_state:
    st.session_state.selected_category = CATEGORIES[0]


# === СБРОС СЧЁТЧИКА В НАЧАЛЕ НОВОГО ДНЯ ===
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
    
    # === ПОЛУЧЕНИЕ КЛЮЧА GIGACHAT ===
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
    
    # === ВЫБОР КАТЕГОРИИ ===
    st.markdown("### 📋 Категория вопроса")
    selected_category = st.selectbox(
        "Выберите тему:",
        CATEGORIES,
        index=CATEGORIES.index(st.session_state.selected_category)
    )
    st.session_state.selected_category = selected_category
    
    st.markdown("---")
    
    # === СЧЁТЧИК ЗАПРОСОВ ===
    st.markdown("### 📊 Ваш тариф")
    remaining = max(0, FREE_LIMIT - st.session_state.questions_today)
    
    if remaining > 0:
        st.info(f"🆓 **Free-тариф**\n\nОсталось вопросов сегодня: **{remaining}** из {FREE_LIMIT}")
    else:
        st.warning(f"⚠️ **Лимит Free-тарифа исчерпан**\n\nВозвращайтесь завтра или оформите Premium 💎")
    
    st.markdown("---")
    
    # === УПРАВЛЕНИЕ ЧАТОМ ===
    st.markdown("### 🗑️ Управление чатом")
    if st.button("🗑️ Очистить историю", use_container_width=True):
        st.session_state.messages = []
        st.rerun()
    
    st.markdown("---")
    
    # === ИНФОРМАЦИЯ ===
    st.markdown("### ℹ️ О сервисе")
    st.markdown("""
    🤖 Работает на базе GigaChat (Сбер)
    
    ⚖️ Ссылается на конкретные статьи законов
    
    📋 Даёт пошаговые инструкции
    
    📄 Генерирует шаблоны документов
    """)
    
    st.markdown("---")
    st.markdown(f"📅 **Сегодня:** {datetime.now().strftime('%d.%m.%Y')}")


# === ОСНОВНАЯ ОБЛАСТЬ — ЧАТ ===

# Приветственное сообщение, если чат пустой
if not st.session_state.messages:
    st.markdown(f"""
    <div style='background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin-bottom: 20px;'>
        <h3>👋 Здравствуйте! Я ваш ИИ-юрист.</h3>
        <p>Сейчас выбрана категория: <b>{selected_category}</b></p>
        <p>Опишите вашу ситуацию, и я помогу разобраться в правовых вопросах:</p>
        <ul>
            <li>📝 Дам краткий ответ на вашу ситуацию</li>
            <li>⚖️ Сошлюсь на конкретные статьи законов</li>
            <li>📋 Составлю пошаговую инструкцию</li>
            <li>📄 Подскажу, какие документы нужны</li>
            <li>🏛️ Расскажу, куда обращаться</li>
            <li>⏰ Укажу важные сроки</li>
        </ul>
        <p><b>Напишите ваш вопрос ниже 👇</b></p>
    </div>
    """, unsafe_allow_html=True)

# Отображаем историю чата
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# === ПОЛЕ ВВОДА ===
if prompt := st.chat_input("💬 Напишите ваш юридический вопрос..."):
    # Проверка лимита Free-тарифа
    if st.session_state.questions_today >= FREE_LIMIT:
        st.error(f"""
        ⚠️ **Лимит Free-тарифа исчерпан!**
        
        Сегодня вы уже задали {FREE_LIMIT} вопроса.
        
        💎 Оформите Premium для безлимитных консультаций (скоро).
        
        🔄 Возвращайтесь завтра — счётчик обновится автоматически.
        """)
        st.stop()
    
    # Добавляем сообщение пользователя
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)
    
    # Получаем ответ от ИИ
    with st.chat_message("assistant"):
        with st.spinner("🤖 Консультируюсь с законами..."):
            # Формируем историю для контекста
            chat_history = [
                {"role": msg["role"], "content": msg["content"]}
                for msg in st.session_state.messages[:-1]
            ]
            
            # Запрашиваем ответ у ИИ
            answer = legal_advisor.ask_legal_advisor(
                question=prompt,
                category=st.session_state.selected_category,
                chat_history=chat_history,
                auth_key=auth_key
            )
            
            # Форматируем ответ
            formatted_answer = legal_advisor.format_answer_for_display(answer)
            
            # Отображаем ответ
            st.markdown(formatted_answer)
            
            # Сохраняем ответ в историю
            st.session_state.messages.append({"role": "assistant", "content": formatted_answer})
            
            # Увеличиваем счётчик вопросов
            st.session_state.questions_today += 1


# === ПОДВАЛ ===
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #888; font-size: 12px;'>
⚖️ <b>Важно:</b> Консультации носят информационный характер и не заменяют очную консультацию юриста.<br>
Для сложных случаев рекомендуется обратиться к профессиональному юристу.
</div>
""", unsafe_allow_html=True)
