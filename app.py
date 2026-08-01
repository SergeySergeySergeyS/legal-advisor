import streamlit as st
from datetime import date
from gigachat import GigaChat

# ============ НАСТРОЙКИ ============
st.set_page_config(page_title="ИИ Юрист", page_icon="⚖️", layout="wide")

# Инициализация GigaChat с параметром model (это и есть наше исправление!)
@st.cache_resource
def get_gigachat():
    return GigaChat(
        credentials=st.secrets["GIGACHAT_CREDENTIALS"],
        model="GigaChat",           # <-- ВОТ ГЛАВНОЕ ИСПРАВЛЕНИЕ
        verify_ssl_certs=False
    )

gigachat = get_gigachat()

# ============ КАТЕГОРИИ ============
CATEGORIES = {
    "🛒 Защита прав потребителей": "Ты — юрист по защите прав потребителей. Отвечай со ссылками на ЗоЗПП.",
    "💼 Трудовые споры": "Ты — юрист по трудовому праву. Отвечай со ссылками на ТК РФ.",
    "🚗 Споры с ГИБДД": "Ты — юрист по административному праву. Отвечай со ссылками на КоАП РФ.",
    "👨‍👩‍👧 Семейное право": "Ты — семейный юрист. Отвечай со ссылками на СК РФ.",
    "📜 Наследство": "Ты — юрист по наследственному праву. Отвечай со ссылками на ГК РФ часть 3.",
    "🏠 Жилищные вопросы": "Ты — жилищный юрист. Отвечай со ссылками на ЖК РФ."
}

# ============ ТАРИФЫ И ЛИМИТЫ ============
TARIFFS = {
    "🆓 Free": {"price": 0, "limit": 3},
    "💎 Premium": {"price": 499, "limit": 100},
    "🏢 Business": {"price": 2990, "limit": 9999}
}

# ============ СЕССИЯ ============
if "messages" not in st.session_state:
    st.session_state.messages = []
if "questions_today" not in st.session_state:
    st.session_state.questions_today = 0
if "last_date" not in st.session_state:
    st.session_state.last_date = date.today().isoformat()
if "tariff" not in st.session_state:
    st.session_state.tariff = "🆓 Free"

# Сброс счётчика в новый день
if st.session_state.last_date != date.today().isoformat():
    st.session_state.questions_today = 0
    st.session_state.last_date = date.today().isoformat()

# ============ БОКОВАЯ ПАНЕЛЬ ============
with st.sidebar:
    st.title("⚙️ Настройки")
    st.success("✅ Ключ GigaChat получен")
    
    st.subheader("📋 Категория вопроса")
    category = st.selectbox("Выберите тему:", list(CATEGORIES.keys()))
    
    st.subheader("📊 Ваш тариф")
    tariff = st.selectbox("Тариф:", list(TARIFFS.keys()), 
                          index=list(TARIFFS.keys()).index(st.session_state.tariff))
    st.session_state.tariff = tariff
    
    limit = TARIFFS[tariff]["limit"]
    remaining = max(0, limit - st.session_state.questions_today)
    st.info(f"Осталось вопросов сегодня: **{remaining}** из {limit}")
    
    st.subheader("🗑️ Управление чатом")
    if st.button("Очистить чат", use_container_width=True):
        st.session_state.messages = []
        st.rerun()
    
    st.divider()
    st.subheader("ℹ️ О сервисе")
    st.caption("🤖 GigaChat (Сбер)")
    st.caption("⚖️ Ссылки на статьи законов")
    st.caption("📋 Пошаговые инструкции")
    st.caption("📄 4 типа документов: Претензия, Жалоба, Исковое заявление, Ходатайство")
    st.caption(f"📅 Сегодня: {date.today().strftime('%d.%m.%Y')}")

# ============ ЗАГОЛОВОК ============
st.title("🎓 Юридический консультант")
st.caption("Ваш персональный ИИ-юрист — отвечает на вопросы простым языком, со ссылками на законы и пошаговыми инструкциями")
st.markdown("**💬 Задайте вопрос → 🤖 Получите консультацию → 📄 Скачайте готовый документ**")

# ============ ИСТОРИЯ ЧАТА ============
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# ============ ВВОД ВОПРОСА ============
if prompt := st.chat_input("Задайте ваш юридический вопрос..."):
    # Проверка лимита
    if st.session_state.questions_today >= TARIFFS[tariff]["limit"]:
        st.error(f"❌ Лимит вопросов на сегодня исчерпан ({limit}). Оформите Premium-тариф!")
        st.stop()
    
    # Добавляем вопрос пользователя
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)
    
    # Формируем системный промпт
    system_prompt = f"""Ты — профессиональный юридический консультант. 
Категория вопроса: {category}.
Отвечай на русском языке, простым и понятным языком.
ОБЯЗАТЕЛЬНО:
1. Давай ссылки на конкретные статьи законов (ТК РФ, ГК РФ, КоАП, ЗоЗПП и т.д.).
2. Пиши пошаговую инструкцию, что делать.
3. В конце укажи, какие документы нужны.
4. Предложи готовый шаблон документа, если уместно."""
    
    # Запрос к GigaChat
    with st.chat_message("assistant"):
        with st.spinner("🤖 Готовлю юридическую консультацию..."):
            try:
                response = gigachat.chat(
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.3
                )
                answer = response.choices[0].message.content
                st.markdown(answer)
                st.session_state.messages.append({"role": "assistant", "content": answer})
                st.session_state.questions_today += 1
            except Exception as e:
                error_msg = f"❌ Произошла ошибка при получении ответа: {e}\n\nПожалуйста, проверьте ключ API и попробуйте снова."
                st.error(error_msg)
