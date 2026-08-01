import streamlit as st
from datetime import date
from gigachat import GigaChat

# ============ НАСТРОЙКИ СТРАНИЦЫ ============
st.set_page_config(page_title="ИИ Юрист", page_icon="⚖️", layout="wide")

# ============ ИНИЦИАЛИЗАЦИЯ GIGACHAT ============
@st.cache_resource
def get_gigachat():
    return GigaChat(
        credentials=st.secrets["GIGACHAT_CREDENTIALS"],
        verify_ssl_certs=False
    )

gigachat = get_gigachat()

# ============ ДИАГНОСТИКА (ПУЛЕНЕПРОБИВАЕМАЯ) ============
st.sidebar.divider()
if st.sidebar.button("🔍 Проверить доступные модели"):
    try:
        models = gigachat.get_models()
        st.sidebar.success("Запрос успешен! Сырой ответ от API:")
        # Выводим как строку, чтобы избежать ошибок атрибутов
        st.sidebar.code(str(models)) 
        
        # Пытаемся безопасно извлечь имена, если это список словарей или объектов
        try:
            # Пробуем разные варианты названий атрибутов
            model_names = [getattr(m, 'id', getattr(m, 'model', str(m))) for m in models.data]
            st.session_state["detected_model"] = model_names[0]
            st.sidebar.info(f"Первая найденная модель: `{model_names[0]}`")
        except:
            st.session_state["detected_model"] = "GigaChat:latest"
            
    except Exception as e:
        st.sidebar.error(f"Ошибка: {e}")

# ============ КАТЕГОРИИ ============
CATEGORIES = {
    "🛒 Защита прав потребителей": "Ты — юрист по защите прав потребителей. Отвечай со ссылками на Закон РФ 'О защите прав потребителей'.",
    "💼 Трудовые споры": "Ты — юрист по трудовому праву. Отвечай со ссылками на Трудовой кодекс РФ (ТК РФ).",
    "🚗 Споры с ГИБДД": "Ты — юрист по административному праву. Отвечай со ссылками на Кодекс РФ об административных правонарушениях (КоАП РФ).",
    "👨‍👩‍👧 Семейное право": "Ты — семейный юрист. Отвечай со ссылками на Семейный кодекс РФ (СК РФ).",
    "📜 Наследство": "Ты — юрист по наследственному праву. Отвечай со ссылками на Гражданский кодекс РФ (часть 3).",
    "🏠 Жилищные вопросы": "Ты — жилищный юрист. Отвечай со ссылками на Жилищный кодекс РФ (ЖК РФ)."
}

# ============ ТАРИФЫ И ЛИМИТЫ ============
TARIFFS = {
    "🆓 Free": {"price": 0, "limit": 3},
    "💎 Premium": {"price": 499, "limit": 100},
    "🏢 Business": {"price": 2990, "limit": 9999}
}

# ============ УПРАВЛЕНИЕ СЕССИЕЙ ============
if "messages" not in st.session_state:
    st.session_state.messages = []
if "questions_today" not in st.session_state:
    st.session_state.questions_today = 0
if "last_date" not in st.session_state:
    st.session_state.last_date = date.today().isoformat()
if "tariff" not in st.session_state:
    st.session_state.tariff = "🆓 Free"
if "detected_model" not in st.session_state:
    # ИСПРАВЛЕНИЕ: пробуем GigaChat:latest как наиболее вероятный рабочий вариант
    st.session_state.detected_model = "GigaChat:latest" 

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
    st.caption(f"📅 Сегодня: {date.today().strftime('%d.%m.%Y')}")

# ============ ОСНОВНОЙ ЭКРАН ============
st.title("🎓 Юридический консультант")
st.markdown("**💬 Задайте вопрос → 🤖 Получите консультацию**")

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# ============ ОБРАБОТКА ВВОДА ============
if prompt := st.chat_input("Задайте ваш юридический вопрос..."):
    if st.session_state.questions_today >= TARIFFS[tariff]["limit"]:
        st.error(f"❌ Лимит вопросов исчерпан ({limit}).")
        st.stop()
    
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)
    
    system_prompt = f"Ты — профессиональный юридический консультант. Категория: {category}. Отвечай со ссылками на законы РФ, пиши пошаговую инструкцию."
    
    with st.chat_message("assistant"):
        with st.spinner("🤖 Готовлю консультацию..."):
            try:
                current_model = st.session_state.get("detected_model", "GigaChat:latest")
                
                response = gigachat.chat({
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": prompt}
                    ],
                    "model": current_model
                })
                answer = response.choices[0].message.content
                st.markdown(answer)
                
                st.session_state.messages.append({"role": "assistant", "content": answer})
                st.session_state.questions_today += 1
                
            except Exception as e:
                st.error(f"❌ Ошибка: {e}")
