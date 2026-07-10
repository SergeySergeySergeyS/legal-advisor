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

# === ТИПЫ ДОКУМЕНТОВ ===
DOC_TYPES = {
    "претензия": "📄 Претензия",
    "жалоба": "📄 Жалоба",
    "иск": "📄 Исковое заявление",
    "ходатайство": "📄 Ходатайство"
}

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

if "generated_docs" not in st.session_state:
    st.session_state.generated_docs = {}


# === СБРОС СЧЁТЧИКА ===
today = datetime.now().date()
if st.session_state.last_reset_date != today:
    st.session_state.questions_today = 0
    st.session_state.last_reset_date = today


# === ЗАГОЛОВОК ===
st.title("🎓 Юридический консультант")
st.markdown("""
**Ваш персональный ИИ-юрист — отвечает на вопросы простым языком, со ссылками на законы и пошаговыми инструкциями**

💬 Задайте вопрос → 🤖 Получите консультацию → 📄 Скачайте готовый документ
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
        st.session_state.generated_docs = {}
        st.rerun()
    
    st.markdown("---")
    
    st.markdown("### ℹ️ О сервисе")
    st.markdown("""
    🤖 GigaChat (Сбер)
    ⚖️ Ссылки на статьи законов
    📋 Пошаговые инструкции
    📄 4 типа документов:
       • Претензия
       • Жалоба
       • Исковое заявление
       • Ходатайство
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
        <p><b>После консультации вы сможете скачать готовый документ:</b></p>
        <ul>
            <li>📄 <b>Претензия</b> — контрагенту (продавцу, исполнителю)</li>
            <li>📄 <b>Жалоба</b> — в госорган (Роспотребнадзор, ГИТ, ГИБДД)</li>
            <li>📄 <b>Исковое заявление</b> — в суд</li>
            <li>📄 <b>Ходатайство</b> — в различные инстанции</li>
        </ul>
        <p><b>Напишите ваш вопрос ниже 👇</b></p>
    </div>
    """, unsafe_allow_html=True)

# История чата
for i, message in enumerate(st.session_state.messages):
    with st.chat_message(message["role"]):
        st.markdown(message["content"])
        
        # Кнопки скачивания документов ПОСЛЕ каждого ответа ИИ
        if message["role"] == "assistant":
            st.markdown("---")
            st.markdown("**📄 Выберите тип документа для скачивания:**")
            
            # 4 кнопки в ряд
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                btn_key = f"btn_pret_{i}"
                if st.button("📄 Претензия", key=btn_key, use_container_width=True):
                    _generate_and_store_doc(i, "претензия", auth_key)
            
            with col2:
                btn_key = f"btn_zhal_{i}"
                if st.button("📄 Жалоба", key=btn_key, use_container_width=True):
                    _generate_and_store_doc(i, "жалоба", auth_key)
            
            with col3:
                btn_key = f"btn_isk_{i}"
                if st.button("📄 Иск", key=btn_key, use_container_width=True):
                    _generate_and_store_doc(i, "иск", auth_key)
            
            with col4:
                btn_key = f"btn_hod_{i}"
                if st.button("📄 Ходатайство", key=btn_key, use_container_width=True):
                    _generate_and_store_doc(i, "ходатайство", auth_key)
            
            # Показываем кнопки скачивания для сгенерированных документов
            for doc_type in ["претензия", "жалоба", "иск", "ходатайство"]:
                doc_key = f"doc_{doc_type}_{i}"
                if doc_key in st.session_state.generated_docs:
                    doc_data = st.session_state.generated_docs[doc_key]
                    st.download_button(
                        label=f"💾 Скачать {doc_data['filename']}",
                        data=doc_data['data'],
                        file_name=doc_data['filename'],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_{doc_type}_{i}",
                        use_container_width=True
                    )

# === ПОЛЕ ВВОДА ===
if prompt := st.chat_input("💬 Напишите ваш юридический вопрос..."):
    if st.session_state.questions_today >= FREE_LIMIT:
        st.error("⚠️ Лимит исчерпан! Возвращайтесь завтра.")
        st.stop()
    
    st.session_state.messages.append({"role": "user", "content": prompt})
    
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
        st.session_state.messages.append({"role": "assistant", "content": formatted_answer})
        st.session_state.questions_today += 1
    
    st.rerun()


# === ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ ГЕНЕРАЦИИ ДОКУМЕНТА ===
def _generate_and_store_doc(message_index: int, doc_type: str, auth_key: str):
    """Генерирует документ и сохраняет его в session_state"""
    doc_key = f"doc_{doc_type}_{message_index}"
    
    with st.spinner(f"📝 Формирую {doc_type}..."):
        try:
            history_for_doc = st.session_state.messages[:message_index+1]
            
            template_data = template_generator.analyze_chat_for_template(
                chat_history=history_for_doc,
                category=st.session_state.selected_category,
                auth_key=auth_key,
                doc_type=doc_type
            )
            
            with tempfile.TemporaryDirectory() as temp_dir:
                doc_path = template_generator.generate_legal_document(
                    template_data=template_data,
                    output_dir=temp_dir,
                    doc_type=doc_type
                )
                
                with open(doc_path, 'rb') as f:
                    doc_bytes = f.read()
                
                doc_type_upper = template_data.get('document_type', doc_type).upper()
                st.session_state.generated_docs[doc_key] = {
                    'data': doc_bytes,
                    'filename': f"{doc_type_upper}.docx"
                }
                
                st.success(f"✅ Документ '{doc_type_upper}' готов!")
                st.rerun()
                
        except Exception as e:
            st.error(f"❌ Ошибка создания документа: {e}")


# === ПОДВАЛ ===
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #888; font-size: 12px;'>
⚖️ <b>Важно:</b> Консультации носят информационный характер и не заменяют очную консультацию юриста.<br>
Для сложных случаев рекомендуется обратиться к профессиональному юристу.
</div>
""", unsafe_allow_html=True)
