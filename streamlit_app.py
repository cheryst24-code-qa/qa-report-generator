# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
import matplotlib.pyplot as plt
import io
import tempfile
import os

def set_col_width(col, width_twips):
    """Устанавливает ширину колонки в таблице DOCX"""
    for cell in col.cells:
        tc = cell._element.tcPr
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), str(int(width_twips)))
        tcW.set(qn('w:type'), 'dxa')
        tc.append(tcW)

def plot_to_buffer():
    """Сохраняет диаграмму в буфер и возвращает его"""
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    plt.close()
    return buf

def add_table_from_df(doc, df):
    """Создаёт таблицу с фиксированной шириной и границами"""
    if df.empty:
        table = doc.add_table(rows=2, cols=len(df.columns))
        for i, col in enumerate(df.columns):
            table.cell(0, i).text = str(col)
            table.cell(1, i).text = ""
    else:
        table = doc.add_table(rows=df.shape[0] + 1, cols=len(df.columns))
    
    table.style = 'Table Grid'
    total_width = Inches(6.5)
    
    # Установка ширины колонок
    num_cols = len(df.columns)
    if num_cols > 0:
        # Первая колонка (обычно ID) — 15% ширины
        first_width_twips = int(total_width.twips * 0.15)
        remaining_width_twips = total_width.twips - first_width_twips
        other_width_twips = int(remaining_width_twips / (num_cols - 1)) if num_cols > 1 else int(remaining_width_twips)
    
        set_col_width(table.columns[0], first_width_twips)
        for i in range(1, num_cols):
            set_col_width(table.columns[i], other_width_twips)

    for i, col_name in enumerate(df.columns):
        cell = table.cell(0, i)
        cell.text = str(col_name)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
            paragraph.paragraph_format.space_after = Pt(2)
            paragraph.paragraph_format.space_before = Pt(2)

    if not df.empty:
        for row_idx, (_, row) in enumerate(df.iterrows()):
            for col_idx, value in enumerate(row):
                cell = table.cell(row_idx + 1, col_idx)
                cell.text = str(value) if pd.notna(value) else ""
                cell.paragraphs[0].paragraph_format.space_after = Pt(2)
                cell.paragraphs[0].paragraph_format.space_before = Pt(2)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

def set_col_width(col, width_twips):
    """Устанавливает ширину колонки в таблице DOCX"""
    for cell in col.cells:
        tc = cell._element.tcPr
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), str(int(width_twips)))
        tcW.set(qn('w:type'), 'dxa')
        tc.append(tcW)

def generate_docx(data, module_data_list, defects_df):
    """Генерирует строго деловой DOCX-отчет"""
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    
    # === ЗАГОЛОВОК ОТЧЕТА ===
    title = doc.add_heading(data["report_title"], 0)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    title_font = title.runs[0].font
    title_font.size = Pt(16)
    title_font.bold = True

    # === ИНФОРМАЦИОННЫЕ ПОЛЯ (в виде таблицы с фиксированной шириной) ===
    info_table = doc.add_table(rows=6, cols=2)
    info_table.style = 'Table Grid'
    total_width = Inches(6.5)
    
    # Устанавливаем ширину колонок: первая колонка — 15%, вторая — 85%
    first_col_width = total_width * 0.15
    second_col_width = total_width * 0.85
    
    for row in info_table.rows:
        row.cells[0].width = first_col_width
        row.cells[1].width = second_col_width
    
    fields = [
        ('Проект:', data["project"]),
        ('Тип приложения:', data["app_type"]),
        ('Версия приложения:', data["version"]),
        ('Период тестирования:', data["test_period"]),
        ('Дата формирования отчёта:', data["report_date"]),
        ('Тест-инженер:', data["engineer"])
    ]
    
    for i, (label, value) in enumerate(fields):
        cell1 = info_table.cell(i, 0)
        cell1.text = label
        cell1.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        for run in cell1.paragraphs[0].runs:
            run.font.bold = True
        
        cell2 = info_table.cell(i, 1)
        cell2.text = value
        cell2.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # === КРАТКОЕ РЕЗЮМЕ (в виде таблицы с фиксированной шириной) ===
    doc.add_heading('1. КРАТКОЕ РЕЗЮМЕ', 1)
    
    summary_table = doc.add_table(rows=8, cols=2)
    summary_table.style = 'Table Grid'
    
    # Устанавливаем ширину колонок: первая колонка — 15%, вторая — 85%
    for row in summary_table.rows:
        row.cells[0].width = first_col_width
        row.cells[1].width = second_col_width
    
    total = data['total_tc']
    pass_pct = data['pass'] / total * 100 if total > 0 else 0
    fail_pct = 100 - pass_pct
    
    summary_fields = [
        ('Статус релиза:', data['release_status']),
        ('Критические дефекты (S1):', str(data['s1'])),
        ('Мажорные дефекты (S2):', str(data['s2'])),
        ('Всего тест-кейсов:', str(data['total_tc'])),
        ('Успешно (Pass):', f"{data['pass']} ({pass_pct:.1f}%)"),
        ('Упали (Fail):', f"{data['fail']} ({fail_pct:.1f}%)"),
        ('Основной риск:', data['risk']),
        ('Рекомендация:', data['recommendation'])
    ]
    
    for i, (label, value) in enumerate(summary_fields):
        cell1 = summary_table.cell(i, 0)
        cell1.text = label
        cell1.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        for run in cell1.paragraphs[0].runs:
            run.font.bold = True
        
        cell2 = summary_table.cell(i, 1)
        cell2.text = value
        cell2.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # === ДИАГРАММЫ ===
    plt.figure(figsize=(5, 4))
    plt.pie([data['pass'], data['fail']], labels=['PASS', 'FAIL'], autopct='%1.1f%%',
            colors=['#4CAF50', '#F44336'], startangle=90)
    plt.title('Рис. 1. Распределение результатов тест-кейсов')
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    plt.close()
    doc.add_picture(buf, width=Inches(5))
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    plt.figure(figsize=(5, 4))
    bars = plt.bar(['Critical (S1)', 'Major (S2)'], [data['s1'], data['s2']],
                   color=['#F44336', '#FF9800'])
    plt.title('Рис. 2. Дефекты по уровню серьёзности')
    plt.ylabel('Количество')
    for bar in bars:
        h = bar.get_height()
        if h > 0:
            plt.text(bar.get_x() + bar.get_width()/2, h + 0.05, str(int(h)), ha='center', va='bottom')
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    plt.close()
    doc.add_picture(buf, width=Inches(5))
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # === КОНТЕКСТ ТЕСТИРОВАНИЯ (в виде таблицы с фиксированной шириной) ===
    doc.add_heading('2. КОНТЕКСТ ТЕСТИРОВАНИЯ', 1)
    context_table = doc.add_table(rows=6, cols=2)
    context_table.style = 'Table Grid'
    
    # Устанавливаем ширину колонок: первая колонка — 15%, вторая — 85%
    for row in context_table.rows:
        row.cells[0].width = first_col_width
        row.cells[1].width = second_col_width
    
    context_fields = [
        ('Устройство / Браузер:', data['device_browser']),
        ('ОС / Платформа:', data['os_platform']),
        ('Сборка / Версия:', data['build']),
        ('Стенд:', f"Тестовое окружение (адрес: {data['env_url']})"),
        ('Инструменты:', data['tools']),
        ('Методология:', data['methodology'])
    ]
    
    for i, (label, value) in enumerate(context_fields):
        cell1 = context_table.cell(i, 0)
        cell1.text = label
        cell1.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        for run in cell1.paragraphs[0].runs:
            run.font.bold = True
        
        cell2 = context_table.cell(i, 1)
        cell2.text = value
        cell2.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # === РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ ПО МОДУЛЯМ ===
    doc.add_heading('3. РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ ПО МОДУЛЯМ', 1)
    for idx, module_info in enumerate(module_data_list):
        title = module_info['title']
        df = module_info['df']
        doc.add_heading(f'3.{idx+1}. {title}', 2)
        add_table_from_df(doc, df)  # <<< Для таблиц модулей используется отдельная функция

    # === АНАЛИЗ ДЕФЕКТОВ ===
    doc.add_heading('4. АНАЛИЗ ДЕФЕКТОВ', 1)
    add_table_from_df(doc, defects_df)  # <<< Для таблицы дефектов используется отдельная функция

    doc.add_paragraph('Последствия:').paragraph_format.space_after = Pt(6)
    doc.add_paragraph(data['consequences']).paragraph_format.space_after = Pt(6)

    # === ОГРАНИЧЕНИЯ ТЕСТИРОВАНИЯ ===
    doc.add_heading('5. ОГРАНИЧЕНИЯ ТЕСТИРОВАНИЯ', 1)
    for line in data['limitations'].split('\n'):
        if line.strip():
            p = doc.add_paragraph()
            p.add_run(f"• {line.strip()}")
            p.paragraph_format.space_after = Pt(2)

    # === ВЫВОД И РЕКОМЕНДАЦИИ ===
    doc.add_heading('6. ВЫВОД И РЕКОМЕНДАЦИИ', 1)
    doc.add_paragraph('Вывод:').paragraph_format.space_after = Pt(6)
    doc.add_paragraph(data['conclusion']).paragraph_format.space_after = Pt(6)
    doc.add_paragraph('Рекомендации:').paragraph_format.space_after = Pt(6)
    for line in data['recommendations_detailed'].split('\n'):
        if line.strip():
            p = doc.add_paragraph()
            p.add_run(f"• {line.strip()}")
            p.paragraph_format.space_after = Pt(2)

    # === ПОДПИСЬ (в виде таблицы с фиксированной шириной) ===
    doc.add_heading('7. ПОДПИСЬ', 1)
    signature_table = doc.add_table(rows=3, cols=2)
    signature_table.style = 'Table Grid'
    
    # Устанавливаем ширину колонок: первая колонка — 15%, вторая — 85%
    for row in signature_table.rows:
        row.cells[0].width = first_col_width
        row.cells[1].width = second_col_width
    
    signature_fields = [
        ('Роль:', data['role']),
        ('ФИО:', data['fullname']),
        ('Дата:', data['signature_date'])
    ]
    
    for i, (label, value) in enumerate(signature_fields):
        cell1 = signature_table.cell(i, 0)
        cell1.text = label
        cell1.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        for run in cell1.paragraphs[0].runs:
            run.font.bold = True
        
        cell2 = signature_table.cell(i, 1)
        cell2.text = value
        cell2.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# === ДАННЫЕ ПО УМОЛЧАНИЮ ===
default_modules = [
    {"title": "Главный экран и навигация", "df": pd.DataFrame([
        ["MAIN-01", "Отображение карточек товаров", "PASS", "—"],
        ["MAIN-02", "Фильтрация по категориям", "PASS", "—"],
        ["NAV-01", "Переход между разделами", "PASS", "—"],
        ["NAV-02", "Поиск товара с опечаткой", "FAIL", "BUG-SEARCH-001 . Не находятся товары при ошибке в 1 символе (например, «мыло» → «мылоо»)"]
    ], columns=["ID", "Сценарий", "Статус", "Комментарий"])},
    
    {"title": "Аутентификация и безопасность", "df": pd.DataFrame([
        ["AUTH-01", "Вход по логину/паролю", "PASS", "—"],
        ["SEC-01", "SQL-инъекция в поле поиска", "FAIL", "BUG-SEC-001 . При вводе `' OR '1'='1` — белый экран, частичный краш"],
        ["SEC-02", "XSS-атака через поле поиска", "FAIL", "BUG-SEC-002 . При вводе `<script>alert(1)</script>` — выполнение скрипта"]
    ], columns=["ID", "Сценарий", "Статус", "Комментарий"])},
    
    {"title": "Каталог и корзина", "df": pd.DataFrame([
        ["CATALOG-01", "Отображение списка товаров", "PASS", "—"],
        ["CART-01", "Добавление в корзину", "PASS", "—"],
        ["CART-02", "Оформление заказа", "PASS", "—"]
    ], columns=["ID", "Сценарий", "Статус", "Комментарий"])},
    
    {"title": "Дополнительные сценарии", "df": pd.DataFrame([
        ["OFFLINE-01", "Работа без интернета", "PASS", "Кэширование работает корректно"],
        ["SPECIAL-01", "Поиск со спецсимволами (@, #, $)", "PASS", "—"]
    ], columns=["ID", "Сценарий", "Статус", "Комментарий"])}
]

default_defects = pd.DataFrame([
    ["BUG-SEARCH-001", "Поиск", "Не работает fuzzy search (поиск с опечатками)", "Major (S2)", "New"],
    ["BUG-SEC-001", "Безопасность", "Уязвимость к SQL-инъекциям в поле поиска", "Critical (S1)", "New"],
    ["BUG-SEC-002", "Безопасность", "Уязвимость к XSS-атакам в поле поиска", "Critical (S1)", "New"]
], columns=["ID", "Модуль", "Заголовок", "Серьёзность", "Статус"])

# === ИНТЕРФЕЙС STREAMLIT ===
st.set_page_config(page_title="Универсальный генератор QA-отчёта", layout="wide")
st.title("📄 Универсальный генератор отчёта о тестировании")

# === ФОРМА ВВОДА ===
with st.form("main_form"):
    report_title = st.text_input(
        "Название отчёта",
        "Отчёт о тестировании мобильного приложения Лемана ПРО"
    )
    
    st.header("1. Краткое резюме")
    col1, col2 = st.columns(2)
    with col1:
        project = st.text_input("Проект", "Лемана ПРО")
        app_type = st.selectbox("Тип приложения", ["Мобильное", "Веб-приложение"])
        version = st.text_input("Версия приложения", "241006.001")
        test_period = st.text_input("Период тестирования", "29–30 ноября 2025 г.")
        report_date = st.text_input("Дата формирования отчёта", "30 ноября 2025 г.")
        engineer = st.text_input("Тест-инженер", "Черкасов Игорь")
    with col2:
        release_status = st.selectbox("Статус релиза", ["НЕ РЕКОМЕНДОВАН К ВЫПУСКУ", "РЕКОМЕНДОВАН К ВЫПУСКУ"], index=0)
        s1 = st.number_input("Критические дефекты (S1)", min_value=0, value=2)
        s2 = st.number_input("Мажорные дефекты (S2)", min_value=0, value=1)
        total_tc = st.number_input("Всего тест-кейсов", min_value=1, value=72)
        pass_tc = st.number_input("Успешно (Pass)", min_value=0, value=69)
        fail_tc = st.number_input("Упали (Fail)", min_value=0, value=3)
    
    risk = st.text_area("Основной риск", "Уязвимости безопасности позволяют нарушителю получить доступ к данным пользователей и вызвать отказ в обслуживании.")
    recommendation = st.text_area("Рекомендация", "Релиз возможен только после устранения всех S1/S2 дефектов и повторного тестирования.")

    st.header("2. Контекст тестирования")
    col3, col4 = st.columns(2)
    with col3:
        device_browser = st.text_input("Устройство / Браузер", "Xiaomi 12")
        os_platform = st.text_input("ОС / Платформа", "Android 15")
        build = st.text_input("Сборка", "lemanna-pro_241006.001.apk")
    with col4:
        env_url = st.text_input("URL стенда", "https://test.lemanna.pro        ")
        tools = st.text_input("Инструменты", "Postman (API), Burp Suite (безопасность), Jira (баг-трекинг)")
        methodology = st.text_input("Методология", "Ручное функциональное тестирование + проверка безопасности")

    st.header("3. Результаты тестирования по модулям")
    
    num_modules = st.slider("Количество модулей", min_value=1, max_value=10, value=4)
    
    module_data_list = []
    for i in range(num_modules):
        with st.expander(f"Модуль 3.{i+1}", expanded=True):
            title = st.text_input(f"Название модуля 3.{i+1}", value=default_modules[i]["title"] if i < len(default_modules) else f"Модуль 3.{i+1}")
            df_key = f"mod_{i}"
            default_df = default_modules[i]["df"] if i < len(default_modules) else pd.DataFrame(columns=["ID", "Сценарий", "Статус", "Комментарий"])
            df = st.data_editor(default_df, num_rows="dynamic", key=df_key)
            module_data_list.append({"title": title, "df": df})

    st.header("4. Анализ дефектов")
    defects = st.data_editor(default_defects, num_rows="dynamic", key="defects")
    consequences = st.text_area("Последствия", "- S1 дефекты позволяют злоумышленнику получить данные других пользователей или вывести приложение из строя.\n- S2 дефект снижает юзабилити: пользователи не найдут товар при опечатке.")

    st.header("5. Ограничения тестирования")
    limitations = st.text_area("Ограничения тестирования", "1. Не тестировалась оплата через Apple Pay (устройство Android).\n2. Не проверена синхронизация с 1С (нет доступа к интеграционному стенду).\n3. Не проведено нагрузочное тестирование (ограничение по времени).")
    
    st.header("6. Вывод и рекомендации")
    conclusion = st.text_area("Вывод", "Сборка 241006.001 содержит критические уязвимости безопасности, делающие её непригодной для выпуска в production. Наличие S1 дефектов нарушает базовые принципы защиты данных пользователей.")
    recommendations_detailed = st.text_area("Рекомендации (подробно)", "Немедленно исправить уязвимости BUG-SEC-001 и BUG-SEC-002.\nРеализовать fuzzy search для повышения юзабилити (BUG-SEARCH-001).\nПровести повторное тестирование после фиксов с фокусом на:\n• Повторную проверку полей ввода на инъекции\n• Тестирование сценариев поиска с опечатками\n• Настроить автоматизированную проверку безопасности (например, OWASP ZAP) в CI/CD.")
    
    st.header("7. Подпись")
    role = st.text_input("Роль", "Тест-инженер")
    fullname = st.text_input("ФИО", "Черкасов Игорь")
    signature_date = st.text_input("Дата", "30.11.2025")

    submitted = st.form_submit_button("📥 Создать отчёт")

if submitted:
    # === ПОДГОТОВКА ДАННЫХ ===
    data = {
        "report_title": report_title,
        "project": project,
        "app_type": app_type,
        "version": version,
        "test_period": test_period,
        "report_date": report_date,
        "engineer": engineer,
        "release_status": release_status,
        "s1": s1,
        "s2": s2,
        "total_tc": total_tc,
        "pass": pass_tc,
        "fail": fail_tc,
        "device_browser": device_browser,
        "os_platform": os_platform,
        "build": build,
        "env_url": env_url,
        "tools": tools,
        "methodology": methodology,
        "risk": risk,
        "recommendation": recommendation,
        "consequences": consequences,
        "limitations": limitations,
        "conclusion": conclusion,
        "recommendations_detailed": recommendations_detailed,
        "role": role,
        "fullname": fullname,
        "signature_date": signature_date,
    }
    
    try:
        # === ГЕНЕРАЦИЯ DOCX ===
        docx_buffer = generate_docx(data, module_data_list, defects)
        st.success("✅ Отчёт готов!")
        
        # === КНОПКА СКАЧИВАНИЯ ===
        st.download_button(
            "📄 Скачать .docx",
            docx_buffer,
            "Отчёт_о_тестировании.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
            
    except Exception as e:
        st.error(f"❌ Ошибка: {e}")