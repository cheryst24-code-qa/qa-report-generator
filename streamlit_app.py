import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import base64
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import json
from datetime import datetime
import numpy as np

# === ФУНКЦИИ ДЛЯ РАБОТЫ С ЧЕРНОВИКАМИ ===
def save_draft(data, module_data_list, defects_df):
    """Сохраняет данные формы в структуру для черновика"""
    draft = {
        "saved_at": datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
        "data": data,
        "modules": [
            {
                "title": m["title"],
                "df": m["df"].to_dict(orient="records") if not m["df"].empty else []
            }
            for m in module_data_list
        ],
        "defects": defects_df.to_dict(orient="records") if not defects_df.empty else []
    }
    return json.dumps(draft, ensure_ascii=False, indent=2)

def load_draft(json_content):
    """Восстанавливает данные из черновика"""
    try:
        draft = json.loads(json_content)
        
        # Восстанавливаем данные формы
        data = draft.get("data", {})
        
        # Восстанавливаем дефекты
        defects_records = draft.get("defects", [])
        defects_df = pd.DataFrame(
            defects_records,
            columns=["ID", "Модуль", "Заголовок", "Серьёзность", "Статус"]
        ) if defects_records else pd.DataFrame(columns=["ID", "Модуль", "Заголовок", "Серьёзность", "Статус"])
        
        # Восстанавливаем модули
        modules = []
        for mod in draft.get("modules", []):
            df_records = mod.get("df", [])
            df = pd.DataFrame(
                df_records,
                columns=["ID", "Сценарий", "Статус", "Комментарий"]
            ) if df_records else pd.DataFrame(columns=["ID", "Сценарий", "Статус", "Комментарий"])
            modules.append({"title": mod["title"], "df": df})
        
        return data, modules, defects_df, draft.get("saved_at", "неизвестно")
    except Exception as e:
        st.error(f"❌ Ошибка загрузки черновика: {str(e)}")
        return None, None, None, None

# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ DOCX ===
def set_col_width(table, col_idx, width_twips):
    """Устанавливает ширину колонки в таблице DOCX (в твипах)"""
    for row in table.rows:
        cell = row.cells[col_idx]
        tc = cell._tc
        tc_pr = tc.get_or_add_tcPr()
        tc_w = OxmlElement('w:tcW')
        tc_w.set(qn('w:w'), str(width_twips))
        tc_w.set(qn('w:type'), 'dxa')
        tc_pr.append(tc_w)

def add_table_from_df(doc, df, col_widths=None):
    """Добавляет таблицу из DataFrame в документ DOCX"""
    if df.empty or len(df.columns) == 0:
        doc.add_paragraph("Нет данных для отображения")
        return
    
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = 'Table Grid'
    
    # Заголовки
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(df.columns):
        hdr_cells[i].text = str(column)
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(10)
                run.font.name = 'Calibri Light'
    
    # Данные
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            cell_text = str(value) if pd.notna(value) else ""
            row_cells[i].text = cell_text
            for paragraph in row_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
                    run.font.name = 'Calibri Light'
    
    # Установка ширины колонок (если задана)
    if col_widths:
        for col_idx, width in enumerate(col_widths):
            set_col_width(table, col_idx, width)

# === ГЕНЕРАЦИЯ DOCX ОТЧЁТА ===
def generate_docx(data, module_data_list, defects_df):
    doc = Document()
    
    # Стиль документа
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri Light'
    font.size = Pt(11)
    
    # Заголовок
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(data["report_title"])
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.name = 'Calibri Light'
    doc.add_paragraph()
    
    # Основная информация (таблица 25%/75%)
    info_table = doc.add_table(rows=7, cols=2)
    info_table.style = 'Table Grid'
    
    fields = [
        ("Проект", data["project"]),
        ("Тип приложения", data["app_type"]),
        ("Версия", data["version"]),
        ("Период тестирования", data["test_period"]),
        ("Дата формирования отчёта", data["report_date"]),
        ("Инженер по тестированию", data["engineer"]),
        ("Статус релиза", data["release_status"]),
    ]
    
    for i, (label, value) in enumerate(fields):
        info_table.cell(i, 0).text = label
        info_table.cell(i, 1).text = str(value)
        # Жирный шрифт для лейблов
        for paragraph in info_table.cell(i, 0).paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Calibri Light'
                run.font.size = Pt(10)
        # Обычный шрифт для значений
        for paragraph in info_table.cell(i, 1).paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Calibri Light'
                run.font.size = Pt(10)
    
    # Установка ширины колонок 25%/75% (25% = 1800 твипов от общей ширины ~7200)
    set_col_width(info_table, 0, 1800)
    set_col_width(info_table, 1, 5400)
    doc.add_paragraph()
    
    # Краткое резюме
    doc.add_paragraph("Краткое резюме", style='Heading 2')
    
    # Метрики (таблица 25%/75%)
    metrics_table = doc.add_table(rows=4, cols=2)
    metrics_table.style = 'Table Grid'
    
    metrics = [
        ("Количество дефектов (S1)", f"{data['s1']}"),
        ("Количество дефектов (S2)", f"{data['s2']}"),
        ("Всего тест-кейсов", f"{data['total_tc']}"),
        ("Пройдено успешно", f"{data['pass']}"),
    ]
    
    for i, (label, value) in enumerate(metrics):
        metrics_table.cell(i, 0).text = label
        metrics_table.cell(i, 1).text = value
        for paragraph in metrics_table.cell(i, 0).paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Calibri Light'
                run.font.size = Pt(10)
        for paragraph in metrics_table.cell(i, 1).paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Calibri Light'
                run.font.size = Pt(10)
    
    set_col_width(metrics_table, 0, 1800)
    set_col_width(metrics_table, 1, 5400)
    doc.add_paragraph()
    
    # Риски и рекомендации
    if data["risk"].strip():
        doc.add_paragraph("Риски", style='Heading 3')
        doc.add_paragraph(data["risk"])
    
    if data["recommendation"].strip():
        doc.add_paragraph("Рекомендации", style='Heading 3')
        doc.add_paragraph(data["recommendation"])
    doc.add_paragraph()
    
    # Диаграммы (заглушки с описанием)
    doc.add_paragraph("Диаграммы", style='Heading 2')
    doc.add_paragraph("Рис. 1. Распределение результатов тестирования")
    doc.add_paragraph("[Диаграмма будет вставлена вручную]")
    doc.add_paragraph()
    doc.add_paragraph("Рис. 2. Распределение дефектов по серьёзности")
    doc.add_paragraph("[Диаграмма будет вставлена вручную]")
    doc.add_paragraph()
    
    # Контекст тестирования
    doc.add_paragraph("Контекст тестирования", style='Heading 2')
    
    context_table = doc.add_table(rows=5, cols=2)
    context_table.style = 'Table Grid'
    
    context_fields = [
        ("Устройство / Браузер", data["device_browser"]),
        ("ОС / Платформа", data["os_platform"]),
        ("Сборка", data["build"]),
        ("URL окружения", data["env_url"]),
        ("Инструменты", data["tools"]),
    ]
    
    for i, (label, value) in enumerate(context_fields):
        context_table.cell(i, 0).text = label
        context_table.cell(i, 1).text = str(value)
        for paragraph in context_table.cell(i, 0).paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Calibri Light'
                run.font.size = Pt(10)
        for paragraph in context_table.cell(i, 1).paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Calibri Light'
                run.font.size = Pt(10)
    
    set_col_width(context_table, 0, 1800)
    set_col_width(context_table, 1, 5400)
    doc.add_paragraph()
    
    if data["methodology"].strip():
        doc.add_paragraph("Методология тестирования", style='Heading 3')
        doc.add_paragraph(data["methodology"])
        doc.add_paragraph()
    
    # Результаты по модулям
    doc.add_paragraph("Результаты тестирования по модулям", style='Heading 2')
    
    for module in module_data_list:
        doc.add_paragraph(module["title"], style='Heading 3')
        add_table_from_df(doc, module["df"], col_widths=[1800, 5400, 1800, 5400])
        doc.add_paragraph()
    
    # Анализ дефектов
    if not defects_df.empty:
        doc.add_paragraph("Анализ дефектов", style='Heading 2')
        add_table_from_df(doc, defects_df)
        doc.add_paragraph()
    
    # Ограничения тестирования
    if data["limitations"].strip():
        doc.add_paragraph("Ограничения тестирования", style='Heading 2')
        # Преобразуем маркированный список в нумерованный
        lines = [line.strip() for line in data["limitations"].split('\n') if line.strip()]
        for line in lines:
            # Убираем маркеры "-", "*" если есть
            clean_line = line.lstrip('-*• ').strip()
            doc.add_paragraph(clean_line, style='List Number')
        doc.add_paragraph()
    
    # Вывод и рекомендации
    doc.add_paragraph("Вывод и рекомендации", style='Heading 2')
    
    if data["consequences"].strip():
        doc.add_paragraph("Последствия дефектов", style='Heading 3')
        doc.add_paragraph(data["consequences"])
    
    if data["conclusion"].strip():
        doc.add_paragraph("Вывод", style='Heading 3')
        doc.add_paragraph(data["conclusion"])
    
    if data["recommendations_detailed"].strip():
        doc.add_paragraph("Рекомендации", style='Heading 3')
        doc.add_paragraph(data["recommendations_detailed"])
    doc.add_paragraph()
    
    # Подпись
    doc.add_paragraph("Подпись", style='Heading 2')
    
    signature_table = doc.add_table(rows=3, cols=2)
    signature_table.style = 'Table Grid'
    
    signature_fields = [
        ("Роль", data["role"]),
        ("ФИО", data["fullname"]),
        ("Дата", data["signature_date"]),
    ]
    
    for i, (label, value) in enumerate(signature_fields):
        signature_table.cell(i, 0).text = label
        signature_table.cell(i, 1).text = str(value)
        for paragraph in signature_table.cell(i, 0).paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Calibri Light'
                run.font.size = Pt(10)
        for paragraph in signature_table.cell(i, 1).paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Calibri Light'
                run.font.size = Pt(10)
    
    set_col_width(signature_table, 0, 1800)
    set_col_width(signature_table, 1, 5400)
    
    # Сохранение в буфер
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# === ГЕНЕРАЦИЯ HTML ОТЧЁТА ===
def escape_html(text):
    if pd.isna(text) or text is None:
        return ""
    return (str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&#x27;"))

def generate_html_report(data, module_data_list, defects_df):
    # Валидация метрик
    validation_errors = []
    total_tc = data["total_tc"]
    pass_tc = data["pass"]
    fail_tc = data["fail"]
    
    if pass_tc + fail_tc != total_tc:
        validation_errors.append("⚠️ Сумма статусов (PASS + FAIL) не равна общему количеству тест-кейсов")
    
    # Подготовка данных для диаграмм
    labels_results = ['PASS', 'FAIL']
    sizes_results = [pass_tc, fail_tc]
    colors_results = ['#4CAF50', '#F44336']
    
    labels_severity = []
    sizes_severity = []
    colors_severity_map = {'S1': '#F44336', 'S2': '#FF9800', 'S3': '#FFC107', 'S4': '#4CAF50'}
    
    if data['s1'] > 0:
        labels_severity.append('S1')
        sizes_severity.append(data['s1'])
    if data['s2'] > 0:
        labels_severity.append('S2')
        sizes_severity.append(data['s2'])
    
    # Создание диаграмм в base64
    def plot_to_base64(fig):
        buf = io.BytesIO()
        fig.savefig(buf, format='png', bbox_inches='tight', dpi=150)
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode('utf-8')
        buf.close()
        plt.close(fig)
        return img_base64
    
    # Диаграмма результатов
    fig1, ax1 = plt.subplots(figsize=(6, 4))
    ax1.pie(sizes_results, labels=labels_results, colors=colors_results, autopct='%1.1f%%', startangle=90)
    ax1.axis('equal')
    chart1_base64 = plot_to_base64(fig1)
    
    # Диаграмма серьёзности (если есть дефекты)
    chart2_base64 = None
    if sizes_severity:
        fig2, ax2 = plt.subplots(figsize=(6, 4))
        colors_sev = [colors_severity_map.get(lbl, '#9E9E9E') for lbl in labels_severity]
        ax2.pie(sizes_severity, labels=labels_severity, colors=colors_sev, autopct='%1.1f%%', startangle=90)
        ax2.axis('equal')
        chart2_base64 = plot_to_base64(fig2)
    
    # Формирование HTML
    html_content = f"""
    <!DOCTYPE html>
    <html lang="ru">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{escape_html(data['report_title'])}</title>
        <style>
            body {{
                font-family: 'Calibri', 'Segoe UI', Arial, sans-serif;
                line-height: 1.6;
                color: #333;
                max-width: 1200px;
                margin: 0 auto;
                padding: 20px;
                background-color: #f9f9f9;
            }}
            .container {{
                background: white;
                padding: 30px;
                border-radius: 8px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }}
            h1 {{
                text-align: center;
                color: #1a365d;
                font-size: 24px;
                margin-bottom: 30px;
                border-bottom: 2px solid #4472C4;
                padding-bottom: 10px;
            }}
            h2 {{
                color: #4472C4;
                border-left: 4px solid #4472C4;
                padding-left: 10px;
                margin-top: 25px;
            }}
            h3 {{
                color: #5b616b;
                margin-top: 20px;
            }}
            .info-table {{
                width: 100%;
                border-collapse: collapse;
                margin: 15px 0;
                font-size: 14px;
            }}
            .info-table th, .info-table td {{
                border: 1px solid #ddd;
                padding: 8px 12px;
                text-align: left;
                vertical-align: top;
            }}
            .info-table th {{
                background-color: #f2f2f2;
                width: 25%;
                font-weight: bold;
            }}
            .status-pass {{
                background-color: #e8f5e9;
                color: #2e7d32;
                font-weight: bold;
            }}
            .status-fail {{
                background-color: #ffebee;
                color: #c62828;
                font-weight: bold;
            }}
            .severity-s1 {{
                background-color: #ffebee;
                color: #c62828;
                font-weight: bold;
            }}
            .severity-s2 {{
                background-color: #fff3e0;
                color: #e65100;
                font-weight: bold;
            }}
            .chart-container {{
                text-align: center;
                margin: 25px 0;
            }}
            .chart-container img {{
                max-width: 100%;
                height: auto;
                border: 1px solid #ddd;
                border-radius: 4px;
            }}
            .chart-caption {{
                font-size: 13px;
                color: #666;
                margin-top: 5px;
            }}
            .validation-error {{
                background-color: #ffebee;
                color: #c62828;
                padding: 10px;
                border-radius: 4px;
                margin: 15px 0;
                border-left: 4px solid #c62828;
            }}
            .limitations li {{
                margin-bottom: 5px;
            }}
            @media print {{
                body {{
                    background-color: white;
                    padding: 0;
                }}
                .container {{
                    box-shadow: none;
                    padding: 15px;
                }}
                .no-print {{
                    display: none;
                }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>{escape_html(data['report_title'])}</h1>
            
            <!-- Валидация -->
            {"".join([f'<div class="validation-error">{err}</div>' for err in validation_errors]) if validation_errors else ""}
            
            <!-- Основная информация -->
            <h2>1. Основная информация</h2>
            <table class="info-table">
                <tr><th>Проект</th><td>{escape_html(data['project'])}</td></tr>
                <tr><th>Тип приложения</th><td>{escape_html(data['app_type'])}</td></tr>
                <tr><th>Версия</th><td>{escape_html(data['version'])}</td></tr>
                <tr><th>Период тестирования</th><td>{escape_html(data['test_period'])}</td></tr>
                <tr><th>Дата формирования отчёта</th><td>{escape_html(data['report_date'])}</td></tr>
                <tr><th>Инженер по тестированию</th><td>{escape_html(data['engineer'])}</td></tr>
                <tr><th>Статус релиза</th><td>{escape_html(data['release_status'])}</td></tr>
            </table>
            
            <!-- Краткое резюме -->
            <h2>2. Краткое резюме</h2>
            <table class="info-table">
                <tr><th>Количество дефектов (S1)</th><td>{data['s1']}</td></tr>
                <tr><th>Количество дефектов (S2)</th><td>{data['s2']}</td></tr>
                <tr><th>Всего тест-кейсов</th><td>{data['total_tc']}</td></tr>
                <tr><th>Пройдено успешно</th><td>{data['pass']}</td></tr>
            </table>
            
            <h3>Риски</h3>
            <p>{escape_html(data['risk'])}</p>
            
            <h3>Рекомендации</h3>
            <p>{escape_html(data['recommendation'])}</p>
            
            <!-- Диаграммы -->
            <h2>3. Диаграммы</h2>
            <div class="chart-container">
                <img src="data:image/png;base64,{chart1_base64}" alt="Распределение результатов">
                <div class="chart-caption">Рис. 1. Распределение результатов тестирования</div>
            </div>
            """
    
    if chart2_base64:
        html_content += f"""
            <div class="chart-container">
                <img src="data:image/png;base64,{chart2_base64}" alt="Распределение дефектов по серьёзности">
                <div class="chart-caption">Рис. 2. Распределение дефектов по серьёзности</div>
            </div>
            """
    
    # Контекст тестирования
    html_content += f"""
            <h2>4. Контекст тестирования</h2>
            <table class="info-table">
                <tr><th>Устройство / Браузер</th><td>{escape_html(data['device_browser'])}</td></tr>
                <tr><th>ОС / Платформа</th><td>{escape_html(data['os_platform'])}</td></tr>
                <tr><th>Сборка</th><td>{escape_html(data['build'])}</td></tr>
                <tr><th>URL окружения</th><td>{escape_html(data['env_url'])}</td></tr>
                <tr><th>Инструменты</th><td>{escape_html(data['tools'])}</td></tr>
            </table>
            
            <h3>Методология тестирования</h3>
            <p>{escape_html(data['methodology'])}</p>
            """
    
    # Результаты по модулям
    html_content += f"""
            <h2>5. Результаты тестирования по модулям</h2>
            """
    
    for module in module_data_list:
        html_content += f"""
            <h3>{escape_html(module['title'])}</h3>
            <table class="info-table">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Сценарий</th>
                        <th>Статус</th>
                        <th>Комментарий</th>
                    </tr>
                </thead>
                <tbody>
            """
        for _, row in module['df'].iterrows():
            status_class = "status-pass" if str(row['Статус']).strip().upper() == "PASS" else "status-fail"
            html_content += f"""
                    <tr>
                        <td>{escape_html(row['ID'])}</td>
                        <td>{escape_html(row['Сценарий'])}</td>
                        <td class="{status_class}">{escape_html(row['Статус'])}</td>
                        <td>{escape_html(row['Комментарий'])}</td>
                    </tr>
            """
        html_content += """
                </tbody>
            </table>
            """
    
    # Анализ дефектов
    if not defects_df.empty:
        html_content += f"""
            <h2>6. Анализ дефектов</h2>
            <table class="info-table">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Модуль</th>
                        <th>Заголовок</th>
                        <th>Серьёзность</th>
                        <th>Статус</th>
                    </tr>
                </thead>
                <tbody>
            """
        for _, row in defects_df.iterrows():
            sev_class = f"severity-{str(row['Серьёзность']).lower()}" if pd.notna(row['Серьёзность']) else ""
            html_content += f"""
                    <tr>
                        <td>{escape_html(row['ID'])}</td>
                        <td>{escape_html(row['Модуль'])}</td>
                        <td>{escape_html(row['Заголовок'])}</td>
                        <td class="{sev_class}">{escape_html(row['Серьёзность'])}</td>
                        <td>{escape_html(row['Статус'])}</td>
                    </tr>
            """
        html_content += """
                </tbody>
            </table>
            """
    
    # Ограничения тестирования
    if data["limitations"].strip():
        html_content += f"""
            <h2>7. Ограничения тестирования</h2>
            <div class="limitations">
                <ol>
        """
        lines = [line.strip() for line in data["limitations"].split('\n') if line.strip()]
        for line in lines:
            clean_line = line.lstrip('-*• ').strip()
            html_content += f"<li>{escape_html(clean_line)}</li>"
        html_content += """
                </ol>
            </div>
            """
    
    # Вывод и рекомендации
    html_content += f"""
            <h2>8. Вывод и рекомендации</h2>
            """
    
    if data["consequences"].strip():
        html_content += f"""
            <h3>Последствия дефектов</h3>
            <p>{escape_html(data['consequences'])}</p>
            """
    
    if data["conclusion"].strip():
        html_content += f"""
            <h3>Вывод</h3>
            <p>{escape_html(data['conclusion'])}</p>
            """
    
    if data["recommendations_detailed"].strip():
        html_content += f"""
            <h3>Рекомендации</h3>
            <p>{escape_html(data['recommendations_detailed'])}</p>
            """
    
    # Подпись
    html_content += f"""
            <h2>9. Подпись</h2>
            <table class="info-table">
                <tr><th>Роль</th><td>{escape_html(data['role'])}</td></tr>
                <tr><th>ФИО</th><td>{escape_html(data['fullname'])}</td></tr>
                <tr><th>Дата</th><td>{escape_html(data['signature_date'])}</td></tr>
            </table>
            
            <div class="no-print" style="margin-top: 30px; text-align: center; color: #666; font-size: 12px;">
                Отчёт сгенерирован автоматически через инструмент тестовой отчётности
            </div>
        </div>
    </body>
    </html>
    """
    
    return html_content

# === ГЕНЕРАЦИЯ XLSX ОТЧЁТА ===
def generate_xlsx_single_sheet(data, module_data_list, defects_df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Отчёт о тестировании"
    
    # Стили
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(name='Calibri Light', size=11, bold=True, color="FFFFFF")
    normal_font = Font(name='Calibri Light', size=11)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    align_left = Alignment(horizontal='left', vertical='top', wrap_text=True)
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Заголовок
    ws.merge_cells('A1:D1')
    ws['A1'] = data["report_title"]
    ws['A1'].font = Font(name='Calibri Light', size=16, bold=True)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 40
    
    row = 3
    
    # Основная информация
    ws.cell(row=row, column=1, value="ОСНОВНАЯ ИНФОРМАЦИЯ")
    ws.cell(row=row, column=1).font = Font(name='Calibri Light', size=14, bold=True, color="4472C4")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 1
    
    info_fields = [
        ("Проект", data["project"]),
        ("Тип приложения", data["app_type"]),
        ("Версия", data["version"]),
        ("Период тестирования", data["test_period"]),
        ("Дата формирования отчёта", data["report_date"]),
        ("Инженер по тестированию", data["engineer"]),
        ("Статус релиза", data["release_status"]),
    ]
    
    for label, value in info_fields:
        ws.cell(row=row, column=1, value=label).font = Font(name='Calibri Light', size=11, bold=True)
        ws.cell(row=row, column=2, value=value).font = normal_font
        ws.cell(row=row, column=1).border = border
        ws.cell(row=row, column=2).border = border
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        row += 1
    
    row += 1
    
    # Краткое резюме
    ws.cell(row=row, column=1, value="КРАТКОЕ РЕЗЮМЕ")
    ws.cell(row=row, column=1).font = Font(name='Calibri Light', size=14, bold=True, color="4472C4")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 1
    
    metrics = [
        ("Количество дефектов (S1)", data["s1"]),
        ("Количество дефектов (S2)", data["s2"]),
        ("Всего тест-кейсов", data["total_tc"]),
        ("Пройдено успешно", data["pass"]),
        ("Упало", data["fail"]),
    ]
    
    for label, value in metrics:
        ws.cell(row=row, column=1, value=label).font = Font(name='Calibri Light', size=11, bold=True)
        ws.cell(row=row, column=2, value=value).font = normal_font
        ws.cell(row=row, column=1).border = border
        ws.cell(row=row, column=2).border = border
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        row += 1
    
    row += 1
    
    # Риски и рекомендации
    if data["risk"].strip():
        ws.cell(row=row, column=1, value="РИСКИ").font = Font(name='Calibri Light', size=12, bold=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        row += 1
        ws.cell(row=row, column=1, value=data["risk"]).font = normal_font
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        ws.cell(row=row, column=1).alignment = align_left
        row += 2
    
    if data["recommendation"].strip():
        ws.cell(row=row, column=1, value="РЕКОМЕНДАЦИИ").font = Font(name='Calibri Light', size=12, bold=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        row += 1
        ws.cell(row=row, column=1, value=data["recommendation"]).font = normal_font
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        ws.cell(row=row, column=1).alignment = align_left
        row += 2
    
    # Результаты по модулям
    ws.cell(row=row, column=1, value="РЕЗУЛЬТАТЫ ПО МОДУЛЯМ")
    ws.cell(row=row, column=1).font = Font(name='Calibri Light', size=14, bold=True, color="4472C4")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 1
    
    for module in module_data_list:
        ws.cell(row=row, column=1, value=module["title"]).font = Font(name='Calibri Light', size=12, bold=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        row += 1
        
        # Заголовки таблицы модуля
        headers = ["ID", "Сценарий", "Статус", "Комментарий"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = align_center
        row += 1
        
        # Данные модуля
        for _, r in module["df"].iterrows():
            ws.cell(row=row, column=1, value=r["ID"]).font = normal_font
            ws.cell(row=row, column=2, value=r["Сценарий"]).font = normal_font
            status_cell = ws.cell(row=row, column=3, value=r["Статус"])
            status_cell.font = normal_font
            
            # Цвет статуса
            status_val = str(r["Статус"]).strip().upper() if pd.notna(r["Статус"]) else ""
            if status_val == "PASS":
                status_cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                status_cell.font = Font(name='Calibri Light', size=11, bold=True, color="006100")
            elif status_val == "FAIL":
                status_cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                status_cell.font = Font(name='Calibri Light', size=11, bold=True, color="9C0006")
            
            ws.cell(row=row, column=4, value=r["Комментарий"] if pd.notna(r["Комментарий"]) else "").font = normal_font
            
            for col in range(1, 5):
                ws.cell(row=row, column=col).border = border
                ws.cell(row=row, column=col).alignment = align_left
            
            row += 1
        
        row += 1
    
    # Анализ дефектов
    if not defects_df.empty:
        ws.cell(row=row, column=1, value="АНАЛИЗ ДЕФЕКТОВ")
        ws.cell(row=row, column=1).font = Font(name='Calibri Light', size=14, bold=True, color="4472C4")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        row += 1
        
        # Заголовки дефектов
        defect_headers = ["ID", "Модуль", "Заголовок", "Серьёзность"]
        for col, header in enumerate(defect_headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = align_center
        row += 1
        
        # Данные дефектов
        for _, r in defects_df.iterrows():
            ws.cell(row=row, column=1, value=r["ID"]).font = normal_font
            ws.cell(row=row, column=2, value=r["Модуль"]).font = normal_font
            ws.cell(row=row, column=3, value=r["Заголовок"]).font = normal_font
            
            sev_cell = ws.cell(row=row, column=4, value=r["Серьёзность"])
            sev_cell.font = normal_font
            
            # Цвет серьёзности
            sev_val = str(r["Серьёзность"]).strip().upper() if pd.notna(r["Серьёзность"]) else ""
            if sev_val == "S1":
                sev_cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                sev_cell.font = Font(name='Calibri Light', size=11, bold=True, color="9C0006")
            elif sev_val == "S2":
                sev_cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
                sev_cell.font = Font(name='Calibri Light', size=11, bold=True, color="9C5700")
            
            for col in range(1, 5):
                ws.cell(row=row, column=col).border = border
                ws.cell(row=row, column=col).alignment = align_left
            
            row += 1
        
        row += 1
    
    # Ограничения
    if data["limitations"].strip():
        ws.cell(row=row, column=1, value="ОГРАНИЧЕНИЯ ТЕСТИРОВАНИЯ")
        ws.cell(row=row, column=1).font = Font(name='Calibri Light', size=14, bold=True, color="4472C4")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        row += 1
        
        lines = [line.strip() for line in data["limitations"].split('\n') if line.strip()]
        for i, line in enumerate(lines, 1):
            clean_line = line.lstrip('-*• ').strip()
            ws.cell(row=row, column=1, value=f"{i}. {clean_line}").font = normal_font
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
            ws.cell(row=row, column=1).alignment = align_left
            row += 1
        
        row += 1
    
    # Вывод и рекомендации
    ws.cell(row=row, column=1, value="ВЫВОД И РЕКОМЕНДАЦИИ")
    ws.cell(row=row, column=1).font = Font(name='Calibri Light', size=14, bold=True, color="4472C4")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 1
    
    if data["consequences"].strip():
        ws.cell(row=row, column=1, value="Последствия дефектов").font = Font(name='Calibri Light', size=12, bold=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        row += 1
        ws.cell(row=row, column=1, value=data["consequences"]).font = normal_font
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        ws.cell(row=row, column=1).alignment = align_left
        row += 2
    
    if data["conclusion"].strip():
        ws.cell(row=row, column=1, value="Вывод").font = Font(name='Calibri Light', size=12, bold=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        row += 1
        ws.cell(row=row, column=1, value=data["conclusion"]).font = normal_font
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        ws.cell(row=row, column=1).alignment = align_left
        row += 2
    
    if data["recommendations_detailed"].strip():
        ws.cell(row=row, column=1, value="Рекомендации").font = Font(name='Calibri Light', size=12, bold=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        row += 1
        ws.cell(row=row, column=1, value=data["recommendations_detailed"]).font = normal_font
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        ws.cell(row=row, column=1).alignment = align_left
        row += 2
    
    # Подпись
    ws.cell(row=row, column=1, value="ПОДПИСЬ")
    ws.cell(row=row, column=1).font = Font(name='Calibri Light', size=14, bold=True, color="4472C4")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 1
    
    signature_fields = [
        ("Роль", data["role"]),
        ("ФИО", data["fullname"]),
        ("Дата", data["signature_date"]),
    ]
    
    for label, value in signature_fields:
        ws.cell(row=row, column=1, value=label).font = Font(name='Calibri Light', size=11, bold=True)
        ws.cell(row=row, column=2, value=value).font = normal_font
        ws.cell(row=row, column=1).border = border
        ws.cell(row=row, column=2).border = border
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        row += 1
    
    # Автоподбор высоты строк
    for row_idx in range(1, row):
        ws.row_dimensions[row_idx].height = 15
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# === ДАННЫЕ ПО УМОЛЧАНИЮ ===
default_modules = [
    {
        "title": "Авторизация и регистрация",
        "df": pd.DataFrame([
            {"ID": "AUTH-001", "Сценарий": "Вход по валидным данным", "Статус": "PASS", "Комментарий": ""},
            {"ID": "AUTH-002", "Сценарий": "Вход по невалидным данным", "Статус": "PASS", "Комментарий": ""},
            {"ID": "AUTH-003", "Сценарий": "Восстановление пароля", "Статус": "FAIL", "Комментарий": "Письмо не приходит на почту"},
        ])
    },
    {
        "title": "Поиск товаров",
        "df": pd.DataFrame([
            {"ID": "SEARCH-001", "Сценарий": "Поиск по точному названию", "Статус": "PASS", "Комментарий": ""},
            {"ID": "SEARCH-002", "Сценарий": "Поиск с опечаткой", "Статус": "FAIL", "Комментарий": "Fuzzy search не работает"},
        ])
    }
]

default_defects = pd.DataFrame([
    {"ID": "BUG-SEC-001", "Модуль": "Авторизация", "Заголовок": "SQL-инъекция в поле логина", "Серьёзность": "S1", "Статус": "Открыт"},
    {"ID": "BUG-SEC-002", "Модуль": "Авторизация", "Заголовок": "Пароли хранятся в открытом виде", "Серьёзность": "S1", "Статус": "Открыт"},
    {"ID": "BUG-SEARCH-001", "Модуль": "Поиск", "Заголовок": "Нет поддержки опечаток", "Серьёзность": "S2", "Статус": "Открыт"},
])

# === ИНТЕРФЕЙС STREAMLIT ===
st.set_page_config(page_title="Генератор отчёта", layout="wide")

# === ЗАГРУЗКА ЧЕРНОВИКА (ВНЕ ФОРМЫ!) ===
st.title("📄 Генератор профессиональных тестовых отчётов")
st.markdown("""
Создавайте отчёты о тестировании в трёх форматах (DOCX, HTML, XLSX) с соблюдением корпоративных стандартов.
Сохраняйте черновики и возвращайтесь к ним позже!
""")

# Кнопка загрузки черновика ВНЕ формы
uploaded_file = st.file_uploader(
    "⬆️ Загрузить черновик (.json)",
    type=["json"],
    label_visibility="visible",
    key="draft_uploader_outside_form"
)
if uploaded_file is not None:
    content = uploaded_file.read().decode("utf-8")
    restored_data, restored_modules, restored_defects, saved_at = load_draft(content)
    if restored_data is not None:
        st.session_state.draft_data = restored_data
        st.session_state.draft_modules = restored_modules
        st.session_state.draft_defects = restored_defects
        st.session_state.draft_saved_at = saved_at
        st.success(f"✅ Черновик загружен! Сохранён: {saved_at}")
        st.rerun()

# === ПОДГОТОВКА ДАННЫХ ДЛЯ ФОРМЫ ===
if "draft_data" in st.session_state and st.session_state.draft_data is not None:
    draft_data = st.session_state.draft_data
    draft_modules = st.session_state.draft_modules
    draft_defects = st.session_state.draft_defects
    draft_saved_at = st.session_state.draft_saved_at
    
    # Подставляем значения из черновика
    report_title_val = draft_data.get("report_title", "Отчёт о тестировании мобильного приложения Лемана ПРО")
    project_val = draft_data.get("project", "Лемана ПРО")
    app_type_val = draft_data.get("app_type", "Мобильное")
    version_val = draft_data.get("version", "241006.001")
    test_period_val = draft_data.get("test_period", "29–30 ноября 2025 г.")
    report_date_val = draft_data.get("report_date", "30 ноября 2025 г.")
    engineer_val = draft_data.get("engineer", "Черкасов Игорь")
    
    release_status_val = draft_data.get("release_status", "НЕ РЕКОМЕНДОВАН К ВЫПУСКУ")
    s1_val = draft_data.get("s1", 2)
    s2_val = draft_data.get("s2", 1)
    total_tc_val = draft_data.get("total_tc", 72)
    pass_tc_val = draft_data.get("pass", 69)
    fail_tc_val = draft_data.get("fail", 3)
    risk_val = draft_data.get("risk", "Уязвимости безопасности позволяют нарушителю получить доступ к данным пользователей и вызвать отказ в обслуживании.")
    recommendation_val = draft_data.get("recommendation", "Релиз возможен только после устранения всех S1/S2 дефектов и повторного тестирования.")
    
    device_browser_val = draft_data.get("device_browser", "Xiaomi 12")
    os_platform_val = draft_data.get("os_platform", "Android 15")
    build_val = draft_data.get("build", "lemanna-pro_241006.001.apk")
    env_url_val = draft_data.get("env_url", "https://test.lemanna.pro")
    tools_val = draft_data.get("tools", "Postman (API), Burp Suite (безопасность), Jira (баг-трекинг)")
    methodology_val = draft_data.get("methodology", "Ручное функциональное тестирование + проверка безопасности")
    
    consequences_val = draft_data.get("consequences", "- S1 дефекты позволяют злоумышленнику получить данные других пользователей или вывести приложение из строя.\n- S2 дефект снижает юзабилити: пользователи не найдут товар при опечатке.")
    limitations_val = draft_data.get("limitations", "1. Не тестировалась оплата через Apple Pay (устройство Android).\n2. Не проверена синхронизация с 1С (нет доступа к интеграционному стенду).\n3. Не проведено нагрузочное тестирование (ограничение по времени).")
    conclusion_val = draft_data.get("conclusion", "Сборка 241006.001 содержит критические уязвимости безопасности, делающие её непригодной для выпуска в production. Наличие S1 дефектов нарушает базовые принципы защиты данных пользователей.")
    recommendations_detailed_val = draft_data.get("recommendations_detailed", "Немедленно исправить уязвимости BUG-SEC-001 и BUG-SEC-002.\nРеализовать fuzzy search для повышения юзабилити (BUG-SEARCH-001).\nПровести повторное тестирование после фиксов с фокусом на:\n- Повторную проверку полей ввода на инъекции\n- Тестирование сценариев поиска с опечатками\n- Настроить автоматизированную проверку безопасности (например, OWASP ZAP) в CI/CD.")
    
    role_val = draft_data.get("role", "QA-инженер")
    fullname_val = draft_data.get("fullname", "Черкасов Игорь")
    signature_date_val = draft_data.get("signature_date", "30.11.2025")
    
    # Очищаем session_state после применения
    del st.session_state.draft_data
    del st.session_state.draft_modules
    del st.session_state.draft_defects
    del st.session_state.draft_saved_at
else:
    # Значения по умолчанию
    report_title_val = "Отчёт о тестировании мобильного приложения Лемана ПРО"
    project_val = "Лемана ПРО"
    app_type_val = "Мобильное"
    version_val = "241006.001"
    test_period_val = "29–30 ноября 2025 г."
    report_date_val = "30 ноября 2025 г."
    engineer_val = "Черкасов Игорь"
    
    release_status_val = "НЕ РЕКОМЕНДОВАН К ВЫПУСКУ"
    s1_val = 2
    s2_val = 1
    total_tc_val = 72
    pass_tc_val = 69
    fail_tc_val = 3
    risk_val = "Уязвимости безопасности позволяют нарушителю получить доступ к данным пользователей и вызвать отказ в обслуживании."
    recommendation_val = "Релиз возможен только после устранения всех S1/S2 дефектов и повторного тестирования."
    
    device_browser_val = "Xiaomi 12"
    os_platform_val = "Android 15"
    build_val = "lemanna-pro_241006.001.apk"
    env_url_val = "https://test.lemanna.pro"
    tools_val = "Postman (API), Burp Suite (безопасность), Jira (баг-трекинг)"
    methodology_val = "Ручное функциональное тестирование + проверка безопасности"
    
    consequences_val = "- S1 дефекты позволяют злоумышленнику получить данные других пользователей или вывести приложение из строя.\n- S2 дефект снижает юзабилити: пользователи не найдут товар при опечатке."
    limitations_val = "1. Не тестировалась оплата через Apple Pay (устройство Android).\n2. Не проверена синхронизация с 1С (нет доступа к интеграционному стенду).\n3. Не проведено нагрузочное тестирование (ограничение по времени)."
    conclusion_val = "Сборка 241006.001 содержит критические уязвимости безопасности, делающие её непригодной для выпуска в production. Наличие S1 дефектов нарушает базовые принципы защиты данных пользователей."
    recommendations_detailed_val = "Немедленно исправить уязвимости BUG-SEC-001 и BUG-SEC-002.\nРеализовать fuzzy search для повышения юзабилити (BUG-SEARCH-001).\nПровести повторное тестирование после фиксов с фокусом на:\n- Повторную проверку полей ввода на инъекции\n- Тестирование сценариев поиска с опечатками\n- Настроить автоматизированную проверку безопасности (например, OWASP ZAP) в CI/CD."
    
    role_val = "QA-инженер"
    fullname_val = "Черкасов Игорь"
    signature_date_val = "30.11.2025"

# === ФОРМА ВВОДА ДАННЫХ ===
with st.form("main_form"):
    # Заголовок и основная информация
    st.subheader("1. Заголовок и основная информация")
    
    report_title = st.text_input("Название отчёта", report_title_val)
    col1, col2, col3 = st.columns(3)
    with col1:
        project = st.text_input("Проект", project_val)
        app_type = st.selectbox("Тип приложения", ["Мобильное", "Веб"], 
                               index=0 if app_type_val == "Мобильное" else 1)
    with col2:
        version = st.text_input("Версия", version_val)
        test_period = st.text_input("Период тестирования", test_period_val)
    with col3:
        report_date = st.text_input("Дата формирования отчёта", report_date_val)
        engineer = st.text_input("Инженер по тестированию", engineer_val)
    
    release_status = st.selectbox(
        "Статус релиза",
        ["НЕ РЕКОМЕНДОВАН К ВЫПУСКУ", "РЕКОМЕНДОВАН К ВЫПУСКУ С ЗАМЕЧАНИЯМИ", "РЕКОМЕНДОВАН К ВЫПУСКУ"],
        index=["НЕ РЕКОМЕНДОВАН К ВЫПУСКУ", "РЕКОМЕНДОВАН К ВЫПУСКУ С ЗАМЕЧАНИЯМИ", "РЕКОМЕНДОВАН К ВЫПУСКУ"].index(release_status_val)
    )
    
    # Метрики
    st.subheader("2. Метрики тестирования")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        s1 = st.number_input("Дефекты S1", min_value=0, value=s1_val)
    with col2:
        s2 = st.number_input("Дефекты S2", min_value=0, value=s2_val)
    with col3:
        total_tc = st.number_input("Всего тест-кейсов", min_value=0, value=total_tc_val)
    with col4:
        pass_tc = st.number_input("Пройдено успешно (PASS)", min_value=0, value=pass_tc_val)
    fail_tc = st.number_input("Упало (FAIL)", min_value=0, value=fail_tc_val)
    
    # Риски и рекомендации
    risk = st.text_area("Риски", risk_val, height=80)
    recommendation = st.text_area("Краткие рекомендации", recommendation_val, height=60)
    
    # Контекст тестирования
    st.subheader("3. Контекст тестирования")
    col1, col2 = st.columns(2)
    with col1:
        device_browser = st.text_input("Устройство / Браузер", device_browser_val)
        os_platform = st.text_input("ОС / Платформа", os_platform_val)
        build = st.text_input("Сборка", build_val)
        env_url = st.text_input("URL окружения", env_url_val)
    with col2:
        tools = st.text_area("Инструменты", tools_val, height=100)
        methodology = st.text_area("Методология тестирования", methodology_val, height=100)
    
    # Результаты по модулям
    st.subheader("4. Результаты по модулям")
    module_data_list = []
    
    for i, mod in enumerate(default_modules):
        with st.expander(f"Модуль {i+1}: {mod['title']}", expanded=False):
            title = st.text_input(f"Название модуля {i+1}", mod["title"], key=f"title_{i}")
            df_edited = st.data_editor(
                mod["df"],
                column_config={
                    "ID": st.column_config.TextColumn("ID", width="small"),
                    "Сценарий": st.column_config.TextColumn("Сценарий", width="medium"),
                    "Статус": st.column_config.SelectboxColumn("Статус", options=["PASS", "FAIL", "SKIP"], width="small"),
                    "Комментарий": st.column_config.TextColumn("Комментарий", width="large"),
                },
                hide_index=True,
                key=f"module_{i}",
                use_container_width=True,
                num_rows="dynamic"
            )
            module_data_list.append({"title": title, "df": df_edited})
    
    # Анализ дефектов
    st.subheader("5. Анализ дефектов")
    defects = st.data_editor(
        default_defects if "draft_defects" not in st.session_state else draft_defects,
        column_config={
            "ID": st.column_config.TextColumn("ID", width="small"),
            "Модуль": st.column_config.TextColumn("Модуль", width="small"),
            "Заголовок": st.column_config.TextColumn("Заголовок", width="medium"),
            "Серьёзность": st.column_config.SelectboxColumn("Серьёзность", options=["S1", "S2", "S3", "S4"], width="small"),
            "Статус": st.column_config.TextColumn("Статус", width="small"),
        },
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic"
    )
    
    # Ограничения и вывод
    st.subheader("6. Ограничения и вывод")
    limitations = st.text_area(
        "Ограничения тестирования (нумерованный список)",
        limitations_val,
        height=120,
        help="Введите по одному ограничению на строку. Маркеры будут автоматически преобразованы в нумерацию."
    )
    consequences = st.text_area("Последствия дефектов", consequences_val, height=100)
    conclusion = st.text_area("Вывод", conclusion_val, height=100)
    recommendations_detailed = st.text_area("Детальные рекомендации", recommendations_detailed_val, height=100)
    
    # Подпись
    st.subheader("7. Подпись")
    col1, col2, col3 = st.columns(3)
    with col1:
        role = st.text_input("Роль", role_val)
    with col2:
        fullname = st.text_input("ФИО", fullname_val)
    with col3:
        signature_date = st.text_input("Дата", signature_date_val)
    
    # === КНОПКИ ФОРМЫ ===
    col1, col2 = st.columns(2)
    with col1:
        save_draft_clicked = st.form_submit_button("💾 Сохранить черновик", type="secondary")
    with col2:
        submitted = st.form_submit_button("📥 Создать отчёт", type="primary")

# === ОБРАБОТКА КНОПОК (ВНЕ ФОРМЫ!) ===
if save_draft_clicked:
    # Собираем данные из текущей формы
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
        "env_url": env_url.strip(),
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
    
    draft_json = save_draft(data, module_data_list, defects)
    st.session_state.draft_to_download = draft_json
    st.session_state.draft_filename = f"черновик_{datetime.now().strftime('%d%m%Y_%H%M%S')}.json"
    st.success("✅ Черновик подготовлен к скачиванию!")
    st.rerun()

# Кнопка скачивания черновика (вне формы)
if "draft_to_download" in st.session_state and st.session_state.draft_to_download is not None:
    st.download_button(
        "⬇️ Скачать черновик",
        st.session_state.draft_to_download,
        st.session_state.draft_filename,
        "application/json",
        use_container_width=True,
        type="primary"
    )
    if st.button("Закрыть", key="close_draft"):
        del st.session_state.draft_to_download
        del st.session_state.draft_filename
        st.rerun()

# === ГЕНЕРАЦИЯ ОТЧЁТА ===
if submitted:
    # Сбор данных
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
        "env_url": env_url.strip(),
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
    
    # Валидация
    validation_errors = []
    if pass_tc + fail_tc != total_tc:
        validation_errors.append("⚠️ Сумма статусов (PASS + FAIL) не равна общему количеству тест-кейсов")
    
    if validation_errors:
        for err in validation_errors:
            st.error(err)
        st.stop()
    
    # Генерация отчётов
    with st.spinner("Генерация отчётов..."):
        # DOCX
        docx_buffer = generate_docx(data, module_data_list, defects)
        
        # HTML
        html_content = generate_html_report(data, module_data_list, defects)
        html_buffer = io.BytesIO(html_content.encode('utf-8'))
        
        # XLSX
        xlsx_buffer = generate_xlsx_single_sheet(data, module_data_list, defects)
    
    # Отображение кнопок загрузки
    st.success("✅ Отчёты успешно сгенерированы!")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.download_button(
            "📄 Скачать DOCX",
            docx_buffer,
            "отчёт_тестирование.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    
    with col2:
        st.download_button(
            "🌐 Скачать HTML",
            html_buffer,
            "отчёт_тестирование.html",
            "text/html",
            use_container_width=True
        )
    
    with col3:
        st.download_button(
            "📊 Скачать XLSX",
            xlsx_buffer,
            "отчёт_тестирование.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    # Предпросмотр HTML
    with st.expander("🔍 Предпросмотр HTML-отчёта"):
        st.components.v1.html(html_content, height=600, scrolling=True)