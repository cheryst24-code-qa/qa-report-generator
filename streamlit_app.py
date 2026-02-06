# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import io
import base64
import traceback
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def set_col_width(col, width_twips):
    """Устанавливает точную ширину колонки в таблице Word"""
    for cell in col.cells:
        tc = cell._element.tcPr
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), str(int(width_twips)))
        tcW.set(qn('w:type'), 'dxa')
        tc.append(tcW)

def add_table_from_df(doc, df):
    """Создаёт таблицу с фиксированной шириной 15%/85% и сеткой"""
    if len(df.columns) == 0:
        doc.add_paragraph("Нет данных для отображения")
        doc.add_paragraph().paragraph_format.space_after = Pt(6)
        return
    
    if df.empty:
        table = doc.add_table(rows=2, cols=len(df.columns))
        for i, col in enumerate(df.columns):
            table.cell(0, i).text = str(col)
            table.cell(1, i).text = ""
    else:
        table = doc.add_table(rows=df.shape[0] + 1, cols=len(df.columns))
        table.style = 'Table Grid'
    
    total_width = Inches(6.5)
    num_cols = len(df.columns)
    
    if num_cols > 0:
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

def generate_docx(data, module_data_list, defects_df):
    """Генерирует отчёт в точном соответствии с образцом из PDF"""
    doc = Document()
    
    # Настройка стиля документа
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    
    # === ЗАГОЛОВОК ОТЧЁТА (центрированный, крупный) ===
    title = doc.add_heading(data["report_title"], 0)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    title_font = title.runs[0].font
    title_font.size = Pt(16)
    title_font.bold = True
    
    # === ТАБЛИЦА С ОСНОВНОЙ ИНФОРМАЦИЕЙ (6 строк × 2 колонки) ===
    total_width_twips = Inches(6.5).twips
    first_col_width_twips = int(total_width_twips * 0.25)
    second_col_width_twips = int(total_width_twips * 0.75)
    
    info_table = doc.add_table(rows=6, cols=2)
    info_table.style = 'Table Grid'
    set_col_width(info_table.columns[0], first_col_width_twips)
    set_col_width(info_table.columns[1], second_col_width_twips)
    
    fields = [
        ('Проект:', data["project"]),
        ('Тип приложения:', data["app_type"]),
        ('Версия приложения:', data["version"]),
        ('Период тестирования:', data["test_period"]),
        ('Дата формирования отчёта:', data["report_date"]),
        ('QA-инженер:', data["engineer"])
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
    
    # === РАЗДЕЛ 1: КРАТКОЕ РЕЗЮМЕ ===
    doc.add_heading('1. КРАТКОЕ РЕЗЮМЕ', 1)
    
    summary_table = doc.add_table(rows=8, cols=2)
    summary_table.style = 'Table Grid'
    set_col_width(summary_table.columns[0], first_col_width_twips)
    set_col_width(summary_table.columns[1], second_col_width_twips)
    
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
    # Диаграмма 1: Распределение результатов
    plt.figure(figsize=(5, 4))
    plt.pie(
        [data['pass'], data['fail']],
        labels=['PASS', 'FAIL'],
        autopct='%1.1f%%',
        colors=['#4CAF50', '#F44336'],
        startangle=90
    )
    plt.title('Рис. 1. Распределение результатов тест-кейсов')
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    plt.close()
    
    doc.add_picture(buf, width=Inches(5))
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    
    # Диаграмма 2: Дефекты по серьёзности
    plt.figure(figsize=(5, 4))
    bars = plt.bar(
        ['Critical (S1)', 'Major (S2)'],
        [data['s1'], data['s2']],
        color=['#F44336', '#FF9800'],
        width=0.5
    )
    plt.title('Рис. 2. Дефекты по уровню серьёзности')
    plt.ylabel('Количество')
    plt.ylim(0, max(data['s1'], data['s2'], 1) * 1.3)
    
    for bar in bars:
        h = bar.get_height()
        if h > 0:
            plt.text(
                bar.get_x() + bar.get_width()/2,
                h + 0.05,
                str(int(h)),
                ha='center',
                va='bottom'
            )
    
    plt.grid(axis='y', alpha=0.3, linestyle='--')
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    plt.close()
    
    doc.add_picture(buf, width=Inches(5))
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    
    # === РАЗДЕЛ 2: КОНТЕКСТ ТЕСТИРОВАНИЯ ===
    doc.add_heading('2. КОНТЕКСТ ТЕСТИРОВАНИЯ', 1)
    
    context_table = doc.add_table(rows=6, cols=2)
    context_table.style = 'Table Grid'
    set_col_width(context_table.columns[0], first_col_width_twips)
    set_col_width(context_table.columns[1], second_col_width_twips)
    
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
    
    # === РАЗДЕЛ 3: РЕЗУЛЬТАТЫ ПО МОДУЛЯМ ===
    doc.add_heading('3. РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ ПО МОДУЛЯМ', 1)
    
    for idx, module_info in enumerate(module_data_list):
        title = module_info['title']
        df = module_info['df']
        
        doc.add_heading(f'3.{idx+1}. {title}', 2)
        add_table_from_df(doc, df)
    
    # === РАЗДЕЛ 4: АНАЛИЗ ДЕФЕКТОВ ===
    doc.add_heading('4. АНАЛИЗ ДЕФЕКТОВ', 1)
    add_table_from_df(doc, defects_df)
    
    # Последствия: просто текст после заголовка без лишних отступов
    p = doc.add_paragraph()
    p.add_run('Последствия: ').bold = True
    p.add_run(data['consequences'])
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # === РАЗДЕЛ 5: ОГРАНИЧЕНИЯ ТЕСТИРОВАНИЯ (нумерованный список!) ===
    doc.add_heading('5. ОГРАНИЧЕНИЯ ТЕСТИРОВАНИЯ', 1)
    
    # ВАЖНО: в образце используется нумерованный список (1., 2., 3.), а не маркированный
    for line in data['limitations'].split('\n'):
        if line.strip():
            # Убираем автоматическую нумерацию, если пользователь уже ввёл её
            clean_line = line.strip()
            if not clean_line[0].isdigit():
                # Если нет нумерации — добавляем вручную
                p = doc.add_paragraph(clean_line, style='List Number')
            else:
                p = doc.add_paragraph(clean_line)
            p.paragraph_format.space_after = Pt(2)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # === РАЗДЕЛ 6: ВЫВОД И РЕКОМЕНДАЦИИ ===
    doc.add_heading('6. ВЫВОД И РЕКОМЕНДАЦИИ', 1)
    
    # Вывод: текст сразу после слова "Вывод:" без переноса строки
    p = doc.add_paragraph()
    p.add_run('Вывод: ').bold = True
    p.add_run(data['conclusion'])
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # Рекомендации: маркированный список
    p = doc.add_paragraph()
    p.add_run('Рекомендации:').bold = True
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    
    for line in data['recommendations_detailed'].split('\n'):
        if line.strip():
            p = doc.add_paragraph(line.strip(), style='List Bullet')
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.space_after = Pt(2)
    
    # === РАЗДЕЛ 7: ПОДПИСЬ (чистая таблица 3×2 без артефактов) ===
    doc.add_heading('7. ПОДПИСЬ', 1)
    
    signature_table = doc.add_table(rows=3, cols=2)
    signature_table.style = 'Table Grid'
    set_col_width(signature_table.columns[0], first_col_width_twips)
    set_col_width(signature_table.columns[1], second_col_width_twips)
    
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
    
    # Сохранение документа
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generate_chart_base64(pass_count, fail_count, s1_count, s2_count):
    """Генерирует диаграммы в base64"""
    plt.figure(figsize=(6, 4.5))
    plt.pie(
        [pass_count, fail_count],
        labels=['PASS', 'FAIL'],
        autopct='%1.1f%%',
        colors=['#4CAF50', '#F44336'],
        startangle=90,
        textprops={'fontsize': 11}
    )
    plt.title('Рис. 1. Распределение результатов тест-кейсов', fontsize=10, pad=15)
    
    buf1 = io.BytesIO()
    plt.savefig(buf1, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    
    plt.figure(figsize=(6, 4.5))
    bars = plt.bar(
        ['Critical (S1)', 'Major (S2)'],
        [s1_count, s2_count],
        color=['#F44336', '#FF9800'],
        width=0.5
    )
    plt.title('Рис. 2. Дефекты по уровню серьёзности', fontsize=10, pad=15)
    plt.ylabel('Количество', fontsize=11)
    plt.ylim(0, max(s1_count, s2_count, 1) * 1.3)
    
    for bar in bars:
        h = bar.get_height()
        if h > 0:
            plt.text(
                bar.get_x() + bar.get_width()/2,
                h + 0.05,
                str(int(h)),
                ha='center',
                va='bottom',
                fontsize=11,
                fontweight='bold'
            )
    
    plt.grid(axis='y', alpha=0.3, linestyle='--')
    
    buf2 = io.BytesIO()
    plt.savefig(buf2, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    
    chart1_base64 = base64.b64encode(buf1.getvalue()).decode('utf-8')
    chart2_base64 = base64.b64encode(buf2.getvalue()).decode('utf-8')
    
    return chart1_base64, chart2_base64

def escape_html(text):
    """Экранирование HTML для безопасности"""
    if not isinstance(text, str):
        return str(text)
    return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;')
                .replace("'", '&#39;'))

def generate_html_report(data, module_data_list, defects_df):
    """Генерирует HTML-отчёт в соответствии с образцом"""
    chart1, chart2 = generate_chart_base64(data['pass'], data['fail'], data['s1'], data['s2'])
    
    total = data['total_tc']
    pass_pct = data['pass'] / total * 100 if total > 0 else 0
    fail_pct = 100 - pass_pct
    
    html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{escape_html(data['report_title'])}</title>
    <style>
        body {{
            font-family: Calibri, 'Segoe UI', sans-serif;
            font-size: 11pt;
            line-height: 1.5;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            color: #000;
        }}
        h1 {{
            text-align: center;
            font-size: 16pt;
            font-weight: bold;
            margin-bottom: 25px;
            margin-top: 0;
        }}
        h2 {{
            font-size: 14pt;
            margin-top: 25px;
            margin-bottom: 12px;
            padding-bottom: 4px;
            border-bottom: 2px solid #000;
        }}
        h3 {{
            font-size: 13pt;
            margin-top: 20px;
            margin-bottom: 10px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 12px 0 18px 0;
            page-break-inside: avoid;
        }}
        th, td {{
            border: 1px solid #000;
            padding: 8px 10px;
            text-align: left;
            vertical-align: top;
        }}
        th {{
            background-color: #f5f5f5;
            font-weight: bold;
        }}
        .info-table td:first-child,
        .summary-table td:first-child,
        .context-table td:first-child,
        .signature-table td:first-child {{
            width: 25%;
            font-weight: bold;
            background-color: #f9f9f9;
        }}
        .status-pass {{ color: #2e7d32; font-weight: bold; }}
        .status-fail {{ color: #d32f2f; font-weight: bold; }}
        .risk {{ color: #d32f2f; font-weight: bold; }}
        .chart-container {{
            text-align: center;
            margin: 25px 0;
            page-break-inside: avoid;
        }}
        .chart-title {{
            font-weight: bold;
            margin-top: 8px;
            font-size: 11pt;
        }}
        ol {{
            padding-left: 20px;
            margin: 10px 0;
        }}
        ul {{
            padding-left: 20px;
            margin: 10px 0;
        }}
        li {{
            margin-bottom: 5px;
        }}
        @media print {{
            body {{
                padding: 15px;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            .chart-container img {{
                max-width: 100% !important;
                height: auto !important;
            }}
            table {{
                page-break-inside: avoid;
            }}
            h2, h3 {{
                page-break-after: avoid;
            }}
        }}
        @page {{
            size: A4;
            margin: 15mm;
        }}
    </style>
</head>
<body>
    <h1>{escape_html(data['report_title'])}</h1>
    
    <table class="info-table">
        <tr><td>Проект:</td><td>{escape_html(data['project'])}</td></tr>
        <tr><td>Тип приложения:</td><td>{escape_html(data['app_type'])}</td></tr>
        <tr><td>Версия приложения:</td><td>{escape_html(data['version'])}</td></tr>
        <tr><td>Период тестирования:</td><td>{escape_html(data['test_period'])}</td></tr>
        <tr><td>Дата формирования отчёта:</td><td>{escape_html(data['report_date'])}</td></tr>
        <tr><td>QA-инженер:</td><td>{escape_html(data['engineer'])}</td></tr>
    </table>
    
    <h2>1. КРАТКОЕ РЕЗЮМЕ</h2>
    <table class="summary-table">
        <tr><td>Статус релиза:</td><td>{escape_html(data['release_status'])}</td></tr>
        <tr><td>Критические дефекты (S1):</td><td>{data['s1']}</td></tr>
        <tr><td>Мажорные дефекты (S2):</td><td>{data['s2']}</td></tr>
        <tr><td>Всего тест-кейсов:</td><td>{data['total_tc']}</td></tr>
        <tr><td>Успешно (Pass):</td><td class="status-pass">{data['pass']} ({pass_pct:.1f}%)</td></tr>
        <tr><td>Упали (Fail):</td><td class="status-fail">{data['fail']} ({fail_pct:.1f}%)</td></tr>
        <tr><td>Основной риск:</td><td class="risk">{escape_html(data['risk'])}</td></tr>
        <tr><td>Рекомендация:</td><td>{escape_html(data['recommendation'])}</td></tr>
    </table>
    
    <div class="chart-container">
        <img src="data:image/png;base64,{chart1}" alt="Распределение результатов тест-кейсов" style="max-width: 100%; height: auto; display: block; margin: 0 auto;">
        <div class="chart-title">Рис. 1. Распределение результатов тест-кейсов</div>
    </div>
    
    <div class="chart-container">
        <img src="data:image/png;base64,{chart2}" alt="Дефекты по уровню серьёзности" style="max-width: 100%; height: auto; display: block; margin: 0 auto;">
        <div class="chart-title">Рис. 2. Дефекты по уровню серьёзности</div>
    </div>
    
    <h2>2. КОНТЕКСТ ТЕСТИРОВАНИЯ</h2>
    <table class="context-table">
        <tr><td>Устройство / Браузер:</td><td>{escape_html(data['device_browser'])}</td></tr>
        <tr><td>ОС / Платформа:</td><td>{escape_html(data['os_platform'])}</td></tr>
        <tr><td>Сборка / Версия:</td><td>{escape_html(data['build'])}</td></tr>
        <tr><td>Стенд:</td><td>Тестовое окружение (адрес: {escape_html(data['env_url'])})</td></tr>
        <tr><td>Инструменты:</td><td>{escape_html(data['tools'])}</td></tr>
        <tr><td>Методология:</td><td>{escape_html(data['methodology'])}</td></tr>
    </table>
    """
    
    # Модули
    # === РАЗДЕЛ 3: РЕЗУЛЬТАТЫ ПО МОДУЛЯМ ===
    html += "<h2>3. РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ ПО МОДУЛЯМ</h2>"
    for idx, module_info in enumerate(module_data_list):
        html += f"<h3>3.{idx+1}. {escape_html(module_info['title'])}</h3>"
        # Исправленные ширины колонок: Сценарий увеличен до 45%, Комментарий уменьшен до 28%
        html += '<table><tr><th style="width: 15%;">ID</th><th style="width: 45%;">Сценарий</th><th style="width: 12%;">Статус</th><th style="width: 28%;">Комментарий</th></tr>'
        df = module_info['df']
        if not df.empty:
            for _, row in df.iterrows():
                # Определяем CSS-класс для цветового выделения статуса
                status_class = "status-pass" if str(row[2]).upper() == "PASS" else "status-fail" if str(row[2]).upper() == "FAIL" else ""
                html += f"<tr><td>{escape_html(row[0])}</td><td>{escape_html(row[1])}</td><td class='{status_class}'>{escape_html(row[2])}</td><td>{escape_html(row[3])}</td></tr>"
        else:
            html += "<tr><td colspan='4' style='text-align:center'>Нет данных</td></tr>"
        html += "</table>"

    # Дефекты
    html += "<h2>4. АНАЛИЗ ДЕФЕКТОВ</h2>"
    html += '<table><tr><th style="width: 15%;">ID</th><th style="width: 15%;">Модуль</th><th>Заголовок</th><th style="width: 20%;">Серьёзность</th><th style="width: 15%;">Статус</th></tr>'
    if not defects_df.empty:
        for _, row in defects_df.iterrows():
            html += f"<tr><td>{escape_html(row[0])}</td><td>{escape_html(row[1])}</td><td>{escape_html(row[2])}</td><td>{escape_html(row[3])}</td><td>{escape_html(row[4])}</td></tr>"
    else:
        html += "<tr><td colspan='5' style='text-align:center'>Нет данных</td></tr>"
    html += "</table>"
    
    # Последствия
    html += f"<p><strong>Последствия:</strong> {escape_html(data['consequences']).replace(chr(10), '<br>').replace('\n', '<br>')}</p>"
    
    # Ограничения (нумерованный список!)
    html += "<h2>5. ОГРАНИЧЕНИЯ ТЕСТИРОВАНИЯ</h2><ol>"
    for line in data['limitations'].split('\n'):
        if line.strip():
            html += f"<li>{escape_html(line.strip())}</li>"
    html += "</ol>"
    
    # Вывод и рекомендации
    html += f"""
    <h2>6. ВЫВОД И РЕКОМЕНДАЦИИ</h2>
    <p><strong>Вывод:</strong> {escape_html(data['conclusion'])}</p>
    <p><strong>Рекомендации:</strong></p>
    <ul>
    """
    for line in data['recommendations_detailed'].split('\n'):
        if line.strip():
            html += f"<li>{escape_html(line.strip())}</li>"
    html += "</ul>"
    
    # Подпись
    html += f"""
    <h2>7. ПОДПИСЬ</h2>
    <table class="signature-table">
        <tr><td>Роль:</td><td>{escape_html(data['role'])}</td></tr>
        <tr><td>ФИО:</td><td>{escape_html(data['fullname'])}</td></tr>
        <tr><td>Дата:</td><td>{escape_html(data['signature_date'])}</td></tr>
    </table>
</body>
</html>"""
    
    buffer = io.BytesIO()
    buffer.write(html.encode('utf-8'))
    buffer.seek(0)
    return buffer

def generate_xlsx_single_sheet(data, module_data_list, defects_df):
    """Генерирует Excel-отчёт"""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Отчёт о тестировании"
    
    COL_WIDTHS = {'A': 22, 'B': 14, 'C': 32, 'D': 12, 'E': 35}
    
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    section_fill = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
    context_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    defects_fill = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
    notes_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    signature_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
    pass_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    fail_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    wrap_left = Alignment(wrap_text=True, vertical="top", horizontal="left")
    wrap_center = Alignment(wrap_text=True, vertical="center", horizontal="center")
    wrap_right = Alignment(wrap_text=True, vertical="top", horizontal="right")
    
    row = 1
    
    # Заголовок
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws.cell(row=row, column=1, value=data["report_title"])
    cell.font = Font(name='Calibri', size=16, bold=True, color="FFFFFF")
    cell.fill = header_fill
    cell.alignment = wrap_center
    for col in range(1, 6):
        ws.cell(row=row, column=col).border = thin_border
    
    row += 2
    
    # Ключевые метрики
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws.cell(row=row, column=1, value="📊 КЛЮЧЕВЫЕ МЕТРИКИ")
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = section_fill
    cell.alignment = wrap_center
    for col in range(1, 6):
        ws.cell(row=row, column=col).border = thin_border
    
    row += 1
    
    summary_rows = [
        ["Проект", data["project"]],
        ["Версия", data["version"]],
        ["Период тестирования", data["test_period"]],
        ["Всего тест-кейсов", str(data["total_tc"])],
        ["Успешно (Pass)", f"{data['pass']} ({data['pass']/data['total_tc']*100:.1f}%)"],
        ["Упали (Fail)", f"{data['fail']} ({data['fail']/data['total_tc']*100:.1f}%)"],
        ["Critical (S1)", str(data["s1"])],
        ["Major (S2)", str(data["s2"])],
        ["Статус релиза", data["release_status"]],
        ["Рекомендация", data["recommendation"]],
    ]
    
    for label, value in summary_rows:
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=1, value=label).border = thin_border
        ws.cell(row=row, column=1, value=label).alignment = wrap_right
        ws.merge_cells(f'B{row}:E{row}')
        cell_value = ws.cell(row=row, column=2, value=value)
        cell_value.border = thin_border
        cell_value.alignment = wrap_left
        row += 1
    
    row += 1
    
    # Контекст тестирования
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws.cell(row=row, column=1, value="⚙️ КОНТЕКСТ ТЕСТИРОВАНИЯ")
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = context_fill
    cell.alignment = wrap_center
    for col in range(1, 6):
        ws.cell(row=row, column=col).border = thin_border
    
    row += 1
    
    context_rows = [
        ["Устройство / Браузер", data["device_browser"]],
        ["ОС / Платформа", data["os_platform"]],
        ["Сборка / Версия", data["build"]],
        ["Стенд", data["env_url"].strip()],
        ["Инструменты", data["tools"]],
        ["Методология", data["methodology"]],
        ["Тест-инженер", data["engineer"]],
        ["Дата формирования", data["report_date"]],
    ]
    
    for label, value in context_rows:
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=1, value=label).border = thin_border
        ws.cell(row=row, column=1, value=label).alignment = wrap_right
        ws.merge_cells(f'B{row}:E{row}')
        cell_value = ws.cell(row=row, column=2, value=value)
        cell_value.border = thin_border
        cell_value.alignment = wrap_left
        row += 1
    
    row += 1
    
    # Результаты по модулям
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws.cell(row=row, column=1, value="✅ РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ ПО МОДУЛЯМ")
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = section_fill
    cell.alignment = wrap_center
    for col in range(1, 6):
        ws.cell(row=row, column=col).border = thin_border
    
    row += 1
    
    test_headers = ["Модуль", "ID", "Сценарий", "Статус", "Комментарий"]
    for col_idx, header in enumerate(test_headers, start=1):
        cell = ws.cell(row=row, column=col_idx, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = wrap_center
    
    row += 1
    
    for module_info in module_data_list:
        module_name = module_info['title']
        df = module_info['df']
        if not df.empty:
            for _, test_row in df.iterrows():
                ws.cell(row=row, column=1, value=module_name).border = thin_border
                ws.cell(row=row, column=1, value=module_name).alignment = wrap_left
                
                ws.cell(row=row, column=2, value=test_row[0]).border = thin_border
                ws.cell(row=row, column=2, value=test_row[0]).alignment = wrap_center
                
                ws.cell(row=row, column=3, value=test_row[1]).border = thin_border
                ws.cell(row=row, column=3, value=test_row[1]).alignment = wrap_left
                
                status_cell = ws.cell(row=row, column=4, value=test_row[2])
                status_cell.border = thin_border
                status_cell.alignment = wrap_center
                if str(test_row[2]).upper() == "PASS":
                    status_cell.fill = pass_fill
                    status_cell.font = Font(color="006100", bold=True)
                elif str(test_row[2]).upper() == "FAIL":
                    status_cell.fill = fail_fill
                    status_cell.font = Font(color="9C0006", bold=True)
                
                ws.cell(row=row, column=5, value=test_row[3]).border = thin_border
                ws.cell(row=row, column=5, value=test_row[3]).alignment = wrap_left
                
                row += 1
    
    row += 1
    
    # Анализ дефектов
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws.cell(row=row, column=1, value="🐞 АНАЛИЗ ДЕФЕКТОВ")
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = defects_fill
    cell.alignment = wrap_center
    for col in range(1, 6):
        ws.cell(row=row, column=col).border = thin_border
    
    row += 1
    
    defect_headers = ["ID", "Модуль", "Заголовок", "Серьёзность", "Статус"]
    for col_idx, header in enumerate(defect_headers, start=1):
        cell = ws.cell(row=row, column=col_idx, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = wrap_center
    
    row += 1
    
    if not defects_df.empty:
        for _, defect_row in defects_df.iterrows():
            for col_idx, value in enumerate(defect_row, start=1):
                cell = ws.cell(row=row, column=col_idx, value=value)
                cell.border = thin_border
                cell.alignment = wrap_left if col_idx in (3, 5) else wrap_center
            row += 1
    else:
        ws.merge_cells(f'A{row}:E{row}')
        cell = ws.cell(row=row, column=1, value="Нет зарегистрированных дефектов")
        cell.alignment = wrap_center
        cell.border = thin_border
        row += 1
    
    row += 1
    
    # Ограничения, вывод, рекомендации
    sections = [
        ("⚠️ ОГРАНИЧЕНИЯ ТЕСТИРОВАНИЯ", data["limitations"]),
        ("💡 ВЫВОД", data["conclusion"]),
        ("📌 РЕКОМЕНДАЦИИ", data["recommendations_detailed"]),
    ]
    
    for title, content in sections:
        ws.merge_cells(f'A{row}:E{row}')
        cell = ws.cell(row=row, column=1, value=title)
        cell.font = Font(bold=True, size=12, color="FFFFFF")
        cell.fill = notes_fill
        cell.alignment = wrap_center
        for col in range(1, 6):
            ws.cell(row=row, column=col).border = thin_border
        row += 1
        
        for line in content.split('\n'):
            if line.strip():
                ws.merge_cells(f'A{row}:E{row}')
                cell = ws.cell(row=row, column=1, value=line.strip())
                cell.alignment = wrap_left
                cell.border = thin_border
                row += 1
        
        row += 1
    
    # Подпись
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws.cell(row=row, column=1, value="Подпись")
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = signature_fill
    cell.alignment = wrap_center
    for col in range(1, 6):
        ws.cell(row=row, column=col).border = thin_border
    
    row += 1
    
    signature_rows = [
        ["Роль", data["role"]],
        ["ФИО", data["fullname"]],
        ["Дата", data["signature_date"]],
    ]
    
    for label, value in signature_rows:
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=1, value=label).border = thin_border
        ws.cell(row=row, column=1, value=label).alignment = wrap_right
        ws.merge_cells(f'B{row}:E{row}')
        cell_value = ws.cell(row=row, column=2, value=value)
        cell_value.border = thin_border
        cell_value.alignment = wrap_left
        row += 1
    
    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width
    
    wb.save(output)
    output.seek(0)
    return output

# === ДАННЫЕ ПО УМОЛЧАНИЮ (точно как в образце PDF) ===
default_modules = [
    {
        "title": "Главный экран и навигация",
        "df": pd.DataFrame([
            ["MAIN-01", "Отображение карточек товаров", "PASS", "—"],
            ["MAIN-02", "Фильтрация по категориям", "PASS", "—"],
            ["NAV-01", "Переход между разделами", "PASS", "—"],
            ["NAV-02", "Поиск товара с опечаткой", "FAIL", "BUG-SEARCH-001. Не находятся товары при ошибке в 1 символе (например, «мыло» → «мылоо»)"]
        ], columns=["ID", "Сценарий", "Статус", "Комментарий"])
    },
    {
        "title": "Аутентификация и безопасность",
        "df": pd.DataFrame([
            ["AUTH-01", "Вход по логину/паролю", "PASS", "—"],
            ["SEC-01", "SQL-инъекция в поле поиска", "FAIL", "BUG-SEC-001. При вводе `' OR '1'='1` — белый экран, частичный краш"],
            ["SEC-02", "XSS-атака через поле поиска", "FAIL", "BUG-SEC-002. При вводе `<script>alert(1)</script>` — выполнение скрипта"]
        ], columns=["ID", "Сценарий", "Статус", "Комментарий"])
    },
    {
        "title": "Каталог и корзина",
        "df": pd.DataFrame([
            ["CATALOG-01", "Отображение списка товаров", "PASS", "—"],
            ["CART-01", "Добавление в корзину", "PASS", "—"],
            ["CART-02", "Оформление заказа", "PASS", "—"]
        ], columns=["ID", "Сценарий", "Статус", "Комментарий"])
    },
    {
        "title": "Дополнительные сценарии",
        "df": pd.DataFrame([
            ["OFFLINE-01", "Работа без интернета", "PASS", "Кэширование работает корректно"],
            ["SPECIAL-01", "Поиск со спецсимволами (@, #, $)", "PASS", "—"]
        ], columns=["ID", "Сценарий", "Статус", "Комментарий"])
    }
]

default_defects = pd.DataFrame([
    ["BUG-SEARCH-001", "Поиск", "Не работает fuzzy search (поиск с опечатками)", "Major (S2)", "New"],
    ["BUG-SEC-001", "Безопасность", "Уязвимость к SQL-инъекциям в поле поиска", "Critical (S1)", "New"],
    ["BUG-SEC-002", "Безопасность", "Уязвимость к XSS-атакам в поле поиска", "Critical (S1)", "New"]
], columns=["ID", "Модуль", "Заголовок", "Серьёзность", "Статус"])

# === ИНТЕРФЕЙС STREAMLIT (структура как в отчёте из PDF) ===
st.set_page_config(page_title="Генератор отчёта как в образце", layout="wide")
st.title("📄 Генератор отчёта о тестировании")

with st.form("main_form"):
    # === ЗАГОЛОВОК ОТЧЁТА ===
    report_title = st.text_input(
        "Название отчёта",
        "Отчёт о тестировании мобильного приложения Лемана ПРО"
    )
    
    # === ОСНОВНАЯ ИНФОРМАЦИЯ (отдельная секция ДО раздела 1, как в PDF) ===
    st.subheader("Основная информация")
    col_info1, col_info2 = st.columns(2)
    
    with col_info1:
        project = st.text_input("Проект", "Лемана ПРО")
        app_type = st.selectbox("Тип приложения", ["Мобильное", "Веб-приложение"], index=0)
        version = st.text_input("Версия приложения", "241006.001")
    
    with col_info2:
        test_period = st.text_input("Период тестирования", "29–30 ноября 2025 г.")
        report_date = st.text_input("Дата формирования отчёта", "30 ноября 2025 г.")
        engineer = st.text_input("Тест-инженер", "Черкасов Игорь")
    
    # === РАЗДЕЛ 1: КРАТКОЕ РЕЗЮМЕ (только данные резюме, как в таблице из PDF) ===
    st.header("1. Краткое резюме")
    col1, col2 = st.columns(2)
    
    with col1:
        release_status = st.selectbox("Статус релиза", ["НЕ РЕКОМЕНДОВАН К ВЫПУСКУ", "РЕКОМЕНДОВАН К ВЫПУСКУ"], index=0)
        s1 = st.number_input("Критические дефекты (S1)", min_value=0, value=2)
        s2 = st.number_input("Мажорные дефекты (S2)", min_value=0, value=1)
    
    with col2:
        total_tc = st.number_input("Всего тест-кейсов", min_value=1, value=72)
        pass_tc = st.number_input("Успешно (Pass)", min_value=0, value=69)
        fail_tc = st.number_input("Упали (Fail)", min_value=0, value=3)
    
    # Риски и рекомендации — под таблицами (как в образце)
    risk = st.text_area(
        "Основной риск",
        "Уязвимости безопасности позволяют нарушителю получить доступ к данным пользователей и вызвать отказ в обслуживании."
    )
    recommendation = st.text_area(
        "Рекомендация",
        "Релиз возможен только после устранения всех S1/S2 дефектов и повторного тестирования."
    )
    
    # === РАЗДЕЛ 2: КОНТЕКСТ ТЕСТИРОВАНИЯ ===
    st.header("2. Контекст тестирования")
    col3, col4 = st.columns(2)
    
    with col3:
        device_browser = st.text_input("Устройство / Браузер", "Xiaomi 12")
        os_platform = st.text_input("ОС / Платформа", "Android 15")
        build = st.text_input("Сборка / Версия", "lemanna-pro_241006.001.apk")
    
    with col4:
        env_url = st.text_input("URL стенда", "https://test.lemanna.pro")
        tools = st.text_input("Инструменты", "Postman (API), Burp Suite (безопасность), Jira (баг-трекинг)")
        methodology = st.text_input("Методология", "Ручное функциональное тестирование + проверка безопасности")
    
    # === РАЗДЕЛ 3: РЕЗУЛЬТАТЫ ПО МОДУЛЯМ ===
    st.header("3. Результаты тестирования по модулям")
    num_modules = st.slider("Количество модулей", min_value=1, max_value=10, value=4)
    
    module_data_list = []
    for i in range(num_modules):
        with st.expander(f"Модуль 3.{i+1}", expanded=True):
            title = st.text_input(
                f"Название модуля 3.{i+1}",
                value=default_modules[i]["title"] if i < len(default_modules) else f"Модуль 3.{i+1}",
                key=f"title_{i}"
            )
            df_key = f"mod_{i}"
            default_df = default_modules[i]["df"] if i < len(default_modules) else pd.DataFrame(columns=["ID", "Сценарий", "Статус", "Комментарий"])
            
            df = st.data_editor(
                default_df,
                num_rows="dynamic",
                key=df_key,
                column_config={
                    "ID": st.column_config.TextColumn("ID", width="small"),
                    "Сценарий": st.column_config.TextColumn("Сценарий", width="medium"),
                    "Статус": st.column_config.SelectboxColumn("Статус", options=["PASS", "FAIL"], width="small"),
                    "Комментарий": st.column_config.TextColumn("Комментарий", width="large")
                }
            )
            module_data_list.append({"title": title, "df": df})
    
    # === РАЗДЕЛ 4: АНАЛИЗ ДЕФЕКТОВ ===
    st.header("4. Анализ дефектов")
    defects = st.data_editor(
        default_defects,
        num_rows="dynamic",
        key="defects",
        column_config={
            "ID": st.column_config.TextColumn("ID", width="small"),
            "Модуль": st.column_config.TextColumn("Модуль", width="small"),
            "Заголовок": st.column_config.TextColumn("Заголовок", width="medium"),
            "Серьёзность": st.column_config.SelectboxColumn("Серьёзность", options=["Critical (S1)", "Major (S2)", "Minor (S3)"], width="small"),
            "Статус": st.column_config.SelectboxColumn("Статус", options=["New", "Open", "Fixed", "Closed"], width="small")
        }
    )
    consequences = st.text_area(
        "Последствия",
        "- S1 дефекты позволяют злоумышленнику получить данные других пользователей или вывести приложение из строя.\n- S2 дефект снижает юзабилити: пользователи не найдут товар при опечатке."
    )
    
    # === РАЗДЕЛ 5: ОГРАНИЧЕНИЯ ТЕСТИРОВАНИЯ ===
    st.header("5. Ограничения тестирования")
    limitations = st.text_area(
        "Ограничения тестирования",
        "1. Не тестировалась оплата через Apple Pay (устройство Android).\n2. Не проверена синхронизация с 1С (нет доступа к интеграционному стенду).\n3. Не проведено нагрузочное тестирование (ограничение по времени)."
    )
    
    # === РАЗДЕЛ 6: ВЫВОД И РЕКОМЕНДАЦИИ ===
    st.header("6. Вывод и рекомендации")
    conclusion = st.text_area(
        "Вывод",
        "Сборка 241006.001 содержит критические уязвимости безопасности, делающие её непригодной для выпуска в production. Наличие S1 дефектов нарушает базовые принципы защиты данных пользователей."
    )
    recommendations_detailed = st.text_area(
        "Рекомендации (подробно)",
        "Немедленно исправить уязвимости BUG-SEC-001 и BUG-SEC-002.\nРеализовать fuzzy search для повышения юзабилити (BUG-SEARCH-001).\nПровести повторное тестирование после фиксов с фокусом на:\n- Повторную проверку полей ввода на инъекции\n- Тестирование сценариев поиска с опечатками\n- Настроить автоматизированную проверку безопасности (например, OWASP ZAP) в CI/CD."
    )
    
    # === РАЗДЕЛ 7: ПОДПИСЬ ===
    st.header("7. Подпись")
    role = st.text_input("Роль", "QA-инженер")
    fullname = st.text_input("ФИО", "Черкасов Игорь")
    signature_date = st.text_input("Дата", "30.11.2025")
    
    submitted = st.form_submit_button("📥 Создать отчёт", type="primary")

# === ГЕНЕРАЦИЯ ОТЧЁТА ===
if submitted:
    validation_errors = []
    
    if pass_tc + fail_tc != total_tc:
        validation_errors.append(
            f"⚠️ Сумма статусов ({pass_tc} PASS + {fail_tc} FAIL = {pass_tc + fail_tc}) "
            f"не равна общему количеству тест-кейсов ({total_tc})"
        )
    
    if total_tc <= 0:
        validation_errors.append("❌ Общее количество тест-кейсов должно быть больше 0")
    
    if s1 < 0 or s2 < 0:
        validation_errors.append("❌ Количество дефектов не может быть отрицательным")
    
    if not report_title.strip():
        validation_errors.append("❌ Название отчёта не может быть пустым")
    
    if validation_errors:
        for error in validation_errors:
            st.error(error)
        st.stop()
    
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
    
    try:
        docx_buffer = generate_docx(data, module_data_list, defects)
        html_buffer = generate_html_report(data, module_data_list, defects)
        xlsx_buffer = generate_xlsx_single_sheet(data, module_data_list, defects)
        
        st.success("✅ Отчёт создан!")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button(
                "📄 DOCX",
                docx_buffer,
                "Отчёт_о_тестировании.docx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
        with col2:
            st.download_button(
                "🌐 HTML",
                html_buffer,
                "Отчёт_о_тестировании.html",
                "text/html",
                use_container_width=True
            )
        with col3:
            st.download_button(
                "📊 XLSX",
                xlsx_buffer,
                "Отчёт_о_тестировании.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    except Exception as e:
        st.error(f"❌ Ошибка генерации: {str(e)}")
        with st.expander("Детали ошибки"):
            st.code(traceback.format_exc())