# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import matplotlib
matplotlib.use('Agg')  # Критически важно для работы в Streamlit Cloud
import matplotlib.pyplot as plt
import io
import base64
import traceback
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

def set_col_width(col, width_twips):
    """Устанавливает ширину колонки в таблице DOCX"""
    for cell in col.cells:
        tc = cell._element.tcPr
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), str(int(width_twips)))
        tcW.set(qn('w:type'), 'dxa')
        tc.append(tcW)

def add_table_from_df(doc, df):
    """Создаёт таблицу с фиксированной шириной и границами"""
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
                run.font.bold = True  # type: ignore
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
    """Генерирует строго деловой DOCX-отчёт"""
    doc = Document()
    # Исправление предупреждения BaseStyle.font: прямое присвоение
    doc.styles['Normal'].font.name = 'Calibri Light'  # type: ignore
    doc.styles['Normal'].font.size = Pt(12)  # type: ignore
    
    # ЗАГОЛОВОК
    title = doc.add_heading(data["report_title"], 0)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    title_font = title.runs[0].font
    title_font.size = Pt(16)
    title_font.bold = True

    # ЕДИНЫЙ РАСЧЁТ ШИРИНЫ ДЛЯ ВСЕХ 2-КОЛОНОЧНЫХ ТАБЛИЦ (в твипсах)
    total_width_twips = Inches(6.5).twips
    first_col_width_twips = int(total_width_twips * 0.25)
    second_col_width_twips = int(total_width_twips * 0.75)

    # ИНФОРМАЦИОННЫЕ ПОЛЯ
    info_table = doc.add_table(rows=6, cols=2)
    info_table.style = 'Table Grid'
    # УСТАНАВЛИВАЕМ ШИРИНУ ТОЛЬКО ЧЕРЕЗ КОЛОНКИ (НЕ ЯЧЕЙКИ!)
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
            run.font.bold = True  # type: ignore
        
        cell2 = info_table.cell(i, 1)
        cell2.text = value
        cell2.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # КРАТКОЕ РЕЗЮМЕ
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
            run.font.bold = True  # type: ignore
        
        cell2 = summary_table.cell(i, 1)
        cell2.text = value
        cell2.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ДИАГРАММЫ (гарантированно работают в Streamlit Cloud)
    plt.figure(figsize=(5, 4))
    plt.pie([data['pass'], data['fail']], labels=['PASS', 'FAIL'], autopct='%1.1f%%',
            colors=['#4CAF50', '#F44336'], startangle=90)
    plt.title('Рис. 1. Распределение результатов тест-кейсов')
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    plt.close()
    doc.add_picture(buf, width=Inches(5))
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    plt.figure(figsize=(5, 4))
    bars = plt.bar(['Critical (S1)', 'Major (S2)'], [data['s1'], data['s2']],
                   color=['#F44336', '#FF9800'])
    plt.title('Рис. 2. Дефекты по уровню серьёзности')
    plt.ylabel('Количество')
    plt.ylim(0, max(data['s1'], data['s2'], 1) * 1.3)
    for bar in bars:
        h = bar.get_height()
        if h > 0:
            plt.text(bar.get_x() + bar.get_width()/2, h + 0.05, str(int(h)), ha='center', va='bottom')
    plt.grid(axis='y', alpha=0.3, linestyle='--')
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    plt.close()
    doc.add_picture(buf, width=Inches(5))
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # КОНТЕКСТ ТЕСТИРОВАНИЯ
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
            run.font.bold = True  # type: ignore
        
        cell2 = context_table.cell(i, 1)
        cell2.text = value
        cell2.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # РЕЗУЛЬТАТЫ ПО МОДУЛЯМ
    doc.add_heading('3. РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ ПО МОДУЛЯМ', 1)
    for idx, module_info in enumerate(module_data_list):
        title = module_info['title']
        df = module_info['df']
        doc.add_heading(f'3.{idx+1}. {title}', 2)
        add_table_from_df(doc, df)

    # АНАЛИЗ ДЕФЕКТОВ
    doc.add_heading('4. АНАЛИЗ ДЕФЕКТОВ', 1)
    add_table_from_df(doc, defects_df)

    doc.add_paragraph('Последствия:').paragraph_format.space_after = Pt(6)
    doc.add_paragraph(data['consequences']).paragraph_format.space_after = Pt(6)

    # ОГРАНИЧЕНИЯ ТЕСТИРОВАНИЯ
    doc.add_heading('5. ОГРАНИЧЕНИЯ ТЕСТИРОВАНИЯ', 1)
    for line in data['limitations'].split('\n'):
        if line.strip():
            p = doc.add_paragraph()
            p.add_run(f"• {line.strip()}")
            p.paragraph_format.space_after = Pt(2)

    # ВЫВОД И РЕКОМЕНДАЦИИ
    doc.add_heading('6. ВЫВОД И РЕКОМЕНДАЦИИ', 1)
    doc.add_paragraph('Вывод:').paragraph_format.space_after = Pt(6)
    doc.add_paragraph(data['conclusion']).paragraph_format.space_after = Pt(6)
    doc.add_paragraph('Рекомендации:').paragraph_format.space_after = Pt(6)
    for line in data['recommendations_detailed'].split('\n'):
        if line.strip():
            p = doc.add_paragraph()
            p.add_run(f"• {line.strip()}")
            p.paragraph_format.space_after = Pt(2)

    # ПОДПИСЬ (ИСПРАВЛЕНА ОПЕЧАТКА: signature_table вместо info_table)
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
            run.font.bold = True  # type: ignore
        
        cell2 = signature_table.cell(i, 1)
        cell2.text = value
        cell2.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generate_chart_base64(pass_count, fail_count, s1_count, s2_count):
    """Генерирует диаграммы и возвращает их как корректные base64 строки"""
    # Диаграмма 1: Распределение результатов
    plt.figure(figsize=(6, 4.5))
    plt.pie([pass_count, fail_count], labels=['PASS', 'FAIL'], autopct='%1.1f%%',
            colors=['#4CAF50', '#F44336'], startangle=90, textprops={'fontsize': 11})
    plt.title('Рис. 1. Распределение результатов тест-кейсов', fontsize=10, pad=15)
    buf1 = io.BytesIO()
    plt.savefig(buf1, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    
    # Диаграмма 2: Дефекты по серьёзности
    plt.figure(figsize=(6, 4.5))
    bars = plt.bar(['Critical (S1)', 'Major (S2)'], [s1_count, s2_count],
                   color=['#F44336', '#FF9800'], width=0.5)
    plt.title('Рис. 2. Дефекты по уровню серьёзности', fontsize=10, pad=15)
    plt.ylabel('Количество', fontsize=11)
    plt.ylim(0, max(s1_count, s2_count, 1) * 1.3)
    for bar in bars:
        h = bar.get_height()
        if h > 0:
            plt.text(bar.get_x() + bar.get_width()/2, h + 0.05, str(int(h)), 
                    ha='center', va='bottom', fontsize=11, fontweight='bold')
    plt.grid(axis='y', alpha=0.3, linestyle='--')
    buf2 = io.BytesIO()
    plt.savefig(buf2, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    
    # === КРИТИЧЕСКИ ВАЖНО: корректное получение данных из буфера ===
    # Используем getvalue() для гарантированного получения полных данных
    chart1_base64 = base64.b64encode(buf1.getvalue()).decode('utf-8')
    chart2_base64 = base64.b64encode(buf2.getvalue()).decode('utf-8')
    
    return chart1_base64, chart2_base64

def generate_html_report(data, module_data_list, defects_df):
    """Генерирует HTML-версию отчёта с корректными диаграммами и форматированием"""
    # Генерация диаграмм
    chart1, chart2 = generate_chart_base64(data['pass'], data['fail'], data['s1'], data['s2'])
    
    # Расчёт процентов
    total = data['total_tc']
    pass_pct = data['pass'] / total * 100 if total > 0 else 0
    fail_pct = 100 - pass_pct
    
    # === КРИТИЧЕСКИ ВАЖНО: экранирование специальных символов в данных ===
    def escape_html(text):
        if not isinstance(text, str):
            return str(text)
        return (text.replace('&', '&amp;')
                    .replace('<', '&lt;')
                    .replace('>', '&gt;')
                    .replace('"', '&quot;')
                    .replace("'", '&#39;'))
    
    # Создание HTML с профессиональным форматированием
    html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{escape_html(data['report_title'])}</title>
    <style>
        body {{
            font-family: 'Calibri Light', Times, serif;
            font-size: 12pt;
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
        /* Фиксированная ширина колонок как в DOCX (25%/75%) */
        .info-table td:first-child,
        .summary-table td:first-child,
        .context-table td:first-child,
        .signature-table td:first-child {{
            width: 25%;
            font-weight: bold;
            background-color: #f9f9f9;
        }}
        .info-table td:last-child,
        .summary-table td:last-child,
        .context-table td:last-child,
        .signature-table td:last-child {{
            width: 75%;
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
        ul {{
            padding-left: 20px;
            margin: 10px 0;
        }}
        li {{
            margin-bottom: 5px;
        }}
        .no-print {{
            display: none;
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
            .no-print {{
                display: none !important;
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
    
    <!-- Информационная таблица (25%/75%) -->
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
    
    <!-- === ГАРАНТИРОВАННО КОРРЕКТНЫЕ ДИАГРАММЫ === -->
    <div class="chart-container">
        <img src="data:image/png;base64,{chart1}" 
             alt="Распределение результатов тест-кейсов" 
             style="max-width: 100%; height: auto; display: block; margin: 0 auto;">
        <div class="chart-title">Рис. 1. Распределение результатов тест-кейсов</div>
    </div>
    
    <div class="chart-container">
        <img src="data:image/png;base64,{chart2}" 
             alt="Дефекты по уровню серьёзности" 
             style="max-width: 100%; height: auto; display: block; margin: 0 auto;">
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
    
    # Результаты по модулям
    html += "<h2>3. РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ ПО МОДУЛЯМ</h2>"
    for idx, module_info in enumerate(module_data_list):
        html += f"<h3>3.{idx+1}. {escape_html(module_info['title'])}</h3>"
        html += '<table><tr><th style="width: 15%;">ID</th><th>Сценарий</th><th style="width: 12%;">Статус</th><th>Комментарий</th></tr>'
        df = module_info['df']
        if not df.empty:
            for _, row in df.iterrows():
                status_class = "status-pass" if str(row[2]).upper() == "PASS" else "status-fail" if str(row[2]).upper() == "FAIL" else ""
                html += f"<tr><td>{escape_html(row[0])}</td><td>{escape_html(row[1])}</td><td class='{status_class}'>{escape_html(row[2])}</td><td>{escape_html(row[3])}</td></tr>"
        else:
            html += "<tr><td colspan='4' style='text-align:center'>Нет данных</td></tr>"
        html += "</table>"
    
    # Анализ дефектов
    html += "<h2>4. АНАЛИЗ ДЕФЕКТОВ</h2>"
    html += '<table><tr><th style="width: 15%;">ID</th><th style="width: 15%;">Модуль</th><th>Заголовок</th><th style="width: 20%;">Серьёзность</th><th style="width: 15%;">Статус</th></tr>'
    if not defects_df.empty:
        for _, row in defects_df.iterrows():
            html += f"<tr><td>{escape_html(row[0])}</td><td>{escape_html(row[1])}</td><td>{escape_html(row[2])}</td><td>{escape_html(row[3])}</td><td>{escape_html(row[4])}</td></tr>"
    else:
        html += "<tr><td colspan='5' style='text-align:center'>Нет данных</td></tr>"
    html += "</table>"
    html += f"<p><strong>Последствия:</strong><br>{escape_html(data['consequences']).replace(chr(10), '<br>')}</p>"
    
    # Ограничения тестирования
    html += "<h2>5. ОГРАНИЧЕНИЯ ТЕСТИРОВАНИЯ</h2><ul>"
    for line in data['limitations'].split('\n'):
        if line.strip():
            html += f"<li>{escape_html(line.strip())}</li>"
    html += "</ul>"
    
    # Вывод и рекомендации
    html += f"""
        <h2>6. ВЫВОД И РЕКОМЕНДАЦИИ</h2>
        <p><strong>Вывод:</strong><br>{escape_html(data['conclusion'])}</p>
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
        
        <div class="no-print" style="margin-top: 30px; padding: 15px; background-color: #e3f2fd; border-radius: 5px; border: 1px solid #90caf9;">
            <h3 style="margin-top: 0;">💡 Как сохранить отчёт как PDF:</h3>
            <ol>
                <li>Нажмите <strong>Ctrl+P</strong> (Windows) или <strong>Cmd+P</strong> (Mac)</li>
                <li>Выберите «Сохранить как PDF»</li>
                <li>Установите ориентацию «Книжная», масштаб «100%»</li>
                <li>Нажмите «Сохранить»</li>
            </ol>
        </div>
    </body>
</html>"""
    
    buffer = io.BytesIO()
    buffer.write(html.encode('utf-8'))
    buffer.seek(0)
    return buffer

def generate_xlsx_single_sheet(data, module_data_list, defects_df):
    """Генерирует профессиональный XLSX-отчёт с корректной структурой"""
    from io import BytesIO
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Отчёт о тестировании"
    
    # Ширины: A=20 (метрики), B=45 (значения)
    COL_WIDTHS = {'A': 20, 'B': 45, 'C': 16, 'D': 14, 'E': 48}
    
    # Стили
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    section_fill = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
    context_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    defects_fill = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
    notes_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    signature_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
    pass_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    fail_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    critical_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    major_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    wrap_left = Alignment(wrap_text=True, vertical="top", horizontal="left")
    wrap_center = Alignment(wrap_text=True, vertical="center", horizontal="center")
    
    row = 1
    
    # === ЗАГОЛОВОК ===
    ws.merge_cells(f'A{row}:B{row}')
    cell = ws.cell(row=row, column=1, value=data["report_title"])
    cell.font = Font(name='Calibri', size=16, bold=True, color="FFFFFF")
    cell.fill = header_fill
    cell.alignment = wrap_center
    for col in range(1, 3):
        ws.cell(row=row, column=col).border = thin_border
    row += 2
    
    # === СВОДКА (2 колонки) ===
    ws.merge_cells(f'A{row}:B{row}')
    cell = ws.cell(row=row, column=1, value="📊 КЛЮЧЕВЫЕ МЕТРИКИ")
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = section_fill
    cell.alignment = wrap_center
    for col in range(1, 3):
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
        ws.cell(row=row, column=1, value=label).alignment = Alignment(wrap_text=True, vertical="top", horizontal="right")
        
        ws.cell(row=row, column=2, value=value).border = thin_border
        ws.cell(row=row, column=2, value=value).alignment = wrap_left
        
        # Подсветка статуса
        if "НЕ РЕКОМЕНДОВАН" in str(value):
            ws.cell(row=row, column=2, value=value).fill = critical_fill
            ws.cell(row=row, column=2, value=value).font = Font(color="FFFFFF", bold=True)
        elif "РЕКОМЕНДОВАН" in str(value):
            ws.cell(row=row, column=2, value=value).fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
            ws.cell(row=row, column=2, value=value).font = Font(color="FFFFFF", bold=True)
        
        row += 1
    row += 1
    
    # === КОНТЕКСТ (2 колонки) ===
    ws.merge_cells(f'A{row}:B{row}')
    cell = ws.cell(row=row, column=1, value="⚙️ КОНТЕКСТ ТЕСТИРОВАНИЯ")
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = context_fill
    cell.alignment = wrap_center
    for col in range(1, 3):
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
        ws.cell(row=row, column=1, value=label).alignment = Alignment(wrap_text=True, vertical="top", horizontal="right")
        
        ws.cell(row=row, column=2, value=value).border = thin_border
        ws.cell(row=row, column=2, value=value).alignment = wrap_left
        row += 1
    row += 1
    
    # === РЕЗУЛЬТАТЫ ТЕСТОВ (5 колонок) ===
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws.cell(row=row, column=1, value="✅ РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ ПО МОДУЛЯМ")
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = section_fill
    cell.alignment = wrap_center
    for col in range(1, 6):
        ws.cell(row=row, column=col).border = thin_border
    row += 1
    
    # Заголовки тестов
    test_headers = ["Модуль", "ID", "Сценарий", "Статус", "Комментарий"]
    for col_idx, header in enumerate(test_headers, start=1):
        cell = ws.cell(row=row, column=col_idx, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = wrap_center
    row += 1
    
    # Данные тестов
    for module_info in module_data_list:
        module_name = module_info['title']
        df = module_info['df']
        if not df.empty:
            for _, test_row in df.iterrows():
                ws.cell(row=row, column=1, value=module_name).border = thin_border
                ws.cell(row=row, column=1, value=module_name).alignment = wrap_left
                
                ws.cell(row=row, column=2, value=test_row[0]).border = thin_border
                ws.cell(row=row, column=2, value=test_row[0]).alignment = wrap_left
                
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
    
    # === ДЕФЕКТЫ (5 колонок) ===
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
                if col_idx == 4:  # Серьёзность
                    sev = str(value)
                    if "Critical" in sev:
                        cell.fill = critical_fill
                        cell.font = Font(color="FFFFFF", bold=True)
                    elif "Major" in sev:
                        cell.fill = major_fill
                        cell.font = Font(color="FFFFFF", bold=True)
            row += 1
    else:
        ws.merge_cells(f'A{row}:E{row}')
        cell = ws.cell(row=row, column=1, value="Нет зарегистрированных дефектов")
        cell.alignment = wrap_center
        cell.border = thin_border
        row += 1
    row += 1
    
    # === ОСТАЛЬНЫЕ СЕКЦИИ (Ограничения, Вывод, Рекомендации, Подпись) ===
    # ... (оставьте как в вашем текущем коде, но используйте 5 колонок A-E) ...
    
    # === УСТАНОВКА ШИРИН ===
    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width
    
    wb.save(output)
    output.seek(0)
    return output

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
        env_url = st.text_input("URL стенда", "https://test.lemanna.pro")
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
    recommendations_detailed = st.text_area("Рекомендации (подробно)", "Немедленно исправить уязвимости BUG-SEC-001 и BUG-SEC-002.\nРеализовать fuzzy search для повышения юзабилити (BUG-SEARCH-001).\nПровести повторное тестирование после фиксов с фокусом на:\n- Повторную проверку полей ввода на инъекции\n- Тестирование сценариев поиска с опечатками\n- Настроить автоматизированную проверку безопасности (например, OWASP ZAP) в CI/CD.")
    
    st.header("7. Подпись")
    role = st.text_input("Роль", "QA-инженер")
    fullname = st.text_input("ФИО", "Черкасов Игорь")
    signature_date = st.text_input("Дата", "30.11.2025")

    submitted = st.form_submit_button("📥 Создать отчёт", type="primary")

if submitted:
    # БАЗОВАЯ ВАЛИДАЦИЯ ДАННЫХ
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
        
    if pass_tc > total_tc or fail_tc > total_tc:
        validation_errors.append("❌ Количество успешных/проваленных тестов не может превышать общее")
    
    if validation_errors:
        for error in validation_errors:
            st.error(error)
        st.stop()
    
    # ПОДГОТОВКА ДАННЫХ
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
        # ГЕНЕРАЦИЯ ОТЧЁТОВ
        docx_buffer = generate_docx(data, module_data_list, defects)
        html_buffer = generate_html_report(data, module_data_list, defects)
        xlsx_buffer = generate_xlsx_single_sheet(data, module_data_list, defects)
        
        st.success("✅ Отчёт готов!")
        
        # КНОПКИ СКАЧИВАНИЯ
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
                "📊 XLSX",
                xlsx_buffer,
                "Отчёт_о_тестировании.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="secondary"
            )
        
        with col3:
            st.download_button(
                "🌐 HTML",
                html_buffer,
                "Отчёт_о_тестировании.html",
                "text/html",
                use_container_width=True
            )
        
        # ИНСТРУКЦИЯ
        st.markdown("""
        <div style="background-color: #3f403f; padding: 15px; border-radius: 8px; margin-top: 20px; border: 1px solid #81c784;">
            <h4>🖨️ Как получить профессиональный PDF:</h4>
            <ol>
                <li>Скачайте файл <strong>HTML</strong></li>
                <li>Откройте в <strong>браузере</strong></li>
                <li>Нажмите <kbd>Ctrl+P</kbd> → «Сохранить как PDF»</li>
                <li>Установите: ориентация «Книжная», масштаб «100%»</li>
                <li>Сохраните — получите отчёт с диаграммами</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
            
    except Exception as e:
        st.error(f"❌ Ошибка генерации отчёта: {str(e)}")
        with st.expander("Показать детали ошибки"):
            st.code(traceback.format_exc())