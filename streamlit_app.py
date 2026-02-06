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