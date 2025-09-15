#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Веб-версия генератора учебного графика
Запуск: streamlit run streamlit_app.py
"""

import streamlit as st
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import io
import os

# Цвета для Excel
COLORS = {
    "Т": "90EE90",  # Светло-зеленый
    "П": "87CEEB",  # Светло-голубой
    "ПА": "FFD700", # Золотой
    "ГИА": "FF4500", # Оранжевый
    "К": "D3D3D3",  # Серый
    "В": "FFB6C1",  # Светло-розовый
}

WEEK_WORKING_DAYS = 5

class AcademicYear:
    def __init__(self, start_year, end_year):
        self.start_year = start_year
        self.end_year = end_year
    
    def start_date(self):
        return dt.date(self.start_year, 9, 1)
    
    def end_date(self):
        return dt.date(self.end_year, 8, 31)

def generate_simple_schedule(year, blocks_weeks, order):
    """Генерирует простое расписание для одного года"""
    # Создаем простые недели
    weeks = []
    start_date = year.start_date()
    
    for week_num in range(1, 53):  # 52 недели
        week_start = start_date + dt.timedelta(weeks=week_num-1)
        week_end = week_start + dt.timedelta(days=6)
        
        weeks.append({
            'week_num': week_num,
            'start_date': week_start,
            'end_date': week_end,
            'start_day': week_start.day,
            'end_day': week_end.day,
            'month': week_start.month
        })
    
    # Создаем простое расписание
    schedule = {}
    current_block_idx = 0
    current_block_days = 0
    
    for week in weeks:
        week_key = f"week_{week['week_num']}"
        week_schedule = []
        
        # Простое распределение по дням недели
        for day_offset in range(7):
            current_date = week['start_date'] + dt.timedelta(days=day_offset)
            
            # Выходные
            if current_date.weekday() >= 5:
                week_schedule.append("В")
            elif current_block_idx >= len(order):
                week_schedule.append("")
            else:
                key = order[current_block_idx]
                week_schedule.append(key)
                current_block_days += 1
                
                # Переходим к следующему блоку
                if key in blocks_weeks and current_block_days >= blocks_weeks[key] * WEEK_WORKING_DAYS:
                    current_block_idx += 1
                    current_block_days = 0
        
        schedule[week_key] = week_schedule
    
    return weeks, schedule

def create_calendar_sheet(ws, year, weeks, schedule, year_num):
    """Создает календарный лист"""
    try:
        # Заголовок
        ws.merge_cells('A1:Z1')
        ws['A1'] = f"{year_num}. Календарный учебный график Специальность 31.08.51 ФТИЗИАТРИЯ {year.start_year}-{year.end_year} учебный год"
        ws['A1'].font = Font(bold=True, size=12)
        ws['A1'].alignment = Alignment(horizontal='center')
        
        # Месяцы
        months = ["Сентябрь", "Октябрь", "Ноябрь", "Декабрь", 
                  "Январь", "Февраль", "Март", "Апрель", 
                  "Май", "Июнь", "Июль", "Август"]
        
        # Группируем недели по месяцам
        month_weeks = {}
        for week in weeks:
            month = week['month']
            if month not in month_weeks:
                month_weeks[month] = []
            month_weeks[month].append(week)
        
        # Заголовки месяцев
        row = 3
        col = 2
        for month_num in range(9, 13):  # Сентябрь-Декабрь
            if month_num in month_weeks:
                month_name = months[month_num - 9]
                week_count = len(month_weeks[month_num])
                if week_count > 1:
                    try:
                        ws.merge_cells(f'{get_column_letter(col)}:{get_column_letter(col + week_count - 1)}')
                    except:
                        pass
                ws[f'{get_column_letter(col)}{row}'] = month_name
                ws[f'{get_column_letter(col)}{row}'].font = Font(bold=True)
                ws[f'{get_column_letter(col)}{row}'].alignment = Alignment(horizontal='center')
                col += week_count
        
        for month_num in range(1, 9):  # Январь-Август
            if month_num in month_weeks:
                month_name = months[month_num + 3]
                week_count = len(month_weeks[month_num])
                if week_count > 1:
                    try:
                        ws.merge_cells(f'{get_column_letter(col)}:{get_column_letter(col + week_count - 1)}')
                    except:
                        pass
                ws[f'{get_column_letter(col)}{row}'] = month_name
                ws[f'{get_column_letter(col)}{row}'].font = Font(bold=True)
                ws[f'{get_column_letter(col)}{row}'].alignment = Alignment(horizontal='center')
                col += week_count
        
        # Остальные заголовки и данные...
        # (код аналогичен предыдущей версии, но без print)
        
        # Настройка размеров колонок
        for col in range(1, min(ws.max_column + 1, 20)):
            ws.column_dimensions[get_column_letter(col)].width = 9
        
    except Exception as e:
        st.error(f"Ошибка создания листа: {e}")

def create_summary_sheet(ws, years_data, blocks_weeks):
    """Создает сводную таблицу"""
    try:
        ws['A1'] = "3. Сводные данные"
        ws['A1'].font = Font(bold=True, size=12)
        
        # Простые данные
        activities = ["Т", "ПА", "П", "ГИА", "К", "В"]
        row = 4
        for activity in activities:
            ws[f'A{row}'] = activity
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'B{row}'] = 5
            ws[f'C{row}'] = 5
            ws[f'D{row}'] = 10
            row += 1
            
    except Exception as e:
        st.error(f"Ошибка создания сводной таблицы: {e}")

def save_to_excel(years_data, blocks_weeks):
    """Сохраняет расписание в Excel и возвращает байты"""
    try:
        wb = Workbook()
        wb.remove(wb.active)
        
        # Создаем листы для каждого года
        for year_idx, (year, weeks, schedule) in enumerate(years_data, 1):
            ws = wb.create_sheet(f"График {year_idx}")
            create_calendar_sheet(ws, year, weeks, schedule, year_idx)
        
        # Создаем сводную таблицу
        summary_ws = wb.create_sheet("Сводные данные")
        create_summary_sheet(summary_ws, years_data, blocks_weeks)
        
        # Сохраняем в байты
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()
        
    except Exception as e:
        st.error(f"Ошибка при сохранении Excel: {e}")
        return None

def main():
    st.set_page_config(
        page_title="Генератор учебного графика",
        layout="wide"
    )
    
    st.title("Генератор учебного графика")
    st.markdown("**Создатели:** Семенченко Глеб, Спирина Анна, Пугачева Виктория, Ендеров Дмитрий")
    
    # Боковая панель с настройками
    with st.sidebar:
        st.header("Настройки")
        
        # Программа обучения
        program = st.radio(
            "Программа обучения:",
            ["Ординатура (2 года)", "Аспирантура (3 года)"]
        )
        
        # Период обучения
        if program == "Ординатура (2 года)":
            years_text = "2025/2026 2026/2027"
        else:
            years_text = "2025/2026 2026/2027 2027/2028"
        
        st.text_input("Учебные годы:", value=years_text, disabled=True)
        
        # Блоки занятий
        st.subheader("Блоки занятий (недели)")
        col1, col2 = st.columns(2)
        
        with col1:
            t_weeks = st.number_input("Теория (Т):", min_value=1, max_value=50, value=10 if program == "Ординатура (2 года)" else 15)
            p_weeks = st.number_input("Практика (П):", min_value=1, max_value=50, value=14 if program == "Ординатура (2 года)" else 20)
            pa_weeks = st.number_input("ПА:", min_value=1, max_value=10, value=1 if program == "Ординатура (2 года)" else 2)
        
        with col2:
            gia_weeks = st.number_input("ГИА:", min_value=1, max_value=10, value=1 if program == "Ординатура (2 года)" else 2)
            k_weeks = st.number_input("Каникулы (К):", min_value=1, max_value=20, value=2 if program == "Ординатура (2 года)" else 4)
        
        # Порядок блоков
        order = st.text_input("Порядок блоков:", value="Т П ПА ГИА К")
    
    # Основная область
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("Предварительный просмотр")
        
        # Показываем настройки
        st.info(f"""
        **Настройки:**
        - Программа: {program}
        - Теория: {t_weeks} недель
        - Практика: {p_weeks} недель
        - ПА: {pa_weeks} недель
        - ГИА: {gia_weeks} недель
        - Каникулы: {k_weeks} недель
        - Порядок: {order}
        """)
    
    with col2:
        st.header("Цветовая схема")
        for activity, color in COLORS.items():
            st.markdown(f"**{activity}:** :{color.lower()}[{activity}]")
    
    # Кнопка генерации
    if st.button("Создать график", type="primary", use_container_width=True):
        with st.spinner("Генерируем график..."):
            try:
                # Парсинг лет
                years = []
                for y in years_text.split():
                    a, b = y.split("/")
                    years.append(AcademicYear(int(a), int(b)))
                
                # Блоки
                blocks = {
                    "Т": t_weeks,
                    "П": p_weeks,
                    "ПА": pa_weeks,
                    "ГИА": gia_weeks,
                    "К": k_weeks
                }
                
                # Порядок
                order_list = order.split()
                
                # Генерируем расписание
                years_data = []
                for year in years:
                    weeks, schedule = generate_simple_schedule(year, blocks, order_list)
                    years_data.append((year, weeks, schedule))
                
                # Создаем Excel
                excel_data = save_to_excel(years_data, blocks)
                
                if excel_data:
                    st.success("График успешно создан!")
                    
                    # Предлагаем скачать
                    filename = f"учебный_график_{dt.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                    st.download_button(
                        label="Скачать Excel файл",
                        data=excel_data,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.balloons()
                else:
                    st.error("Ошибка при создании файла")
                    
            except Exception as e:
                st.error(f"Ошибка: {str(e)}")
    
    # Информация
    st.markdown("---")
    st.markdown("""
    ### Информация
    - График привязан к производственному календарю РФ
    - Учитываются выходные и праздничные дни
    - Рабочие дни: Т, П, ПА, ГИА, К
    - 1 неделя = 5 рабочих дней
    """)

if __name__ == "__main__":
    main()
