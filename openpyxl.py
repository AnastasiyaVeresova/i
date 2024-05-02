import openpyxl

# Загрузка файла Excel
file = 'book.xlsx'
workbook = openpyxl.load_workbook(file)

# Сопоставление названий листов с соответствующими таблицами
sheet_tables = {
    'Таблица ячеек трансформатора 100-330': 'Таблица Т1',
    'Таблица2': 'Таблица Т2',
    # Добавьте другие соответствия здесь
}

# Функция для отображения таблицы данных
def display_table(sheet_name):
    sheet = workbook[sheet_tables[sheet_name]]
    table_data = []
    for row in sheet.iter_rows(values_only=True):
        table_data.append(row)
    # Отображение данных таблицы
    for row in table_data:
        print(row)

# Список заголовков таблиц
table_headers = ['Таблица ячеек трансформатора 100-330', 'Таблица2']

# Обработка наведения на заголовок
for header in table_headers:
    print(f'Наведение на заголовок: {header}')
    display_table(header)
