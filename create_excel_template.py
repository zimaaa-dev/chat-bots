import pandas as pd
from openpyxl import Workbook

# Создаем DataFrame с примерами вопросов
data = {
    'ID вопроса': [1, 2, 3, 4, 5],
    'Вопрос': [
        'Сколько вам лет?',
        'Вы согласны с условиями?',
        'Выберите вариант ответа',
        'Введите дату рождения',
        'Введите сумму'
    ],
    'Ответы число': [None, None, None, None, None],
    'Ответы текст': [None, None, None, None, None],
    'Клавиатура': [0, 1, 2, 3, 4],
    'Тип': ['int', 'bool', 'str', 'date', 'float'],
    'Условие': [None, None, None, None, None],
    'Варианты ответов': [None, 'ДА:НЕТ', 'Вариант А:Вариант Б:Вариант В', None, None],
    'Описание': [None, None, None, None, None],
    'Картинка': [None, None, None, None, None],
    'Список': [None, None, None, None, None],
    'Пусто': [None, None, None, None, None],
    'Пусто2': [None, None, None, None, None],
}

df = pd.DataFrame(data)

# Создаем Excel файл с двумя листами
with pd.ExcelWriter('files_excel/example_price_list.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Otvet', index=False)
    
    # Создаем лист для формул
    wb = writer.book
    ws_formules = wb.create_sheet('Formules')
    ws_formules['A1'] = 'ID вопроса'
    ws_formules['B1'] = 'Формула'
    
print('Пример Excel файла создан: files_excel/example_price_list.xlsx')