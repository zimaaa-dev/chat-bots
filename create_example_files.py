import pandas as pd
import os

# Создаем папку files_excel если ее нет
os.makedirs('files_excel', exist_ok=True)

# Пример данных для Excel файла
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

# Создаем несколько примеров файлов
for i in range(1, 4):
    filename = f'files_excel/files_start_{i}.xlsx'
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Otvet', index=False)
        
        # Создаем лист для формул
        wb = writer.book
        ws_formules = wb.create_sheet('Formules')
        ws_formules['A1'] = 'ID вопроса'
        ws_formules['B1'] = 'Формула'
    
    print(f'Создан файл: {filename}')

print('Все файлы созданы успешно!')