import pandas as pd

def process_file(file_path):
    try:
        # Загрузим файл с указанием заголовков
        input_df = pd.read_excel(file_path, header=0)  # предполагаем, что заголовки на первой строке

        # Выводим столбцы, чтобы убедиться в их наличии
        print("Столбцы в DataFrame:")
        print(input_df.columns)

        # Печатаем первые несколько строк для проверки данных
        print("\nПервые строки данных:")
        print(input_df.head())

        # Удаление столбцов с ненужными или пустыми значениями
        input_df = input_df.dropna(axis=1, how='all')  # Удаляет пустые столбцы
        print("\nПосле удаления пустых столбцов:")
        print(input_df.columns)

        # Если нужно, переименовываем столбцы для удобства
        if len(input_df.columns) == 13:  # Убедимся, что количество столбцов соответствует
            input_df.columns = [
                'Вопрос', 'Оценка', 'Вес', 'Оценка с учетом веса, %',
                'Не соответствует ожиданиям', 'Значительно ниже ожиданий', 'Ниже ожиданий',
                'Частично ниже ожиданий', 'Соответствует ожиданиям', 'Выше ожиданий',
                'Превосходит все ожидания', 'Не взаимодействовал(-а)'
            ]
        else:
            print(f"Предполагаемое количество столбцов не совпадает. Количество: {len(input_df.columns)}")

        # Проверим, что все нужные столбцы есть
        required_columns = [
            'Вопрос', 'Оценка', 'Вес', 'Оценка с учетом веса, %',
            'Не соответствует ожиданиям', 'Значительно ниже ожиданий', 'Ниже ожиданий',
            'Частично ниже ожиданий', 'Соответствует ожиданиям', 'Выше ожиданий',
            'Превосходит все ожидания', 'Не взаимодействовал(-а)'
        ]

        missing_columns = [col for col in required_columns if col not in input_df.columns]
        if missing_columns:
            print(f"Не найдены следующие столбцы: {', '.join(missing_columns)}")
        else:
            print("Все необходимые столбцы присутствуют.")

        # Печатаем итоговый DataFrame
        print("\nИтоговые данные:")
        print(input_df[required_columns])

        # Возвращаем DataFrame
        return input_df

    except Exception as e:
        print(f"Ошибка при обработке файла: {e}")

# Путь к вашему файлу Excel
file_path = "/home/sirius/Рабочий стол/chatBot_sirius/CSI.XLSX"

# Обработка файла
processed_data = process_file(file_path)
