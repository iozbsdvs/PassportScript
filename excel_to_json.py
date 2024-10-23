# excel_to_json.py

import pandas as pd
import json
import logging


def excel_to_json(excel_file, json_file):
    try:
        # Читаем данные из Excel-файла
        df = pd.read_excel(excel_file, sheet_name='Support')

        # Убираем лишние пробелы в названиях колонок
        df.columns = [col.strip() for col in df.columns]

        # Проверка, что нужные колонки существуют
        required_columns = ['Имя сервера', 'Сайзинг\ncpu/ram/hdd sys/hdd app', 'IP адрес']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            raise ValueError(f"Отсутствуют следующие колонки в Excel файле: {missing_columns}")

        # Переименовываем колонки для удобства
        df = df.rename(columns={
            'Имя сервера': 'Имя сервера',
            'Сайзинг\ncpu/ram/hdd sys/hdd app': 'Сайзинг',
            'IP адрес': 'IP адрес'
        })

        # Отбираем нужные колонки
        df = df[['Имя сервера', 'Сайзинг', 'IP адрес']]

        # Удаляем записи с отсутствующими именами серверов
        df = df.dropna(subset=['Имя сервера'])

        # Преобразуем DataFrame в список словарей
        data = df.to_dict(orient='records')

        # Сохраняем данные в JSON-файл
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

        logging.info(f"Данные успешно сохранены в файле {json_file}")

    except Exception as e:
        logging.error(f"Ошибка при обработке Excel файла: {e}")
        raise e  # Пробрасываем исключение вверх
