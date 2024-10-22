# excel_extractor.py

import pandas as pd
import logging


def extract_data_from_excel(excel_file):
    try:
        df = pd.read_excel(excel_file, sheet_name='Support')

        # Убираем лишние пробелы в названиях колонок
        df.columns = [col.strip() for col in df.columns]

        # Проверка, что нужные колонки существуют
        required_columns = ['Имя сервера', 'Сайзинг\ncpu/ram/hdd sys/hdd app', 'IP адрес']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            raise ValueError(f"Отсутствуют следующие колонки в Excel файле: {missing_columns}")

        # Переименовываем и отбираем нужные колонки
        df = df[['Имя сервера', 'Сайзинг\ncpu/ram/hdd sys/hdd app', 'IP адрес']]
        df = df.rename(columns={'Сайзинг\ncpu/ram/hdd sys/hdd app': 'Сайзинг'})
        df = df.dropna()

        logging.info(f"Извлечено {len(df)} записей из Excel файла")

        return df

    except Exception as e:
        logging.error(f"Ошибка при обработке Excel файла: {e}")
        raise e  # Пробрасываем исключение вверх
