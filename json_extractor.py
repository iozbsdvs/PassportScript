# json_extractor.py

import json
import pandas as pd
import logging
from typing import Any, Dict, List


def setup_logger() -> None:
    """
    Настраивает конфигурацию логирования для модуля.
    """
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )


def extract_data_from_json(json_file: str) -> pd.DataFrame:
    """
    Извлекает данные о виртуальных машинах из JSON-файла и преобразует их в DataFrame.

    :param json_file: Путь к JSON-файлу.
    :return: DataFrame с информацией о виртуальных машинах.
    :raises FileNotFoundError: Если JSON-файл не найден.
    :raises json.JSONDecodeError: Если JSON-файл содержит некорректный JSON.
    :raises Exception: Для остальных ошибок при обработке файла.
    """
    try:
        logging.info(f"Загрузка данных из JSON-файла: {json_file}")
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)

        vm_list: List[Dict[str, Any]] = []

        for section in data:
            section_name = section.get('Раздел', '')
            if not section_name:
                logging.warning("Отсутствует название раздела. Пропуск раздела.")
                continue

            for item in section.get('Данные', []):
                naimenovanie = item.get('Наименование', '')
                role = item.get('Роль', '')
                for vm in item.get('ВМ', []):
                    vm_name = vm.get('Имя сервера', '')
                    ip_address = vm.get('IP адрес', '')
                    sizing = vm.get('Сайзинг', '')

                    if not vm_name:
                        logging.warning("Отсутствует 'Имя сервера'. Пропуск записи ВМ.")
                        continue

                    vm_entry = {
                        'Раздел': section_name,
                        'Наименование': naimenovanie,
                        'Роль': role,
                        'Имя сервера': vm_name,
                        'IP адрес': ip_address,
                        'Сайзинг': sizing
                    }
                    vm_list.append(vm_entry)

        if not vm_list:
            logging.warning("Нет данных для преобразования в DataFrame.")

        # Создаем DataFrame из списка ВМ
        df_json = pd.DataFrame(vm_list)
        logging.info("Преобразование данных в DataFrame успешно завершено.")
        return df_json

    except FileNotFoundError as fnfe:
        logging.error(f"JSON-файл не найден: {json_file}")
        raise fnfe
    except json.JSONDecodeError as jde:
        logging.error(f"Ошибка декодирования JSON-файла {json_file}: {jde}")
        raise jde
    except Exception as e:
        logging.error(f"Ошибка при обработке JSON-файла: {e}")
        raise e


def save_dataframe_to_excel(df: pd.DataFrame, excel_file: str) -> None:
    """
    Сохраняет DataFrame в Excel-файл.

    :param df: DataFrame для сохранения.
    :param excel_file: Путь к выходному Excel-файлу.
    :raises Exception: Если не удалось сохранить файл.
    """
    try:
        logging.info(f"Сохранение DataFrame в Excel-файл: {excel_file}")
        df.to_excel(excel_file, index=False)
        logging.info(f"DataFrame успешно сохранен в файле {excel_file}")
    except Exception as e:
        logging.error(f"Не удалось сохранить Excel-файл {excel_file}: {e}")
        raise e


def main():
    """
    Основная функция для извлечения данных из JSON и сохранения их в Excel.
    """
    setup_logger()

    # Пути к файлам (замените на ваши пути или используйте аргументы командной строки)
    json_input = 'result.json'  # JSON-файл с данными из HTML
    excel_output = 'json_data.xlsx'  # Excel-файл для сохранения данных

    try:
        # Извлечение данных из JSON
        df = extract_data_from_json(json_input)

        # Сохранение данных в Excel
        save_dataframe_to_excel(df, excel_output)

    except Exception as e:
        logging.exception(f"Произошла ошибка при выполнении скрипта: {e}")

