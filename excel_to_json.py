# excel_to_json.py

import json
import logging
from typing import List, Dict, Any
import pandas as pd

logger = logging.getLogger(__name__)


def excel_to_json(excel_file: str, json_file: str) -> None:
    """
    Преобразует данные из Excel-файла в JSON-файл.

    :param excel_file: Путь к исходному Excel-файлу.
    :param json_file: Путь к выходному JSON-файлу.
    :raises FileNotFoundError: Если Excel-файл не найден.
    :raises ValueError: Если отсутствуют требуемые колонки.
    :raises Exception: Для остальных ошибок при обработке файла.
    """
    try:
        logger.info(f"Чтение Excel-файла: {excel_file}")
        df = load_excel(excel_file, sheet_name='Support')

        logger.debug("Убираем лишние пробелы в названиях колонок")
        df.columns = [col.strip() for col in df.columns]

        logger.info("Проверка наличия необходимых колонок")
        required_columns = ['Имя сервера', 'Сайзинг\ncpu/ram/hdd sys/hdd app', 'IP адрес']
        check_required_columns(df, required_columns)

        logger.debug("Переименование колонок для удобства")
        df = rename_columns(df)

        logger.debug("Отбор необходимых колонок")
        df = select_columns(df)

        logger.info("Удаление записей с отсутствующими именами серверов")
        df = drop_missing_server_names(df)

        logger.info("Преобразование DataFrame в список словарей")
        data = df.to_dict(orient='records')

        logger.info(f"Сохранение данных в JSON-файл: {json_file}")
        save_json(data, json_file)

    except FileNotFoundError as fnfe:
        logger.error(f"Excel-файл не найден: {fnfe}")
        raise
    except ValueError as ve:
        logger.error(f"Ошибка в данных Excel-файла: {ve}")
        raise
    except Exception as e:
        logger.error(f"Ошибка при обработке Excel-файла: {e}")
        raise


def load_excel(file_path: str, sheet_name: str = 'Support') -> pd.DataFrame:
    """
    Загружает данные из Excel-файла.

    :param file_path: Путь к Excel-файлу.
    :param sheet_name: Название листа для чтения.
    :return: DataFrame с данными.
    :raises FileNotFoundError: Если файл не найден.
    :raises Exception: Для остальных ошибок при чтении файла.
    """
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    except FileNotFoundError as fnfe:
        logger.error(f"Excel-файл не найден: {file_path}")
        raise
    except Exception as e:
        logger.error(f"Не удалось прочитать Excel-файл {file_path}: {e}")
        raise


def check_required_columns(df: pd.DataFrame, required_columns: List[str]) -> None:
    """
    Проверяет наличие необходимых колонок в DataFrame.

    :param df: DataFrame для проверки.
    :param required_columns: Список требуемых колонок.
    :raises ValueError: Если отсутствуют необходимые колонки.
    """
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Отсутствуют следующие колонки в Excel-файле: {missing_columns}")
    logger.debug("Все необходимые колонки присутствуют")


def rename_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Переименовывает колонки DataFrame для удобства.

    :param df: Исходный DataFrame.
    :return: DataFrame с переименованными колонками.
    """
    return df.rename(columns={
        'Имя сервера': 'Имя сервера',
        'Сайзинг\ncpu/ram/hdd sys/hdd app': 'Сайзинг',
        'IP адрес': 'IP адрес'
    })


def select_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Отбирает необходимые колонки из DataFrame.

    :param df: Исходный DataFrame.
    :return: DataFrame с выбранными колонками.
    """
    return df[['Имя сервера', 'Сайзинг', 'IP адрес']]


def drop_missing_server_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    Удаляет записи с отсутствующими именами серверов.

    :param df: Исходный DataFrame.
    :return: DataFrame без записей с отсутствующими именами серверов.
    """
    return df.dropna(subset=['Имя сервера'])


def save_json(data: List[Dict[str, Any]], json_file: str) -> None:
    """
    Сохраняет данные в формате JSON в файл.

    :param data: Данные для сохранения.
    :param json_file: Путь к выходному JSON-файлу.
    :raises Exception: Если не удалось сохранить файл.
    """
    try:
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        logger.info(f"Данные успешно сохранены в файле {json_file}")
    except Exception as e:
        logger.error(f"Не удалось сохранить JSON-файл {json_file}: {e}")
        raise
