# main.py

import logging
import sys
from html_to_json import parse_html_to_json
from excel_to_json import excel_to_json
from compare_json import compare_json


def setup_logging() -> None:
    """
    Настраивает конфигурацию логирования для приложения.
    """
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )


def main() -> None:
    """
    Основная функция приложения, которая выполняет парсинг HTML, извлечение данных из Excel,
    сравнение JSON-файлов и генерацию выходного Excel-файла.
    """
    setup_logging()

    # Пути к файлам (замените на ваши пути или используйте аргументы командной строки)
    html_file = 'page.html'  # HTML файл для парсинга
    html_json_file = 'result.json'  # JSON-файл, генерируемый из HTML

    excel_file = 'data.xlsx'  # Excel-файл с данными виртуальных машин
    excel_json_file = 'excel_data.json'  # JSON-файл, генерируемый из Excel

    output_excel_file = 'comparison_result.xlsx'  # Имя выходного файла

    try:
        # Парсинг HTML и генерация JSON
        logging.info("Парсинг HTML и генерация JSON файла")
        parse_html_to_json(html_file, html_json_file)

        # Извлечение данных из Excel и генерация JSON
        logging.info("Извлечение данных из Excel и генерация JSON файла")
        excel_to_json(excel_file, excel_json_file)

        # Сравнение JSON-файлов и генерация выходного Excel файла
        logging.info("Сравнение данных и генерация выходного Excel файла")
        compare_json(html_json_file, excel_json_file, output_excel_file)

        logging.info(f"Скрипт успешно выполнен. Результаты сохранены в файле {output_excel_file}")

    except Exception as e:
        logging.error(f"Произошла ошибка при выполнении скрипта: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
