# main.py

import logging
from html_to_json import parse_html_to_json
from excel_extractor import extract_data_from_excel
from json_extractor import extract_data_from_json
from comparator import compare_data


def main():
    # Настройка логирования
    logging.basicConfig(level=logging.INFO, format='%(message)s')

    # Пути к файлам (замените на ваши пути)
    html_file = 'page.html'  # HTML файл для парсинга
    json_file = 'result.json'  # JSON-файл, генерируемый из HTML
    excel_file = 'data.xlsx'  # Excel-файл с данными виртуальных машин
    output_excel_file = 'comparison_result.xlsx'  # Имя выходного файла

    try:
        # Парсинг HTML и генерация JSON
        logging.info("Парсинг HTML и генерация JSON файла")
        parse_html_to_json(html_file, json_file)

        # Извлечение данных из JSON файла
        logging.info("Извлечение данных из JSON файла")
        df_json = extract_data_from_json(json_file)

        # Извлечение данных из Excel файла
        logging.info("Извлечение данных из Excel файла")
        df_excel = extract_data_from_excel(excel_file)

        # Проверяем, что данные были успешно извлечены
        if df_json.empty:
            logging.error("Ошибка: данные из JSON файла пусты.")
            return
        if df_excel.empty:
            logging.error("Ошибка: данные из Excel файла пусты.")
            return

        # Сравнение данных и генерация выходного Excel файла
        logging.info("Сравнение данных и генерация выходного Excel файла")
        compare_data(df_json, df_excel, output_excel_file)

        logging.info(f"Скрипт успешно выполнен. Результаты сохранены в файле {output_excel_file}")

    except Exception as e:
        logging.error(f"Произошла ошибка при выполнении скрипта: {e}")


if __name__ == "__main__":
    main()
