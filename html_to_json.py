import json
import re
import logging
from typing import Any, Dict, List, Optional
from bs4 import BeautifulSoup


def parse_html_to_json(html_file: str, json_file: str) -> None:
    """
    Парсит HTML-файл и сохраняет данные в формате JSON.

    :param html_file: Путь к HTML-файлу.
    :param json_file: Путь к выходному JSON-файлу.
    :raises FileNotFoundError: Если HTML-файл не найден.
    :raises Exception: Для остальных ошибок при парсинге.
    """
    try:
        # Настройка логирования
        logging.basicConfig(level=logging.DEBUG, format='%(message)s')

        # Чтение HTML-кода из файла
        html_content = load_html(html_file)

        # Парсинг HTML-кода
        soup = BeautifulSoup(html_content, 'html.parser')

        # Поиск всех <div class='innerCell'>
        inner_cells = soup.find_all('div', class_='innerCell')

        # Результирующий список
        result: List[Dict[str, Any]] = []

        for cell in inner_cells:
            # Извлечение заголовка раздела
            section_title = get_section_title(cell)
            if not section_title:
                logging.warning("Заголовок секции не найден. Пропуск секции.")
                continue

            logging.info(f"Обработка раздела: {section_title}")
            data: List[Dict[str, Any]] = []

            # Поиск таблицы внутри текущей ячейки
            table = cell.find('table')
            if table:
                # Парсинг таблицы в матрицу данных
                table_data = parse_html_table(table)

                if not table_data:
                    logging.warning(f"Таблица в разделе '{section_title}' пуста.")
                    continue

                # Извлечение и нормализация заголовков столбцов
                headers = table_data[0]
                normalized_headers = [normalize_header(header) for header in headers]
                logging.debug(f"Исходные заголовки: {headers}")
                logging.debug(f"Нормализованные заголовки: {normalized_headers}")
                header_indices = {header: idx for idx, header in enumerate(normalized_headers)}

                current_naimenovanie: str = ''
                current_role: str = ''

                for row in table_data[1:]:
                    # Обновление 'Наименование' и 'Роль', если они присутствуют в строке
                    current_naimenovanie = update_field('наименование', header_indices, row, current_naimenovanie)
                    current_role = update_field('роль', header_indices, row, current_role)

                    # Извлечение данных ВМ
                    vm_entry = extract_vm_entry(header_indices, row)

                    # Проверка наличия 'Имя сервера' и добавление в данные
                    if 'Имя сервера' in vm_entry and vm_entry['Имя сервера']:
                        existing_entry = find_existing_entry(data, current_naimenovanie, current_role)
                        if existing_entry:
                            existing_entry['ВМ'].append(vm_entry)
                        else:
                            data.append({
                                'Наименование': current_naimenovanie,
                                'Роль': current_role,
                                'ВМ': [vm_entry]
                            })
            else:
                logging.warning(f"Таблица не найдена в разделе: {section_title}")

            # Добавление раздела в результирующий список
            result.append({
                'Раздел': section_title,
                'Данные': data
            })

        # Сохранение результата в JSON-файл
        save_json(result, json_file)
        logging.info(f"Данные успешно сохранены в файле {json_file}")

    except FileNotFoundError as fnfe:
        logging.error(f"HTML-файл не найден: {fnfe}")
        raise
    except Exception as e:
        logging.exception(f"Ошибка при парсинге HTML файла: {e}")
        raise


def load_html(file_path: str) -> str:
    """
    Загружает HTML-контент из файла.

    :param file_path: Путь к HTML-файлу.
    :return: Содержимое HTML-файла.
    :raises FileNotFoundError: Если файл не найден.
    :raises Exception: Для остальных ошибок при чтении файла.
    """
    try:
        logging.info(f"Чтение HTML-файла: {file_path}")
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read()
    except FileNotFoundError as fnfe:
        logging.error(f"HTML-файл не найден: {file_path}")
        raise
    except Exception as e:
        logging.error(f"Не удалось прочитать HTML-файл {file_path}: {e}")
        raise


def get_section_title(cell: Any) -> Optional[str]:
    """
    Извлекает заголовок раздела из ячейки.

    :param cell: BeautifulSoup объект ячейки.
    :return: Заголовок раздела или None, если не найден.
    """
    h3 = cell.find('h3')
    if h3:
        return h3.get_text(strip=True)
    return None


def parse_html_table(table: Any) -> List[List[str]]:
    """
    Парсит HTML-таблицу в список строк, корректно обрабатывая rowspan и colspan.
    Возвращает список строк, где каждая строка - это список значений ячеек.

    :param table: BeautifulSoup объект таблицы.
    :return: Список строк, каждая строка - список значений ячеек.
    """
    rows = table.find_all('tr')
    table_data: List[List[str]] = []
    rowspan_dict: Dict[int, Dict[str, Any]] = {}  # Словарь для отслеживания ячеек с rowspan

    # Определение максимального количества колонок для заполнения пустыми ячейками
    num_columns = max(len(row.find_all(['td', 'th'])) for row in rows) if rows else 0

    for row in rows:
        row_cells: List[str] = []
        cells = row.find_all(['td', 'th'])
        cell_idx = 0

        while cell_idx < len(cells) or len(row_cells) < num_columns:
            # Проверка наличия незавершенных rowspan
            if len(row_cells) in rowspan_dict:
                cell_info = rowspan_dict[len(row_cells)]
                row_cells.append(cell_info['value'])
                cell_info['remaining'] -= 1
                if cell_info['remaining'] == 0:
                    del rowspan_dict[len(row_cells) - 1]
                continue

            if cell_idx >= len(cells):
                # Если ячеек больше нет, заполняем пустыми значениями
                row_cells.append('')
                continue

            cell = cells[cell_idx]
            cell_text = cell.get_text(strip=True)
            colspan = int(cell.get('colspan', 1))
            rowspan = int(cell.get('rowspan', 1))

            # Добавление значения ячейки, учитывая colspan
            for _ in range(colspan):
                row_cells.append(cell_text)

            # Обработка rowspan
            if rowspan > 1:
                for i in range(colspan):
                    rowspan_dict[len(row_cells) - colspan + i] = {'value': cell_text, 'remaining': rowspan - 1}

            cell_idx += 1

        logging.debug(f"Обработанная строка: {row_cells}")
        table_data.append(row_cells)

    return table_data


def normalize_header(header: str) -> str:
    """
    Нормализует заголовок столбца: приводит к нижнему регистру, удаляет все пробельные символы и дефисы.

    :param header: Исходный заголовок.
    :return: Нормализованный заголовок.
    """
    header = header.strip().lower().replace('-', '')
    header = re.sub(r'\s+', '', header)  # Удаляем все виды пробельных символов
    return header


def update_field(field_name: str, header_indices: Dict[str, int], row: List[str], current_value: str) -> str:
    """
    Обновляет значение поля ('Наименование' или 'Роль') на основе текущей строки таблицы.

    :param field_name: Имя поля для обновления.
    :param header_indices: Словарь индексов заголовков.
    :param row: Текущая строка таблицы.
    :param current_value: Текущее значение поля.
    :return: Обновленное значение поля.
    """
    if field_name in header_indices:
        idx = header_indices[field_name]
        if idx < len(row) and row[idx]:
            logging.debug(f"Обновление '{field_name}': {row[idx]}")
            return row[idx]
    return current_value


def extract_vm_entry(header_indices: Dict[str, int], row: List[str]) -> Dict[str, str]:
    """
    Извлекает информацию о ВМ из строки таблицы.

    :param header_indices: Словарь индексов заголовков.
    :param row: Текущая строка таблицы.
    :return: Словарь с информацией о ВМ.
    """
    vm_entry: Dict[str, str] = {}

    # Унификация ключа 'Имя сервера'
    if 'доменноеимя' in header_indices:
        idx = header_indices['доменноеимя']
        if idx < len(row):
            vm_entry['Имя сервера'] = row[idx]
    elif 'имясервера' in header_indices:
        idx = header_indices['имясервера']
        if idx < len(row):
            vm_entry['Имя сервера'] = row[idx]

    # Обработка 'IP адрес'
    ip_headers = ['ipадрес', 'ipaddress', 'ip', 'ip_адрес']
    for ip_header in ip_headers:
        if ip_header in header_indices:
            idx = header_indices[ip_header]
            if idx < len(row):
                vm_entry['IP адрес'] = row[idx]
                break
    else:
        vm_entry['IP адрес'] = ''

    # Обработка 'Сайзинг'
    sizing_headers = ['сайзинг', 'sizing']
    for sizing_header in sizing_headers:
        if sizing_header in header_indices:
            idx = header_indices[sizing_header]
            if idx < len(row):
                vm_entry['Сайзинг'] = row[idx]
                break
    else:
        vm_entry['Сайзинг'] = ''

    logging.debug(f"Извлеченная ВМ: {vm_entry}")
    return vm_entry


def find_existing_entry(data: List[Dict[str, Any]], naimenovanie: str, role: str) -> Optional[Dict[str, Any]]:
    """
    Ищет существующую запись с заданным 'Наименование' и 'Роль'.

    :param data: Список данных.
    :param naimenovanie: Текущее 'Наименование'.
    :param role: Текущая 'Роль'.
    :return: Ссылка на существующую запись или None.
    """
    return next(
        (item for item in data if item['Наименование'] == naimenovanie and item['Роль'] == role),
        None
    )


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
        logging.info(f"Данные успешно сохранены в файле {json_file}")
    except Exception as e:
        logging.error(f"Не удалось сохранить JSON файл {json_file}: {e}")
        raise
