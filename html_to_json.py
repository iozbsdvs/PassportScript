import json
import re
from bs4 import BeautifulSoup
import logging


def parse_html_to_json(html_file, json_file):
    try:
        # Настраиваем логирование
        logging.basicConfig(level=logging.DEBUG, format='%(message)s')

        # Читаем HTML-код из файла
        with open(html_file, 'r', encoding='utf-8') as file:
            html_content = file.read()

        # Парсим HTML-код
        soup = BeautifulSoup(html_content, 'html.parser')

        # Ищем все <div class='innerCell'>
        inner_cells = soup.find_all('div', class_='innerCell')

        # Результирующий список
        result = []

        for cell in inner_cells:
            # Находим заголовок h3 внутри текущей ячейки
            h3 = cell.find('h3')
            if h3:
                section_title = h3.get_text(strip=True)
                logging.info(f"Обработка раздела: {section_title}")

                data = []  # Инициализируем data перед проверкой наличия таблицы

                # Ищем таблицу внутри этой ячейки
                table = cell.find('table')
                if table:
                    # Преобразуем таблицу в матрицу данных
                    table_data = parse_html_table(table)

                    # Извлекаем заголовки столбцов и нормализуем их
                    headers = table_data[0]
                    normalized_headers = [normalize_header(header) for header in headers]
                    logging.debug(f"Исходные заголовки: {headers}")
                    logging.debug(f"Нормализованные заголовки: {normalized_headers}")
                    header_indices = {header: idx for idx, header in enumerate(normalized_headers)}

                    current_naimenovanie = ''
                    current_role = ''

                    for row in table_data[1:]:
                        # Обновляем 'Наименование' и 'Роль', если они присутствуют в строке
                        if 'наименование' in header_indices:
                            idx = header_indices['наименование']
                            if row[idx]:
                                current_naimenovanie = row[idx]
                        if 'роль' in header_indices:
                            idx = header_indices['роль']
                            if row[idx]:
                                current_role = row[idx]

                        # Собираем данные ВМ
                        vm_entry = {}

                        # Унифицируем ключ 'Имя сервера'
                        if 'доменноеимя' in header_indices:
                            idx = header_indices['доменноеимя']
                            vm_entry['Имя сервера'] = row[idx]
                        elif 'имясервера' in header_indices:
                            idx = header_indices['имясервера']
                            vm_entry['Имя сервера'] = row[idx]

                        # Обрабатываем 'IP адрес'
                        ip_headers = ['ipадрес', 'ipaddress', 'ip', 'ip_адрес']
                        for ip_header in ip_headers:
                            if ip_header in header_indices:
                                idx = header_indices[ip_header]
                                vm_entry['IP адрес'] = row[idx]
                                break  # Прекращаем поиск после первого совпадения

                        # Обрабатываем 'Сайзинг'
                        sizing_headers = ['сайзинг', 'sizing']
                        for sizing_header in sizing_headers:
                            if sizing_header in header_indices:
                                idx = header_indices[sizing_header]
                                vm_entry['Сайзинг'] = row[idx]
                                break

                        # Проверяем, что vm_entry содержит 'Имя сервера'
                        if 'Имя сервера' in vm_entry and vm_entry['Имя сервера']:
                            # Проверяем, есть ли уже запись с текущими 'Наименование' и 'Роль'
                            existing_entry = next(
                                (item for item in data if
                                 item['Наименование'] == current_naimenovanie and item['Роль'] == current_role),
                                None
                            )
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

                # Добавляем в результирующий список
                result.append({
                    'Раздел': section_title,
                    'Данные': data
                })

        # Преобразуем результат в JSON
        json_result = json.dumps(result, ensure_ascii=False, indent=4)

        # Сохраняем JSON в файл
        with open(json_file, 'w', encoding='utf-8') as f:
            f.write(json_result)

        logging.info(f"Данные успешно сохранены в файле {json_file}")

    except Exception as e:
        logging.error(f"Ошибка при парсинге HTML файла: {e}")
        raise e  # Пробрасываем исключение вверх


def parse_html_table(table):
    """
    Парсит HTML-таблицу в список строк, корректно обрабатывая rowspan и colspan.
    Возвращает список строк, где каждая строка - это список значений ячеек.
    """
    rows = table.find_all('tr')
    table_data = []
    rowspan_dict = {}  # Словарь для отслеживания ячеек с rowspan

    for row in rows:
        row_cells = []
        cells = row.find_all(['td', 'th'])
        col_index = 0  # Индекс текущей колонки в строке

        # Индекс ячейки в визуальном представлении строки
        cell_idx = 0

        while cell_idx < len(cells):
            cell = cells[cell_idx]
            # Проверяем, есть ли незавершенные rowspan с предыдущих строк
            while len(row_cells) in rowspan_dict:
                rowspan_cell = rowspan_dict[len(row_cells)]
                row_cells.append(rowspan_cell['value'])
                rowspan_cell['remaining'] -= 1
                if rowspan_cell['remaining'] == 0:
                    del rowspan_dict[len(row_cells) - 1]

            # Получаем текст ячейки
            cell_text = cell.get_text(strip=True)
            colspan = int(cell.get('colspan', 1))
            rowspan = int(cell.get('rowspan', 1))

            # Добавляем значение ячейки в текущую строку, учитывая colspan
            for _ in range(colspan):
                row_cells.append(cell_text)

            # Если rowspan > 1, сохраняем ячейку для заполнения в следующих строках
            if rowspan > 1:
                for i in range(colspan):
                    rowspan_dict[len(row_cells) - colspan + i] = {'value': cell_text, 'remaining': rowspan - 1}

            cell_idx += 1

        # Заполняем оставшиеся ячейки строки из rowspan_dict или пустыми значениями
        num_columns = max(len(row.find_all(['td', 'th'])) for row in rows)
        while len(row_cells) < num_columns:
            if len(row_cells) in rowspan_dict:
                rowspan_cell = rowspan_dict[len(row_cells)]
                row_cells.append(rowspan_cell['value'])
                rowspan_cell['remaining'] -= 1
                if rowspan_cell['remaining'] == 0:
                    del rowspan_dict[len(row_cells) - 1]
            else:
                row_cells.append('')  # Пустая ячейка

        logging.debug(f"Обработанная строка: {row_cells}")
        table_data.append(row_cells)

    return table_data


def normalize_header(header):
    """
    Нормализует заголовок столбца: приводит к нижнему регистру, удаляет все пробельные символы и дефисы.
    """
    header = header.strip().lower().replace('-', '')
    header = re.sub(r'\s+', '', header)  # Удаляем все виды пробельных символов
    return header