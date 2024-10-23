# compare_json.py

import json
import logging
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from utils import adjust_column_widths


def compare_json(json_file_1, json_file_2, output_excel_file):
    try:
        # Настраиваем логирование на максимально подробный уровень
        logging.basicConfig(level=logging.DEBUG, format='%(message)s')

        # Загружаем данные из JSON-файлов
        logging.info(f"Загружаю данные из JSON-файла 1: {json_file_1}")
        with open(json_file_1, 'r', encoding='utf-8') as f1:
            data1 = json.load(f1)

        logging.info(f"Загружаю данные из JSON-файла 2: {json_file_2}")
        with open(json_file_2, 'r', encoding='utf-8') as f2:
            data2 = json.load(f2)

        # Преобразуем данные из JSON-файла 1 (паспорт) в словарь
        logging.info("Преобразую данные из JSON-файла 1 в словарь для быстрого доступа")
        dict1 = {}
        for section in data1:
            section_name = section.get('Раздел', '')
            logging.debug(f"Раздел: {section_name}")
            for item in section.get('Данные', []):
                for vm in item.get('ВМ', []):
                    server_name = vm.get('Имя сервера', '').strip().lower()
                    if server_name:
                        dict1[server_name] = {
                            'IP адрес': vm.get('IP адрес', ''),
                            'Сайзинг': vm.get('Сайзинг', ''),
                            'Источник': f"Раздел: {section_name}"
                        }

        # Преобразуем данные из JSON-файла 2 (сайзинг) в словарь
        logging.info("Преобразую данные из JSON-файла 2 в словарь для быстрого доступа")
        dict2 = {}
        for item in data2:
            server_name = item.get('Имя сервера', '').strip().lower()
            if server_name:
                dict2[server_name] = {
                    'IP адрес': item.get('IP адрес', ''),
                    'Сайзинг': item.get('Сайзинг', ''),
                    'Источник': "Excel файл"
                }

        # Получаем множество всех серверов
        servers1 = set(dict1.keys())
        servers2 = set(dict2.keys())
        all_servers = servers1.union(servers2)
        logging.info(f"Всего серверов в паспорте: {len(servers1)}")
        logging.info(f"Всего серверов в сайзинге: {len(servers2)}")
        logging.info(f"Общее количество уникальных серверов для сравнения: {len(all_servers)}")

        # Подготавливаем данные для записи в Excel
        matched_rows = []        # Серверы, присутствующие в обоих файлах
        unmatched_rows_red = []  # Серверы отсутствующие в паспорте
        unmatched_rows_blue = [] # Серверы отсутствующие в сайзинге

        for server in sorted(all_servers):
            logging.info(f"\nСравнение сервера: {server}")
            row = {'Имя сервера': server}
            red_cells = []
            blue_cells = []
            full_row_color = None

            item1 = dict1.get(server)
            item2 = dict2.get(server)

            # Добавляем данные из паспорта
            if item1:
                row['IP адрес в паспорте'] = item1.get('IP адрес', '')
                row['Сайзинг в паспорте'] = item1.get('Сайзинг', '')
                row['Источник в паспорте'] = item1.get('Источник', '')
                logging.debug(f"  Данные из паспорта: IP адрес: {row['IP адрес в паспорте']}, "
                              f"Сайзинг: {row['Сайзинг в паспорте']}, Источник: {row['Источник в паспорте']}")
            else:
                row['IP адрес в паспорте'] = ''
                row['Сайзинг в паспорте'] = ''
                row['Источник в паспорте'] = ''
                logging.debug("  Сервер отсутствует в паспорте")

            # Добавляем данные из сайзинга
            if item2:
                row['IP адрес в сайзинге'] = item2.get('IP адрес', '')
                row['Сайзинг в сайзинге'] = item2.get('Сайзинг', '')
                logging.debug(f"  Данные из сайзинга: IP адрес: {row['IP адрес в сайзинге']}, "
                              f"Сайзинг: {row['Сайзинг в сайзинге']}, Источник: {item2.get('Источник', '')}")
            else:
                row['IP адрес в сайзинге'] = ''
                row['Сайзинг в сайзинге'] = ''
                logging.debug("  Сервер отсутствует в сайзинге")

            # Логика подсветки
            if item1 and item2:
                # Сервер есть в обоих файлах
                discrepancies = False
                # Проверка IP адресов
                ip1 = str(row['IP адрес в паспорте']).strip()
                ip2 = str(row['IP адрес в сайзинге']).strip()
                if ip1 != ip2:
                    discrepancies = True
                    red_cells.append('IP адрес в паспорте')

                # Проверка сайзинга
                sizing1 = str(row['Сайзинг в паспорте']).strip()
                sizing2 = str(row['Сайзинг в сайзинге']).strip()
                if sizing1 != sizing2:
                    discrepancies = True
                    red_cells.append('Сайзинг в паспорте')

                if discrepancies:
                    matched_rows.append({'data': row, 'red_cells': red_cells, 'blue_cells': blue_cells, 'full_row_color': None})
                else:
                    matched_rows.append({'data': row, 'red_cells': red_cells, 'blue_cells': blue_cells, 'full_row_color': None})
            else:
                # Сервер отсутствует в одном из файлов
                if not item1:
                    full_row_color = 'red'
                    logging.info("  Сервер отсутствует в паспорте")
                if not item2:
                    full_row_color = 'blue'
                    logging.info("  Сервер отсутствует в сайзинге")

                if not item1 and not item2:
                    full_row_color = 'red'  # При отсутствии в обоих, выделяем красным

                row_entry = {'data': row, 'red_cells': red_cells, 'blue_cells': blue_cells, 'full_row_color': full_row_color}

                if full_row_color == 'red':
                    unmatched_rows_red.append(row_entry)
                elif full_row_color == 'blue':
                    unmatched_rows_blue.append(row_entry)

        # Объединяем списки: сначала совпадающие, затем отсутствующие в паспорте (красным), затем отсутствующие в сайзинге (синим)
        all_rows = matched_rows + unmatched_rows_red + unmatched_rows_blue

        # Записываем результаты в Excel-файл
        wb = Workbook()
        ws = wb.active

        # Порядок колонок без 'Источник в сайзинге'
        headers = ['Имя сервера',
                   'IP адрес в сайзинге', 'IP адрес в паспорте',
                   'Сайзинг в сайзинге', 'Сайзинг в паспорте',
                   'Источник в паспорте']
        ws.append(headers)

        # Определяем цвета для подсветки
        red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        blue_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')  # Светло-синий

        for entry in all_rows:
            row_data = entry['data']
            red_cells = entry.get('red_cells', [])
            blue_cells = entry.get('blue_cells', [])
            full_row_color = entry.get('full_row_color')

            ws_row = [row_data.get(header, '') for header in headers]
            ws.append(ws_row)
            row_num = ws.max_row

            if full_row_color == 'red':
                for col in range(1, len(headers) + 1):
                    ws.cell(row=row_num, column=col).fill = red_fill
            elif full_row_color == 'blue':
                for col in range(1, len(headers) + 1):
                    ws.cell(row=row_num, column=col).fill = blue_fill
            else:
                # Для совпадающих серверов с несоответствиями, выделяем только конкретные ячейки красным
                for col_name in red_cells:
                    if col_name in headers:
                        col_idx = headers.index(col_name) + 1
                        ws.cell(row=row_num, column=col_idx).fill = red_fill

        # Настраиваем ширину колонок
        adjust_column_widths(ws)

        # Сохраняем Excel-файл
        wb.save(output_excel_file)
        logging.info(f"\nРезультаты сравнения сохранены в файле {output_excel_file}")

    except Exception as e:
        logging.error(f"Ошибка при сравнении JSON-файлов: {e}")
        raise e  # Пробрасываем исключение вверх