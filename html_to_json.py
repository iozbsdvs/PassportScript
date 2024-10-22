import json
from bs4 import BeautifulSoup
import logging


def parse_html_to_json(html_file, json_file):
    try:
        # Читаем HTML-код из файла
        with open(html_file, 'r', encoding='utf-8') as file:
            html_content = file.read()

        # Парсим HTML-код
        soup = BeautifulSoup(html_content, 'html.parser')

        # Ищем все <div class="innerCell">
        inner_cells = soup.find_all('div', class_='innerCell')

        # Имена заголовков разделов, которые нам нужны
        target_titles = [
            'Система управления базами данных Pangolin SE',
            'Список виртуальных серверов Ведомства',
            'Список виртуальных серверов НСУД'
        ]

        # Результирующий список
        result = []

        for cell in inner_cells:
            # Находим заголовок h3 внутри текущей ячейки
            h3 = cell.find('h3')
            if h3 and h3.get_text(strip=True) in target_titles:
                section_title = h3.get_text(strip=True)
                # Ищем таблицу внутри этой ячейки
                table = cell.find('table')
                if table:
                    # Извлекаем заголовки столбцов
                    header_row = table.find('tr')
                    if not header_row:
                        continue  # Пропускаем, если нет строк в таблице
                    header_cells = header_row.find_all(['td', 'th'])
                    headers = [hc.get_text(strip=True) for hc in header_cells]
                    # Создаем словарь соответствия названий столбцов и их индексов
                    header_indices = {header: idx for idx, header in enumerate(headers)}
                    # Список записей
                    data = []
                    rows = table.find_all('tr')[1:]  # Пропускаем строку с заголовками
                    for row in rows:
                        cells = row.find_all('td')
                        if len(cells) < len(headers):
                            continue  # Пропускаем неполные строки
                        # Инициализируем словарь для текущей строки
                        row_data = {}
                        for header, idx in header_indices.items():
                            cell_text = cells[idx].get_text(separator='\n', strip=True)
                            row_data[header] = cell_text
                        # Обрабатываем данные в зависимости от наличия нужных столбцов
                        vm_list = []
                        if 'Доменное имя' in header_indices and 'IP адрес' in header_indices:
                            # Разбиваем доменные имена и IP адреса на списки
                            domain_list = row_data['Доменное имя'].split('\n')
                            ip_list = row_data['IP адрес'].split('\n')
                            # Сопоставляем доменные имена и IP адреса
                            for domain, ip in zip(domain_list, ip_list):
                                vm_list.append({
                                    'Доменное имя': domain,
                                    'IP адрес': ip
                                })
                        elif 'Имя сервера' in header_indices and 'IP адрес' in header_indices:
                            # Обрабатываем таблицу "Список виртуальных серверов Ведомства"
                            vm_list.append({
                                'Имя сервера': row_data['Имя сервера'],
                                'IP адрес': row_data['IP адрес'],
                                'Сайзинг': row_data.get('Сайзинг', row_data.get('/cpu/ram/hdd sys/hdd app/', ''))
                            })
                        else:
                            continue  # Пропускаем строки без необходимых данных
                        # Добавляем в список данные о сервере
                        data.append({
                            'Наименование': row_data.get('Наименование', row_data.get('Назначение сервера', '')),
                            'Роль': row_data.get('Роль', ''),
                            'ВМ': vm_list
                        })
                else:
                    continue  # Пропускаем, если таблица не найдена
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
