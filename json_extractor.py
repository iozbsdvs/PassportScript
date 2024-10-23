# json_extractor.py

import json
import pandas as pd
import logging


def extract_data_from_json(json_file):
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)

        vm_list = []

        for section in data:
            section_name = section.get('Раздел', '')
            for item in section.get('Данные', []):
                naimenovanie = item.get('Наименование', '')
                role = item.get('Роль', '')
                for vm in item.get('ВМ', []):
                    vm_name = vm.get('Имя сервера', '')
                    ip_address = vm.get('IP адрес', '')
                    sizing = vm.get('Сайзинг', '')

                    vm_list.append({
                        'Раздел': section_name,
                        'Наименование': naimenovanie,
                        'Роль': role,
                        'Имя сервера': vm_name,
                        'IP адрес': ip_address,
                        'Сайзинг': sizing
                    })

        # Создаем DataFrame из списка ВМ
        df_json = pd.DataFrame(vm_list)
        return df_json

    except Exception as e:
        logging.error(f"Ошибка при обработке JSON файла: {e}")
        raise e  # Пробрасываем исключение вверх
