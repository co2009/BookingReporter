import argparse
from datetime import datetime

def parse_args():
    parser = argparse.ArgumentParser(description='Программа с параметром даты в формате yyyy_mm')
    parser.add_argument(
        'date',
        type=str,
        help='Дата в формате yyyy_mm (например: 2024_10)'
    )
    parser.add_argument(
        '--suffix',
        type=str,
        default='',
        help='Суффикс для имени файла (будет добавлен через "_")'
    )    
    args = parser.parse_args()
    
    # Проверка формата даты
    try:
        year, month = map(int, args.date.split('_'))
        datetime(year=year, month=month, day=1)  # Проверяем, что дата валидна
    except (ValueError, IndexError):
        parser.error("Неверный формат даты. Используйте yyyy_mm (например: 2024_10)")
    
    return args

def name_plus_suffix( name: str, suffix : str) -> str :
    return name +  (f'_{suffix}' if suffix else '')