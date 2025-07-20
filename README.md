# Booking Management System

## Описание
Скрипт `booking.py` предназначен для формирования отчетов о бронировании

## Установка
1. Убедитесь, что у вас установлен Python 3.10 или новее
2. Установите необходимые зависимости:
```bash
pip install pandas numpy openpyxl xlrd xlsxwriter
```

## Структура проекта
- config/
   - apartments.xlsx     # Информация об апартаментах
   - platforms.xlsx     # Данные о платформах бронирования
- xls/
   - bookings_YYYY_MM.xlsx  # Данные бронирований
   - expenses_YYYY_MM.xlsx  # Данные расходов
- output/                # Генерируемые отчеты
- booking.py             # Основной скрипт

## Запуск
```bash
python3 booking.py YYYY_MM
```

Например:
```bash
python3 booking.py 2025_06
```

Перед запуском необходимо скопировать в папку xls файлы bookings_YYYY_MM.xlsx (например, bookings_2025_06.xlsx) и expenses_YYYY_MM.xlsx (например, expenses_2025_06.xlsx).
