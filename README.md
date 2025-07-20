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

```
├── config/
│ ├── apartments.xlsx # Информация об апартаментах
│ └── platforms.xlsx # Данные о платформах бронирования
├── xls/
│ ├── bookings_YYYY_MM.xlsx # Данные бронирований
│ └── expenses_YYYY_MM.xlsx # Данные расходов
├── output/ # Генерируемые отчеты
├── booking.py # Основной скрипт
└── xlsx-copy.py # Копирование xlsx-файлов в рабочую папку (xls)
```

## Запуск
```bash
python3 xlsx-copy.py YYYY_MM
python3 booking.py YYYY_MM
```

Например:
```bash
python3 xlsx-copy.py 2025_06
python3 booking.py 2025_06
```

xlsx-copy копирует выгруженные с RealtyCalendar файлы bookings.xls и expenses.xlsx в папку xls/. При этом к имени файла добавляется суффикс _YYYY_MM. Файл bookongs.xls при копирование конвертируется в xlsx-формат. Результатом являются два файла в папке xls/ (например, bookings_2025_06.xlsx и expenses_2025_06.xlsx). Исходные файлы удаляются.
