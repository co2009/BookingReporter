import pandas as pd
import numpy as np
from pathlib import Path
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, TypeAlias, List, Set
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl
from datetime import datetime
import argparse
import calendar

import xlrd
import re
import os

log : Set[str] = set()
warnings : Set[str] = set()

def nan_to_zero(value : float) -> float:
    return 0.0 if np.isnan(value) else value

def find_kpb(s, default_value: int) -> int:
    match = re.search(r'(\d+)КПБ', s)
    if match:
        return int(match.group(1))

    return default_value

def extract_service_percent(s):
    match = re.search(r'КОМПЛАТ=(\d+\.?\d*)%', s)  # \d+\.?\d* — одна точка или ни одной
    if not match:
        match = re.search(r'КОМПЛАТ=(\d+\.?\d*)', s)  # \d+\.?\d* — одна точка или ни одной
        if match:
            log.add(f'Нет символа % для КОМПЛАТ для "{s}"')
        
        return -1
    
    service_fee = float(match.group(1))
    if not 0 <= service_fee < 100:
        return -1
    
    return service_fee

def extract_extra_pay(s):
    match = re.search(r'ДОП(П?)ЛАТА=(\d+)', s)
    return int(match.group(2)) if match else 0

def extract_discount(s):
    match = re.search(r'СКИДКА=(\d+)', s)
    return int(match.group(1)) if match else 0

def extract_platform_name(s):
    match = re.search(r'Забронировано через\s+([^\s\n]+)', s.strip(), re.IGNORECASE)
    if match:
        return match.group(1)
    return ""    

def read_platforms_info(xlsx_fpath : Path) -> dict[str,float]:
    df = pd.read_excel(xlsx_fpath)

    platforms_info = dict[str,float]()
    for row in df.itertuples():
        platforms_info[row.Платформа] = row.Комиссия

    return platforms_info

default_clothes_cost : int = 1000

Address: TypeAlias = str
Alias: TypeAlias = str

@dataclass
class Apartment:
    owner_percent : int
    fix_price: int
    cleaning: int
    clothes: int
    comment: str
    alias: Alias

Apartments: TypeAlias = Dict[Address, Apartment]
Aliases: TypeAlias = Dict[Alias, Address]

def read_apartments(xlsx_fpath : Path) -> (Apartments, Aliases):
    df = pd.read_excel(xlsx_fpath).fillna({"Примечание": "", "Белье" : default_clothes_cost, "Уборка" : 1300, "ПроцентХозяину" : 70, "Фикс" : 0})

    apartments : Apartments = {}
    aliases : Aliases = {}
    for row in df.itertuples():
        address = Address(row.Квартира).strip()
        apartment = Apartment(int(row.ПроцентХозяину), int(row.Фикс), int(row.Уборка), int(row.Белье), str(row.Примечание), Alias(row.Псевдоним).strip())
        apartments[address] = apartment
        aliases[apartment.alias] = address

    return apartments, aliases

@dataclass
class Expense:
    cost : int # cost = full_cost / parts
    full_cost : int # сумма на несколько квартир
    parts : int     # количество квартир, на которые бьется full_cost
    category : str
    date : str
    comment : str

@dataclass
class ApartmentExpenses:
    expenses : List[Expense] = field(default_factory=list)
    total_cost : int = 0

    # def __post_init__(self):
    #     if self.expenses is None:  # Инициализируем, если None
    #         self.expenses = []     # Теперь у каждого экземпляра свой список!    

Expenses : TypeAlias = Dict[Alias,ApartmentExpenses]

@dataclass
class ExpenseAccounts:
    vera_land: Expenses = field(default_factory=lambda: defaultdict(ApartmentExpenses))
    owner: Expenses = field(default_factory=lambda: defaultdict(ApartmentExpenses))

def is_fix_pay_to_veraland(apartment : Apartment):
    return apartment.owner_percent >= 100

def is_fix_pay_to_owner(apartment : Apartment):
    return apartment.owner_percent <= 0

def extend_expenses(expenses : Expenses, func, apartments: Apartments, aliases : Aliases) -> ExpenseAccounts:
    for alias, apartment_expenses in expenses.items():
        address = aliases.get(alias, None)
        apartment = apartments.get(address)
        if func(apartment) and apartment.fix_price > 0:
            expense = Expense(
                cost = apartment.fix_price,
                full_cost = apartment.fix_price,
                parts = 1,
                category = "Расходники",
                date = None,
                comment="VeraLand",
            )
            apartment_expenses.expenses.append(expense)
            apartment_expenses.total_cost += expense.cost


def read_expenses(xlsx_fpath : Path, global_apartments: Apartments, global_aliases : Aliases) -> ExpenseAccounts:
    df = pd.read_excel(xlsx_fpath).fillna({"Комментарий": "", "Квартира": ""})
    expense_accounts = ExpenseAccounts()

    categories_owner = {"Стартовое вложение", "КУ"}
    categories_vera_land = {"Расходники","Прочие расходы агенства","Уборка коридора","Продвижение"}

    for _, row in df.iterrows():
        apartments = [apt.strip() for apt in row.Квартира.split(";") if apt.strip()]
        
        # Если квартиры не указаны, используем пустую строку
        if not apartments:
            log.add(f'Расходы: Нет квартиры для {", ".join(f"{k}:{v}" for k, v in row.items())}')
            continue

        # Берем количество квартир и полную стоимость 
        apartments_count = len(apartments)
        full_cost = row.Сумма

        # Проверяем категорию на предмет и выбираем, на кого писать расход
        expenses = None
        category : str = row.Статья
        if category in categories_owner:
            expenses = expense_accounts.owner
        elif category in categories_vera_land:
            expenses = expense_accounts.vera_land
        else :
            continue

        comment = row.Комментарий
        if not comment and category == "Уборка коридора":
            comment = category

        # Создаем объект Expense для каждой квартиры
        expense = Expense(
            cost = full_cost // apartments_count,
            full_cost = full_cost,
            parts = apartments_count,
            category = category,
            date = row.Дата,
            comment=comment,
        )

        if not expense.comment:
            log.add(f'Нет цели расхода для {", ".join(f"{k}:{v}" for k, v in row.items())}')

        # Добавляем запись для каждой квартиры
        for apartment in apartments:
            if not apartment in global_aliases:
                log.add(f'Нет квартиры {apartment} для {", ".join(f"{k}:{v}" for k, v in row.items())}')
                continue

            expenses[apartment].expenses.append(expense)
            expenses[apartment].total_cost += expense.cost

    extend_expenses(expense_accounts.owner, is_fix_pay_to_veraland, global_apartments, global_aliases)
    extend_expenses(expense_accounts.vera_land, is_fix_pay_to_owner, global_apartments, global_aliases)

    return expense_accounts

def calculate_nights(row):
    try:
        check_in = datetime.strptime(row['Заезд'], '%d.%m.%Y')
        check_out = datetime.strptime(row['Выезд'], '%d.%m.%Y')
        return (check_out - check_in).days
    except:
        return None  # В случае ошибки в формате даты

@dataclass
class KpbAndCleaningCounts:
    kpbs : int = 0
    cleanings: int = 0


# Функция для подсчета КПБ и уборок с суммами
def count_kpbs_with_costs(comments_series, kpb_and_cleaning_counts : KpbAndCleaningCounts):
    kpb_count = 0
    kpb_total_cost = 0
    
    # Считаем КПБ
    for comment in comments_series:
        kpb_matches = re.findall(r'(\d+)КПБ\((\d+)\)', str(comment))
        for count, cost in kpb_matches:
            kpb_count += int(count)
            kpb_total_cost += int(cost)
    kpb_and_cleaning_counts.kpbs = kpb_count
    return kpb_total_cost, f"КПБ: {kpb_count}"

# Функция для подсчета КПБ и уборок с суммами
def count_cleanings_with_costs(comments_series, kpb_and_cleaning_counts : KpbAndCleaningCounts):
    cleaning_count = 0
    cleaning_total_cost = 0
    
    # Считаем уборки
    for comment in comments_series:
        cleaning_matches = re.findall(r'Уборка\((\d+)\)', str(comment))
        for cost in cleaning_matches:
            cleaning_count += 1
            cleaning_total_cost += int(cost)
    
    kpb_and_cleaning_counts.cleanings = cleaning_count
    return cleaning_total_cost, f"Уборок: {cleaning_count}"

def print_expense(title : str, expenses : Expenses, apartment_alias : Alias):
    print(title,end=": ")
    apartment_expenses : ApartmentExpenses = expenses.get(apartment_alias,ApartmentExpenses())
    print(apartment_expenses.total_cost,end=": ")
    print("; ".join(f"{expense.comment} ({expense.cost})" for expense in apartment_expenses.expenses),end=": ")
    print()

def add_expense(df_main : pd.DataFrame, expense_column_tag : str, expenses_rows, cost : int, comment : str):
        row_dict = {col: '' for col in df_main.columns}
                
        # Заполняем данные конкретного расхода
        row_dict.update({expense_column_tag: cost,'Комментарии': comment,})
                
        expenses_rows.append(row_dict)

def make_expenses_df(
        address : Address, 
        expense_column_tag : str,
        add_kpb_n_cleaning : bool,
        kpb_and_cleaning_counts : KpbAndCleaningCounts,
        df_main : pd.DataFrame, 
        apartment_expenses : ApartmentExpenses) -> pd.DataFrame:

    expenses_rows = []

    if add_kpb_n_cleaning:
        add_expense( df_main, expense_column_tag, expenses_rows, *count_kpbs_with_costs(df_main['Комментарии'], kpb_and_cleaning_counts))
        add_expense( df_main, expense_column_tag, expenses_rows, *count_cleanings_with_costs(df_main['Комментарии'], kpb_and_cleaning_counts))

    for expense in apartment_expenses.expenses:
        add_expense( df_main, expense_column_tag, expenses_rows, expense.cost, expense.comment)

    if expenses_rows:
        return pd.DataFrame(expenses_rows)[df_main.columns]
    return pd.DataFrame(columns=df_main.columns)
    

def summary_to_pivot(
        summary_data, apartment_alias : str, owner_tag : str, vera_land_tag : str, bookings_count : int, 
        kpb_and_cleaning_counts  : KpbAndCleaningCounts, days_in_month : int, pivot_data):
    pivot_dict = {}
    pivot_dict["Адрес"] = apartment_alias
    pivot_dict["Заездов"] = bookings_count
    pivot_dict["От гостя"] = summary_data["От гостя"]
    pivot_dict["Платформе"] = summary_data["Платформе"]
    pivot_dict["Доп плата"] = summary_data["Доп плата"]
    pivot_dict["Выручка"] = summary_data["Выручка"]
    pivot_dict["VeraLand"] = summary_data[vera_land_tag]
    pivot_dict["Расход VL"] = summary_data["Расход VL"]
    pivot_dict["Итог VL"] = summary_data[vera_land_tag] - summary_data["Расход VL"]
    pivot_dict["Собственнику"] = summary_data[owner_tag]
    pivot_dict["Расход"] = summary_data["Расход"]
    pivot_dict["КПБ"] = kpb_and_cleaning_counts.kpbs
    pivot_dict["Уборок"] = kpb_and_cleaning_counts.cleanings
    pivot_dict["Ночей"] = int(summary_data["Ночей"])
    pivot_dict["Занято(%)"] = int(summary_data["Ночей"] / days_in_month * 100.0)
    pivot_data.append(pivot_dict)

def save_to_excel(xlsx_path, result_df):
    with pd.ExcelWriter(xlsx_path, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        result_df.to_excel(writer, sheet_name="Отчет", index=False)
        book = writer.book
        sheet = writer.sheets["Отчет"]

        # Определяем формат для последней строки
        bold_format = book.add_format({
            'bold': True,
            'font_size': 12,  # Увеличиваем размер шрифта
        })
        
        # Применяем формат к последней строке
        last_row = len(result_df)  # Нумерация строк в Excel начинается с 1
        for col_num, value in enumerate(result_df.columns):
            value = result_df.iloc[-1, col_num]
            write_value = value if not pd.isna(value) else ''
            sheet.write(last_row, col_num, write_value, bold_format)

        for idx, col in enumerate(result_df.columns):
            max_len = max(result_df[col].astype(str).map(len).max(), len(col))
            sheet.set_column(idx, idx, max_len + 2)


def make_reports(fname_suffix: str, days_in_month : int):
    print("\n")

    platforms_info = read_platforms_info('config/platforms.xlsx')
    print(platforms_info)

    apartments, aliases = read_apartments('config/apartments.xlsx')
    print(*apartments.items(),sep="\n")
    print(aliases.items())

    expense_accounts = read_expenses(f'xls/expenses_{fname_suffix}.xlsx', apartments, aliases)
    print("Расходы собственников: ", *expense_accounts.owner.items(), sep="\n")

    xls_path = Path(__file__).absolute().parent / f'xls/bookings_{fname_suffix}.xlsx'
    print("\nFile Path:", xls_path)

    df_whole = pd.read_excel(xls_path).fillna({"Источник": ""})

    print(df_whole)

    reports_dict = defaultdict(list)

    for _, row in df_whole.iterrows():
        address : Address = row.Объект
        apartment : Apartment = apartments.get(address)

        if not apartment:
            log.add(f"Нет апартамента для {address}")
            continue

        raw_comment = (str)(row.Примечания)
        comment = raw_comment.upper()

        total_pay = row.Сумма

        kpb_cost = apartment.clothes if apartment else default_clothes_cost
        kpb_count = find_kpb(comment, 1 if total_pay > 0 else 0)

        cleaning = 0 if ("УБОХОЗ" in comment) else apartment.cleaning if apartment else 1300

        expense = kpb_count * kpb_cost + cleaning 
        report_comment = ""
        if kpb_count > 0:
            report_comment += f"{kpb_count}КПБ({kpb_count * kpb_cost})"
        if cleaning > 0:
            if len(report_comment):
                report_comment += ","
            report_comment += f"Уборка({cleaning})"

        platform : str = row.Источник
        if not platform:
            platform = extract_platform_name(comment)

        if not platform or platform == 'manual':
            if row.Менеджер == "bookings_widget@tutt.ru":
                platform = "Модуль бронирования"

        if not platform:
            log.add(f'Нет платформы для {", ".join(f"{k}:{v}" for k, v in row.items())}')

        service_percent = extract_service_percent(comment)

        if service_percent < 0.0 and platform in platforms_info:
            service_percent = platforms_info[platform]

        if service_percent < 0.0:
            service_percent = 15.0

        if total_pay <= 0:
            if kpb_count > 0:
                warnings.add(f'Расходы без доходов для {", ".join(f"{k}:{v}" for k, v in row.items())}')
            else:
                continue

        extra_pay : int = extract_extra_pay(comment)
        extra_pay -= extract_discount(comment)

        service_fee = row.Сумма * service_percent / 100.0
        net_pay = total_pay - service_fee + extra_pay
        
        owner_percent = apartment.owner_percent if apartment else 70
        owner_title = f"Собственнику {owner_percent}%"
        vera_land_title = f"VeraLand {100 - owner_percent}%"

        report_row = {
            "Источник брони": platform,
            "От гостя": int(total_pay),
            "Платформе": int(service_fee),
            "%": service_percent,
            "Доп плата": int(extra_pay),
            "Выручка": int(net_pay),
            vera_land_title: int(net_pay * (100.0 - owner_percent) / 100.0),
            "Расход VL": int(expense) if is_fix_pay_to_owner(apartment) else 0,
            owner_title: int(net_pay * owner_percent / 100.0),
            "Расход": 0 if is_fix_pay_to_owner(apartment) else int(expense),
            "Комментарии": report_comment,
            "Заезд": row.Заезд,
            "Выезд": row.Выезд,
            "Ночей": calculate_nights(row),
        }
        reports_dict[address].append(report_row)

    # Создаем DataFrame для каждого объекта
    dataframes = {k: pd.DataFrame(v) for k, v in reports_dict.items()}

    out_folder  = Path(f"./output/{fname_suffix}")
    out_folder.mkdir(parents=True, exist_ok=True)

    pivot_data = []

    print("Отчеты по объектам:")
    for address, df_obj in dataframes.items():
        apartment : Apartment = apartments.get(address)
        owner_tag = f"Собственнику {apartment.owner_percent}%"
        vera_land_tag = f"VeraLand {100 - apartment.owner_percent}%"
        fix_pay_to_owner : bool = is_fix_pay_to_owner(apartment)
        fix_pay_to_veraland : bool = is_fix_pay_to_veraland(apartment)

        # Создаем DataFrame с одной строкой
        empty_row = pd.DataFrame([[np.nan]*len(df_obj.columns)], columns=df_obj.columns)

        # Создаем DataFrame с расходами
        temp_owner_expense_column_tag = "Расход" if not fix_pay_to_owner else "Расход VL"
        temp_vera_land_expense_column_tag = "Расход VL" if not fix_pay_to_veraland else "Расход"

        kpb_and_cleaning_counts = KpbAndCleaningCounts()
        expense_of_owner_df = make_expenses_df(
            address, temp_owner_expense_column_tag, True, kpb_and_cleaning_counts, df_obj, expense_accounts.owner.get(apartment.alias,ApartmentExpenses()))
        expense_of_vera_land_df = make_expenses_df(
            address, temp_vera_land_expense_column_tag, False, kpb_and_cleaning_counts, df_obj, expense_accounts.vera_land.get(apartment.alias,ApartmentExpenses()))

        expense_of_owner_sum_raw = nan_to_zero(expense_of_owner_df[temp_owner_expense_column_tag].sum())
        expense_of_vera_land_sum_raw = nan_to_zero(expense_of_vera_land_df[temp_vera_land_expense_column_tag].sum())
        total_expense_sum_raw = expense_of_owner_sum_raw + expense_of_vera_land_sum_raw
        expense_of_owner_sum =  0 if fix_pay_to_owner else total_expense_sum_raw if fix_pay_to_veraland else expense_of_owner_sum_raw
        expense_of_vera_land_sum =  0 if fix_pay_to_veraland else total_expense_sum_raw if fix_pay_to_owner else expense_of_vera_land_sum_raw

        # Создаем итоговую строку с суммами
        summary_data = {
            'Источник брони': 'ИТОГО',
            **df_obj.select_dtypes(include='number').sum().to_dict(),
            'Расход': expense_of_owner_sum,
            'Расход VL': expense_of_vera_land_sum,
            '%': np.nan,
            **{col: '' for col in df_obj.select_dtypes(exclude='number').columns if col not in ['Источник брони']}
        }
        
        # Создаем итоговую строку расходов собственника
        summary_owner_expenses_data = {
            **{col: '' for col in df_obj.columns},
            temp_owner_expense_column_tag : expense_of_owner_sum_raw,
            'Комментарии': "Итого"
        }

        summary_owner_expenses_df = pd.DataFrame([summary_owner_expenses_data])[df_obj.columns]

        # Создаем итоговую строку расходов vera_land
        summary_vera_land_expenses_data = {
            **{col: '' for col in df_obj.columns},
            temp_vera_land_expense_column_tag : expense_of_vera_land_sum_raw,
            'Комментарии': "Итого"
        }

        summary_vera_land_expenses_df = pd.DataFrame([summary_vera_land_expenses_data])[df_obj.columns]

        if is_fix_pay_to_veraland(apartment):
            summary_data[vera_land_tag] = apartment.fix_price

        if is_fix_pay_to_owner(apartment):
            summary_data[owner_tag] = apartment.fix_price

        summary_df = pd.DataFrame([summary_data])[df_obj.columns]

        summary_to_pivot(summary_data, apartment.alias, owner_tag, vera_land_tag, len(df_obj), kpb_and_cleaning_counts, days_in_month, pivot_data)

        # Создаем совсем итоговую строку с суммой на руки
        final_data = {
            **{col: np.nan for col in df_obj.select_dtypes(include='number').columns},
            **{col: "" for col in df_obj.select_dtypes(exclude='number').columns if col not in ['Источник брони']},
            'Источник брони': 'ПЕРЕВЕСТИ',
            owner_tag : summary_data[owner_tag] - summary_data['Расход'],
            vera_land_tag : summary_data[vera_land_tag] - summary_data['Расход VL'],
        }

        if is_fix_pay_to_veraland(apartment):
            final_data[vera_land_tag] = apartment.fix_price

        final_df = pd.DataFrame([final_data])[df_obj.columns]

        # Приводим числовые колонки к целым числам
        result_df = pd.concat([
            df_obj, empty_row, 
            expense_of_owner_df, summary_owner_expenses_df, empty_row, 
            expense_of_vera_land_df, summary_vera_land_expenses_df, empty_row, 
            summary_df, empty_row, 
            final_df], ignore_index=True)

        print(f"\nОбъект: {address}", end = "")
        if apartment.comment:
            print(f"  ({apartment.comment})", end = "")
        if apartment.fix_price:
            print(f", Фикс: {apartment.fix_price}", end = "")
        print()
        print(result_df)
        print("Дополнительные расходы:")
        print_expense("VeraLand", expense_accounts.vera_land, apartment.alias)
        print_expense("Собственник", expense_accounts.owner, apartment.alias)

        xlsx_name = apartment.alias + "_" + fname_suffix + ".xlsx"
        xlsx_path : Path = out_folder / xlsx_name
        print("Отчет: ", xlsx_path)

        save_to_excel(xlsx_path, result_df)

    if pivot_data:

        pivot_df = pd.DataFrame(pivot_data)

        summary_pivot_data = [{
            'Адрес': 'ИТОГО',
            **pivot_df.select_dtypes(include='number').sum().to_dict(),
            'Занято(%)': pivot_df['Занято(%)'].mean(),
            **{col: '' for col in pivot_df.select_dtypes(exclude='number').columns if col not in ['Адрес']}
        }]

        summary_pivot_df = pd.DataFrame(summary_pivot_data, columns = pivot_df.columns)
        empty_pivot_df = pd.DataFrame([['']*len(pivot_df.columns)], columns=pivot_df.columns)
        final_pivot_df = pd.concat([pivot_df, empty_pivot_df, summary_pivot_df], ignore_index=True)

    print("\nСводная таблица:\n", final_pivot_df)

    xlsx_pivot_name = "_" + fname_suffix + ".xlsx"
    xlsx_pivot_path : Path = out_folder / xlsx_pivot_name
    print("Сводный отчет: ", xlsx_pivot_path)
    save_to_excel(xlsx_pivot_path, final_pivot_df)


def parse_args():
    parser = argparse.ArgumentParser(description='Программа с параметром даты в формате yyyy_mm')
    parser.add_argument(
        'date',
        type=str,
        help='Дата в формате yyyy_mm (например: 2024_10)'
    )
    args = parser.parse_args()
    
    # Проверка формата даты
    try:
        year, month = map(int, args.date.split('_'))
        datetime(year=year, month=month, day=1)  # Проверяем, что дата валидна
    except (ValueError, IndexError):
        parser.error("Неверный формат даты. Используйте yyyy_mm (например: 2024_10)")
    
    return args

if __name__ == "__main__":

    args = parse_args()
    yyyy_mm : str = args.date
    year, month = map(int, args.date.split('_'))
    days_in_month = calendar.monthrange(year, month)[1]

    make_reports(yyyy_mm, days_in_month)

    if warnings:
        print("\033[33m")
        print(*warnings, sep="\n")
        print("\033[0m")

    if log:
        print("\033[31m")
        print(*log, sep="\n")
        print("\033[0m")

