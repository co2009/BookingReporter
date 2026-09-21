import utilities

import os
from pathlib import Path
import pandas as pd
from bs4 import BeautifulSoup

def ask_yes_no(question: str, default: bool = True) -> bool:
    choices = " [Y/n] " if default else " [y/N] "
    prompt = question + choices
    
    while True:
        answer = input(prompt).strip().lower()
        
        if not answer:
            return default
        if answer in ('y', 'yes'):
            return True
        if answer in ('n', 'no'):
            return False
        
        print("Please input 'y' or 'n'")
        
def get_downloads_path() -> Path: 
    return os.path.join(os.path.expanduser("~"), "Downloads")


def check_path(fpath : Path) -> bool:
    return True

def parse_excel_xml(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'lxml-xml')
    
    data = []
    for row in soup.find_all('Row'):
        cells = [cell.Data.text if cell.Data else '' for cell in row.find_all('Cell')]
        data.append(cells)
    
    return pd.DataFrame(data[1:], columns=data[0]) if data else pd.DataFrame()

def save_to_excel(xlsx_path, df):
    with pd.ExcelWriter(xlsx_path, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        df.to_excel(writer, sheet_name="Отчет", index=False)

def copy_concrete_file(fpath : Path, ext : Path, dest_fname : Path, dest_suffix : str) -> bool:
    full_path : Path =  fpath.with_suffix(ext)
    print(f"Try to copy {full_path}")
    
    if not full_path.exists():
        print(f"{full_path} does not exist")
        return False

    dest_path = Path(f'xls/{dest_fname}_{dest_suffix}.xlsx').resolve()
    if dest_path.exists():
        if not ask_yes_no(f"{dest_path} already exists. Overwrite?"):
            return False
           
    if( ext == ".xls"):
        df = parse_excel_xml(full_path)
        print(f"Total: {len(df)} records")
        save_to_excel(dest_path, df)
        os.remove(full_path)
    else:
        os.replace(full_path, dest_path)
        
    print(f"file moved from {full_path} to {dest_path}")
    return True

def copy_file(fname : Path, ext : Path, dest_fname : Path, dest_suffix : str) -> bool:
    path : Path = get_downloads_path() / Path(fname)
    
    if copy_concrete_file( path, ext, dest_fname, dest_suffix):
        return False
    
    return True

if __name__ == "__main__":
    args = utilities.parse_args()
    yyyy_mm : str = args.date
    suffix : str = args.suffix
    
    copy_file("expenses", ".xlsx", "expenses", utilities.name_plus_suffix(yyyy_mm, suffix))
    copy_file("bookings-report-21-09-2026", ".xlsx", "bookings", utilities.name_plus_suffix(yyyy_mm, suffix))
