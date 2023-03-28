from tkinter import ttk, filedialog
from tkinter.messagebox import showerror
from tkinter import *
from pathlib import Path
import json
import os
import tkinter as tk
import pandas as pd

def load_json() -> dict | None:
    filepath = filedialog.askopenfilenames(defaultextension='json')
    json_dict = {}
    for item in filepath:
        if Path(item).suffix != '.json':
            showerror(title='Ошибка', message='Файл должен быть с разрешением .json', detail=f'{os.path.basename(item)}')
            return None
    for item in filepath:
        with open(item, encoding='utf-8') as fp:
            json_dict.update(json.load(fp))
    return json_dict

def save_json(json_dict: dict):
    with open('test_new_json.json', 'w', encoding='utf-8') as fp:
        json.dump(json_dict, fp, indent=4, ensure_ascii=False)

def get_names(json_dict: dict) -> list:
    return list(json_dict.keys())

def write_file(col, data):
    filepath = os.getcwd() + f'\Отчет.xlsx'
    df1 = pd.DataFrame(data, columns=col)
    df1.to_excel(filepath, index=False)

def main_window():

    json_dict = {}
    sorted_json_dict = {}
    names = []
    selected = ()
    selected_names = []


    def load_json_save():
        nonlocal json_dict
        nonlocal names
        json_dict = load_json()
        names = get_names(json_dict)
        listbox_name.insert(END, *names)
        
    
    def save_json_save():
        save_json(json_dict)

    
    def listbox_name_selected(event):
        nonlocal selected
        nonlocal selected_names
        selected = listbox_name.curselection()
        print([listbox_name.get(i) for i in selected])
        print(selected)
        selected_names = [listbox_name.get(i) for i in selected]

    
    def make_report():
        nonlocal json_dict
        nonlocal sorted_json_dict
        date_from = entry_from.get().replace(' ', '')
        date_to = entry_to.get().replace(' ', '')
        date_from_year, date_from_month, date_from_day, date_to_year, date_to_month, date_to_day = None, None, None, None, None, None
        try:
            date_from_day, date_from_month, date_from_year = date_from.split('.')
            date_to_day, date_to_month, date_to_year = date_to.split('.')
        except ValueError:
            try:
                date_from_day, date_from_month, date_from_year = date_from.split('/')
                date_to_day, date_to_month, date_to_year = date_to.split('/')
            except ValueError:
                pass
        if date_from_year != None:
            date_from_year = int(date_from_year)
            date_from_month = int(date_from_month)
            date_from_day = int(date_from_day)
            date_to_year = int(date_to_year)
            date_to_month = int(date_to_month)
            date_to_day = int(date_to_day)
            
        print(date_from, date_to)

        for year in range(date_from_year, date_to_year + 1):
            rng_mnth = (1, 12 + 1)
            if date_from_year == date_to_year: 
                rng_mnth = (date_from_month, date_to_month + 1)
                if date_from_month == date_to_month: 
                    if date_from_day == date_to_day:
                        rng_mnth = (date_from_month, date_to_month + 1)
            else: 
                if year == date_from_year: 
                    rng_mnth = (date_from_month, 12 + 1)
                elif year == date_to_year: 
                    rng_mnth = (1, date_to_month + 1)
            for month in range(*rng_mnth):
                rng_day = (1, 31 + 1)
                if date_from_day != date_to_day:
                    if date_from_year != date_to_year and date_from_month != date_to_month:
                        if year == date_from_year: 
                            if month == date_from_month:
                                rng_day = (date_from_day, 31 + 1)
                        elif year == date_to_year:
                            if month == date_to_month:
                                rng_day = (1, date_to_day + 1)
                    if date_from_year == date_to_year:
                        if date_from_month != date_to_month:
                            if month == date_from_month:
                                    rng_day = (date_from_day, 31 + 1)
                            elif month == date_to_month:
                                    rng_day = (1, date_to_day + 1)
                        if date_from_month == date_to_month:
                            rng_day = (date_from_day, date_to_day + 1)
                if date_from_day == date_to_day:
                    if date_from_year == date_to_year:
                        if date_from_month != date_to_month:
                            if month == date_from_month and year == date_from_year: 
                                rng_day = (date_from_day, 31 + 1)
                            elif month == date_to_month and year == date_to_year:
                                rng_day = (1, date_to_day + 1)
                        if date_from_month == date_to_month:
                            rng_day = (date_from_day, date_to_day + 1)
                    if date_from_year != date_to_year:
                        if date_from_month != date_to_month or date_from_month == date_to_month:
                            if month == date_from_month and year == date_from_year: 
                                rng_day = (date_from_day, 31 + 1)   
                            elif month == date_to_month and year == date_to_year:
                                rng_day = (1, date_to_day + 1)
                                
                for day in range(*rng_day):
                    print(f'{day}-{month}-{year}, rng_m:{rng_mnth}, rng_d:{rng_day}')
        print(sorted_json_dict)


    root = Tk()
    root.title('Создание отчета')
    root.minsize(400, 300)
    root.maxsize(400, 300)
    width, height = 400, 300
    root.geometry(f"{width}x{height}+{int(root.winfo_screenwidth() / 2) - int(width / 2)}+{int(root.winfo_screenheight() / 2) - int(height / 2)}")


    
    btn = ttk.Button(root, command=load_json_save, text='Открыть')
    btn.pack()
    btn_save = ttk.Button(root, command=save_json_save, text='Save')
    btn_save.pack(pady=10)


    listbox_name = Listbox(root, relief=FLAT, width=50, selectmode='multiple')
    listbox_name.bind('<<ListboxSelect>>', listbox_name_selected)
    listbox_name.pack()
    label_from = tk.Label(root, text='От: ')
    label_from.pack(side=LEFT)
    entry_from = tk.Entry(root)
    entry_from.pack(side=LEFT)
    label_to = tk.Label(root, text='До: ')
    label_to.pack(side=LEFT)
    entry_to = tk.Entry(root)
    entry_to.pack(side=LEFT)
    make_report_btn = ttk.Button(root, text='Создать', command=make_report)
    make_report_btn.pack(side=LEFT)

    

    root.mainloop()

def main():
    main_window()


if __name__ == '__main__':
    main()