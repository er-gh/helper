from tkinter import ttk, filedialog
from tkinter.messagebox import showerror
from tkinter import *
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import *
from dateutil.relativedelta import relativedelta
import tkinter as tk
import pandas as pd
import json
import os
import copy


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
        tmp_date_from = entry_from.get().replace(' ', '')
        tmp_date_to = entry_to.get().replace(' ', '')
        date_from, date_to, date_counter = {}, {}, {}
        try:
            date_from['day'], date_from['month'], date_from['year'] = tmp_date_from.split('.')
            date_to['day'], date_to['month'], date_to['year'] = tmp_date_to.split('.')
        except ValueError:
            try:
                date_from['day'], date_from['month'], date_from['year'] = tmp_date_from.split('/')
                date_to['day'], date_to['month'], date_to['year'] = tmp_date_to.split('/')
            except ValueError:
                pass

        if len(date_from) > 0:
            date_from = datetime(year=int(date_from['year']), month=int(date_from['month']), day=int(date_from['day']))
            date_to = datetime(year=int(date_to['year']), month=int(date_to['month']), day=int(date_to['day']))

        
        for name in selected_names:
            temp_check_task = {}
            date_counter = date_from
            while date_counter <= date_to:
                try:
                    js_d = json_dict[name][str(date_counter.year)][str(date_counter.month)][str(date_counter.day)]
                    for task_num, value in js_d.items():
                        print(f'task_num:{task_num}', [f'{k}:{v}' for k, v in value.items()])
                        
                        if task_num in temp_check_task:
                            temp_check_task[task_num]['task'] = value['task']
                            temp_check_task[task_num]['solved'] += value['solved']
                            temp_check_task[task_num]['time'] += value['time']
                        else:
                            temp_check_task[task_num] = dict(value.items())
                except KeyError as e:
                    int_e = int(str(e).replace("'", ''))
                    if int_e == date_counter.year: 
                        date_counter += relativedelta(years=+1)
                        date_counter = datetime(year=date_counter.year, month=1, day=1)
                    if date_counter.month != date_counter.day and int_e == date_counter.month: 
                        date_counter += relativedelta(months=+1)
                        date_counter = datetime(year=date_counter.year, month=date_counter.month, day=1)
                    print(f'cont: {e}, :{int_e}', date_counter, name)
                date_counter += relativedelta(days=+1)

            sorted_json_dict[name] = temp_check_task
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
    entry_from.insert(0, '10/03/2023')
    entry_from.pack(side=LEFT)
    label_to = tk.Label(root, text='До: ')
    label_to.pack(side=LEFT)
    entry_to = tk.Entry(root)
    entry_to.insert(0, '17/04/2023')
    entry_to.pack(side=LEFT)
    make_report_btn = ttk.Button(root, text='Создать', command=make_report)
    make_report_btn.pack(side=LEFT)

    
    root.mainloop()


def main():
    main_window()


if __name__ == '__main__':
    main()
