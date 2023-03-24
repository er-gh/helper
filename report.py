from tkinter import ttk, filedialog
from tkinter.messagebox import showerror
from tkinter import *
from pathlib import Path
import json
import os
import tkinter as tk

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

def main_window():

    json_dict = {}
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

        print(date_from, date_to)
        print(f'year:{date_from_year}, month:{date_from_month}, day:{date_from_day}')
        print(f'year:{date_to_year}, month:{date_to_month}, day:{date_to_day}')
        print([json_dict[name]['2023'] for name in selected_names])


        for year in range(int(date_to_year) - int(date_from_year) + 1):
            year = int(date_from_year) + year
            try:
                for month in range(1, 13 - abs(int(date_to_month) - int(date_from_month))):
                    month = int(date_from_month)
                    try:
                        for day in range(1, 32):
                            #print(f'year:{year}, month:{month}, day:{day}')
                            try:
                                print([json_dict[name][str(year)][str(month)][str(day)] for name in selected_names])
                            except KeyError:
                                continue
                    except KeyError:
                        break
            except KeyError:
                break
#    10/02/2022


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