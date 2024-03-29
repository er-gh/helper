import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog
from tkinter.messagebox import showwarning
from datetime import datetime, timedelta
import pystray
from PIL import Image
import json
import glob
import os


def read_file() -> pd.DataFrame:
    filepath = filedialog.askopenfilename(defaultextension='xlsx')
    if filepath != '':
        excel_data = pd.read_excel(filepath)
        return pd.DataFrame(excel_data)

def write_file(col, data, sheet_name):
    tmp_name = name
    tmp_name = tmp_name.replace(' ', '_')
    filepath = os.getcwd() + f'\{datetime.now().year}_{datetime.now().month}_{datetime.now().day}_{tmp_name}.xlsx'
    df1 = pd.DataFrame(data, columns=col)
    df1.to_excel(filepath, sheet_name=sheet_name, index=False)

def diff_time(time_start: datetime, time_stop: datetime) -> timedelta:
    start = timedelta(minutes=time_start.minute, hours=time_start.hour, days=time_start.day)
    stop = timedelta(minutes=time_stop.minute, hours=time_stop.hour, days=time_stop.day)
    return (stop - start)

def sum_time(*times) -> timedelta:
    sum = timedelta()
    for time in times: sum += time
    return sum

def dt_to_td(time: datetime) -> timedelta:
    return timedelta(hours=time.hour, minutes=time.minute)

def load_json() -> dict | None:
    json_file = glob.glob('*.json')
    if len(json_file) > 1: 
        showwarning("", "В папке может быть только 1 файл с расширением .json")
    global name
    try:
        with open(json_file[0], encoding='utf-8') as fp:
            json_dict = json.load(fp)
            name = list(json_dict.keys())[0]
        return json_dict
    except IndexError:
        return None
    except UnboundLocalError:
        print("unbound")
        return None
    except FileNotFoundError:
        print("fnfe")
        return None

def save_json(dct: dict):
    tmp_name = name
    tmp_name = tmp_name.replace(' ', '_')
    with open(f'{tmp_name}.json', 'w', encoding='utf-8') as fp:
        json.dump(dct, fp, indent=4, ensure_ascii=False)

def save_in_json(dct_check: dict, dct: dict, name: str) -> None:
    json_dict = {}
    date = datetime.now()
    year = str(date.year)
    month = str(date.month)
    day = str(date.day)
    res = None
    if dct_check != None and len(dct_check) > 0: res = check_dict(dct_check, name)
    else: dct_check = {}

    json_dict[name] = {
            year: {
                month: {
                    day: {} } } }

    if res == '':
        print('res = ""')
    else: 
        json_dict[name][year][month][day].update(dct)
    if res == 'y':
        dct_check[name].update(json_dict[name])
        print(f'res = "y"')
    if res == 'm':
        dct_check[name][year].update(json_dict[name][year])
        print('res = "m"')
    if res == 'd':
        dct_check[name][year][month].update(json_dict[name][year][month])
        print('res = "d"')
    print(json_dict)
    json_dict.update(dct_check)
    save_json(json_dict)

def check_dict(dct: dict, name: str) -> str:
    year = datetime.now().year
    month = datetime.now().month
    day = datetime.now().day
    try:
        dct[name][str(year)][str(month)][str(day)]
        return ''
    except KeyError as e:
        e = str(e).replace("'","")
        if year == int(e): return "y"
        if month == int(e): return "m"
        if day == int(e): return "d"
    

name = ''
to_json_dict = {}
from_json_dict = load_json()
data = read_file()
try:
    data = data.fillna('').values.tolist()
    tmp = list()
    for item in data:
        tmp.append(list(filter(lambda s: s != '', item)))
    data = tmp
except AttributeError:
    data = []
selected_item = ()
str_time_start, str_time_stop = '', ''
time_start, time_stop = 0, 0


def window():

    def login_window(root_win):

        def cancel(window):
            win_destroy(window)

        def win_destroy(window):
            window.grab_release()
            window.destroy()

        def save_name(window):
            nonlocal result
            name = entry.get().strip()
            if name != '': result = name
            cancel(window)

        window = Toplevel(root_win)
        window.title('Введите ФИО')
        width, height = 400, 225
        window.geometry(f"{width}x{height}+{int(root_win.winfo_screenwidth() / 2) - int(width / 2)}+{int(root_win.winfo_screenheight() / 2) - int(height / 2)}")
        window.resizable(False, False)
        window["background"] = 'white'
        result = None

        window.focus_force()

        frame_entry = tk.Frame(window)
        frame_entry.config(background='white')
        frame_entry.pack(fill=BOTH, expand=True)
        entry = ttk.Entry(frame_entry, width=25, font=('', 14))
        entry.place(relx=0.5, rely=0.4, anchor=CENTER)

        frame = tk.Frame(window)
        frame.pack(fill=BOTH, side=BOTTOM)
        button_ok = ttk.Button(frame, text='Продолжить', command=lambda: save_name(window))
        button_ok.pack(side=RIGHT, padx=10, pady=10)

        window.protocol("WM_DELETE_WINDOW", lambda: cancel(window))
        window.grab_set()
        window.wait_window()
        return result

    def start():
        global selected_item
        global str_time_start
        global time_start
        listBox.select_set(selected_item[0])
        time_start = datetime.now()
        str_time_start = f'{time_start.hour:02}:{time_start.minute:02}'
        if len(selected_item) > 0:
            label_start['text'] = f'{data[selected_item[0]][1]} {str_time_start}'
            label_stop['text'] = ''
            btn_start['state'] = DISABLED
            btn_stop['state'] = NORMAL
            listBox['state'] = DISABLED
            btn_save['state'] = DISABLED
            entry['state'] = NORMAL
    
    def stop():
        global selected_item
        global str_time_stop
        global time_stop
        global edited_tasks
        global to_json_dict
        global from_json_dict
        if name == "":
            showwarning(message="Пользователь не выбран")
        else:
            time_stop = datetime.now()
            str_time_stop = f'{time_stop.hour:02}:{time_stop.minute:02}'
            label_stop['text'] = f'{data[selected_item[0]][1]} {str_time_stop}'

            if len(data[selected_item[0]]) > 2:
                if ((entry.get() == data[selected_item[0]][2]) or (entry.get() == '' and data[selected_item[0]][2] == 0)):
                    pass
                else:
                    ask()
            else:
                dt_diff_time = datetime.strptime(str(diff_time(time_start, time_stop)), '%H:%M:%S')
                minutes_time = dt_to_td(dt_diff_time).total_seconds() / 60
                try:
                    data[selected_item[0]].append(int(entry.get()))
                except ValueError:
                    if entry.get().strip() != '': showwarning(message="Можно вводить только цифры")
                    data[selected_item[0]].append(0)

                data[selected_item[0]].append(int(minutes_time))

            if len(entry.get()) > 0:
                try:
                    if int(entry.get()):
                        listBox.itemconfig(selected_item[0], bg='green')
                except ValueError:
                    listBox.itemconfig(selected_item[0], bg='yellow')
                entry.delete(0, END)
            elif data[selected_item[0]][2] != 0:
                listBox.itemconfig(selected_item[0], bg='green')
            else:
                listBox.itemconfig(selected_item[0], bg='yellow')

            to_json_dict[str(selected_item[0])] = {
                "task": data[selected_item[0]][1],
                "solved": data[selected_item[0]][2],
                "time": data[selected_item[0]][3]
            }

            if from_json_dict != None:
                try:
                    from_json_dict[name][str(datetime.now().year)][str(datetime.now().month)][str(datetime.now().day)][str(selected_item[0])] = to_json_dict[str(selected_item[0])]
                except KeyError:
                    pass

            save_in_json(from_json_dict, to_json_dict, name)

            label_task_info.config(text=f'Сделано: {int(data[selected_item[0]][2])}')
            label_time_info.config(text=f'Время(мин): {int(data[selected_item[0]][3])}')

            btn_start['state'] = NORMAL
            btn_stop['state'] = DISABLED
            listBox['state'] = NORMAL
            entry['state'] = DISABLED
            btn_save['state'] = NORMAL


    def ask():
        result = askChangeAddCancel(root)
        try:
            if result: 
                data[selected_item[0]][2] = int(entry.get()) + int(data[selected_item[0]][2])
            elif result == False:
                if entry.get() == '': data[selected_item[0]][2] = 0
                else: data[selected_item[0]][2] = int(entry.get())
        except ValueError:
            data[selected_item[0]][2] = 0
        if result != None:
            dt_diff_time = datetime.strptime(str(diff_time(time_start, time_stop)), '%H:%M:%S')
            minutes_time = dt_to_td(dt_diff_time).total_seconds() / 60
            data[selected_item[0]][3] = int(data[selected_item[0]][3]) + int(minutes_time)
        

    def askChangeAddCancel(root_win):
        window = Toplevel(root_win)
        window.title(f'{entry.get()}')
        width, height = 400, 225
        window.geometry(f"{width}x{height}+{int(root_win.winfo_screenwidth() / 2) - int(width / 2)}+{int(root_win.winfo_screenheight() / 2) - int(height / 2)}")
        window.resizable(False, False)
        window["background"] = 'white'

        result = None

        def cancel(window):
            win_destroy(window)

        def add(window):
            nonlocal result
            result = True
            win_destroy(window)

        def change(window):
            nonlocal result
            result = False
            win_destroy(window)

        def win_destroy(window):
            window.grab_release()
            window.destroy()
        
        window.protocol("WM_DELETE_WINDOW", lambda: cancel(window))

        label = ttk.Label(window, text="Добавить или обновить информацию?")
        label.config(background='white', font=("", 12))
        label.pack(anchor='center',side='top', pady=5)

        label_task = ttk.Label(window, text=f'Задание: {data[selected_item[0]][1]}')
        label_task.config(background='white')
        label_task.pack(anchor=W, padx=5, pady=5)

        label_done = ttk.Label(window)
        try:
            label_done.config(text=f'Сделано: {int(data[selected_item[0]][2])}')
        except ValueError:
            label_done.config(text=f'Сделано: ')
        label_done.config(background='white')
        label_done.pack(anchor=W, padx=5, pady=5)

        label_time = ttk.Label(window, text=f'Время(мин): {int(data[selected_item[0]][3])}')
        label_time.config(background='white')
        label_time.pack(anchor=W, padx=5, pady=5)

        frame = ttk.Frame(window)
        frame.pack(side=BOTTOM, fill=X)

        for c in range(3): frame.columnconfigure(index=c, weight=1)

        close_button = ttk.Button(frame, text='Закрыть', command=lambda: cancel(window))
        close_button.grid(column=2, row=0, padx=5, pady=7, sticky=NSEW)

        update_button = ttk.Button(frame, text="Обновить", command=lambda: change(window))
        update_button.grid(column=1, row=0, padx=5, pady=7, sticky=NSEW)

        add_button = ttk.Button(frame, text="Добавить", command=lambda: add(window))
        add_button.grid(column=0, row=0, padx=5, pady=7, sticky=NSEW)

        window.grab_set()
        window.wait_window()
        return result

    
    def save():
        date = f'{datetime.now().strftime("%d|%m|%y")}'
        write_file(data=data, col=['№п/п', 'Задание', 'Сделано', 'Время(мин)'], sheet_name=date)

    def onListboxItemSelect(event):
        global selected_item
        if len(listBox.curselection()) > 0:
            selected_item = listBox.curselection()
            btn_start['state'] = NORMAL
            item = data[selected_item[0]][1]
            item_label.config(text=item)
            try:
                task = data[selected_item[0]][2]
                time = data[selected_item[0]][3]
                try:
                    label_task_info.config(text=f'Сделано: {int(task)}')
                    label_time_info.config(text=f'Время(мин): {int(time)}')
                except ValueError:
                    label_task_info.config(text=f'Сделано: ')
                    label_time_info.config(text=f'Время(мин): {int(time)}')
            except IndexError:
                label_task_info.config(text=f'')
                label_time_info.config(text=f'')


    def onQuitWindow(icon):
        if len(to_json_dict) > 0: save()
        icon.stop()
        root.destroy()

    def onShowWindow(icon):
        icon.stop()
        root.after(0, root.deiconify)

    def onDeleteWindow():
        root.withdraw()
        image = Image.open("icon.png")
        menu = pystray.Menu(pystray.MenuItem('Открыть', onShowWindow, default=True), pystray.MenuItem('Закрыть', onQuitWindow))
        icon = pystray.Icon("Имя", image, "Отчет", menu)
        icon.run()

    def checkStateOnStart():
        for i in range(len(data)):
            if len(data[i]) > 2:
                if data[i][2] > 0: listBox.itemconfig(i, bg='green')
                else: listBox.itemconfig(i, bg='yellow')


    root = Tk()
    root.title('')
    root.minsize(400, 225)
    width, height = 800, 450
    root.geometry(f"{width}x{height}+{int(root.winfo_screenwidth() / 2) - int(width / 2)}+{int(root.winfo_screenheight() / 2) - int(height / 2)}")

    global name
    if name == '':
        root.withdraw()
        result = login_window(root)
        if result == None: 
            root.destroy()
            return
        name = result
        root.deiconify()
        root.focus_force()


    f_left = Frame(root)
    f_right = Frame(root)
    f_right_right = Frame(f_right)

    f_left.pack(side=LEFT, anchor=N, fill=Y)
    f_right.pack(side=RIGHT, fill=BOTH)
    f_right_right.pack(side=BOTTOM, padx=5, pady=10)


    item_label = Label(root)
    item_label.pack()


    btn_start = ttk.Button(f_right_right, text='Начать', command=start)
    btn_stop = ttk.Button(f_right_right, text='Стоп', command=stop)

    btn_start['state'] = DISABLED
    btn_stop['state'] = DISABLED
    btn_start.pack(side=LEFT)
    btn_stop.pack(side=RIGHT)

    str_data = []
    for i in range(len(data)):
        temp_str_data = f'{i + 1:3d} : {data[i][1]}'
        str_data.append(temp_str_data)
    data_var = Variable(value=str_data)

    listBox = Listbox(f_left, listvariable=data_var, relief=FLAT, width=50)
    listBox.bind('<<ListboxSelect>>', onListboxItemSelect)


    scroll_y = Scrollbar(f_left, command=listBox.yview)
    scroll_x = Scrollbar(f_left, command=listBox.xview, orient='horizontal')
    scroll_y.pack(side=RIGHT, fill=Y)
    scroll_x.pack(side=BOTTOM, fill=X)

    listBox.config(yscrollcommand=scroll_y.set)
    listBox.config(xscrollcommand=scroll_x.set)
    listBox.pack(expand=1, fill=Y)


    entry = Entry(root)
    entry.pack()

    label_task_info = ttk.Label(text='')
    label_time_info = ttk.Label(text='')
    label_task_info.pack()
    label_time_info.pack()

    btn_save = ttk.Button(root, text='Сохранить', command=save)
    btn_save.pack(side=BOTTOM, pady=10)

    btn_save['state'] = DISABLED
    entry['state'] = DISABLED


    label_name = ttk.Label(f_right, text=f'{name}')
    label_name.pack()

    
    label_start = ttk.Label(text='', master=f_right)
    label_stop = ttk.Label(text='', master=f_right)
    label_start.pack()
    label_stop.pack()

    checkStateOnStart()

    root.protocol("WM_DELETE_WINDOW", onDeleteWindow)
    root.iconphoto(False, tk.PhotoImage(file='icon.png'))

    root.mainloop()


def main():
    window()


if __name__ == '__main__':
    main()
