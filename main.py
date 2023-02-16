import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog
from tkinter.messagebox import askyesno, showwarning
from datetime import datetime, timedelta
import pystray
from PIL import Image


def read_file() -> pd.DataFrame:
    #filepath = filedialog.askopenfilename(defaultextension='xlsx')
    filepath = 'test.xlsx'
    if filepath != '':
        excel_data = pd.read_excel(filepath)
        return pd.DataFrame(excel_data)

def write_file(col, data, sheet_name):
    filepath = filedialog.asksaveasfilename(defaultextension='xlsx')
    if filepath != '':
        df1 = pd.DataFrame(data, columns=col)
        df1.to_excel(f"{filepath}", sheet_name=sheet_name, index=False)

def diff_time(time_start: datetime, time_stop: datetime) -> timedelta:
    start = timedelta(minutes=time_start.minute, hours=time_start.hour, days=time_start.day)
    stop = timedelta(minutes=time_stop.minute, hours=time_stop.hour, days=time_stop.day)
    return (stop - start)

def sum_time(*times) -> timedelta:
    sum = timedelta()
    for time in times: sum += time
    return sum
    

data = read_file()
data = data.fillna('').values.tolist()
tmp = list()
for item in data:
    tmp.append(list(filter(None, item)))
data = tmp
names = [1, 2, 3, 4, 5, 6]
selected_item = ()
edited_tasks = ['edited_tasks', 'solved_tasks', 'solved_time']
edited_tasks = [0, 0, []]
str_time_start, str_time_stop = '', ''
time_start, time_stop = 0, 0


def window():

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
        if combo.get() == "":
            showwarning(message="Пользователь не выбран")
        else:
            time_stop = datetime.now()
            str_time_stop = f'{time_stop.hour:02}:{time_stop.minute:02}'
            label_stop['text'] = f'{data[selected_item[0]][1]} {str_time_stop}'

            if len(data[selected_item[0]]) > 2:
                if (entry.get() == '' and data[selected_item[0]][2] == '') or (entry.get() == data[selected_item[0]][2]):
                    pass
                else:
                    ask()
            else:
                td_diff_time = diff_time(time_start, time_stop)
                data[selected_item[0]].append(entry.get())
                data[selected_item[0]].append(str(td_diff_time)[:-3])
                edited_tasks[0] += 1
                edited_tasks[1] += int(entry.get())
                edited_tasks[2].append(td_diff_time)

            if len(entry.get()) > 0:
                listBox.itemconfig(selected_item[0], bg='green')
                entry.delete(0, END)
            else:
                listBox.itemconfig(selected_item[0], bg='yellow')

            btn_start['state'] = NORMAL
            btn_stop['state'] = DISABLED
            listBox['state'] = NORMAL
            entry['state'] = DISABLED
            btn_save['state'] = NORMAL

    def ask():
        result = askyesno(title="Подтверждение действия", 
                            message="Изменить значение?", 
                            detail=f"Текущее значение: {data[selected_item[0]][2]}")
        if result: data[selected_item[0]][2] = entry.get()

    def data_last():
        data.append([])
        data.append(['Всего', f'{edited_tasks[0]}', f'{edited_tasks[1]}', f'{str(sum_time(*edited_tasks[2]))[:-3]}'])
    
    def save():
        date = f'{datetime.now().strftime("%d|%m|%y")}'
        data_last()
        write_file(data=data, col=['№п/п', 'Задание', 'Решено', 'Время'], sheet_name=date)

    def onListboxItemSelect(event):
        global selected_item
        if len(listBox.curselection()) > 0:
            selected_item = listBox.curselection()
            btn_start['state'] = NORMAL
            item = data[selected_item[0]][1]
            item_label.config(text=item)

    def onComboboxSelect(event):
        if len(listBox.curselection()) > 0:
            listBox.select_set(selected_item[0])

    def onQuitWindow(icon):
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

    root = Tk()
    root.title('')
    root.minsize(400, 225)
    root.geometry('800x450')


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

    data_var = Variable(value=data)
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

    btn_save = ttk.Button(root, text='Сохранить', command=save)
    btn_save.pack(side=BOTTOM, pady=10)

    btn_save['state'] = DISABLED
    entry['state'] = DISABLED


    combo = ttk.Combobox(f_right, values=names)
    combo.bind("<<ComboboxSelected>>", onComboboxSelect)
    combo['state'] = 'readonly'
    combo.pack()

    
    label_start = ttk.Label(text='', master=f_right)
    label_stop = ttk.Label(text='', master=f_right)
    label_start.pack()
    label_stop.pack()


    #root.protocol("WM_DELETE_WINDOW", onDeleteWindow)
    root.iconphoto(False, tk.PhotoImage(file='icon.png'))

    root.mainloop()


def main():
    window()


if __name__ == '__main__':
    main()
