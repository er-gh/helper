import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog
from tkinter.messagebox import askyesno, showwarning
from datetime import datetime
import pystray
from PIL import Image


def read_file() -> pd.DataFrame:
    filepath = filedialog.askopenfilename(defaultextension='xlsx')
    if filepath != '':
        excel_data = pd.read_excel(filepath)
        return pd.DataFrame(excel_data)

def write_file(index, col, data, sheet_name):
    filepath = filedialog.asksaveasfilename(defaultextension='xlsx')
    if filepath != '':
        df1 = pd.DataFrame(data, index=index, columns=col)
        df1.to_excel(f"{filepath}", sheet_name=sheet_name)


data = read_file()
data = data.fillna('').values.tolist()
tmp = list()
for item in data:
    tmp.append(list(filter(None, item)))
data = tmp
names = [1, 2, 3, 4, 5, 6]
selected_item = ()
time_start, time_stop = 0, 0


def window():

    def start():
        global selected_item
        global time_start
        listBox.select_set(selected_item[0])
        cur_datetime = datetime.now()
        time_start = f'{cur_datetime.hour:02}:{cur_datetime.minute:02}'
        if len(selected_item) > 0:
            label_start['text'] = f'{data[selected_item[0]][1]} {time_start}'
            label_stop['text'] = ''
            btn_start['state'] = DISABLED
            btn_stop['state'] = NORMAL
            listBox['state'] = DISABLED
            btn_save['state'] = DISABLED
            entry['state'] = NORMAL
    
    def stop():
        global selected_item
        global time_stop
        if combo.get() == "":
            showwarning(message="Пользователь не выбран")
        else:
            cur_datetime = datetime.now()
            time_stop = f'{cur_datetime.hour:02}:{cur_datetime.minute:02}'
            label_stop['text'] = f'{data[selected_item[0]][1]} {time_stop}'

            if len(data[selected_item[0]]) > 2:
                if (entry.get() == '' and data[selected_item[0]][2] == '') or (entry.get() == data[selected_item[0]][2]):
                    pass
                else:
                    ask()
            else:
                data[selected_item[0]].append(entry.get())
                data[selected_item[0]].append(time_start)
                data[selected_item[0]].append(time_stop)
                data[selected_item[0]].append(combo.get())

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
    
    def save():
        indexes = []
        date = f'{datetime.now().day:02}|{datetime.now().month:02}|{datetime.now().year:02}'
        for i in range(len(data)):
            indexes.append(data[i][0])
        write_file(data=data, index=indexes, col=['№п/п', 'Задание', 'Решено', 'Начало', 'Конец', 'ФИО'], sheet_name=date)

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


    root.protocol("WM_DELETE_WINDOW", onDeleteWindow)
    root.iconphoto(False, tk.PhotoImage(file='icon.png'))

    root.mainloop()


def main():
    window()


if __name__ == '__main__':
    main()
