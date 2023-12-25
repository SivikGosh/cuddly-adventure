import os.path

import ffmpeg
import pandas
import os
import re
from tkinter import filedialog, Tk, ttk


dataframe = pandas.DataFrame({
    "№ п/п": [],
    "Тип носителя": [],
    "Номер машинного носителя": [],
    "Наименование контента": [],
    "Наименование файла": [],
    "Формат файла": [],
    "Размер файла (Мб)": [],
    "Продолжительность (сек.)": [],
})
duration = 0
rows = []


def get_duration(file):
    global duration
    try:
        duration = ffmpeg.probe(file)["streams"][0]["duration"]
    except ffmpeg._run.Error:
        duration = 0
    finally:
        return int(float(duration))


def get_size(file):
    size = os.path.getsize(file)
    return size / 1048576


def get_name(file):
    file = os.path.basename(file).split('.')
    name = file[0]
    form = file[1]
    return name, form


def recurse(folder):
    os.chdir(folder)
    files = os.listdir('./')
    full_duration = 0
    file_count = 0

    for file in files:
        if re.findall(r'^.+\.[a-zA-Z0-9]+$', file):
            duration = get_duration(file)
            size = get_size(file)
            file_name, file_format = get_name(file)

            row = {
                "Наименование файла": file_name,
                "Формат файла": file_format,
                "Размер файла (Мб)": round(float(size), 2),
                "Продолжительность (сек.)": duration,
            }

            rows.append(row)

            full_duration += duration
            file_count += 1

    label['text'] = f'Общий хронометраж {full_duration} сек.\n' \
                    f'Записано {file_count} файлов'

    return rows


def main():
    folder = filedialog.askdirectory()
    rows = recurse(folder)
    for row in rows:
        global dataframe
        dataframe = dataframe._append(row, ignore_index=True)
    destination = filedialog.askdirectory()
    dataframe.to_excel(f'{destination}/таблица.xlsx', index=False)


if __name__ == "__main__":
    root = Tk()
    root.title("Приложение на Tkinter")
    root.geometry("600x800")
    root.resizable(False, False)
    style = ttk.Style()
    style.configure('W.TButton', font=('Tahoma', 10, 'bold'), )
    btn = ttk.Button(
        text="Выбрать папку",
        command=main,
        padding=10,
        width=100,
        style='W.TButton'
    )
    btn.pack()
    label = ttk.Label(root, padding=10)
    label.pack()
    root.mainloop()
