import os

from tkinter import *
from tkinter import ttk
from tkinter.messagebox import showerror, showinfo
from generate import generate_birks


# Создание окна
window = Tk()
window.title("Генерация бирок")
window.minsize(900, 700)

canvas = Canvas(window)
scrollbar = Scrollbar(window, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=scrollbar.set)

canvas.grid(row=0, column=0, sticky="nsew")
scrollbar.grid(row=0, column=1, sticky="ns")

window.grid_rowconfigure(0, weight=1)
window.grid_columnconfigure(0, weight=1)

frame = Frame(canvas)
canvas.create_window((0, 0), window=frame, anchor="nw")

# Создание заголовка
label = ttk.Label(frame, text="Выберите количество бирок и введите данные", padding=10, font=("Arial", 14))
label.grid(column=2, row=0)

# Радио кнопки для выбора больших или маленьких бирок
var = IntVar()
var.set(1)
radio_1 = Radiobutton(frame, text='Маленькая бирка', variable=var, value=1).grid(column=0, row=2)
radio_2 = Radiobutton(frame, text='Большая бирка', variable=var, value=2).grid(column=1, row=2)


# Функция для получения результата выбора пользователя
def change():
    match var.get():
        case 1: return 1
        case 2: return 2


count_birk_label = ttk.Label(frame, text="Введите количество бирок", font=("Arial", 10))
count_birk_label.grid(column=2, row=1)
count_birk_enter = ttk.Entry(frame)
count_birk_enter.grid(column=2, row=2)


def interfase_enter_date():
    try:
        count_birk = int(count_birk_enter.get())
        count_line = 4

        for line in range(count_birk):
            ttk.Label(frame, text="Введите номер", font=("Arial", 10)).grid(column=0, row=count_line, pady=25,
                                                                            sticky=N)

            globals()[f'number_prot_{line}'] = ttk.Entry(frame)
            globals()[f'number_prot_{line}'].grid(column=0, row=count_line + 1, padx=8, sticky=N)

            ttk.Label(frame, text="Введите до скольки кВ", font=("Arial", 10)).grid(column=1, row=count_line, pady=25,
                                                                                    sticky=N)

            globals()[f'count_kv_{line}'] = ttk.Entry(frame)
            globals()[f'count_kv_{line}'].grid(column=1, row=count_line + 1, padx=8, sticky=N)

            ttk.Label(frame, text="Введите дату следующего испытания в формате - DD.MM.YYYY",
                      font=("Arial", 10)).grid(column=2, row=count_line, pady=25, sticky=N)

            globals()[f'data_{line}'] = ttk.Entry(frame)
            globals()[f'data_{line}'].grid(column=2, row=count_line + 1, padx=8, sticky=N)
            
            count_line += 1

        frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))
        Button(frame, text="Заполнить", command=registr_info).grid(column=3, row=count_line-2)
    except:
        showerror("message", f"Вы ввели не правильное значение количества бирок!")
        return
        

def registr_info():
    count_birk = int(count_birk_enter.get())
    date_birk = []

    for count_value in range(count_birk):
        number_prot = globals()[f'number_prot_{count_value}'].get()
        count_kv = globals()[f'count_kv_{count_value}'].get()
        data = globals()[f'data_{count_value}'].get()

        date_birk.append([int(number_prot), int(count_kv), data])

    min_max_radio = change()

    if min_max_radio == 1:
        doc = generate_birks(count_birk=count_birk, date=date_birk, flag_size=True)
    else:
        doc = generate_birks(count_birk=count_birk, date=date_birk, flag_size=False)
        
    doc.save("result.docx")
    showinfo("message", "Бирки успешно сгенерированы, файл сохранен")
    
    os.startfile("result.docx")


Button(frame, text="Ввести данные", command=interfase_enter_date).grid(column=3, row=2)


if __name__ == "__main__":
    window.mainloop()