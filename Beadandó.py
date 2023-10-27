import tkinter as tk
from tkinter import ttk, messagebox
import os
import openpyxl
import sajat_modul
from sajat_modul import WidgetGridConfigurer 

def openxlsx():
    filepath = r'C:\Users\8hmar\beadando\data.xlsx'

    if not os.path.exists(filepath):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        heading = ["Keresztnév", "Vezetéknév", "Személyigazolvány szám", "Nem", "Életkor", "Származás",
                               "Lakcím", "Vallás", "Magyar állampolgár"]
        sheet.append(heading)
        workbook.save(filepath)
    else:
        workbook = openpyxl.load_workbook(filepath)

    workbook.save(filepath)
    workbook.close()
    os.system(f'start excel.exe "{filepath}"')


def enterdata():
    accepted = accept_var.get()

    if accepted == "Elfogadva":
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()

        if firstname and lastname and not any(char.isdigit() for char in firstname) and not any(
                char.isdigit() for char in lastname):
            szemelyigszam = id_num_entry.get()
            if szemelyigszam and len(szemelyigszam) == 8:

                gender = gender_combobox.get()
                age = age_spinbox.get()
                nationality = nationality_combobox.get()
                magyarallamp = reg_status_var.get()
                address = address_entry.get()
                religion = religion_combobox.get()

                filepath = r'C:\Users\8hmar\beadando\data.xlsx'

                if not os.path.exists(filepath):
                    workbook = openpyxl.Workbook()
                    sheet = workbook.active
                    heading = ["Keresztnév", "Vezetéknév", "Személyigazolvány szám", "Nem", "Életkor", "Származás",
                               "Lakcím", "Vallás", "Magyar állampolgár"]
                    sheet.append(heading)
                    workbook.save(filepath)

                workbook = openpyxl.load_workbook(filepath)
                sheet = workbook.active
                sheet.append([firstname, lastname, szemelyigszam, gender, age, nationality, address, religion,
                              magyarallamp])
                workbook.save(filepath)
                messagebox.showinfo(title="Success", message="Az adatokat sikeresen rögzítette!")
            else:
                messagebox.showwarning(title="Error", message="Nem adta meg a Személyigazolványszámát, vagy rosszul adta meg!"
                                                              "\nKérem ellenőrizze!")
        else:
            messagebox.showwarning(title="Error", message="Nem adta meg a Vezeték és Keresztnevét vagy számokat adott"
                                                          " meg, kérem ellenőrizze!")
    else:
        messagebox.showwarning(title="Error", message="Nem fogadta el az Általános Szerződési Feltételeket!")


window = tk.Tk()
window.title("Népszámláló adatbekérő ablak")

frame = tk.Frame(window)
frame.pack()

# Saving user info
user_info_frame = tk.LabelFrame(frame, text="Személyes adatai")
user_info_frame.grid(row=0, column=0, padx=20, pady=20)

last_name_label = tk.Label(user_info_frame, text="Vezetéknév")
last_name_label.grid(row=0, column=0)
first_name_label = tk.Label(user_info_frame, text="Keresztnév")
first_name_label.grid(row=0, column=1)

last_name_entry = tk.Entry(user_info_frame)
first_name_entry = tk.Entry(user_info_frame)
last_name_entry.grid(row=1, column=0)
first_name_entry.grid(row=1, column=1)

gender_label = tk.Label(user_info_frame, text="Neme")
gender_combobox = ttk.Combobox(user_info_frame, values=["Úr", "Hölgy"], state="readonly")
gender_label.grid(row=0, column=2)
gender_combobox.grid(row=1, column=2)

age_label = tk.Label(user_info_frame, text="Életkor")
age_spinbox = tk.Spinbox(user_info_frame, from_=18, to=110)
age_label.grid(row=2, column=0, )
age_spinbox.grid(row=3, column=0)

id_num_label = tk.Label(user_info_frame, text="Személyigazolvány szám")
id_num_label.grid(row=2, column=1)
id_num_entry = tk.Entry(user_info_frame)
id_num_entry.grid(row=3, column=1)

address_label = tk.Label(user_info_frame, text="Lakcíme")
address_entry = tk.Entry(user_info_frame)
address_label.grid(row=2, column=2)
address_entry.grid(row=3, column=2)

widget_configurer = WidgetGridConfigurer(user_info_frame)
widget_configurer.configure_widgets()

# Saving Curse Info
course_frame = tk.LabelFrame(frame)
course_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

registered_label = tk.Label(course_frame, text="Magyar állampolgár")
reg_status_var = tk.StringVar(value="Nem")
registered_check = tk.Checkbutton(course_frame, text="Igen", variable=reg_status_var,
                                  onvalue="Magyar állampolgár", offvalue="Nem magyar állampolgár")
registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

nationality_label = tk.Label(course_frame, text="Származás")
nationality_combobox = ttk.Combobox(course_frame, values=["Albánia", "Andorra", "Ausztria", "Azerbajdzsán",
                                                          "Belgium", "Bosznia-Hercegovina", "Bulgária", "Ciprus",
                                                          "Csehország", "Dánia", "Észtország", "Fehéroroszország",
                                                          "Finnország", "Franciaország", "Görögország", "Grúzia",
                                                          "Hollandia", "Horvátország", "Írország", "Izland",
                                                          "Kazahsztán", "Koszovó", "Lengyelország", "Lettország",
                                                          "Liechtenstein", "Litvánia", "Luxemburg", "Macedónia",
                                                          "Magyarország", "Málta", "Moldova", "Monaco", "Montenegró",
                                                          "Norvégia", "Németország", "Olaszország", "Oroszország",
                                                          "Portugália", "Románia", "San Marino", "Spanyolország",
                                                          "Svájc", "Svédország", "Szlovákia", "Szlovénia",
                                                          "Törökország)", "Ukrajna", "Vatikánváros", "Kína", "Korea",
                                                          "Japán", "India", "USA"], state="readonly")
nationality_label.grid(row=0, column=1)
nationality_combobox.grid(row=1, column=1)

religion_label = tk.Label(course_frame, text="Vallás")
religion_combobox = ttk.Combobox(course_frame, values=["Kereszténység", "Iszlám", "Vallástalan", "Hinduizmus",
                                                       "Buddhizmus", "Népi vallások", "Más vallások"],state="readonly")
religion_label.grid(row=0, column=2)
religion_combobox.grid(row=1, column=2)

widget_configurer = WidgetGridConfigurer(course_frame)
widget_configurer.configure_widgets()

# Accept terms
terms_frame = tk.LabelFrame(frame, text="Hozzájárulási igazolás")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tk.StringVar(value="Nincs elfogadva")
term_check = tk.Checkbutton(terms_frame, text="Hozzá járulok az adataim felvételéhez.",
                            variable=accept_var, onvalue="Elfogadva", offvalue="Nincs elfogadva")
term_check.grid(row=0, column=0)

widget_configurer = WidgetGridConfigurer(terms_frame)
widget_configurer.configure_widgets()

# Button
buttons_frame = tk.LabelFrame(frame)
buttons_frame.grid(row=3, column=0, padx=20, pady=20)

button = tk.Button(buttons_frame, text="Befejezés", command=enterdata)
button.grid(row=3, column=0, sticky="news", padx=30, pady=20)

button = tk.Button(buttons_frame, text="Megnyitás", command=openxlsx)
button.grid(row=3, column=1, sticky="news", padx=30, pady=20)

window.mainloop()
