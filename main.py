import math
from tkinter import messagebox
import tkinter as tk
from tkinter import ttk

import numpy as np
from pyautocad import Autocad, APoint  # Импортируйте необходимые классы из библиотеки Autocad

from initial_data import WordEquationReplacer, second_, third_
#from autocad import AutoCADLines, AutoCADLines_1


class App:
    def __init__(self, root):
        print("Инициализация началась")
        self.root = root
        self.root.title("Ввод данных")

        # Создание виджета Canvas для прокрутки
        self.canvas = tk.Canvas(root)
        self.scroll_y = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scroll_y.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scroll_y.pack(side="right", fill="y")

        # Добавление обработчика события прокрутки мыши
        self.canvas.bind_all("<MouseWheel>", self.on_mouse_wheel)

        # Создание полей ввода
        self.entries = {}
        entry_labels = [
            "Количество городского населения на отчетный год в районе одной из станций с грузовыми операциями",
            "Плотность сельского населения на  отчетный год",
            "Ежегодный прирост численности сельского населения",
            "Ежегодный прирост численности городского населения",
            "Коэффициент подвижности населения",
            "Лесосырье",
            "Продукция заводской деревообработки",
            "Пиломатериалы",
            "Круглый лес",
            "дрова и другие отходы",
            "Плотность посевов на отчетный год",
            "удельный вес посевных площадей зерна на отчетный срок",
            "прирост посевных площадей зерна на расчетный срок",
            "зерно, урожайность на расчетный срок",
            "зерно, потребность для городского населения",
            "зерно, потребность для сельского населения",
            "зерновые фуражные культуры, удельный вес на отчетный срок",
            "прирост фуража на расчетный срок",
            "урожайность фуража на расчетный срок",
            "картофель, удельный вес на расчетный срок",
            "картофель, прирост на расчетный срок",
            "картофель, урожайность на расчетный срок",
            "картофель, потребность для городского населения",
            "картофель, потребность для сельского населения",
            "остальные культуры, удельный вес на расчетный срок",
            "остальные культуры, прирост на расчетный срок",

            "Крупный_рогатый_скот_количество_на_100_га",
            "Крупный_рогатый_скот_прирост",
            "Крупный_рогатый_скот_потребность_в_картофеле",
            "Крупный_рогатый_скот_потребность_в_фуражном_зерне",
            "Мелкий_скот_количество_на_100_га",
            "Мелкий_скот_скот_прирост_скот",
            "Мелкий_скот_потребность_в_картофеле",
            "Мелкий_скот_потребность_в_фуражном_зерне",
            "Нормы_высева_на_расчетный_срок_зерновые_продовольственные",
            "Нормы_высева_на_расчетный_срок_зерновые_фуражные",
            "Нормы_высева_на_расчетный_срок_картофель",
            "товары народного потребления",
            "каменный уголь",
            "лесоматериалы",
            "минеральные удобрения",
            "нефтепродукты",
            "Нормы_расходов_грузов_ввоза_для_нужд_населения_товары_народного_потребления",
            "Нормы_расходов_грузов_каменный уголь",
            "Нормы_расходов_грузов_лесоматериалы",
            "Нормы расходов грузов ввоза для нужд сельского хозяйства минеральные удобрения",
            "Нормы расходов грузов ввоза нефтепродукты",
            "транзитные грузы от А к Б_каменный уголь",
            "транзитные грузы от А к Б_руда",
            "транзитные грузы от А к Б_нефтепродукты",
            "грузы капитального строительства",
            "транзитные грузы от А к Б_товары_народного_потребления",
            "транзитные грузы от А к Б_прочие грузы",
            "транзитные грузы от Б к А_Металлы",
            "транзитные грузы от Б к А_Машины и металлоизделия",
            "транзитные грузы от Б к А_Лесоматериалы",
            "транзитные грузы от Б к А_Сельскохозяйственные грузы",
            "транзитные грузы от Б к А_ минеральные удобрения",
            "транзитные грузы от А к Б_прочие грузы"



        ]

        # Добавление полей ввода в три столбца с выбором для определенных полей
        for index, label in enumerate(entry_labels):
            ttk.Label(self.scrollable_frame, text=label + ":").grid(row=index, column=0, sticky="w", pady=5)
            entry = ttk.Entry(self.scrollable_frame)
            # if label == "redaction":
            # entry = ttk.Combobox(self.scrollable_frame, values=["new", "old"])
            # elif label == "rail_type":
            # entry = ttk.Combobox(self.scrollable_frame, values=["Р65", "Р50"])
            # elif label == "material_of_sleepers":
            # entry = ttk.Combobox(self.scrollable_frame, values=["Дерево", "Железобетон"])
            # else:
            # entry = ttk.Entry(self.scrollable_frame)

            entry.grid(row=index, column=1, pady=5)
            entry.bind('<Return>', self.focus_next)  # Привязка клавиши Enter к функции
            self.entries[label] = entry  # Сохраняем ссылки на поля ввода

        # Кнопка для копирования значений
        self.copy_button = ttk.Button(self.scrollable_frame, text="Скопировать", command=self.copy_to_clipboard)
        self.copy_button.grid(row=len(entry_labels) + 1, column=0, columnspan=2, pady=5)

        # Кнопка для вставки значений
        self.paste_button = ttk.Button(self.scrollable_frame, text="Вставить", command=self.paste_from_clipboard)
        self.paste_button.grid(row=len(entry_labels) + 2, column=0, columnspan=2, pady=5)

        # Кнопка для подтверждения ввода
        self.confirm_button = ttk.Button(self.scrollable_frame, text="Подтвердить", command=self.confirm)
        self.confirm_button.grid(row=len(entry_labels), column=0, columnspan=2, pady=20)
        # Устанавливаем размер окна (ширина x высота)
        self.root.geometry("500x600")  # Значение можно изменить по вашему желанию
        print("Инициализация закончилась")

    def on_mouse_wheel(self, event):
        """Обработка прокрутки колесика мыши"""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")  # Прокрутка на 1 единицу

    def focus_next(self, event):
        """Переход к следующему полю ввода"""
        current_widget = event.widget
        current_widget_index = list(self.entries.values()).index(current_widget)

        # Переход к следующему полю ввода
        next_widget = list(self.entries.values())[(current_widget_index + 1) % len(self.entries)]
        next_widget.focus_set()

    def copy_to_clipboard(self):
        """Копирует значения из полей ввода в буфер обмена"""
        clipboard_data = "\n".join([f"{label}: {entry.get()}" for label, entry in self.entries.items()])
        self.root.clipboard_clear()
        self.root.clipboard_append(clipboard_data)
        messagebox.showinfo("Копирование", "Данные скопированы в буфер обмена.")

    def paste_from_clipboard(self):
        """Вставляет значения из буфера обмена в поля ввода"""
        clipboard_data = self.root.clipboard_get().strip().split('\n')
        for line, label in zip(clipboard_data, self.entries.keys()):
            try:
                key, value = line.split(':', 1)
                if key.strip() == label:
                    self.entries[label].delete(0, tk.END)  # Очистить текущее значение
                    self.entries[label].insert(0, value.strip())  # Вставить новое значение
            except ValueError:
                continue  # Игнорируем строки, которые не в формате "label: value"

    def confirm(self):
        print("Инициализация давно закончилась")
        # Считываем значения из полей ввода в глобальный словарь
        data = {}
        for field, entry in self.entries.items():
            value = entry.get().strip()  # Убираем пробелы

            if field in [
                "Количество городского населения на отчетный год в районе одной из станций с грузовыми операциями",
                "Плотность сельского населения на  отчетный год",
                "Ежегодный прирост численности сельского населения",
                "Ежегодный прирост численности городского населения",
                "Коэффициент подвижности населения",
                "Лесосырье",
                "Продукция заводской деревообработки",
                "Пиломатериалы",
                "Круглый лес",
                "дрова и другие отходы",
            "Плотность посевов на отчетный год",
            "удельный вес посевных площадей зерна на отчетный срок",
            "прирост посевных площадей зерна на расчетный срок",
            "зерно, урожайность на расчетный срок",
            "зерно, потребность для городского населения",
            "зерно, потребность для сельского населения",
            "зерновые фуражные культуры, удельный вес на отчетный срок",
            "прирост фуража на расчетный срок",
            "урожайность фуража на расчетный срок",
            "картофель, удельный вес на расчетный срок",
            "картофель, прирост на расчетный срок",
            "картофель, урожайность на расчетный срок",
            "картофель, потребность для городского населения",
            "картофель, потребность для сельского населения",
            "остальные культуры, удельный вес на расчетный срок",
            "остальные культуры, прирост на расчетный срок",
            "Крупный_рогатый_скот_количество_на_100_га",
            "Крупный_рогатый_скот_прирост",
            "Крупный_рогатый_скот_потребность_в_картофеле",
            "Крупный_рогатый_скот_потребность_в_фуражном_зерне",
            "Мелкий_скот_количество_на_100_га",
            "Мелкий_скот_скот_прирост_скот",
            "Мелкий_скот_потребность_в_картофеле",
            "Мелкий_скот_потребность_в_фуражном_зерне",
            "Нормы_высева_на_расчетный_срок_зерновые_продовольственные",
            "Нормы_высева_на_расчетный_срок_зерновые_фуражные",
            "Нормы_высева_на_расчетный_срок_картофель",
            "товары народного потребления",
            "каменный уголь",
            "лесоматериалы",
            "минеральные удобрения",
            "нефтепродукты",
            "Нормы_расходов_грузов_ввоза_для_нужд_населения_товары_народного_потребления",
            "Нормы_расходов_грузов_каменный уголь",
            "Нормы_расходов_грузов_лесоматериалы",
            "Нормы расходов грузов ввоза для нужд сельского хозяйства минеральные удобрения",
            "Нормы расходов грузов ввоза нефтепродукты",

            "транзитные грузы от А к Б_каменный уголь",
            "транзитные грузы от А к Б_руда",
            "транзитные грузы от А к Б_нефтепродукты",
            "грузы капитального строительства",
            "транзитные грузы от А к Б_товары_народного_потребления",
            "транзитные грузы от А к Б_прочие грузы",

            "транзитные грузы от Б к А_Металлы",
            "транзитные грузы от Б к А_Машины и металлоизделия",
            "транзитные грузы от Б к А_Лесоматериалы",
            "транзитные грузы от Б к А_Сельскохозяйственные грузы",
            "транзитные грузы от Б к А_ минеральные удобрения",
            "транзитные грузы от Б к А_прочие грузы"
                    ]:
                # Преобразуем числовые значения
                try:
                    if field in [
              "Коэффициент подвижности населения",
            "Продукция заводской деревообработки",
            "Радиус кривой на участке поворота",
            "Ежегодный прирост численности сельского населения",

            "Ежегодный прирост численности городского населения",
            "Лесосырье",
                        "зерно, урожайность на расчетный срок",
            "зерно, потребность для городского населения",
                    "зерно, потребность для сельского населения",
                        "урожайность фуража на расчетный срок",
                        "картофель, урожайность на расчетный срок",
                        "картофель, потребность для городского населения",
                        "картофель, потребность для сельского населения",
                        "Крупный_рогатый_скот_потребность_в_картофеле",
                        "Крупный_рогатый_скот_потребность_в_фуражном_зерне",
                        "Мелкий_скот_потребность_в_картофеле",
                        "Мелкий_скот_потребность_в_фуражном_зерне",
                        "Нормы_высева_на_расчетный_срок_зерновые_продовольственные",
                        "Нормы_высева_на_расчетный_срок_зерновые_фуражные",
                        "Нормы_высева_на_расчетный_срок_картофель",
                        "товары народного потребления",
                        "каменный уголь",
                        "лесоматериалы",
                        "минеральные удобрения",
                        "нефтепродукты",
                        "Нормы_расходов_грузов_ввоза_для_нужд_населения_товары_народного_потребления",
                        "Нормы_расходов_грузов_каменный уголь",
                        "Нормы_расходов_грузов_лесоматериалы",
                        "Нормы расходов грузов ввоза для нужд сельского хозяйства минеральные удобрения",
                        "Нормы расходов грузов ввоза нефтепродукты",

                        "транзитные грузы от А к Б_каменный уголь",
                        "транзитные грузы от А к Б_руда",
                        "транзитные грузы от А к Б_нефтепродукты",
                        "грузы капитального строительства",
                        "транзитные грузы от А к Б_товары_народного_потребления",
                        "транзитные грузы от А к Б_прочие грузы",

                        "транзитные грузы от Б к А_Металлы",
                        "транзитные грузы от Б к А_Машины и металлоизделия",
                        "транзитные грузы от Б к А_Лесоматериалы",
                        "транзитные грузы от Б к А_Сельскохозяйственные грузы",
                        "транзитные грузы от Б к А_ минеральные удобрения",
                        "транзитные грузы от Б к А_прочие грузы"
                    ]:
                        value = float(value)
                    else:
                        value = int(value)
                except ValueError:
                    messagebox.showwarning("Ошибка", f"Введите корректное числовое значение для {field}.")
                    return

            data[field] = value  # Сохраняем значения

        # Формирование итоговых переменных
        self.gg = 9.81
        self.yvodi = 1

        ## ВВодимс
        self.pepls_movie = data["Коэффициент подвижности населения"]
        self.townsman_up = data["Ежегодный прирост численности городского населения"]
        self.villager_up = data["Ежегодный прирост численности сельского населения"]
        self.villager_p = data["Плотность сельского населения на  отчетный год"]
        self.townsman_station = data["Количество городского населения на отчетный год в районе одной из станций с грузовыми операциями"]

        self.wood = data["Лесосырье"]
        self.wood_products = data["Продукция заводской деревообработки"]
        self.Lumber = data["Пиломатериалы"]
        self.roundwood = data["Круглый лес"]
        self.firewood = data["дрова и другие отходы"]

        self.plot_posevov = data["Плотность посевов на отчетный год"]
        self.ves_posevov = data["удельный вес посевных площадей зерна на отчетный срок"]
        self.up_posevov = data["прирост посевных площадей зерна на расчетный срок"]
        self.eat_posevov = data["зерно, урожайность на расчетный срок"]
        self.need_city_wheat = data["зерно, потребность для городского населения"]
        self.need_vilolage_wheat = data["зерно, потребность для сельского населения"]
        self.ves_furaj = data["зерновые фуражные культуры, удельный вес на отчетный срок"]
        self.up_furaj = data["прирост фуража на расчетный срок"]
        self.eat_furaj = data["урожайность фуража на расчетный срок"]
        self.ves_potato = data["картофель, удельный вес на расчетный срок"]
        self.up_potato = data["картофель, прирост на расчетный срок"]
        self.eat_potato = data["картофель, урожайность на расчетный срок"]
        self.need_city_potato = data["картофель, потребность для городского населения"]
        self.need_village_potato = data["картофель, потребность для сельского населения"]
        self.ves_culture = data["остальные культуры, удельный вес на расчетный срок"]
        self.up_culture = data["остальные культуры, прирост на расчетный срок"]

        self.cattle_potato = data["Крупный_рогатый_скот_потребность_в_картофеле"]
        self.cattle_furaj  = data["Крупный_рогатый_скот_потребность_в_фуражном_зерне"]
        self.small_cattle_potato = data["Мелкий_скот_потребность_в_картофеле"]
        self.small_cattle_furaj = data["Мелкий_скот_потребность_в_фуражном_зерне"]
        self.norma_viseva_wheat = data["Нормы_высева_на_расчетный_срок_зерновые_продовольственные"]
        self.norma_viseva_furaj = data["Нормы_высева_на_расчетный_срок_зерновые_фуражные"]
        self.norma_viseva_potato = data["Нормы_высева_на_расчетный_срок_картофель"]
        self.for_people = data["товары народного потребления"]
        self.coal = data["каменный уголь"]
        self.timber = data["лесоматериалы"]
        self.minerals = data["минеральные удобрения"]
        self.oil = data["нефтепродукты"]
        self.cattle_large = data["Крупный_рогатый_скот_количество_на_100_га"]
        self.cattle_small  = data["Мелкий_скот_количество_на_100_га"]

        self.prirost_large = data["Крупный_рогатый_скот_прирост"]
        self.prirost_small  = data["Мелкий_скот_скот_прирост_скот"]




        second = second_(pepls_movie = self.pepls_movie, townsman_up =  self.townsman_up, villager_up = self.villager_up, villager_p = self.villager_p, townsman_station = self.townsman_station, wood=self.wood, wood_products=self.wood_products, Lumber=self.Lumber, roundwood=self.roundwood, firewood=self.firewood)

        third = third_(plot_posevov = self.plot_posevov, ves_posevov =  self.ves_posevov, up_posevov = self.up_posevov, eat_posevov = self.eat_posevov, need_city_wheat = self.need_city_wheat, need_vilolage_wheat=self.need_vilolage_wheat, ves_furaj=self.ves_furaj, up_furaj=self.up_furaj, eat_furaj=self.eat_furaj, ves_potato=self.ves_potato, up_potato=self.up_potato, eat_potato=self.eat_potato, need_city_potato=self.need_city_potato, need_village_potato=self.need_village_potato, ves_culture=self.ves_culture, up_culture=self.up_culture)
        self.up_A_posevov = round(self.ves_posevov * third.Fct_pos_A / 100 * (1+ self.up_posevov / 100) , 2)
        self.up_A_furaj = round(self.ves_furaj * third.Fct_pos_A / 100 * (1+ self.up_furaj / 100), 2)
        self.up_A_potato = round(self.ves_potato * third.Fct_pos_A / 100 * (1+ self.up_potato / 100), 2)
        self.up_A_culture = round(self.ves_culture * third.Fct_pos_A / 100 * (1+ self.up_culture / 100), 2)

        self.up_B_posevov = round(self.ves_posevov * third.Fct_pos_B / 100 * (1+self.up_posevov / 100), 2)
        self.up_B_furaj = round(self.ves_furaj * third.Fct_pos_B / 100 * (1+self.up_furaj / 100), 2)
        self.up_B_potato = round(self.ves_potato * third.Fct_pos_B / 100 * (1+self.up_potato / 100), 2)
        self.up_B_culture = round(self.ves_culture * third.Fct_pos_B / 100 * (1+self.up_culture / 100), 2)

        self.A_all_ploshad= self.up_A_posevov + self.up_A_furaj + self.up_A_potato + self.up_A_culture
        self.B_all_ploshad  = self.up_B_posevov + self.up_B_furaj +self.up_B_potato + self.up_B_culture

        self.val_B_potato = self.up_B_potato * self.eat_potato
        self.val_B_furaj = self.up_B_furaj * self.eat_furaj
        self.val_B_posevov = self.up_B_posevov * self.eat_posevov
        self.val_B_all = self.val_B_potato+self.val_B_furaj+self.val_B_posevov

        self.val_A_potato = self.up_A_potato * self.eat_potato
        self.val_A_furaj = self.up_A_furaj * self.eat_furaj
        self.val_A_posevov = self.up_A_posevov * self.eat_posevov
        self.val_A_all = self.val_A_potato+self.val_A_furaj+self.val_A_posevov

        self.potreblenie_A_City_wheat = self.need_city_wheat * round(second.ApTownman, 2)
        self.potreblenie_A_City_potato = self.need_city_potato * round(second.ApTownman, 2)
        self.potreblenie_A_Vill_wheat = self.need_vilolage_wheat * round(second.ApVillA / 1000, 2)
        self.potreblenie_A_Vill_potato = self.need_village_potato * round(second.ApVillA / 1000, 2)
        self.potreblenie_B_Vill_wheat = self.need_vilolage_wheat * round(second.ApVillB / 1000, 2)
        self.potreblenie_B_Vill_potato = self.need_village_potato * round(second.ApVillB / 1000, 2)

        self.sigma_potreblenie_B = self.potreblenie_B_Vill_wheat + self.potreblenie_B_Vill_potato

        self.sigma_wheat_A = self.potreblenie_A_City_wheat + self.potreblenie_A_Vill_wheat
        self.sigma_potato_A = self.potreblenie_A_City_potato + self.potreblenie_A_Vill_potato
        self.sigma_potreblenie_A = self.sigma_wheat_A + self.sigma_potato_A

        self.cattle_small_A=second.FctA*self.cattle_small/1000
        self.cattle_small_B = second.FctB * self.cattle_small/1000
        self.cattle_large_A=second.FctA*self.cattle_large/1000
        self.cattle_large_B = second.FctB * self.cattle_large/1000

        self.Nras_A_large=self.cattle_large_A*(1+self.prirost_large/100)
        self.Nras_A_small=self.cattle_small_A*(1+self.prirost_small/100)
        self.Nras_B_large=self.cattle_large_B*(1+self.prirost_large/100)
        self.Nras_B_small=self.cattle_small_B*(1+self.prirost_small/100)

        self.potreba_large_furaj_A=round(self.Nras_A_large,2)*self.cattle_furaj
        self.potreba_large_potato_A=round(self.Nras_A_large,2)*self.cattle_potato
        self.potreba_large_furaj_B=round(self.Nras_B_large,2)*self.cattle_furaj
        self.potreba_large_potato_B=round(self.Nras_B_large,2)*self.cattle_potato

        self.potreba_small_furaj_A=round(self.Nras_A_small,2)*self.small_cattle_furaj
        self.potreba_small_potato_A=round(self.Nras_A_small,2)*self.small_cattle_potato
        self.potreba_small_furaj_B=round(self.Nras_B_small,2)*self.small_cattle_furaj
        self.potreba_small_potato_B=round(self.Nras_B_small,2)*self.small_cattle_potato

        self.sigma_B_potato = self.potreba_small_potato_B+self.potreba_large_potato_B
        self.sigma_B_furaj = self.potreba_small_furaj_B + self.potreba_large_furaj_B
        self.sigma_A_potato = self.potreba_small_potato_A+self.potreba_large_potato_A
        self.sigma_A_furaj = self.potreba_small_furaj_A + self.potreba_large_furaj_A

        self.sigma_B_potato_furaj = self.sigma_B_potato + self.sigma_B_furaj
        self.sigma_A_furaj_potato = self.sigma_A_potato + self.sigma_A_furaj

        self.semen_A_wheat = self.up_A_posevov * self.norma_viseva_wheat
        self.semen_A_potato = self.up_A_potato * self.norma_viseva_potato
        self.semen_A_furaj = self.up_A_furaj * self.norma_viseva_furaj
        self.semen_B_wheat = self.up_A_posevov * self.norma_viseva_wheat
        self.semen_B_potato = self.up_A_potato * self.norma_viseva_potato
        self.semen_B_furaj = self.up_B_furaj * self.norma_viseva_furaj

        self.sigma_A_semen = self.semen_A_wheat + self.semen_A_potato + self.semen_A_furaj
        self.sigma_B_semen = self.semen_B_wheat + self.semen_B_potato + self.semen_B_furaj

        self.fond_A_all = self.val_A_potato*0.1+self.val_A_furaj*0.1+self.val_A_posevov*0.1
        self.fond_B_all = self.val_A_potato*0.1+self.val_A_furaj*0.1+self.val_A_posevov*0.1

        self.results_A_wheat = self.semen_A_wheat + self.sigma_wheat_A + self.val_A_posevov*0.1
        self.results_A_furaj = self.semen_A_furaj + self.sigma_A_furaj + self.val_A_furaj*0.1
        self.results_A_potato = self.semen_A_potato + self.sigma_A_potato + self.val_A_potato*0.1
        self.first_sigma_A = self.results_A_wheat + self.results_A_furaj + self.results_A_potato

        self.results_B_wheat = self.semen_B_wheat + self.potreblenie_B_Vill_wheat + self.val_B_posevov * 0.1
        self.results_B_furaj = self.semen_B_furaj + self.sigma_B_furaj + self.val_B_furaj * 0.1
        self.results_B_potato = self.semen_B_potato + self.sigma_B_potato + self.val_B_potato * 0.1
        self.first_sigma_B = self.results_B_wheat + self.results_B_furaj + self.results_B_potato

        self.overage_A_potato = self.val_A_potato - self.results_A_potato
        self.overage_A_furaj = self.val_A_furaj - self.results_A_furaj
        self.overage_A_wheat = self.val_A_posevov - self.results_A_wheat
        self.overage_B_furaj = self.val_B_furaj - self.results_B_furaj
        self.overage_B_potato = self.val_B_potato - self.results_B_potato
        self.overage_B_wheat = self.potreblenie_B_Vill_wheat - self.results_B_wheat

        messagebox.showinfo("Успех", "Все данные успешно введены!")

        self.root.destroy()  # Закрыть окно после подтверждения
        self.root.quit()  # Завершает цикл обработки событий Tkinter

        """autocader_1 = AutoCADLines_1(100,
                                     H=self.H_1,
                                     hp=self.hp,
                                     yb=self.yb,
                                     ynuv=self.ynuv,
                                     B=self.B,
                                     W10=self.W10,
                                     β=self.β,
                                     Vv=self.Vv,
                                     h10=self.h10,
                                     h20=self.h20,
                                     lam=self.λ,
                                     D=self.D,
                                     hn=self.hn,
                                     deltaZ=self.ΔZ,
                                     yyk=self.yyk,
                                     Qk=self.Qk,
                                     Dcp=self.Dcp,
                                     t1=self.t1,
                                     qk=self.qk,
                                     dcp=self.dcp,
                                     t2=self.t2,
                                     Vd=self.Vd,
                                     a=self.a,
                                     Ksh=self.Ksh,
                                     Knag=self.Knag,
                                     k3=self.k3,
                                     n=self.nn,
                                     arrow_size=0.5)
        autocader_1.draw_lines()"""

        # Создание экземпляра WordEquationReplacer
        replacer = WordEquationReplacer('Praktich_zanyatie_2_EI.docx',
                                        AoVillA=f'{round(second.AoVillA/1000,2)}',
                                        AoVillB = f'{round(second.AoVillB/1000,2)}',
                                        AoTownmanA=f'{self.townsman_station}',
                                        ApVillA = f'{round(second.ApVillA/1000,2)}',
                                        ApVillB = f'{round(second.ApVillB/1000,2)}',
                                        ApCityman=f'{round(second.ApTownman,2)}',
                                        pepleA=f'{round(second.pepleA/1000,2)}',
                                        pepleB=f'{round(second.pepleB/1000,2)}',
                                        true_wood=f'{second.true_wood*1000}',
                                        true_wood_products=f'{round(second.true_wood_products*1000,2)}',
                                        true_Lumber=f'{round(second.true_Lumber*1000,2)}',
                                        true_roundwood=f'{round(second.true_roundwood*1000,2)}',
                                        true_firewood=f'{round(second.true_firewood*1000,2)}',
                                        ALL=f'{round(second.all*1000,2)}'
                                                                )

        replacer1 = WordEquationReplacer('Praktich_zanyatie_3_EI.docx',
                                        ves_posevov=f'{self.ves_posevov}',
                                        ves_furaj=f'{self.ves_furaj}',
                                        ves_potato=f'{self.ves_potato}',
                                        ves_culture=f'{self.ves_culture}',

                                        up_posevov=f'{self.up_posevov}',
                                        up_furaj=f'{self.up_furaj}',
                                        up_potato=f'{self.up_potato}',
                                        up_culture=f'{self.up_culture}',

                                        ploshad_posevov=f'{round(self.ves_posevov*third.Fct_pos_A / 100,2)}',
                                        ploshad_furaj=f'{round(self.ves_furaj*third.Fct_pos_A / 100,2)}',
                                        ploshad_potato=f'{round(self.ves_potato*third.Fct_pos_A / 100, 2)}',
                                        ploshad_culture=f'{round(self.ves_culture*third.Fct_pos_A / 100,2)}',

                                        B_ploshad_posevov=f'{round(self.ves_posevov * third.Fct_pos_B / 100, 2)}',
                                        B_ploshad_furaj=f'{round(self.ves_furaj * third.Fct_pos_B / 100, 2)}',
                                        B_ploshad_potato=f'{round(self.ves_potato * third.Fct_pos_B / 100, 2)}',
                                        B_ploshad_culture=f'{round(self.ves_culture * third.Fct_pos_B / 100, 2)}',

                                         up_A_posevov=f'{self.up_A_posevov}',
                                         up_A_furaj=f'{self.up_A_furaj}',
                                         up_A_potato=f'{self.up_A_potato}',
                                         up_A_culture=f'{self.up_A_culture}',

                                         up_B_posevov=f'{self.up_B_posevov}',
                                         up_B_furaj=f'{self.up_B_furaj}',
                                         up_B_potato=f'{self.up_B_potato}',
                                         up_B_culture=f'{self.up_B_culture}',




                                        Fct_pos_B=f'{round(third.Fct_pos_B,2)}',
                                        Fct_pos_A=f'{round(third.Fct_pos_A,2)}',

                                         A_all=f'{round(self.A_all_ploshad,2)}',
                                         B_all=f'{round(self.B_all_ploshad,2)}',

                                         eat_posevov=f'{self.eat_posevov}',
                                         eat_furaj=f'{self.eat_furaj}',
                                         eat_potato=f'{self.eat_potato}',

                                         val_B_potato=f'{round(self.val_B_potato,2)}',
                                         val_B_furaj=f'{round(self.val_B_furaj,2)}',
                                         val_B_posevov=f'{round(self.val_B_posevov,2)}',
                                         val_B_all=f'{round(self.val_B_all,2)}',

                                         val_A_potato=f'{round(self.val_A_potato,2)}',
                                         val_A_furaj=f'{round(self.val_A_furaj,2)}',
                                         val_A_posevov=f'{round(self.val_A_posevov,2)}',
                                         val_A_all=f'{round(self.val_A_all,2)}'

                                        )

        replacer2 = WordEquationReplacer('Praktich_zanyatie_4_EI.docx',
                                         ApVillA=f'{round(second.ApVillA / 1000, 2)}',
                                         ApVillB=f'{round(second.ApVillB / 1000, 2)}',
                                         ApCityman=f'{round(second.ApTownman, 2)}',

                                         NormaVillPotato=f'{self.need_village_potato}',
                                         NormaVillWheat=f'{self.need_vilolage_wheat}',
                                         NormaCitymanWheat=f'{self.need_city_wheat}',
                                         NormaCitymanPotato=f'{self.need_city_potato}',

                                         potreblenie_A_City_wheat=f'{round(self.potreblenie_A_City_wheat,2)}',
                                         potreblenie_A_City_potato=f'{round(self.potreblenie_A_City_potato,2)}',
                                         potreblenie_A_Vill_wheat=f'{round(self.potreblenie_A_Vill_wheat,2)}',
                                         potreblenie_A_Vill_potato=f'{round(self.potreblenie_A_Vill_potato,2)}',
                                         potreblenie_B_Vill_wheat = f'{round(self.potreblenie_B_Vill_wheat,2)}',
                                         potreblenie_B_Vill_potato = f'{round(self.potreblenie_B_Vill_potato,2)}',

                                         sigma_potato_A = f'{round(self.sigma_potato_A,2)}',
                                         sigma_wheat_A=f'{round(self.sigma_wheat_A,2)}',

                                         cattle_large=f'{round(self.cattle_large,2)}',
                                         cattle_small=f'{round(self.cattle_small,2)}',
                                         cattle_small_A=f'{round(self.cattle_small_A,2)}',
                                         cattle_small_B=f'{round(self.cattle_small_B,2)}',
                                         cattle_large_A=f'{round(self.cattle_large_A,2)}',
                                         cattle_large_B=f'{round(self.cattle_large_B,2)}',
                                         FA=f'{round(second.FctA,2)}',
                                         FB=f'{round(second.FctB,2)}',
                                         prirost_large=f'{self.prirost_large}',
                                         prirost_small=f'{self.prirost_small}',
                                         Nras_A_large=f'{round(self.Nras_A_large,2)}',
                                         Nras_A_small=f'{round(self.Nras_A_small,2)}',
                                         Nras_B_large=f'{round(self.Nras_B_large,2)}',
                                         Nras_B_small=f'{round(self.Nras_B_small,2)}',

                                         cattle_potato = f'{self.cattle_potato}',
                                         cattle_furaj  = f'{self.cattle_furaj}',
                                         small_cattle_potato = f'{self.small_cattle_potato}',
                                         small_cattle_furaj = f'{self.small_cattle_furaj}',

                                         potreba_large_furaj_A = f'{round(self.potreba_large_furaj_A,2)}',
                                         potreba_large_potato_A=f'{round(self.potreba_large_potato_A,2)}',
                                         potreba_large_furaj_B=f'{round(self.potreba_large_furaj_B,2)}',
                                         potreba_large_potato_B=f'{round(self.potreba_large_potato_B,2)}',

                                         potreba_small_furaj_A=f'{round(self.potreba_small_furaj_A,2)}',
                                         potreba_small_potato_A=f'{round(self.potreba_small_potato_A,2)}',
                                         potreba_small_furaj_B=f'{round(self.potreba_small_furaj_B,2)}',
                                         potreba_small_potato_B=f'{round(self.potreba_small_potato_B,2)}',

                                         sigma_A_furaj = f'{round(self.sigma_A_furaj,2)}',
                                         sigma_A_potato=f'{round(self.sigma_A_potato,2)}',
                                         sigma_B_furaj=f'{round(self.sigma_B_furaj,2)}',
                                         sigma_B_potato=f'{round(self.sigma_B_potato,2)}',

                                         norma_viseva_potato=f'{self.norma_viseva_potato}',
                                         norma_viseva_furaj=f'{self.norma_viseva_furaj}',
                                         norma_viseva_wheat=f'{self.norma_viseva_wheat}',

                                         semen_A_wheat=f'{round(self.semen_A_wheat,2)}',
                                         semen_A_potato=f'{round(self.semen_A_potato,2)}',
                                         semen_A_furaj=f'{round(self.semen_A_furaj,2)}',
                                         semen_B_wheat=f'{round(self.semen_B_wheat,2)}',
                                         semen_B_potato=f'{round(self.semen_B_potato,2)}',
                                         semen_B_furaj=f'{round(self.semen_B_furaj,2)}',
                                         sigma_A_semen=f'{round(self.sigma_A_semen,2)}',
                                         sigma_B_semen=f'{round(self.sigma_B_semen,2)}',

                                         up_A_posevov=f'{self.up_A_posevov}',
                                         up_A_furaj=f'{self.up_A_furaj}',
                                         up_A_potato=f'{self.up_A_potato}',
                                         up_A_together=f'{self.up_A_culture}',

                                         up_B_posevov=f'{self.up_B_posevov}',
                                         up_B_furaj=f'{self.up_B_furaj}',
                                         up_B_potato=f'{self.up_B_potato}',
                                         up_B_together=f'{self.up_B_culture}',

                                         eat_posevov=f'{self.eat_posevov}',
                                         eat_potato=f'{self.eat_potato}',
                                         eat_furaj=f'{self.eat_furaj}',

                                         val_A_potato=f'{round(self.val_A_potato,2)}',
                                         val_A_furaj=f'{round(self.val_A_furaj,2)}',
                                         val_A_posevov=f'{round(self.val_A_posevov,2)}',
                                         val_B_all=f'{round(self.val_A_posevov,2)}',

                                         val_B_posevov=f'{round(self.val_B_posevov,2)}',
                                         val_B_furaj=f'{round(self.val_B_furaj,2)}',
                                         val_B_potato=f'{round(self.val_B_potato,2)}',
                                         val_A_all=f'{round(self.val_B_potato,2)}',

                                         sigma_potreblenie_B=f'{round(self.sigma_potreblenie_B,2)}',
                                         sigma_potreblenie_A=f'{round(self.sigma_potreblenie_A,2)}',

                                         sigma_B_potato_furaj=f'{round(self.sigma_B_potato_furaj,2)}',
                                         sigma_A_furaj_potato=f'{round(self.sigma_A_furaj_potato,2)}',

                                         fond_B_posevov=f'{round(self.val_B_potato*0.1,2)}',
                                         fond_B_furaj=f'{round(self.val_B_furaj*0.1,2)}',
                                         fond_B_potato=f'{round(self.val_B_posevov*0.1,2)}',
                                         fond_B_all=f'{round(self.fond_B_all,2)}',
                                         fond_A_posevov=f'{round(self.val_A_potato*0.1,2)}',
                                         fond_A_furaj=f'{round(self.val_A_furaj*0.1,2)}',
                                         fond_A_potato=f'{round(self.val_A_posevov*0.1,2)}',
                                         fond_A_all=f'{round(self.fond_A_all,2)}',

                                         results_A_wheat=f'{round(self.results_A_wheat,2)}',
                                         results_A_furaj=f'{round(self.results_A_furaj,2)}',
                                         results_A_potato=f'{round(self.results_A_potato,2)}',
                                         first_sigma_A=f'{round(self.first_sigma_A,2)}',

                                         results_B_wheat=f'{round(self.results_B_wheat,2)}',
                                         results_B_furaj=f'{round(self.results_B_furaj,2)}',
                                         results_B_potato=f'{round(self.results_B_potato,2)}',
                                         first_sigma_B=f'{round(self.first_sigma_B,2)}',

                                         overage_A_potato=f'{round(self.overage_A_potato,2)}',
                                         overage_A_furaj=f'{round(self.overage_A_furaj,2)}',
                                         overage_A_wheat=f'{round(self.overage_A_wheat,2)}',
                                         overage_B_furaj=f'{round(self.overage_B_furaj,2)}',
                                         overage_B_potato=f'{round(self.overage_B_potato,2)}',
                                         overage_B_wheat=f'{round(self.overage_B_wheat,2)}',
                                         )

        replacer3=WordEquationReplacer('Praktich_zanyatie_5_EI.docx',

                                         pepleA=f'{round(second.pepleA / 1000, 2)}',
                                         pepleB=f'{round(second.pepleB / 1000, 2)}',

                                            )

        """autocader = AutoCADLines(100, horda=self.horda_2d,
                                 d=self.d,
                                 dsqrt=self.dsqrt,
                                 H=self.H,
                                 n=self.n,
                                 hbm=self.hбм,
                                 wl=self.WL,
                                 kt=self.Kt,
                                 b0=b0,  # Нужно убедиться, что это определено
                                 t=self.t,
                                 tsqrt=self.t**2,
                                 α=self.α_degrees,  # Преобразование в градусы
                                 RR=self.R,
                                 μ=self.μ_degrees,
                                 XR=self.XR,
                                 YR=self.YR,
                                 NN=self.N,
                                 b1=self.b1,

                                 xn1=self.xn1,
                                 xn2=self.xn2,
                                 xn3=self.xn3,
                                 xn4=self.xn4,
                                 xn5=self.xn5,
                                 xn6=self.xn6,
                                 xn7=self.xn6,

                                 xcp0=self.xcp0,
                                 xcp1=self.xcp1,
                                 xcp2=self.xcp2,
                                 xcp3=self.xcp3,
                                 xcp4=self.xcp4,
                                 xcp5=self.xcp5,
                                 xcp6=self.xcp6,

                                 β0=math.degrees(self.β0),
                                 β1=math.degrees(self.β1),
                                 β2=math.degrees(self.β2),
                                 β3=math.degrees(self.β3),
                                 β4=math.degrees(self.β4),
                                 β5=math.degrees(self.β5),
                                 β6=math.degrees(self.β6),

                                 yп1=self.yп1,
                                 yп2=self.yп2,
                                 yп3=self.yп3,
                                 yп4=self.yп4,
                                 yп5=self.yп5,
                                 yп6=self.yп6,
                                 yп7=self.yп7,

                                 yk0=self.yk0,
                                 yk1=self.yk1,
                                 yk2=self.yk2,
                                 yk3=self.yk3,
                                 yk4=self.yk4,
                                 yk5=self.yk5,
                                 yk6=self.yk6,

                                 h0=self.h0,
                                 h1=self.h1,
                                 h2=self.h2,
                                 h3=self.h3,
                                 h4=self.h4,
                                 h5=self.h5,
                                 h6=self.h6,

                                 g0=self.g0,
                                 g1=self.g1,
                                 g2=self.g2,
                                 g3=self.g3,
                                 g4=self.g4,
                                 g5=self.g5,
                                 g6=self.g6,

                                 hhmax2=max(self.h0, self.h1, self.h2, self.h3, self.h4, self.h5, self.h6) ** 2,
                                 km=self.Km,
                                 xnmax=max(self.xn1, self.xn2, self.xn3, self.xn4, self.xn5, self.xn6, self.xn7),
                                 arrow_size=0.5)
        autocader.draw_lines()"""

        # Обработка документа
        replacer.process_document()
        replacer.save_document('Черновая_заготовка_1.docx')
        replacer1.process_document()
        replacer1.save_document('Черновая_заготовка_2.docx')
        replacer2.process_document()
        replacer2.save_document('Черновая_заготовка_3.docx')
        replacer3.process_document()
        replacer3.save_document('Черновая_заготовка_4.docx')

# Запуск приложения
if __name__ == "__main__":
    root = tk.Tk()

    app = App(root)
    root.mainloop()