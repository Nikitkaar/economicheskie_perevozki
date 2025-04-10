import math
from tkinter import messagebox
import tkinter as tk
from tkinter import ttk

import numpy as np
from pyautocad import Autocad, APoint  # Импортируйте необходимые классы из библиотеки Autocad

from initial_data import WordEquationReplacer, second_, third_
from  drawing import  AutoCADLinesPlacer
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


                        "Нормы_расходов_грузов_ввоза_для_нужд_населения_товары_народного_потребления",
                        "Нормы_расходов_грузов_каменный уголь",
                        "Нормы_расходов_грузов_лесоматериалы",
                        "Нормы расходов грузов ввоза для нужд сельского хозяйства минеральные удобрения",
                        "Нормы расходов грузов ввоза нефтепродукты",
                        "товары народного потребления",

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
        self.pepls_movie = data.get("Коэффициент подвижности населения")
        self.townsman_up = data.get("Ежегодный прирост численности городского населения")
        self.villager_up = data.get("Ежегодный прирост численности сельского населения")
        self.villager_p = data.get("Плотность сельского населения на  отчетный год")
        self.townsman_station = data.get("Количество городского населения на отчетный год в районе одной из станций с грузовыми операциями")

        self.wood = data.get("Лесосырье")
        self.wood_products = data.get("Продукция заводской деревообработки")
        self.Lumber = data.get("Пиломатериалы")
        self.roundwood = data.get("Круглый лес")
        self.firewood = data.get("дрова и другие отходы")

        self.plot_posevov = data.get("Плотность посевов на отчетный год")
        self.ves_posevov = data.get("удельный вес посевных площадей зерна на отчетный срок")
        self.up_posevov = data.get("прирост посевных площадей зерна на расчетный срок")
        self.eat_posevov = data.get("зерно, урожайность на расчетный срок")
        self.need_city_wheat = data.get("зерно, потребность для городского населения")
        self.need_vilolage_wheat = data.get("зерно, потребность для сельского населения")
        self.ves_furaj = data.get("зерновые фуражные культуры, удельный вес на отчетный срок")
        self.up_furaj = data.get("прирост фуража на расчетный срок")
        self.eat_furaj = data.get("урожайность фуража на расчетный срок")
        self.ves_potato = data.get("картофель, удельный вес на расчетный срок")
        self.up_potato = data.get("картофель, прирост на расчетный срок")
        self.eat_potato = data.get("картофель, урожайность на расчетный срок")
        self.need_city_potato = data.get("картофель, потребность для городского населения")
        self.need_village_potato = data.get("картофель, потребность для сельского населения")
        self.ves_culture = data.get("остальные культуры, удельный вес на расчетный срок")
        self.up_culture = data.get("остальные культуры, прирост на расчетный срок")

        self.cattle_potato = data.get("Крупный_рогатый_скот_потребность_в_картофеле")
        self.cattle_furaj  = data.get("Крупный_рогатый_скот_потребность_в_фуражном_зерне")
        self.small_cattle_potato = data.get("Мелкий_скот_потребность_в_картофеле")
        self.small_cattle_furaj = data.get("Мелкий_скот_потребность_в_фуражном_зерне")
        self.norma_viseva_wheat = data.get("Нормы_высева_на_расчетный_срок_зерновые_продовольственные")
        self.norma_viseva_furaj = data.get("Нормы_высева_на_расчетный_срок_зерновые_фуражные")
        self.norma_viseva_potato = data.get("Нормы_высева_на_расчетный_срок_картофель")

        self.cattle_large = data.get("Крупный_рогатый_скот_количество_на_100_га")
        self.cattle_small  = data.get("Мелкий_скот_количество_на_100_га")

        self.prirost_large = data.get("Крупный_рогатый_скот_прирост")
        self.prirost_small  = data.get("Мелкий_скот_скот_прирост_скот")

        self.norma_rashoda_tovar_potrebl = data.get("Нормы_расходов_грузов_ввоза_для_нужд_населения_товары_народного_потребления")
        self.norma_rashoda_tovar_coal = data.get("Нормы_расходов_грузов_каменный уголь")
        self.norma_rashoda_forest = data.get("Нормы_расходов_грузов_лесоматериалы")
        self.norma_rashoda_village = data.get("Нормы расходов грузов ввоза для нужд сельского хозяйства минеральные удобрения")
        self.norma_rashoda_oil = data.get("Нормы расходов грузов ввоза нефтепродукты")

        self.tranzit_AB_coal = data.get("транзитные грузы от А к Б_каменный уголь")
        self.tranzit_AB_rock = data.get("транзитные грузы от А к Б_руда")
        self.tranzit_AB_oil = data.get("транзитные грузы от А к Б_нефтепродукты")
        self.tranzit_AB_build = data.get("грузы капитального строительства")
        self.tranzit_AB_tovar_narod = data.get("транзитные грузы от А к Б_товары_народного_потребления")
        self.tranzit_AB_prochie = data.get("транзитные грузы от А к Б_прочие грузы")

        self.tranzit_BA_Metall = data.get("транзитные грузы от Б к А_Металлы")
        self.tranzit_BA_car = data.get("транзитные грузы от Б к А_Машины и металлоизделия")
        self.tranzit_BA_wood = data.get("транзитные грузы от Б к А_Лесоматериалы")
        self.tranzit_BA_sh_hozyaystvo = data.get("транзитные грузы от Б к А_Сельскохозяйственные грузы")
        self.tranzit_BA_tovar_narod_mineral = data.get("транзитные грузы от Б к А_ минеральные удобрения")
        self.tranzit_BA_prochie = data.get("транзитные грузы от Б к А_прочие грузы")

        self.results_tabl_tranzit_vivoz = self.tranzit_BA_Metall + self.tranzit_BA_car + self.tranzit_BA_wood
        + self.tranzit_BA_sh_hozyaystvo + self.tranzit_BA_tovar_narod_mineral + self.tranzit_BA_prochie

        self.results_tabl_tranzit_vvoz = self.tranzit_AB_coal + self.tranzit_AB_prochie + self.tranzit_AB_rock + self.tranzit_AB_oil
        + self.tranzit_AB_build + self.tranzit_AB_tovar_narod

        self.wagons =  data = {
    "Полувагон_8": {
        "Число_осей": 8,
        "Грузоподъемность": 125,
        "Масса_тары": 43.7,
        "Длина": 20
    },
    "Полувагон_4": {
        "Число_осей": 4,
        "Грузоподъемность": 63,
        "Масса_тары": 21.9,
        "Длина": 14
    },
    "Цистерна_8": {
        "Число_осей": 8,
        "Грузоподъемность": 120,
        "Масса_тары": 51,
        "Длина": 21
    },
    "Цистерна_4": {
        "Число_осей": 4,
        "Грузоподъемность": 60,
        "Масса_тары": 21.7,
        "Длина": 12
    },
    "Крытый_вагон_4": {
        "Число_осей": 4,
        "Грузоподъемность": 62,
        "Масса_тары": 22.2,
        "Длина": 15
    },
    "Платформа_4": {
        "Число_осей": 4,
        "Грузоподъемность": 63,
        "Масса_тары": 20.8,
        "Длина": 14
    }
}

        def qbr(self, wagons, name, alfa):
            return round(wagons[name]["Масса_тары"] + alfa * wagons[name]["Грузоподъемность"],2)

        self.qbr_coal_4 = qbr(self, self.wagons,"Полувагон_4", 1)
        self.qbr_coal_8 =  qbr(self, self.wagons,"Полувагон_8",  1)

        self.qbr_oil_4 = qbr(self, self.wagons,"Цистерна_4", 1)
        self.qbr_oil_8 = qbr(self, self.wagons,"Цистерна_8", 1)

        self.qbr_Build_4 = qbr(self, self.wagons,"Крытый_вагон_4", 0.9)
        self.qbr_Build_4_ = qbr(self, self.wagons,"Платформа_4", 1)

        self.qbr_Metall_4 = qbr(self, self.wagons,"Платформа_4", 1)

        self.qbr_narod_potrebl_and_prochie_4 = qbr(self, self.wagons,"Крытый_вагон_4", 0.65)

        self.qbr_Mineral_4 = qbr(self, self.wagons,"Цистерна_4", 1)
        self.qbr_Mineral_8 = qbr(self, self.wagons,"Цистерна_8", 1)

        self.qbr_sh_4 = qbr(self, self.wagons,"Цистерна_4", 0.7)
        self.qbr_sh_4 = qbr(self, self.wagons,"Цистерна_4", 0.9)

        def β_4(self, fi4, qbr4, qbr8):
            return round((fi4*qbr4)/(fi4*qbr4+(1-fi4)*qbr8),2)

        def β_8(self, fi4, qbr4, qbr8):
            return round(((1-fi4)*qbr8)/(fi4*qbr4+(1-fi4)*qbr8),2)

        second = second_(pepls_movie = self.pepls_movie, townsman_up =  self.townsman_up,
                         villager_up = self.villager_up, villager_p = self.villager_p,
                         townsman_station = self.townsman_station, wood=self.wood, wood_products=self.wood_products,
                         Lumber=self.Lumber, roundwood=self.roundwood, firewood=self.firewood)

        third = third_(FctA=second.FctA, FctB=second.FctB,plot_posevov = self.plot_posevov,
                       ves_posevov =  self.ves_posevov, up_posevov = self.up_posevov, eat_posevov = self.eat_posevov,
                       need_city_wheat = self.need_city_wheat, need_vilolage_wheat=self.need_vilolage_wheat,
                       ves_furaj=self.ves_furaj, up_furaj=self.up_furaj, eat_furaj=self.eat_furaj,
                       ves_potato=self.ves_potato, up_potato=self.up_potato, eat_potato=self.eat_potato,
                       need_city_potato=self.need_city_potato, need_village_potato=self.need_village_potato,
                       ves_culture=self.ves_culture, up_culture=self.up_culture)



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





        self.vsego_A_forest = second.pepleA/1000 * self.norma_rashoda_forest
        self.vsego_A_coal = second.pepleA/1000 * self.norma_rashoda_tovar_coal
        self.vsego_A_tovar = second.pepleA/1000 * self.norma_rashoda_tovar_potrebl
        self.vsego_A_mineral = round(self.A_all_ploshad, 2) * self.norma_rashoda_village
        self.vsego_A_oil = round(self.A_all_ploshad, 2) * self.norma_rashoda_oil




        self.vsego_B_forest = second.pepleB/1000 * self.norma_rashoda_forest
        self.vsego_B_coal = second.pepleB/1000 * self.norma_rashoda_tovar_coal
        self.vsego_B_tovar = second.pepleB/1000 * self.norma_rashoda_tovar_potrebl
        self.vsego_B_mineral = round(self.B_all_ploshad, 2) * self.norma_rashoda_village
        self.vsego_B_oil = round(self.B_all_ploshad, 2) * self.norma_rashoda_oil

        def overager(x, v, c):
            imp = 0
            exp = 0
            if x < 0:
                imp = x
            else:
                exp = x
            if v < 0:
                imp += v
            else:
                exp += v
            if c < 0:
                imp += c
            else:
                exp += c
            return round(-imp,2), round(exp,2)  # Возвращаем значения imp и exp

        self.sh_V_vvoz, self.sh_V_vivoz = overager(self.overage_A_potato,self.overage_A_furaj,self.overage_A_wheat)

        self.sh_G_vvoz, self.sh_G_vivoz = overager(self.overage_B_potato,self.overage_B_furaj,self.overage_B_wheat)
        self.summ_vvoza_A = self.vsego_A_forest + self.vsego_A_coal + self.vsego_A_tovar + self.vsego_A_mineral + self.vsego_A_oil + self.sh_G_vvoz
        self.summ_vvoza = self.vsego_B_forest+self.vsego_B_coal+self.vsego_B_tovar+self.vsego_B_mineral+self.vsego_B_oil+self.sh_V_vvoz
        self.summ_summ = self.summ_vvoza + self.summ_vvoza_A
        self.t_ovar = self.vsego_A_tovar + self.vsego_B_tovar
        self.s_h = self.sh_G_vvoz + self.sh_V_vvoz
        self.m_ineral = self.vsego_B_mineral + self.vsego_A_mineral
        self.f_orest = self.vsego_B_forest + self.vsego_A_forest
        self.o_il = self.vsego_A_oil + self.vsego_B_oil
        self.c_oal = self.vsego_B_coal + self.vsego_A_coal






        self.itogo_vivoz_sh = self.sh_G_vivoz + self.sh_V_vivoz
        self.itogo_vivozzz_stolbik_BA = self.sh_V_vivoz + round(second.all, 2)

        self.itogo_vivoz_vsego = self.itogo_vivozzz_stolbik_BA + self.sh_G_vivoz

        self.Σ_BA_1 = (self.tranzit_BA_prochie + self.tranzit_BA_car + self.tranzit_BA_Metall +
                       self.tranzit_BA_wood + self.tranzit_BA_tovar_narod_mineral + self.tranzit_BA_sh_hozyaystvo)*1000
        self.Σ_BA_2 = self.Σ_BA_1 + self.sh_G_vivoz

        self.Σ_BA_3 = round(self.Σ_BA_2 + second.all + self.sh_V_vivoz, 2)

        self.Σ_AB_3 = (self.tranzit_AB_coal*1000 + self.tranzit_AB_oil*1000 + self.tranzit_AB_tovar_narod*1000
                       + self.tranzit_AB_rock*1000 + self.tranzit_AB_build*1000 + self.tranzit_AB_prochie*1000)
        self.Σ_AB_2 = (self.Σ_AB_3 + self.vsego_B_coal + self.vsego_B_oil + self.vsego_B_forest + self.vsego_B_mineral
                       + self.sh_V_vvoz + self.vsego_B_tovar)
        self.Σ_AB_1 = (self.Σ_AB_2 + self.vsego_A_coal + self.vsego_A_oil + self.vsego_A_forest +
                       self.vsego_A_mineral + self.sh_G_vvoz + self.vsego_A_tovar)

        self.l = 31.15

        self.Σ_ГрБА = round((self.Σ_BA_1 + self.Σ_BA_2 + self.Σ_BA_3) * self.l, 2)/1000

        self.Σ_ГрАБ = round((self.Σ_AB_3 + self.Σ_AB_1 + self.Σ_AB_2) * self.l, 2)/1000

        self.П1 = round(self.pepls_movie * second.pepleA/100000,2)
        self.П2 = round(second.pepleB * self.pepls_movie/100000,2)

        self.Пгр = round(2.5 * self.pepls_movie * (second.pepleA+second.pepleB)/100000 *  self.l * 2 / 1000,2)

        self.ГпрII = round((self.Σ_ГрБА + self.Пгр)/ (3*self.l),2)
        self.ГпрГВ = round((self.Σ_ГрАБ + self.Пгр)/ (3*self.l),2)

        self.mestnie_and_tranzit = round((self.results_tabl_tranzit_vvoz+self.results_tabl_tranzit_vivoz)*1000+self.summ_summ+self.itogo_vivoz_vsego, 2)

        self.vag4coal = round((self.tranzit_AB_coal * 1000 + self.c_oal)/(63), 2)
        self.vag4narod_potrebl = round(self.tranzit_AB_tovar_narod * 1000/(62),2)
        self.vag4oil = round((self.tranzit_AB_oil * 1000 + self.o_il)/(60),2)
        self.vag4build = round(self.tranzit_AB_build * 1000/(62),2)

        self.vag8coal = round((self.tranzit_AB_coal * 1000 + self.c_oal)/(125), 2)
        self.vag8oil = round((self.tranzit_AB_oil * 1000 + self.o_il)/(120),2)
        self.vag4_build = round(self.tranzit_AB_build * 1000/(63),2)

        self.bruttocoal4 = round(self.vag4coal * self.qbr_coal_4, 2)
        self.bruttocoal8 = round(self.vag8coal * self.qbr_coal_8, 2)
        self.bruttooil8 = round(self.vag8oil * self.qbr_oil_8, 2)
        self.bruttoBuild4 = round(self.vag4build*self.qbr_Build_4, 2)
        self.bruttooil4 = round(self.vag4oil * self.qbr_oil_4, 2)
        self.bruttoBuild4_ = round(self.vag4_build*self.qbr_Build_4_, 2)
        self.bruttoNarod4 = round(self.vag4narod_potrebl*self.qbr_narod_potrebl_and_prochie_4, 2)

        self.itogi = round((self.tranzit_BA_car * 1000) + (self.f_orest+second.all+self.tranzit_BA_wood*1000)+(self.tranzit_AB_coal*1000 + self.c_oal)
                                         + (self.tranzit_AB_oil*1000 + self.o_il) + (self.tranzit_AB_build*1000) + (self.tranzit_BA_Metall*1000)
                                         + (self.tranzit_AB_tovar_narod*1000) + (self.tranzit_BA_tovar_narod_mineral *1000+ self.m_ineral) + (self.tranzit_BA_sh_hozyaystvo*1000 + self.s_h + self.itogo_vivoz_sh), 2)

        self.itogvagona4 = round(self.vag4build+self.vag4oil+self.vag4build+self.vag4_build+self.vag4narod_potrebl+self.vag4coal,2)
        self.itogvagona8 = round(self.vag8coal+self.vag8oil,2)
        self.itogVesvagona4 = round(self.bruttooil4 + self.bruttoBuild4 + self.bruttoBuild4_ + self.bruttoNarod4,2)
        self.itogVesvagona8 = round(self.bruttocoal8 + self.bruttooil8,2)

        self.Vesvagona4=round(self.itogVesvagona4/self.itogvagona4, 2)
        self.Vesvagona8 =round(self.itogVesvagona8/self.itogvagona8, 2)
        self.β4 =round(self.itogVesvagona4/(self.itogVesvagona8+self.itogVesvagona4), 2)
        self.β8 =round(self.itogVesvagona8/(self.itogVesvagona8+self.itogVesvagona4), 2)
        self.fi8 = round(self.itogvagona4/(self.itogvagona4+self.itogvagona8),2)
        self.fi4 = round(self.itogvagona8/(self.itogvagona4+self.itogvagona8),2)
        self.teta = round((self.itogVesvagona8 + self.itogVesvagona4)/self.itogi,2)

        messagebox.showinfo("Успех", "Все данные успешно введены!")

        self.root.destroy()  # Закрыть окно после подтверждения
        self.root.quit()  # Завершает цикл обработки событий Tkinter

        autocader_1 = AutoCADLinesPlacer(

            sh_G_vivoz=f"{round(self.sh_G_vivoz, 2)}",
            sh_V_vivoz=f"{round(self.sh_V_vivoz, 2)}",
            tranzit_BA_sh_hozyaystvo=self.tranzit_BA_sh_hozyaystvo*1000,
            sh_BA_3=f"{round(self.sh_G_vivoz + self.sh_V_vivoz + self.tranzit_BA_sh_hozyaystvo*1000, 2)}",
            sh_BA_2=f"{round(self.sh_G_vivoz + self.tranzit_BA_sh_hozyaystvo*1000, 2)}",

             tranzit_BA_tovar_narod_mineral = self.tranzit_BA_tovar_narod_mineral*1000,

            forest_vivoz=f"{round(second.all, 2)}",
            tranzit_BA_wood=self.tranzit_BA_wood*1000,
            forest_BA_3=f"{self.tranzit_BA_wood*1000 + round(second.all, 2)}",

            tranzit_BA_Metall=self.tranzit_BA_Metall*1000,
            tranzit_BA_car=self.tranzit_BA_car*1000,
            tranzit_BA_prochie=self.tranzit_BA_prochie*1000,

            Σ_BA_1=f'{round(self.Σ_BA_1,2)}',
            Σ_BA_2=f'{round(self.Σ_BA_2,2)}',
            Σ_BA_3=f'{round(self.Σ_BA_3,2)}',

            tranzit_AB_coal=self.tranzit_AB_coal*1000,
            vsego_A_coal=f'{round(self.vsego_A_coal, 2)}',
            vsego_B_coal=f'{round(self.vsego_B_coal, 2)}',
            coal_AB_3=f"{self.tranzit_AB_coal*1000}",
            coal_AB_2=f"{round(self.tranzit_AB_coal*1000 + self.vsego_B_coal,2)}",
            coal_AB_1=f"{round(self.tranzit_AB_coal*1000 + self.vsego_B_coal + self.vsego_A_coal,2)}",

            tranzit_AB_oil=self.tranzit_AB_oil*1000,
            vsego_A_oil=f'{round(self.vsego_A_oil, 2)}',
            vsego_B_oil=f'{round(self.vsego_B_oil, 2)}',
            oil_AB_3=f"{round(self.tranzit_AB_oil*1000,2)}",
            oil_AB_2=f"{round(self.tranzit_AB_oil*1000 + self.vsego_B_oil,2)}",
            oil_AB_1=f"{round(self.tranzit_AB_oil*1000 + self.vsego_B_oil + self.vsego_A_oil,2)}",

            vsego_B_forest=f'{round(self.vsego_B_forest, 2)}',
            vsego_A_forest=f'{round(self.vsego_A_forest, 2)}',
            forest_AB_2=f"{round(self.vsego_B_forest,2)}",
            forest_AB_1=f"{round(self.vsego_B_forest + self.vsego_A_forest,2)}",

            vsego_B_mineral=f'{round(self.vsego_B_mineral, 2)}',
            vsego_A_mineral=f'{round(self.vsego_A_mineral, 2)}',
            mineral_AB_2=f"{round(self.vsego_B_mineral,2)}",
            mineral_AB_1=f"{round(self.vsego_B_mineral + self.vsego_A_mineral,2)}",

            sh_V_vvoz=f"{round(self.sh_V_vvoz, 2)}",
             sh_G_vvoz=f"{round(self.sh_G_vvoz, 2)}",
            sh_AB_2=f"{round(self.sh_V_vvoz,2)}",
            sh_AB_1=f"{round(self.sh_V_vvoz + self.sh_G_vvoz,2)}",

            tranzit_AB_tovar_narod=self.tranzit_AB_tovar_narod*1000,
            vsego_B_tovar=f'{round(self.vsego_B_tovar, 2)}',
            vsego_A_tovar=f'{round(self.vsego_A_tovar, 2)}',
            tovar_3=f"{round(self.tranzit_AB_tovar_narod*1000,2)}",
            tovar_2=f"{round(self.vsego_B_tovar + self.tranzit_AB_tovar_narod*1000,2)}",
            tovar_1=f"{round(self.vsego_B_tovar + self.tranzit_AB_tovar_narod*1000 + self.vsego_A_tovar,2)}",

            tranzit_AB_rock=self.tranzit_AB_rock*1000,
            tranzit_AB_build=self.tranzit_AB_build*1000,
            tranzit_AB_prochie=self.tranzit_AB_prochie*1000,

            Σ_AB_1=f'{round(self.Σ_AB_1,2)}',
            Σ_AB_2=f'{round(self.Σ_AB_2,2)}',
            Σ_AB_3=f'{round(self.Σ_AB_3,2)}',

        )
        autocader_1.place_text()

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
                                         val_A_all=f'{round(self.val_A_posevov+self.val_A_furaj+self.val_A_posevov,2)}',

                                         val_B_posevov=f'{round(self.val_B_posevov,2)}',
                                         val_B_furaj=f'{round(self.val_B_furaj,2)}',
                                         val_B_potato=f'{round(self.val_B_potato,2)}',
                                         val_B_all=f'{round(self.val_B_potato+self.val_B_furaj+self.val_B_posevov,2)}',

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

                                         FA=f'{round(third.FctA, 2)}',
                                         FB=f'{round(third.FctB, 2)}',
                                         )

        replacer3=WordEquationReplacer('Praktich_zanyatie_5_EI.docx',

                                       pepleA=f'{round(second.pepleA / 1000, 2)}',

                                       pepleB=f'{round(second.pepleB / 1000, 2)}',
                                       norma_rashoda_tovar_potrebl=f'{self.norma_rashoda_tovar_potrebl}',
                                       norma_rashoda_tovar_coal=f'{self.norma_rashoda_tovar_coal}',
                                       norma_rashoda_forest=f'{self.norma_rashoda_forest}',
                                       norma_rashoda_village=f'{self.norma_rashoda_village}',
                                       norma_rashoda_oil=f'{self.norma_rashoda_oil}',

                                       tranzit_AB_coal=f'{self.tranzit_AB_coal}',
                                       tranzit_AB_rock=f'{self.tranzit_AB_rock}',
                                       tranzit_AB_oil=f'{self.tranzit_AB_oil}',
                                       tranzit_AB_build=f'{self.tranzit_AB_build}',
                                       tranzit_AB_tovar_narod=f'{self.tranzit_AB_tovar_narod}',
                                       tranzit_AB_prochie=f'{self.tranzit_AB_prochie}',

                                       tranzit_BA_coal=f'{self.tranzit_BA_Metall}',
                                       tranzit_BA_rock=f'{self.tranzit_BA_car}',
                                       tranzit_BA_oil=f'{self.tranzit_BA_wood}',
                                       tranzit_BA_build=f'{self.tranzit_BA_sh_hozyaystvo}',
                                       tranzit_BA_tovar_narod=f'{self.tranzit_BA_tovar_narod_mineral}',
                                       tranzit_BA_prochie=f'{self.tranzit_BA_prochie}',

                                       vsego_A_forest=f'{round(self.vsego_A_forest,2)}',
                                       vsego_A_coal=f'{round(self.vsego_A_coal,2)}',
                                       vsego_A_tovar=f'{round(self.vsego_A_tovar,2)}',
                                       vsego_A_mineral=f'{round(self.vsego_A_mineral,2)}',
                                       vsego_A_oil=f'{round(self.vsego_A_oil,2)}',
                                       vsego_B_forest=f'{round(self.vsego_B_forest,2)}',
                                       vsego_B_coal=f'{round(self.vsego_B_coal,2)}',
                                       vsego_B_tovar=f'{round(self.vsego_B_tovar,2)}',
                                       vsego_B_mineral=f'{round(self.vsego_B_mineral,2)}',
                                       vsego_B_oil=f'{round(self.vsego_B_oil,2)}',

                                       A_all=f'{round(self.A_all_ploshad, 2)}',
                                       B_all=f'{round(self.B_all_ploshad, 2)}',
                                       forest_vivoz= f"{round(second.all, 2)}",
                                       sh_V_vivoz=f"{round(self.sh_V_vivoz,2)}",
                                       sh_G_vivoz=f"{round(self.sh_G_vivoz,2)}",
                                       sh_V_vvoz=f"{round(self.sh_V_vvoz,2)}",
                                       sh_G_vvoz=f"{round(self.sh_G_vvoz,2)}",

                                       itogo_vivozzz_stolbik_BA=f"{round(self.itogo_vivozzz_stolbik_BA,2)}",
                                       itogo_vivoz_sh=f"{self.itogo_vivoz_sh}",
                                       itogo_vivoz_vsego=f"{self.itogo_vivoz_vsego}",

                                       summ_vvoza = f'{round(self.summ_vvoza,2)}',
                                       summ_vvoza_A=f'{round(self.summ_vvoza_A,2)}',
                                       summ_summ = f'{round(self.summ_summ,2)}',
                                       t_ovar = f"{round(self.t_ovar,2)}",
                                       s_h = f"{round(self.s_h,2)}",
                                       m_ineral = f"{round(self.m_ineral,2)}",
                                       f_orest = f"{round(self.f_orest,2)}",
                                       o_il = f"{round(self.o_il,2)}",
                                       c_oal=f"{round(self.c_oal,2)}",




                                       #self.itogo_vivoz_sh = self.sh_G_vivoz + self.sh_V_vivoz
                                       #self.itogo_vivozzz_stolbik_BA = self.sh_V_vivoz + round(second.all, 2)
                                       #self.itogo_vivoz_vsego = self.itogo_vivozzz_stolbik_BA + self.sh_G_vivoz
                                       )

        replacer4 = WordEquationReplacer('Praktich_zanyatie_6_EI.docx',

                                         tranzit_AB_coal=f'{self.tranzit_AB_coal}',
                                         tranzit_AB_rock=f'{self.tranzit_AB_rock}',
                                         tranzit_AB_oil=f'{self.tranzit_AB_oil}',
                                         tranzit_AB_build=f'{self.tranzit_AB_build}',
                                         tranzit_AB_tovar_narod=f'{self.tranzit_AB_tovar_narod}',
                                         tranzit_AB_prochie=f'{self.tranzit_AB_prochie}',
                                         results_tabl_tranzit=f'{self.results_tabl_tranzit_vivoz}',

                                         tranzit_BA_Metall=f'{self.tranzit_BA_Metall}',
                                         tranzit_BA_car=f'{self.tranzit_BA_car}',
                                         tranzit_BA_wood=f'{self.tranzit_BA_wood}',
                                         tranzit_BA_sh_hozyaystvo=f'{self.tranzit_BA_sh_hozyaystvo}',
                                         tranzit_BA_tovar_narod_mineral=f'{self.tranzit_BA_tovar_narod_mineral}',
                                         tranzit_BA_prochie=f'{self.tranzit_BA_prochie}',

                                         results_tabl_tranzit_vvoz=f'{self.results_tabl_tranzit_vvoz}',

                                         Σ_BA_1=f'{round(self.Σ_BA_1,2)}',
                                         Σ_BA_2=f'{round(self.Σ_BA_2,2)}',
                                         Σ_BA_3=f'{round(self.Σ_BA_3,2)}',
                                         Σ_AB_1=f'{round(self.Σ_AB_1,2)}',
                                         Σ_AB_2=f'{round(self.Σ_AB_2,2)}',
                                         Σ_AB_3=f'{round(self.Σ_AB_3,2)}',

                                         Σ_ГрБА=f'{round(self.Σ_ГрБА,2)}',

                                         Σ_ГрАБ=f'{round(self.Σ_ГрАБ,2)}',

                                         Г1=f'{round(self.Σ_ГрБА / (self.l * 3), 2)}',
                                         Г2=f'{round(self.Σ_ГрАБ / (self.l * 3), 2)}',

                                         П1=f'{round(second.pepleA,2)} * {self.pepls_movie/100} = {self.П1}',
                                         П2=f'{round(second.pepleB,2)} * {self.pepls_movie/100} = {self.П2}',

                                         ПгрA=f'{2.5}*{round(self.pepls_movie * (second.pepleA+second.pepleB)/100000,2)}*{2*self.l}/1000={self.Пгр}',
                                         #ПгрB=f'{2.5} * {round(second.pepleB * self.pepls_movie/100000,2)} *  31.15 / 1000= {round(2.5 * second.pepleB * self.pepls_movie/100000 *  31.15 * 2 / 1000,2)}',

                                         ГпрIIГ=f'{self.ГпрII}',
                                         ГпрГВ=f'{self.ГпрГВ}',
                                         )

        replacer5 = WordEquationReplacer('Praktich_zanyatie_7_EI.docx',
                                         results_tabl_tranzit=f'{round(self.results_tabl_tranzit_vivoz,2)}',
                                         results_tabl_tranzit_vvoz=f'{self.results_tabl_tranzit_vvoz}',
                                         tranzit_vvoz_and_vivoz=f'{self.results_tabl_tranzit_vvoz+self.results_tabl_tranzit_vivoz}',
                                         tranzit_vvoz_and_vivoz_percent=f'{((self.results_tabl_tranzit_vvoz + self.results_tabl_tranzit_vivoz)/self.mestnie_and_tranzit)*100}',


                                         itogo_vivoz_vsego=f"{self.itogo_vivoz_vsego}",
                                         summ_summ=f'{round(self.summ_summ, 2)}',
                                         itogo_vivoz_vsego_percent=f"{(self.itogo_vivoz_vsego/self.mestnie_and_tranzit)*100}",
                                         summ_summ_percent=f'{(round(self.summ_summ, 2)/self.mestnie_and_tranzit)*100}',
                                         summ_itogo_mestnih = f"{round(self.summ_summ+self.itogo_vivoz_vsego, 2)}",

                                         mestnie_and_tranzit = f"{self.mestnie_and_tranzit}",


                                         oba_gruzooborota = f'{round(self.mestnie_and_tranzit*3*self.l,2)}',

                                         tranzitGruzooborot_vvoz_and_vivoz=f'{round((self.results_tabl_tranzit_vvoz + self.results_tabl_tranzit_vivoz)/1000*self.l*3,2)}',
                                         tranzitGruzooborot_vvoz_and_vivoz_percent=f'{((self.results_tabl_tranzit_vvoz + self.results_tabl_tranzit_vivoz * self.l * 3)/(self.ГпрII+self.ГпрГВ))*100}',

                                         forest =f'{self.f_orest+round(second.all, 2)+self.tranzit_BA_wood*1000}',
                                         coal=f'{round(self.tranzit_AB_coal*1000 + self.c_oal,2)}',
                                         oil=f'{self.tranzit_AB_oil*1000 + self.o_il}',
                                         build=f'{self.tranzit_AB_build*1000}',
                                         Metall=f'{self.tranzit_BA_Metall*1000}',
                                         narod_potrebl=f'{self.tranzit_AB_tovar_narod*1000}',
                                         narod_potrebl_=f'{round((self.tranzit_BA_car * 1000) + (self.f_orest+second.all+self.tranzit_BA_wood*1000) + (self.tranzit_BA_Metall*1000)
                                         + (self.tranzit_AB_tovar_narod*1000) + (self.tranzit_BA_tovar_narod_mineral *1000+ self.m_ineral) + (self.tranzit_BA_sh_hozyaystvo*1000 + self.s_h + self.itogo_vivoz_sh), 2)}',
                                         mineral=f'{self.tranzit_BA_tovar_narod_mineral *1000+ self.m_ineral}',
                                         s_h=f'{self.tranzit_BA_sh_hozyaystvo*1000 + self.s_h + self.itogo_vivoz_sh}',
                                         car=f'{self.tranzit_BA_car*1000}',
                                         itogi=f'{self.itogi}',


                                         itogi16_netto=f"{round(self.tranzit_AB_tovar_narod*1000+self.tranzit_AB_build*1000+self.tranzit_AB_oil*1000 + self.o_il+self.tranzit_AB_coal*1000 + self.c_oal,2)}",
                                         netto16_4_coal=f"{round(β_4(self, 0.8, self.qbr_coal_4, self.qbr_coal_8)*(self.tranzit_AB_coal*1000 + self.c_oal),2)}",
                                         netto16_8_coal=f"{round(β_8(self, 0.8, self.qbr_coal_4, self.qbr_coal_8) * (self.tranzit_AB_coal * 1000 + self.c_oal),2)}",

                                         netto16_4_oil=f"{round(β_4(self, 0.8, self.qbr_oil_4, self.qbr_coal_8) * (self.tranzit_AB_oil * 1000 + self.o_il),2)}",
                                         netto16_8_oil=f"{round(β_8(self, 0.8, self.qbr_oil_8, self.qbr_coal_8) * (self.tranzit_AB_oil * 1000 + self.o_il),2)}",

                                         netto16_4_Build=f"{round(β_4(self, 1, self.qbr_Build_4, 0) * (self.tranzit_AB_build * 1000/2),2)}",
                                         netto16_4_Build_=f"{round(β_4(self, 1, self.qbr_Build_4_, 0) * (self.tranzit_AB_build * 1000/2),2)}",

                                         promtovari_prochie_4=f"{round(β_4(self, 1, self.qbr_narod_potrebl_and_prochie_4, 0) * (self.tranzit_BA_car * 1000) + (self.f_orest+second.all+self.tranzit_BA_wood*1000) + (self.tranzit_BA_Metall*1000)
                                         + (self.tranzit_AB_tovar_narod*1000) + (self.tranzit_BA_tovar_narod_mineral *1000+ self.m_ineral) + (self.tranzit_BA_sh_hozyaystvo*1000 + self.s_h + self.itogo_vivoz_sh), 2)}",

                                         vag4narod_potrebl=f'{self.vag4narod_potrebl}',
                                         vag4coal=f'{self.vag4coal}',
                                         vag4oil=f'{self.vag4oil}',
                                         vag4build=f'{self.vag4build}',

                                         vag8coal=f'{self.vag8coal}',
                                         vag8oil=f'{self.vag8oil}',
                                         vag4_build=f'{self.vag4_build}',

                                         qbr_coal_4 = f'{round(self.qbr_coal_4,2)}',
                                         qbr_coal_8 = f'{round(self.qbr_coal_8,2)}',

                                         qbr_oil_4 = f'{round(self.qbr_oil_4,2)}',
                                         qbr_oil_8 = f'{round(self.qbr_oil_8,2)}',

                                         qbr_Build_4 = f'{round(self.qbr_Build_4,2)}',
                                         qbr_Build_4_ = f'{round(self.qbr_Build_4_,2)}',

                                         qbr_Metall_4 = f'{round(self.qbr_coal_4,2)}',

                                         qbr_narod_potrebl_and_prochie_4 = f'{round(self.qbr_narod_potrebl_and_prochie_4,2)}',

                                         qbr_Mineral_4 = f'{round(self.qbr_Mineral_4,2)}',
                                         qbr_Mineral_8 = f'{round(self.qbr_Mineral_8,2)}',

                                         qbr_sh_4 = f'{round(self.qbr_sh_4,2)}',

                                         bruttocoal4= f'{self.bruttocoal4}',

                                         bruttocoal8=f'{self.bruttocoal8}',
                                         bruttooil4=f'{self.bruttooil4}',
                                         bruttooil8=f'{self.bruttooil8}',
                                         bruttoBuild4=f'{self.bruttoBuild4}',
                                         bruttoBuild4_=f'{self.bruttoBuild4_}',
                                         bruttoNarod4=f'{self.bruttoNarod4}',

                                         itogvagona4=f'{self.itogvagona4}',
                                         itogvagona8=f'{self.itogvagona8}',
                                         itogVesvagona4=f'{self.itogVesvagona4}',
                                         itogVesvagona8=f'{self.itogVesvagona8}',
                                         itogiVesa=f'{round(self.itogVesvagona8 + self.itogVesvagona4, 2)}',

                                         Vesvagona4=f'{self.Vesvagona4}',
                                         Vesvagona8=f'{self.Vesvagona8}',
                                         β4=f'{self.β4}',
                                         β8=f'{self.β8}',
                                         fi8=f'{self.fi8}',
                                         fi4=f'{self.fi4}',
                                         teta=f'{self.teta}',
                                         )




        # Обработка документа
        replacer.process_document()
        replacer.save_document('Черновая_заготовка_1.docx')
        replacer1.process_document()
        replacer1.save_document('Черновая_заготовка_2.docx')
        replacer2.process_document()
        replacer2.save_document('Черновая_заготовка_3.docx')
        replacer3.process_document()
        replacer3.save_document('Черновая_заготовка_4.docx')
        replacer4.process_document()
        replacer4.save_document('Черновая_заготовка_5.docx')
        replacer5.process_document()
        replacer5.save_document('Черновая_заготовка_6.docx')

        print("Σ_БА=", self.Σ_ГрБА)
        print("Σ_АБ=", self.Σ_ГрАБ)
        print("Σ_BA_1=", self.Σ_BA_1)
        print("Σ_BA_2=", self.Σ_BA_2)
        print("Σ_BA_3=", self.Σ_AB_3)
        print("Σ_AB_2=", self.Σ_AB_2)
        print("Σ_AB_1=", self.Σ_AB_1)
        print("Σ_AB_3=", self.Σ_AB_3)
        print("Σ_БА=", self.Σ_ГрБА)
        print("Σ_БА=", self.Σ_ГрБА)
        print("Σ_БА=", self.Σ_ГрБА)
        print("Σ_БА=", self.Σ_ГрБА)
        print("Σ_БА=", self.Σ_ГрБА)
        print("потребление леса=", self.f_orest)
        print("Транзит леса весь=", self.tranzit_BA_wood)
        print("Вывоз леса с завода=", round(second.all, 2))


# Запуск приложения
if __name__ == "__main__":
    root = tk.Tk()

    app = App(root)
    root.mainloop()