from pyautocad import Autocad, APoint  # Импортируйте необходимые классы из библиотеки Autocad


class AutoCADLinesPlacer:
    def __init__(self, **kwargs):
        self.data = kwargs
        self.coordinates = {
            'sh_G_vivoz': (114, 268),
            'sh_V_vivoz': (151, 268),
            'tranzit_BA_sh_hozyaystvo': (75, 264),
            'sh_BA_2': (112, 264),
            'sh_BA_3': (149, 264),

            'tranzit_BA_tovar_narod_mineral': (75, 250),

            'forest_vivoz': (151, 242),
            'tranzit_BA_wood': (75, 235),
            'forest_BA_3': (149, 235),

            'Σ_BA_1': (75, 273),
            'Σ_BA_2': (112, 273),
            'Σ_BA_3': (149, 273),

            'tranzit_BA_Metall': (75, 221),
            'tranzit_BA_car': (75, 207),
            'tranzit_BA_prochie': (75, 192),

            'tranzit_AB_coal': (149, 170),
            'vsego_A_coal': (102, 170),
            'vsego_B_coal': (139, 170),
            'coal_AB_3': (149, 164),
            'coal_AB_2': (112, 164),
            'coal_AB_1': (75, 164),

            'tranzit_AB_oil': (149, 155),
            'vsego_A_oil': (102, 155),
            'vsego_B_oil': (139, 155),
            'oil_AB_3': (149, 149),
            'oil_AB_2': (112, 149),
            'oil_AB_1': (75, 149),

            'vsego_B_forest': (139, 141),
            'vsego_A_forest': (102, 141),
            'forest_AB_2': (112, 137),
            'forest_AB_1': (75, 137),

            'vsego_B_mineral': (139, 128),
            'vsego_A_mineral': (102, 128),
            'mineral_AB_2': (112, 122),
            'mineral_AB_1': (75, 122),

            'sh_V_vvoz': (139, 114),
            'sh_G_vvoz': (102, 114),
            'sh_AB_2': (112, 108),
            'sh_AB_1': (75, 108),

            'tranzit_AB_tovar_narod': (149, 93),
            'vsego_B_tovar': (139, 100),
            'vsego_A_tovar': (102, 100),
            'tovar_3': (159, 100),
            'tovar_2': (112, 93),
            'tovar_1': (75, 93),

            'tranzit_AB_rock': (149, 79),
            'tranzit_AB_build': (149, 65),
            'tranzit_AB_prochie': (149, 50),

            'Σ_AB_1': (75, 41),
            'Σ_AB_2': (112, 41),
            'Σ_AB_3': (149, 41)

            # Добавьте остальные координаты
        }

    def place_text(self):
        """Размещает текст в AutoCAD по заданным координатам"""
        try:
            # Подключение к AutoCAD
            self.acad = Autocad(create_if_not_exists=True)
            self.acad.Visible = True

            # Размещаем каждый параметр
            for param, value in self.data.items():
                if param in self.coordinates:
                    x, y = self.coordinates[param]
                    point = APoint(x, y)
                    self.acad.model.AddText(str(value), point, 2.5)  # 2.5 - высота текста
                    print(f"Размещен параметр {param} = {value} в точке ({x}, {y})")

        except Exception as e:
            print(f"Ошибка: {e}")

# Использование
#placer = AutoCADLinesPlacer(
    #tranzit_AB_coal=50,
    #tranzit_AB_rock=30
#)
#placer.place_text()