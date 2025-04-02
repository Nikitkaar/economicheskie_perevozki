import logging
import math
import re
from docx import Document


class WordEquationReplacer:
    def __init__(self, file_path, **kwargs):
        self.doc = Document(file_path)
        self.replacements = kwargs
        logging.debug(f"Инициализация WordEquationReplacer: {self.replacements}")

    import re

    def replace_text_in_paragraph(self, p):
        #Заменяет текст в переданном параграфе и логирует процесс.#
        for old, new in self.replacements.items():
            # Экранируем текст, чтобы избежать ошибок с регулярными выражениями
            sanitized_old = re.escape(old)

            # Используем re.sub для замены, просто указываем возможные разделители
            # Здесь мы обрабатываем возможные пробелы, запятые и точки
            pattern = re.compile(rf'\b{sanitized_old}\b')  # Слово должно быть целым, используем границы слова

            # Обработаем текст в параграфе
            if pattern.search(p.text):
                for run in p.runs:
                    if pattern.search(run.text):
                        old_text = run.text
                        run.text = pattern.sub(new, run.text)
                        logging.debug(f"Заменено в параграфе: '{old}' на '{new}' в тексте '{old_text}'")
            else:
                logging.debug(f"'{old}' не найдено в тексте: '{p.text}'")

    def replace_text_in_cell(self, cell):
        #Заменяет текст в переданной ячейке.#
        for paragraph in cell.paragraphs:
            self.replace_text_in_paragraph(paragraph)

    def process_document(self):
        #Обрабатывает документ, заменяя текст и обрабатывая таблицы.#
        logging.info("Обработка документа начата.")

        # Заменяем текст в параграфах
        for p in self.doc.paragraphs:
            logging.debug(f"Обрабатываем параграф: '{p.text}'")
            self.replace_text_in_paragraph(p)

        # Заменяем текст в таблицах
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    logging.debug(f"Обрабатываем ячейку: '{cell.text}'")
                    self.replace_text_in_cell(cell)

        logging.info("Обработка документа завершена.")

    def save_document(self, new_file_path):
        #Сохраняет документ в новый файл.
        self.doc.save(new_file_path)
        logging.info(f"Документ сохранен как: {new_file_path}")


class Counter:
    # Статический словарь с данными о локомотивах
    locomotives_data = {
        "ВЛ60": {
            "Конструктивная скорость": 120,
            "Радиус колеса": 625,
            "Число осей": 3,
            "Длина жесткой базы": 4600,
            "Поперечные разбеги": {"крайних": 1.0, "средней у трехосной тележки": 15.5}
        },
        "ВЛ60к": {
            "Конструктивная скорость": 100,
            "Радиус колеса": 625,
            "Число осей": 3,
            "Длина жесткой базы": 4600,
            "Поперечные разбеги": {"крайних": 1.0, "средней у трехосной тележки": 15.5},
        },
        "ЧС4т": {
                "Конструктивная скорость": 160,
                "Радиус колеса": 625,
                "Число осей": 3,
                "Длина жесткой базы": 4600,
                "Поперечные разбеги": {"крайних": 1.3, "средней у трехосной тележки": 1.3}
                # Добавьте остальные локомотивы по аналогии
        }}

    def __init__(self, **kwargs):
        self.Loc = kwargs.get("Loc")
        self.VGruz = kwargs.get("VGruz")
        self.VEmpty = kwargs.get("VEmpty")
        self.rail = kwargs.get("rail")
        self.iznos = kwargs.get("iznos")
        self.shpali = kwargs.get("shpali")
        self.screpleniya = kwargs.get("screpleniya")
        self.rezina = kwargs.get("rezina")
        self.ballast = kwargs.get("ballast")
        self.circle_moment = kwargs.get("circle_moment")
        self.sdvijka = kwargs.get("sdvijka")
        self.R = kwargs.get("R") * 1000
        self.ugolPovorotaNAcurve = kwargs.get("ugolPovorotaNAcurve")
        self.VpryamoNAPerevode = kwargs.get("VpryamoNAPerevode")
        self.VbokovoyNAPerevode = kwargs.get("VbokovoyNAPerevode")
        self.jo = kwargs.get("jo")
        self.γo = kwargs.get("γo")
        self.µ = kwargs.get("µ")
        self.𝜂1 = Counter.locomotives_data[f"{self.Loc}"]["Поперечные разбеги"]["крайних"]
        self.𝜂2 = Counter.locomotives_data[f"{self.Loc}"]["Поперечные разбеги"]["крайних"]
        self.λ = Counter.locomotives_data[f"{self.Loc}"]["Длина жесткой базы"]
        self.L = Counter.locomotives_data[f"{self.Loc}"]["Длина жесткой базы"]
        self.r = Counter.locomotives_data[f"{self.Loc}"]["Радиус колеса"]

        self.qmax = 1509
        self.τ = 70  # угол наклона рабочей поверхности гребня колеса к горизонту, равный для локомотивных бандажей – 70°; self.r
        self.t = 13
        self.S = 1520

        self.b = round((self.L * (self.r + self.t) * math.tan(math.radians(self.τ))) / (
                self.R + (self.S / 2) - (self.r + self.t) * math.tan(math.radians(self.τ))),2)
        self.fh = round(((self.L + self.b) ** 2) / (2 * self.R),2)
        self.Sopt = round(self.qmax + self.fh - self.𝜂1 + 4,2)
        self.bb = round((self.L * self.r * math.tan(math.radians(self.τ))) / (2 * self.R),2)
        self.fhh = round(((self.L + 2 * self.bb) ** 2) / (8 * self.R),2)
        self.fv = round(((self.L - 2 * self.bb) ** 2) / (8 * self.R),2)
        self.δ_min = 7
        self.Smin = round(self.qmax + self.fh - self.fv + (self.δ_min / 2) + 4,2)

