# klass0="A500C"
# klass240="A240"
VD_filename = "VD.xlsx"
import os
import pathlib  # библиотека для работы с путями к файлам в файловой системе
from pathlib import Path
from Service.Error_reports import *

import openpyxl  # библиотека для работы с файлами ms excell

# Пути к папкам
curent_dir = pathlib.Path.cwd()  # Текущая рабочая папка
Ved_det_Path = Path(curent_dir, "Ведомость деталей")  # Путь в папку с вед_деталей


# чтение файла ведомости деталей
def read_VD_file(Ved_det_Path, filename):
    '''чтение файла ведомости деталей
    :argument filename - имя файда из которого читаются данные
    :return data -словарь вида <марка позиции>: <длина позиции>'''
    try:
        os.chdir(Ved_det_Path)
        book = openpyxl.open(filename, read_only=True)
        sheet = book.active
        pos = [sheet[i][0].value for i in range(1, sheet.max_row + 1)]
        lenght = [sheet[i][1].value for i in range(1, sheet.max_row + 1)]
        data = dict(zip(pos, lenght))
        book.close()
    except FileNotFoundError:
        error_report_No_exit(f"Файл {filename} не найден.")
        return {}
    except:
        error_report_end_exit(f"Ошибка в файле {filename}. Откорректируйте файл и перезапустите программу")

    return data


def read_strings(firststring: str, secondstring: str, veddetaley, koefficient=1.0) -> dict:
    klassB500 = "B500"
    if "\P" in firststring:
        secondstring = firststring[firststring.find("\P") + 2:]
        firststring = firststring[:firststring.find("\P")]

    def find_digit(string: str, first_pozition=0) -> str:
        """ функция выводит непрерывную последовательность из
        цифровых символов начиная с <first_pozition> Например результатом работы
        ункции find_digit("вв201в3333цц",2) будет "201"  """

        rezult = ""
        while first_pozition < len(string) and (string[first_pozition] == "." or string[first_pozition].isdigit()):
            rezult += string[first_pozition]
            first_pozition += 1
        return rezult

    def strip_string(string: str) -> str:
        """ Функция удаляет перый апостраф в строке и удаляет все пробелы"""
        if string:
            if string[0] == "'":
                string = string[1:]
            if "\pxqc;" in string:
                string = string.replace("\pxqc;", "")
            string = string.strip()
        return string.replace(" ", "")

    def find_markpozition(firststring: str, veddetaley) -> str:
        """ Функция находит в строке марку"""
        if "Бетон" in firststring or "бетон" in firststring:
            markpozition = firststring
            return markpozition
        else:
            firststring = strip_string(firststring)
            markpozition = firststring
        if "x" in markpozition:
            markpozition = markpozition[markpozition.find("x") + 1:]
        if "х" in markpozition:
            markpozition = markpozition[markpozition.find("х") + 1:]

        if markpozition not in veddetaley:

            if "-" in markpozition:
                try:
                    if 1 <= int(find_digit(markpozition, markpozition.find("-") + 1)) <= 99:
                        markpozition = markpozition[:markpozition.find("-")] + "-" + find_digit(markpozition,
                                                                                                markpozition.find(
                                                                                                    "-") + 1)
                    else:
                        markpozition = markpozition[:markpozition.find("-")]
                except ValueError:
                    markpozition = markpozition[:markpozition.find("-")]
            elif "(L=" in markpozition:
                markpozition = markpozition[:markpozition.find("(L=")]

        return markpozition

    def find_diameter(firststring: str,
                      korrect_diameters=(4, 5, 6, 8, 10, 12, 14, 16, 18, 20, 22, 25, 28, 32, 36, 40)) -> int:
        """Функция находит в строке диаметр"""
        firststring = strip_string(firststring)
        diameter = ""
        if "-" in firststring:
            diameter = firststring[:firststring.find("-")]
            if "x" in diameter:
                diameter = diameter[diameter.find("x") + 1:]
            if "х" in diameter:
                diameter = diameter[diameter.find("х") + 1:]
            if "г" in diameter:
                diameter = diameter[:diameter.find("г")]
            if "B" in diameter:
                diameter = diameter[:diameter.find("B")]
            if "В" in diameter:
                diameter = diameter[:diameter.find("В")]
            while len(diameter) > 1 and not diameter[0].isdigit():
                diameter = diameter[1:]
        elif "(L=" in firststring:
            diameter = firststring[:firststring.find("(L=")]
            if "x" in diameter:
                diameter = diameter[diameter.find("x") + 1:]
            if "х" in diameter:
                diameter = diameter[diameter.find("х") + 1:]
            if "г" in diameter:
                diameter = diameter[:diameter.find("г")]
            if "B" in diameter:
                diameter = diameter[:diameter.find("B")]
            if "В" in diameter:
                diameter = diameter[:diameter.find("В")]
            while len(diameter) > 1 and not diameter[0].isdigit():
                diameter = diameter[1:]

        else:
            diameter = firststring
            if "x" in diameter:
                diameter = diameter[diameter.find("x") + 1:]
            if "х" in diameter:
                diameter = diameter[diameter.find("х") + 1:]
            if "г" in diameter:
                diameter = diameter[:diameter.find("г")]
            if "B" in diameter:
                diameter = diameter[:diameter.find("B")]
            if "В" in diameter:
                diameter = diameter[:diameter.find("В")]
            while len(diameter) > 1 and not diameter[0].isdigit():
                diameter = diameter[1:]

        try:
            diameter = int(diameter)
        except:
            return -1
        if diameter in korrect_diameters:
            return diameter
        else:
            return -1

    def find_lenght(firststring: str, veddetaley) -> float:
        """Функция находит в строке длину позиции"""
        veddetaley
        markpoz = find_markpozition(firststring, veddetaley)
        firststring = strip_string(firststring)
        lenght = ""

        if markpoz in veddetaley:
            return veddetaley[markpoz]
        if "..." in firststring or "..." in firststring or chr(8230) in firststring or "," in firststring:
            return -1
        if "-" in firststring:
            lenght = find_digit(firststring, firststring.find("-") + 1)
        elif "(L=" in firststring:
            lenght = find_digit(firststring, firststring.find("(L=") + 3)
        try:
            if lenght and int(lenght) > 99:
                return int(lenght)
        except ValueError:
            if '.' in lenght:
                return float(lenght)
        except:
            return -1
        else:
            return -1

    def find_function(firststring: str, veddetaley) -> bool:
        # 0-в пог метрах, 1- штучный
        return find_markpozition(firststring, veddetaley) in veddetaley or find_lenght(firststring, veddetaley) <= 99

    def find_count(firststring: str, secondstring: str) -> float:
        """ Функция вычисляет количество"""
        firststring = strip_string(firststring)
        secondstring = strip_string(secondstring)
        if "шт." in firststring:
            count = find_digit(firststring, firststring.find("шт.") + 3)
        elif "V=" in firststring:
            count = find_digit(firststring, firststring.find("V=") + 2)
        elif "шт." in secondstring:
            count = find_digit(secondstring, secondstring.find("шт.") + 3)
        elif "V=" in secondstring:
            count = find_digit(secondstring, secondstring.find("V=") + 2)
        else:
            count = 1
        try:
            if "x" in firststring or "х" in firststring:
                count = int(find_digit(firststring)) * int(count)
            return float(count)
        except:
            return -1

    def find_klass(firststring: str, veddetaley, klass0="A500C", klass240="A240", klassB500="BpI") -> str:
        """ Функция определяет класс арматуры"""
        if "г" in find_markpozition(firststring, veddetaley):
            return klass240
        if "B" in find_markpozition(firststring, veddetaley) or "В" in find_markpozition(firststring, veddetaley):
            return klassB500
        else:
            return klass0

    # print(f"***** {find_markpozition(firststring)}   {find_count(firststring,secondstring)=}")
    # firststring=strip_string(firststring)
    secondstring = strip_string(secondstring)
    return {"markpozition": find_markpozition(firststring, veddetaley),
            "diameter": find_diameter(firststring),
            "lenght": find_lenght(firststring, veddetaley),
            "count": find_count(firststring, secondstring) * koefficient if find_markpozition(firststring,
                                                                                              veddetaley) not in veddetaley else find_count(
                firststring, secondstring),
            "klass": find_klass(firststring, veddetaley),
            "func": int(find_function(firststring, veddetaley))
            }


# TEST
# veddetaley = read_VD_file(Ved_det_Path, VD_filename)
if __name__ == "__main__":
    # veddetaley = read_VD_file(Ved_det_Path, VD_filename)

    f = "10г-4950\P шт.10"
    f1 = "Р10-1\Pшаг 200, шт.6"
    s = "'апрапрен ен"
    print(read_strings(s, "", read_VD_file(Ved_det_Path, VD_filename)))
