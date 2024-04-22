import win32com.client
from ReadAutocadData.read_autocad_data_by_polygon import select_object_in_rect
import pythoncom
import os
import openpyxl
from Service.Error_reports import *
import pathlib
from pathlib import Path

# framework_Path = Path(curent_dir, "Каркасы")  # Путь в папку с каркасами
#     embedded_parts_Path = Path(curent_dir, "Закладные детали")  # Путь в папку с закладными
#     Specification_Path = Path(curent_dir, "Спецификации и ВРС")  # Путь в папку с результирующими спецификациями
#     Ved_det_Path = Path(curent_dir, "Ведомость деталей")  # Путь в папку с вед_деталей
class Parameters:
    def __init__(self, acad_block_parameters: dict, framework_Path:str, embedded_parts_Path:str, specification_Path:str, ved_det_Path:str):
        self.framework_Path=framework_Path
        self.embedded_parts_Path=embedded_parts_Path
        self.specification_Path = specification_Path
        self.ved_det_Path = ved_det_Path
        for key, value in acad_block_parameters.items():
            self.__setattr__(key.lower(), value)
        self.ved_det_dict = self.read_VD_file()

        #Формирование словаря с обозначениями и стандартами и классами
        temp = {}
        try:
            for poz in self.arm_standart.split(';'):
                sufix, data = poz.split(':')
                standart, klass = data.split(',')
                standart = standart[1:]
                klass = klass[:-1]
                temp[sufix] = (standart, klass)

            self.arm_standart = temp
        except:
            raise ValueError('Ошибка в блоке параметров в атрибуте "arm_standart"')
        # формирование списка точек контура
        try:
            self.contour = [float(i) for i in self.p1_coord.split(", ")] + [float(i) for i in self.p2_coord.split(", ")] + [float(i) for i in self.p3_coord.split(", ")] + [float(i) for i in self.p4_coord.split(", ")] + [float(i) for i in self.p1_coord.split(", ")]
        except:
            raise ValueError('Ошибка в координатах контура')

    def __add__(self, other):
        if isinstance(other, self.__class__) and self == other:
            if other.level in self.list_poz_str:
                self.list_poz_str[other.level].append((other.multiple_data, other.list_poz_str[other.level][0][1]))
            else:
                self.list_poz_str[other.level] = [(other.multiple_data, other.list_poz_str[other.level][0][1])]
            # print("функция __add__", self.list_poz)
            return self
        else:  # Если объекты не совместимы или отличаются параметры (имя конструкции, коэффициэнт умножения для ВРС,
            # шифр, стандарт арматуры,...то игнорирует суммирование и возвращает исходный объект
            print(f'Объекты {self} и {other} не ,были просуммированы так как они не совместимы между собой')
            return self

    def __eq__(self, other):
        atr_list = ('constr_name', 'multiplevrs', 'projectcode', 'arm_standart', 'multiplevrs', 'specification_head',
                    'specification_type', 'vd_file_name',
                    'vrs_type')  # Список атрибутов при равенстве которых будет считаться что объекты равны
        return all(self.__getattribute__(i) == other.__getattribute__(i) for i in atr_list)

    def add_list_poz(self, list_poz: dict):
        # print("функция add_list_poz")
        self.list_poz_str = list_poz
        print(f'{list_poz=}')

    def create_object_list(self):
        pass

    def read_VD_file(self):
        '''чтение файла ведомости деталей
        :argument filename - имя файда из которого читаются данные
        :return data -словарь вида <марка позиции>: <длина позиции>'''
        filename = self.vd_file_name

        try:
            os.chdir(self.ved_det_Path)
        except:
            error_report_No_exit(f"Папка {self.ved_det_Path} не найдена.")

        try:
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

    def __setattr__(self, key, value):
        try:
            if '.' in value:
                self.__dict__[key] = float(value)
            elif value in ('0', '1', True, False):
                self.__dict__[key] = bool(value)
            elif value.isdigit():
                self.__dict__[key] = int(value)
            else:
                self.__dict__[key] = value
        except:
            self.__dict__[key] = value

    def __str__(self):
        rezult='\n___Начало блока Параметры:___\n'
        for atr_name in self.__dict__:
            rezult += f'{atr_name} : {self.__dict__[atr_name]} \n'
        rezult+='___Конец блока Параметры___\n'
        return rezult



if __name__ == '__main__':
    pass
