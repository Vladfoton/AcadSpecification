import win32com.client
from read_autocad_data_by_polygon import select_object_in_rect
import pythoncom


class Parameters:
    def __init__(self, acad_block_parameters: dict):
        for key, value in acad_block_parameters.items():
            self.__setattr__(key.lower(), value)
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
            if other.level in self.list_poz:
                self.list_poz[other.level].append((other.multiple_data, other.list_poz[other.level][0][1]))
            else:
                self.list_poz[other.level] = [(other.multiple_data, other.list_poz[other.level][0][1])]
            # print("функция __add__", self.list_poz)
            return self
        else:  # Если объекты не совместимы или отличаются параметры (имя конструкции, коэффициэнт умножения для ВРС,
            # шифр, стандарт арматуры,...то игнорирует суммирование и возвращает исходный объект
            return self

    def __eq__(self, other):
        atr_list = ('constr_name', 'multiplevrs', 'projectcode', 'arm_standart', 'multiplevrs', 'specification_head',
                    'specification_type', 'vd_file_name',
                    'vrs_type')  # Список атрибутов при равенстве которых будет считаться что объекты равны
        return all(self.__getattribute__(i) == other.__getattribute__(i) for i in atr_list)

    def add_list_poz(self, list_poz: dict):
        # print("функция add_list_poz")
        self.list_poz = list_poz

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
        pass


if __name__ == '__main__':
    pass
