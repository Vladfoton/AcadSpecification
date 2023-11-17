import win32com.client


class Parameters:
    def __init__(self, acad_block_parameters: dict):
        for key, value in acad_block_parameters.items():
            self.__setattr__(key.lower(), value)
        temp = {}
        try:
            for poz in self.arm_standart.split(';'):
                sufix, data = poz.split(':')
                standart, klass = data.split(',')
                standart = standart[1:]
                klass = klass[:-1]
                temp[sufix] = (standart, klass)

            self.arm_standart = temp
        except :
            raise ValueError('Ошибка в блоке параметров в атрибуте "arm_standart"')

    def __add__(self,other):
        if isinstance(other, self.__class__):

            self.list_poz.extend(other.list_poz)
            print("функция __add__", self.list_poz)

    def add_list_poz(self,list_poz):
        print("функция add_list_poz")
        self.list_poz = list_poz


#TODO Добавить сравнение объектов и сделать суммирование только если данные с одного конструктивного элемента




    def __setattr__(self, key, value):
        try:
            if '.' in value:
                self.__dict__[key] = float(value)
            else:
                if value.isdigit():
                    self.__dict__[key] = bool(value)
                else:
                    self.__dict__[key] = value
        except:
            self.__dict__[key] = value





    # def find_and_read_parameters(self):
    #     print(self.acad_block_parameters.GetAttributes())
    #     return {atr_data.TagString: atr_data.TextString for atr_data in self.acad_block_parameters.GetAttributes()}



    @staticmethod
    def set_arm_standart(arm_standart:str)->dict:
        return {arm_data[0]: (arm_data[1], arm_data[2]) for arm_data in arm_standart.split(',')}

        pass

    def __str__(self):
        pass

    def add(self, textstring: str):
        pass


if __name__ == '__main__':
    pass