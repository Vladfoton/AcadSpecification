
class Item():
    def __init__(self, markposition:str, syze:float, lenght:float, material:str, standart:str,count:int, function:bool, massa1p:float, unit:str, level_ID:str = '0', const_ID:str = '0', object_list = []):
        self.markposition = markposition  # Маркировка в чертежах и спецификациях
        self.syze = syze  # Характерный типоразмер (Диаметр, толщина, и.т.п.)[мм]

        self.objectlist = object_list  # Список элементов из которых состоит элемент. Если список пустой то элемент состоит из самого себя. Если нет
        self.material=material  # Материал (Класс арматуры, класс бетона,и т.п.)
        self.standart = standart  # Обозначение стандарта (ГОСТ, серия,ТУ,и т.п)
        self.function = function  # Способ подсчета позиции (Арматура в п.м. (0),арматура штучная (1)

        if object_list:
            self.massa1p=sum([obj.total_massa for obj in self.objectlist])
        else:
            self.massa1p = massa1p
        if self.function:
            self.count=count  # Количество ед [шт]
            self.lenght = lenght  # Длина [мм]
            self.massa = self.massa1p * self.lenght/1000 # Масса 1 ед измерения (масса 1м.п. или масса 1 шт)[кг]
        else:
            self.lenght = lenght * count
            self.count =1# Количество ед [шт]
            self.massa =  self.massa1p * self.lenght / 1000   # Масса 1 ед измерения (масса 1м.п. или масса 1 шт)[кг]


        self.total_massa=self.count*self.massa  # общий вес
        self.unit=unit  # ед.измерения(кг, м.п, м3,и.т.п)

        self.const_ID = const_ID  # Принадлежность к конструкции (марка  или номер конструкции)
        #                          0 - входит в основную конструкцию.
        #                          <markposition> - входит в состав <markposition>
        self.level_ID=level_ID  # Принадлежность отметке Level_ID (для групповых спецификаций)
        self.atr_eq_list = ('syze', 'material', 'standart','function', 'unit', 'const_ID',
                    'level_ID')  # Список атрибутов при равенстве которых будет считаться что экземпляры равны и с ними можно проводить операции сложения




    def __str__(self):
        rezult = '\n___Позиция:___\n'
        for atr_name in self.__dict__:
            rezult += f'{atr_name} : {self.__dict__[atr_name]} \n'
        rezult += '___Конец вывода позиции___\n'
        return rezult

    def __eq__(self, other):
        return isinstance(other, self.__class__) and all(self.__getattribute__(i) == other.__getattribute__(i) for i in self.atr_eq_list)

    def __add__(self, other):
        if self == other:
            if self.function and self.lenght == other.lenght:
                self.count+=other.count
                self.total_massa += other.total_massa
            elif not self.function:
                self.lenght += other.lenght*other.count
                self.massa = self.lenght * self.massa1p/1000
                self.total_massa = self.count * self.massa
            else:  # Если объекты не совместимы или отличаются параметры ,...то игнорирует суммирование и возвращает исходный объект
                print(f'Объекты {self} и {other} не ,были просуммированы так как тип подсчета стоит штучные, но при этом различаются размеры. Просуммировать невозможно. Возможно нужно поменять тип подсчета на м.п.')

            return self
        else:  # Если объекты не совместимы или отличаются параметры ,...то игнорирует суммирование и возвращает исходный объект
            print(f'Объекты {self} и {other} не ,были просуммированы так как они не совместимы между собой')
            return self

if __name__ == "__main__":
    dd= Item('aa',10, 3000, 'A500','ГОСТ 888', 10, False, 0.617,'шт')
    dd1 = Item('aa', 10, 2000, 'A500', 'ГОСТ 888', 20, False, 0.617, 'шт')
    dd+=dd1
    dd3 = Item('КР1',0, 1500, 'Каркас КР1','Серия 4444', 5, False, 0,'шт', object_list =[dd])
    dd4 = Item('КП1',0, 2000, 'Каркас КП1','Серия 4444', 2, True, 0,'шт', object_list =[dd1,dd3])
    dd5 = Item('КП1', 0, 2000, 'Каркас КП1', 'Серия 4444', 3, True, 0, 'шт', object_list=[dd1, dd3])
    dd4+=dd5
    print(dd4)
