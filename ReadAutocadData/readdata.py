'''
Программа считывает выделенные позиции в пространстве модели Autocad? выбирает среди них
поз "AcDbBlockReference" "Мультивыноска v1.1" (блок "Мультивыноска v1.1") и
поз. "AcDbMLeader" (Стандартная выноска Autocad) и выводит их содержимое в файл exls

'''
import copy
import os
import pathlib
import win32com.client
from ReadAutocadData.Parameters import Parameters
from ReadAutocadData.read_autocad_data_by_polygon import select_object_in_rect

app = win32com.client.Dispatch("AutoCAD.Application")
aDoc = app.ActiveDocument
msp = aDoc.ModelSpace

aDoc.SendCommand("_REGEN ")

def find_and_read_parameters(framework_Path:str, embedded_parts_Path:str, specification_Path:str, ved_det_Path:str) -> Parameters:
    '''

    :param sset:  Набор объектов Autocad
    :param elem_list: список элементов объектов Parameters
    :return: Дополненный список объектов Parameters
    '''
    #__________чтение данных из Autocad________
    app = win32com.client.Dispatch("AutoCAD.Application")   # Подключение к AutoCAD.Application
    aDoc = app.ActiveDocument
    msp = aDoc.ModelSpace

    aDoc.SendCommand("_REGEN ")  #Регенерация пространства модели в Autocad
    input('Выбери  группу элементов c зонами армирования и нажми ENTER')
    sset = aDoc.PickfirstSelectionSet

    elem_list=[]
    parameter_list = []  # Список блоков parameters
    for acad_obj in sset:
        if acad_obj.EntityName == "AcDbBlockReference":
            if acad_obj.EffectiveName == "parameters":
                parameter_list.append(acad_obj)

    for parameter_elem in parameter_list:
        temp = Parameters(
            {atr_data.TagString: atr_data.TextString for atr_data in parameter_elem.GetAttributes()}, framework_Path, embedded_parts_Path, specification_Path, ved_det_Path)
        list_poz = select_object_in_rect(temp.contour)
        list_poz = {temp.level: [(temp.multiple_data, list_poz)]}
        temp.add_list_poz(list_poz)
        const_element = temp
        if const_element in elem_list:
            # print(elem_list[elem_list.index(const_element)].list_poz, const_element.list_poz)
            elem_list[elem_list.index(const_element)] += const_element
        else:
            elem_list.append(const_element)
    return elem_list


    # return temp



def read_autocad_selection(EntityName=("AcDbBlockReference", "AcDbMLeader"),
                           EffectiveName="Мультивыноска v1.1") -> list:
    sset = aDoc.PickfirstSelectionSet
    list_poz = []
    for t in sset:
        if t.EntityName in EntityName[0] and t.EffectiveName == EffectiveName:
            str1 = t.GetAttributes()[0].TextString
            str2 = t.GetAttributes()[1].TextString
            # print(str1, str2)
            list_poz.append((str1, str2))
        elif t.EntityName == EntityName[1]:
            str1 = t.TextString
            str2 = ""
            # print(str1, str2)
            list_poz.append((str1, str2))
    return list_poz


if __name__ == '__main__':
#     elem_list = []
    input('Выбери  группу элементов и нажми ENTER')
    sset = aDoc.PickfirstSelectionSet
    elem_list = find_and_read_parameters(sset)

    for element in elem_list:
        # print(element, type(element))
        print(f'Марка конструкции : {element.constr_name}')
        print(f"Позиции: {element.__getattribute__('list_poz')}", '\n')
        # print(f'{element.contour=}, {type(element.contour[0])}')
        # if "__" not in i:
        #     print(i, element.__getattribute__(i), type(element.__getattribute__(i)), sep=" --> ")
