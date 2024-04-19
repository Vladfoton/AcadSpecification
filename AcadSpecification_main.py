import os
import pathlib
from pathlib import Path

# from ReadAutocadData import *
import win32com.client
from ReadAutocadData.Parameters import Parameters
from ReadAutocadData.read_autocad_data_by_polygon import select_object_in_rect
from ReadAutocadData.readdata import *


def get_file_list_from_folder(path, folder_name: str) -> list:
    '''
    функция возвращает список файлов в папке path/foldername.
    Если такой папки не существует то создает пустую папку и возвращает пустой список
    :param path: рабочий каталог
    :param folder_name: Имя папки
    :return: Список файлов в каталоге
    '''
    try:
        os.chdir(path)
        folder_path = Path(path, folder_name)
        os.chdir(folder_path)
    except FileNotFoundError:
        os.mkdir(folder_path)
        os.chdir(folder_path)
        print(
            f'Внимание! Каталог "{folder_name}" в папке <{path}> не был обнаружен. Создан пустой каталог <{folder_path}>. При необходимости занесите туда информацию')
    return os.listdir(path=".")



if __name__ == '__main__':
    # os.chdir("..")
    curent_dir = pathlib.Path.cwd()  # Текущая рабочая папка
    os.chdir(curent_dir)
    # Пути к файлам
    framework_Path = Path(curent_dir, "Каркасы")  # Путь в папку с каркасами
    embedded_parts_Path = Path(curent_dir, "Закладные детали")  # Путь в папку с закладными
    Specification_Path = Path(curent_dir, "Спецификации и ВРС")  # Путь в папку с результирующими спецификациями
    Ved_det_Path = Path(curent_dir, "Ведомость деталей")  # Путь в папку с вед_деталей

    filelistFramework = get_file_list_from_folder(curent_dir, 'Каркасы')
    filelistemb_parts = get_file_list_from_folder(curent_dir, "Закладные детали")
    list_specification = []  # список спецификаций сборочных единиц
    list_of_embeded_parts = []  # список закладных деталей
    list_objects = []  # список обрабатываемых объектов
    VRSRound = 2  # сколько знаков после разделителя при округлении в ВРС
    list_zones = []  # список объектов Parameters

 #_________________
    app = win32com.client.Dispatch("AutoCAD.Application")   # Подключение к AutoCAD.Application
    aDoc = app.ActiveDocument
    msp = aDoc.ModelSpace

    aDoc.SendCommand("_REGEN ")  #Регенерация пространства модели в Autocad
    input('Выбери  группу элементов c зонами армирования и нажми ENTER')
    sset = aDoc.PickfirstSelectionSet
    elem_list = find_and_read_parameters(sset)

    #test print
    for element in elem_list:
                print(element)
