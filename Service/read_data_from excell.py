import os
import openpyxl  # библиотека для работы с файлами ms excell

def read_position_Emb_Parts_from_Exls(path,mark_emb_parts,level_ID=None):
    os.chdir(path)
    book = openpyxl.open(str(mark_emb_parts)+".xlsx", read_only=True)
    sheet = book.active
    standart = sheet[1][1].value
    mark=sheet[2][1].value
    list_parts=[]
    for c in range(4,sheet.max_row+1):
        Category_poz=sheet[c][0].value
        mark_poz = sheet[c][1].value
        standart_poz = sheet[c][2].value
        size_poz = sheet[c][3].value
        mass1_poz = sheet[c][4].value
        count_poz = sheet[c][5].value
        material_poz = sheet[c][6].value
        material_standart = sheet[c][7].value
        list_parts.append(Different_position(Category_poz,mark_poz,standart_poz,size_poz,mass1_poz,count_poz,material_poz,material_standart,mark_emb_parts, standart,mark,level_ID=level_ID))
    return list_parts