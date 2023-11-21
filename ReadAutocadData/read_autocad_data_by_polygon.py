''' выбор объектов внутри полигона'''

from pyautocad import Autocad, APoint, ACAD
import win32com.client
import pythoncom

acad = win32com.client.Dispatch("AutoCAD.Application")
acadDoc = acad.ActiveDocument
acadModel = acad.ActiveDocument.ModelSpace


def select_object_in_rect(coord_list: list):
    def APoint(x, y, z=0):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

    def aDouble(xyz):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (xyz))

    def aVariant(vObject):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, (vObject))

    aaa = acadDoc.ActiveSelectionSet
    aaa.clear()
    # coord_list = list(p1) + list(p2) + list(p3) + list(p4)
    # print(f'{coord_list=}')
    aaa.SelectByPolygon(7, aDouble(coord_list))
    # print(f'{aaa.Count=}')
    # for a in aaa:
    #     print(a.EntityName)
    list_poz =[]
    for acad_obj in aaa:
        if acad_obj.EntityName == "AcDbBlockReference":
            if acad_obj.EffectiveName == "Мультивыноска v1.1":
                str1 = acad_obj.GetAttributes()[0].TextString
                str2 = acad_obj.GetAttributes()[1].TextString
                print(str1, str2)
                list_poz.append(f'{str1}\\P{str2}' if str2 else str1)
        elif acad_obj.EntityName == "AcDbMLeader":
            str1 = acad_obj.TextString
            print(str1)
            list_poz.append(str1)

    return list_poz


if __name__ == '__main__':
    # select_object_in_rect([0, 0, 0.0, 50000, 0, 0.0, 50000, 30000, 0.0, 0, 30000, 0.0, 0, 0, 0.0])
    select_object_in_rect([55940.533, 100198.425, 0.0, 105940.533, 100198.425, 0.0, 105940.533, 130198.425, 0.0, 55940.533, 130198.425, 0.0, 55940.533, 100198.425, 0.0])
    pass
