''' выбор объектов внутри полигона'''

from pyautocad import Autocad, APoint, ACAD
import win32com.client
import pythoncom

acad = win32com.client.Dispatch("AutoCAD.Application")
acadDoc = acad.ActiveDocument
acadModel = acad.ActiveDocument.ModelSpace


def select_object_in_rect(p1: tuple, p2: tuple, p3: tuple, p4: tuple):
    def APoint(x, y, z=0):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

    def aDouble(xyz):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (xyz))

    def aVariant(vObject):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, (vObject))

    aaa = acadDoc.ActiveSelectionSet
    coord_list = list(p1) + list(p2) + list(p3) + list(p4)
    # print(f'{coord_list=}')
    aaa.SelectByPolygon(7, aDouble(coord_list))
    print(f'{aaa.Count=}')
    for a in aaa:
        print(a.EntityName)
    # aaa = aVariant(aaa)
    # print(aaa)
    # # acad.ZoomExtents()
    return aaa


if __name__ == '__main__':
    select_object_in_rect((0, 0, 0), (25000, 0, 0), (25000, 10000, 0), (0, 10000, 0))
    pass
