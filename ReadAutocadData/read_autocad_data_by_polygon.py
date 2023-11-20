''' выбор объектов внутри полигона'''

from pyautocad import Autocad, APoint, ACAD
import win32com.client
import pythoncom

acad = win32com.client.Dispatch("AutoCAD.Application")
acadDoc = acad.ActiveDocument
acadModel = acad.ActiveDocument.ModelSpace

def APoint(x, y, z=0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def aDouble(xyz):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (xyz))

def aVariant(vObject):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, (vObject))

aaa = acadDoc.ActiveSelectionSet
aaa.SelectByPolygon(7, aDouble([-949916.5235,1656009.9275,0, -949873.1973,1665494.9022,0, -914996.7868,1665335.5911,0, -915038.986,1656097.3407,0, -949916.5235,1656009.9275,0]))
print(aaa.Count)
for a in aaa:
    print(a.EntityName)
aaa = aVariant(aaa)
# acad.ZoomExtents()