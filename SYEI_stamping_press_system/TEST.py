import win32com.client as win32
import math
# def import_part(path, file_name):
#     catapp = win32.Dispatch('CATIA.Application')
#     documents1 = catapp.Documents
#     partDocument1 = documents1.Open(path + "\\" + file_name + ".stp")
#
# import_part("C:\\Users\\User\\Desktop\\佳馨學姊分析\\2_0_in\\SP\\", "SP_2_0_in")
#


def create_line():
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Sheet.1")
    drawingViews1 = drawingSheet1.Views
    drawingDimensions1 = drawingViews1.Item("Front view")
    name = ['Point1', 'Point2']
    point = [-50, 100, -100, 100]
    dimension = drawingDimensions1.Dimensions.Add(2, name, point, 1)


create_line()