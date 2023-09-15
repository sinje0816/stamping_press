import win32com.client as win32
import win32com.client
import main_program as mprog
import parameter as par


def switch_to_window_by_name(window_name):
    try:
        catia = win32com.client.Dispatch("CATIA.Application")
        windows = catia.Windows
        found_window = None

        for win in windows:
            if win.Caption == window_name:
                found_window = win
                break

        if found_window:
            found_window.Activate()
        else:
            print("Window not found:", window_name)
    except Exception as e:
        print("Error:", e)

def copybody():
    catapp = win32.Dispatch('CATIA.Application')
    part_document1 = catapp.ActiveDocument
    selection1 = part_document1.Selection
    selection1.Clear()
    part1 = part_document1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("PartBody")
    selection1.Add(body1)
    selection1.Copy()

def pastebody(count):
    catapp = win32.Dispatch('CATIA.Application')
    part_document1 = catapp.ActiveDocument
    selection1 = part_document1.Selection
    selection1.Clear()
    part1 = part_document1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("PartBody")
    selection1.Add(body1)
    selection1.Paste()
    body = bodies1.Item("Body."+str(9+count))
    body.name = "T-solt"+str(count)

def removebody(count):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part = partDocument1.Part
    shapeFactory1 = part.ShapeFactory
    bodies = part.Bodies
    partbody = bodies.Item("PartBody")
    part.InWorkObject = partbody
    body = bodies.Item("T-solt"+str(count))
    shapeFactory1.AddNewRemove(body)
    part.Update()

def changetranslate(offset_value):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("PartBody")
    hybridShapes1 = body1.HybridShapes
    hybridShapeTranslate1 = hybridShapes1.Item("Translate.10")
    length1 = hybridShapeTranslate1.Distance
    length1.Value = offset_value
    hybridShapeDirection = hybridShapeTranslate1.Direction
    realParam = hybridShapeDirection.GetX()
    realParam.Value = 0
    realParam = hybridShapeDirection.GetY()
    realParam.Value = 1
    realParam = hybridShapeDirection.GetZ()
    realParam.Value = 0
    part1.Update()

def changerotate(rotate_value):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("PartBody")
    hybridShapes1 = body1.HybridShapes
    hybridShapeRotate = hybridShapes1.Item("Rotate.3")
    angle1 = hybridShapeRotate.Angle
    angle1.Value = rotate_value
    part1.Update()

def create_t_solt(translate, count, pierce, clearance_hole, direction, LV):
    try:
        if pierce == '否':
            mprog.partdeactivate('讓孔')
            mprog.partdeactivate('讓孔倒圓角')
    except:
        pass
    try:
        if clearance_hole == '否':
            mprog.partdeactivate('Mirror.4')
    except:
        pass
    changetranslate(translate)
    mprog.Update()
    copybody()
    switch_to_window_by_name("plate.CATPart")
    if direction == '縱向':
        pastebody(count + par.plate_all_parameter['B'])
    else:
        pastebody(count + par.plate_all_parameter['A'] + LV)
    removebody(count)
    mprog.Update()
    switch_to_window_by_name("T.CATPart")
    mprog.close_window()
