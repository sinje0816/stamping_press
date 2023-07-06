import win32com.client as win32
import excel_parameter_change as epc
import main_program as mprog


# 0706作業
# 長料功能
def pad(value):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    shapeFactory1 = part1.ShapeFactory
    reference1 = part1.CreateReferenceFromName("")
    pad1 = shapeFactory1.AddNewPadFromRef(reference1, value)
    bodies1 = part1.Bodies
    body1 = bodies1.Item("PartBody")
    sketches1 = body1.Sketches
    sketch1 = sketches1.Item("Sketch.1")
    reference2 = part1.CreateReferenceFromObject(sketch1)
    pad1.SetProfileElement(reference2)
    part1.Update()

i = 0
excel = epc.ExcelOp('FRAME5')
excel.part_parameter('FRAME5', i)
mprog.partbodyfeatureactivate('SN1-25_X')
mprog.partbodyfeatureactivate('before80_I', 0)
mprog.partbodyfeatureactivate('before80_R', 0)
mprog.partbodyfeatureactivate('Y', 0)
mprog.activatefeature('2-M8通孔', 0)
mprog.activatefeature('TY牌電磁閥用', 0)
mprog.activatefeature('M6通孔LED', 0)
mprog.activatefeature('PT1/2', 0)