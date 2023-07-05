import win32com.client as win32

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

pad(50)