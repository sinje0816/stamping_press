import win32com.client as win32

# 孔洞開關
def activatehole():
    catapp = win32.Dispatch('CATIA.Application')
    activedoc = catapp.ActiveDocument
    part = activedoc.Part
    bodies = part.Bodies
    body = bodies.Item("PartBody")
    shapes = body.Shapes
    target = '3-M8通'
    remove = shapes.Item(target)
    sub_body = remove.Body
    sub_shapes = sub_body.Shapes
    shapes_cout = sub_shapes.Count

    # here to start activating process
    part.Activate(remove)
    for i in range(1, shapes_cout + 1):
        item_obj = sub_shapes.Item(i)
        part.Activate(item_obj)

    part.Update()
