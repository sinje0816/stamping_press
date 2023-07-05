import win32com.client as win32


# 孔洞開關
def activatehfeatrue(feature):
    catapp = win32.Dispatch('CATIA.Application')
    activedoc = catapp.ActiveDocument
    part = activedoc.Part
    bodies = part.Bodies
    body = bodies.Item("PartBody")
    shapes = body.Shapes
    target = feature
    remove = shapes.Item(target)
    sub_body = remove.Body
    sub_shapes = sub_body.Shapes
    shapes_count = sub_shapes.Count

    # here to start activating process
    part.Activate(remove)
    for i in range(1, shapes_count + 1):
        item_obj = sub_shapes.Item(i)

        # 草圖激活
        try:
            sketch = item_obj.Sketch
            part.Activate(sketch)
        except:
            pass

        part.Activate(item_obj)

    part.Update()

activatehfeatrue('3-M8通')