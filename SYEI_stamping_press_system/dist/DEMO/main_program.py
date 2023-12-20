import os
import win32com.client as win32
import datetime, time
import file_path as fp
import zipfile as zf


# 開啟CATIA
class set_CATIA_workbench_env:
    def __init__(self):
        # self.catapp = win32.Dispatch("CATIA.Application")
        self.env_name = {'Part_Design': 'PrtCfg', 'Product_Assembly': 'Assembly',
                         'Generative_Sheetmetal_Design': 'SmdNewDesignWorkbench', 'Drafting': 'Drw'}
        self.catapp = win32.Dispatch("CATIA.Application")
        # add some CATIA-specific settings here (Seeing CATIA Automation manual-Application section)
        self.catapp.DisplayFileAlerts = False
        self.catapp.Visible = True

    def Part_Design(self):
        self.catapp.Visible = True
        self.catapp.StartWorkbench(self.env_name[self.Part_Design.__name__])
        try:
            temp = self.catapp.ActiveDocument
            temp.close()
        except:
            pass
        return

    def Product_Assembly(self):
        self.catapp.Visible = True
        self.catapp.StartWorkbench(self.env_name[self.Product_Assembly.__name__])
        try:
            temp = self.catapp.ActiveDocument
            temp.close()
        except:
            pass
        return

    def Generative_Sheetmetal_Design(self):
        self.catapp.Visible = True
        self.catapp.StartWorkbench(self.env_name[self.Generative_Sheetmetal_Design.__name__])
        try:
            temp = self.catapp.ActiveDocument
            temp.close()
        except:
            pass
        return

    def Drafting(self):
        self.catapp.Visible = True
        self.catapp.StartWorkbench(self.env_name[self.Drafting.__name__])
        try:
            temp = self.catapp.ActiveDocument
            temp.close()
        except:
            pass
        return


# 開啟組合檔
def assembly_create():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    productDocument1 = documents1.Add("Product")


# 匯入組合檔
def import_file_Product(path, file_name):  # (資料夾路徑，檔案名稱)
    catapp = win32.Dispatch('CATIA.Application')
    productDocument1 = catapp.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    products1Variant = products1
    combination_file = []  # 組合檔
    combination_file.append(path + '\\' + file_name + '.CATProduct')
    products1Variant.AddComponentsFromFiles(combination_file, "All")


# 匯入零件檔
def import_file_Part(path, file_name):  # (資料夾路徑，檔案名稱)
    catapp = win32.Dispatch('CATIA.Application')
    productDocument1 = catapp.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    products1Variant = products1
    combination_file = []  # 組合檔
    combination_file.append(path + '\\' + file_name + '.CATPart')
    products1Variant.AddComponentsFromFiles(combination_file, "All")


# 匯入組合檔
def import_file_Product(path, file_name):  # (資料夾路徑，檔案名稱)
    catapp = win32.Dispatch('CATIA.Application')
    productDocument1 = catapp.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    products1Variant = products1
    combination_file = []  # 組合檔
    combination_file.append(path + '\\' + file_name + '.CATProduct')
    products1Variant.AddComponentsFromFiles(combination_file, "All")


def folder_file_name():  # 抓取檔案名稱
    path = os.listdir(fp.system_root + fp.part)
    x = []
    for file_name in path:
        x.append('or y == \'' + file_name.split('.')[0] + '\'')
        print(' '.join(x))


# frame名稱
def frame_name():
    x = []
    for a in range(1, 44):
        x.append('FRAME' + str(a))
    print(x)


# 偏移組合檔_0(0表示OPPOSITE)
def add_offset_assembly_0(element1, element2, element3, element4, element5, dist,
                          relation):  # (組合檔1, 零件1, 組合檔2, 零件2, 距離, 結合面)
    catapp = win32.Dispatch("CATIA.Application")
    productdoc = catapp.ActiveDocument
    product = productdoc.Product
    product = product.ReferenceProduct
    constraints = product.Connections("CATIAConstraints")
    ref1 = product.CreateReferenceFromName(
        "Product1/%s.1/%s/!PartBody/!%s" % (element1, element2, relation))  # (將%指定檔內容移入%s)
    ref2 = product.CreateReferenceFromName(
        "Product1/%s.1/%s.1/%s/!PartBody/!%s" % (element3, element4, element5, relation))
    constraint = constraints.AddBiEltCst(1, ref1, ref2)  # (1表示偏移拘束, ref1, ref2)
    length = constraint.Dimension
    length.value = dist
    constraint.Orientation = 0
    product.Update()
    return True


# 偏移組合檔_1(1表示SAME)
def add_offset_assembly_1(element1, element2, element3, element4, element5, dist,
                          relation):  # (組合檔1, 零件1, 組合檔2, 零件2, 距離, 結合面)
    catapp = win32.Dispatch("CATIA.Application")
    productdoc = catapp.ActiveDocument
    product = productdoc.Product
    product = product.ReferenceProduct
    constraints = product.Connections("CATIAConstraints")
    ref1 = product.CreateReferenceFromName(
        "Product1/%s.1/%s/!PartBody/!%s" % (element1, element2, relation))  # (將%指定檔內容移入%s)
    ref2 = product.CreateReferenceFromName(
        "Product1/%s.1/%s.1/%s/!PartBody/!%s" % (element3, element4, element5, relation))
    constraint = constraints.AddBiEltCst(1, ref1, ref2)  # (1表示偏移拘束, ref1, ref2)
    length = constraint.Dimension
    length.value = dist
    constraint.Orientation = 1
    product.Update()
    return True


def combination_file_open(target, dir):
    # 連結CATIA
    catapp = win32.Dispatch("CATIA.Application")
    document = catapp.Documents
    # 將路徑設為目錄的文字宣告
    directory = str(dir)
    # directory = '\\'.join(directory.split('/'))
    print(directory)
    # gvar.folderdir = directory
    # 定義零件檔檔名
    part_dir = directory + target + '.CATProduct'
    print(part_dir)
    # partdoc = document.Open("%s%s.%s" % (directory,target,"CATPart"))
    # 開啟該零件檔
    partdoc = document.Open(part_dir)
    return target + '.CATPart'


def param_change(file_name, target, value):
    catapp = win32.Dispatch("CATIA.Application")
    productDocument = catapp.ActiveDocument
    product = productDocument.Product
    product = product.ReferenceProduct
    documents = catapp.Documents
    partDocument = documents.Item(file_name + ".CATPart")
    part = partDocument.Part
    parameter = part.Parameters
    # 按照變數尺寸修改板件長度
    length = parameter.Item(target)
    length.Value = value
    productDocument = catapp.ActiveDocument


# FRAME結合尺寸parameter, parameter_dimension, combined_number
def combined_dimension(combination_number, combined_dimension):  # 組合編號, 組合尺寸
    catapp = win32.Dispatch("CATIA.Application")
    productDocument = catapp.ActiveDocument
    product = productDocument.Product
    constraints = product.Connections("CATIAConstraints")
    constraint = constraints.Item(combination_number)
    length = constraint.Dimension
    length.Value = combined_dimension
    constraint.Orientation = 1
    product.Update()


# 零件檔修改
def import_part(path, file_name):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(path + "\\" + file_name + ".CATPart")


def import_product(path, file_name):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(path + "\\" + file_name + ".CATProduct")


# 零件檔存檔
def save_file_part(path, file_name):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    file_name = file_name + '.CATPart'
    partDocument1 = document.Item(file_name)
    # print(path + '\\' + file_name)
    partDocument1.SaveAs(path + '\\' + file_name)
    partDocument1.Close()


def save_file_stp(path, file_name):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    stp_file_name = file_name + ".stp"
    file_name = file_name + '.CATPart'
    partDocument1 = document.Item(file_name)
    # print(path + '\\' + file_name)
    partDocument1.SaveAs(path + '\\' + stp_file_name)
    # partDocument1.Close()


def save_stpfile_part(path, file_name):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    part_file_name = file_name + '.CATPart'
    file_name = file_name + '.stp'
    partDocument1 = document.Item(file_name)
    # print(path + '\\' + file_name)
    partDocument1.SaveAs(path + '\\' + part_file_name)
    # partDocument1.Close()


# 組合檔存檔
def save_file_product(path, file_name):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    file_name = file_name + '.CATProduct'
    partDocument1 = document.Item(file_name)
    # print(path + '\\' + file_name)
    partDocument1.SaveAs(path + '\\' + file_name)


# 新增資料夾
def new_Folder():
    time = datetime.datetime.now()
    print(time.day, time.hour, time.minute, time.second)
    dir = 'stamping_press' + '_' + str(time.month) + '_' + str(time.day) + '_' + str(time.hour) + '_' + str(
        time.minute) + '_' + str(time.second)
    path = 'C:\\Users\\USER\\Desktop' + '\\' + dir
    os.mkdir(path)
    return path, dir


# 偏移組合檔(零件結合)
def add_offset_assembly(element1, element2, dist, relation, binding_conditions,
                        name):  # (組合檔1, 零件1, 組合檔2, 零件2, 距離, 結合面, 拘束條件)
    catapp = win32.Dispatch('CATIA.Application')
    productDocument = catapp.ActiveDocument
    product = productDocument.Product
    constraints = product.Connections("CATIAConstraints")
    ref1 = product.CreateReferenceFromName("Product1/%s/!PartBody/!%s" % (element1, relation))  # (將%指定檔內容移入%s)
    ref2 = product.CreateReferenceFromName("Product1/%s/!PartBody/!%s" % (element2, relation))
    constraint = constraints.AddBiEltCst(1, ref1, ref2)  # (1表示偏移拘束, ref1, ref2)
    length = constraint.Dimension
    length.value = dist
    constraint.Orientation = binding_conditions  # (1表示SAME, 0表示OPPOSITE)
    constraint.Name = name
    return True


# 偏移組合檔(標準件與零件結合)
def add_offset_product_assembly(element1, element2, element3, dist, relation, binding_conditions,
                                name):  # (組合檔1, 零件1, 組合檔2, 零件2, 距離, 結合面, 拘束條件)
    catapp = win32.Dispatch('CATIA.Application')
    productDocument = catapp.ActiveDocument
    product = productDocument.Product
    constraints = product.Connections("CATIAConstraints")
    ref1 = product.CreateReferenceFromName("Product1/%s/!%s/%s" % (element1, element2, relation))  # (將%指定檔內容移入%s)
    ref2 = product.CreateReferenceFromName("Product1/%s/!PartBody/%s" % (element3, relation))
    constraint = constraints.AddBiEltCst(1, ref1, ref2)  # (1表示偏移拘束, ref1, ref2)
    length = constraint.Dimension
    length.value = dist
    constraint.Orientation = binding_conditions  # (1表示SAME, 0表示OPPOSITE)
    constraint.Name = name
    return True


def base_lock(element1, element2, name):  # 定海神針, 固定基準零件
    catapp = win32.Dispatch('CATIA.Application')
    productDocument = catapp.ActiveDocument
    product = productDocument.Product
    constraints = product.Connections("CATIAConstraints")
    ref = product.CreateReferenceFromName("Product1/%s/!Product1/%s/" % (element1, element2))
    constraint = constraints.AddMonoEltCst(0, ref)
    constraint.Name = name
    return True


def update():
    catapp = win32.Dispatch('CATIA.Application')
    productDocument = catapp.ActiveDocument
    product = productDocument.Product
    product.Update()
    specsAndGeomWindow = catapp.ActiveWindow
    viewer3D = specsAndGeomWindow.ActiveViewer
    viewer3D.Reframe()


# 組合拘束隱藏
def hide_ass_all_Constraint():
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    drawingdocument = catapp.ActiveDocument
    selection1 = drawingdocument.Selection
    selection1.Search("Type='Assembly Design'.Constraint,all")
    selection1.VisProperties.SetShow(1)
    selection1.Clear()
    selection1.Search("Type='Part Design'.Plane,all")
    selection1.VisProperties.SetShow(1)
    selection1.Clear()


def axis_system():  # 開啟座標軸
    catapp = win32.Dispatch('CATIA.Application')
    Part = catapp.ActiveDocument.Part
    AxisSystems = Part.AxisSystems
    AxisSystem = AxisSystems.Add()
    Hybridbodys = Part.HybridBodies
    Part.Update()


def scaling(X):  # 零件等比例縮小
    catapp = win32.Dispatch('CATIA.Application')
    partDocument = catapp.ActiveDocument
    part = partDocument.Part
    shapeFactory = part.ShapeFactory
    axisSystems = part.AxisSystems
    axisSystem = axisSystems.Item("Axis System.1")
    reference = part.CreateReferenceFromBRepName(
        "FVertex:(Vertex:(Neighbours:(Face:(Brp:(AxisSystem.1;2);None:();Cf11:());Face:(Brp:(AxisSystem.1;3);None:();Cf11:());Face:(Brp:(AxisSystem.1;1);None:();Cf11:()));Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        axisSystem)
    scaling = shapeFactory.AddNewScaling2(reference, X)
    part.InWorkObject = scaling
    part.Update()


# 關閉實體外所有東西
def Close_All():
    catapp = win32.Dispatch('CATIA.Application')
    partdocument1 = catapp.ActiveDocument
    selection1 = partdocument1.Selection
    selection1.Search(
        "((((((((CATStFreeStyleSearch.OpenBodyFeature + CATPrtSearch.OpenBodyFeature) + CATGmoSearch.OpenBodyFeature) + CATSpdSearch.OpenBodyFeature) + CATPrtSearch.Sketch + CATPrtSearch.Plane) + CATPrtSearch.MfConstraint) + CATPrtSearch.AxisSystem) + CATPrtSearch.Point) + CATPrtSearch.Line),all")
    visPropertySet1 = selection1.VisProperties
    visPropertySet1.SetShow(1)


def del_ass_all_Constraint():
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    drawingdocument = catapp.ActiveDocument
    selection1 = drawingdocument.Selection
    selection1.Search("Type='Assembly Design'.Constraint,all")
    selection1.VisProperties.SetShow(1)
    selection1.Delete()


def saveas(save_dir, target, data_type):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    try:
        saveas = document.Item('%s%s' % (target, data_type))
        saveas.SaveAs('%s\%s%s' % (save_dir, target, data_type))
    except:
        saveas = catapp.ActiveDocument
        saveas.SaveAs('%s\%s%s' % (save_dir, target, data_type))
    finally:
        saveas.Save()


def PDF_save(save_dir, name):
    catapp = win32.Dispatch("CATIA.Application")
    partdoc = catapp.ActiveDocument
    # 儲存檔案後關閉
    partdoc.ExportData("%s\%s.%s" % (save_dir, name, "pdf"), "pdf")


def constaint_value_change(constraints_name, value, Orientation):
    catapp = win32.Dispatch('CATIA.Application')
    productDocument = catapp.ActiveDocument
    product = productDocument.Product
    product = product.ReferenceProduct
    constraints = product.Connections("CATIAConstraints")
    constraint = constraints.Item(str(constraints_name))
    length = constraint.Dimension
    length.Value = value
    constraint.Orientation = Orientation


def OPEN_Drawing():
    catapp = win32.Dispatch('CATIA.Application')
    documents = catapp.Documents
    drawingDocument = documents.Open(fp.system_root + fp.part + fp.Exploded_Views)


def OPEN_detail_drawing():
    catapp = win32.Dispatch('CATIA.Application')
    documents = catapp.Documents
    drawingDocument = documents.Open(fp.system_root + fp.part + fp.detail_drawing)


def OPEN_Welding_diagram():
    catapp = win32.Dispatch('CATIA.Application')
    documents = catapp.Documents
    drawingDocument = documents.Open(fp.system_root + fp.part + fp.Welding_diagram)


def Close_all_window():
    catapp = win32.Dispatch('CATIA.Application')


# 零件圖存檔
def save_detail_drawing(path, file_name):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    file_name = file_name + '.CATDrawing'
    partDocument2 = document.Item(file_name)
    # print(path + '\\' + file_name)
    partDocument2.SaveAs(path + '\\' + file_name)
    partDocument2.Close()


def switch_window():
    catapp = win32.Dispatch("CATIA.Application")
    specsAndGeomWindow1 = catapp.ActiveWindow
    specsAndGeomWindow1.ActivateNext()


def hide_show_part(part_name, hide_show):  # 隱藏零件
    catapp = win32.Dispatch("CATIA.Application")
    productDocument1 = catapp.ActiveDocument
    selection1 = productDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    product1 = productDocument1.Product
    products1 = product1.Products
    product2 = products1.Item(part_name + '.1')
    products1 = product2.Parent
    selection1.Add(product2)
    visPropertySet1 = visPropertySet1.Parent
    visPropertySet1.SetShow(hide_show)  # 隱藏1, 顯示0
    selection1.Clear()


# 啟動remove下body特徵
def activatefeature(feature, howmany):
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
    if howmany == 0:
        for i in range(1, shapes_count + 1):
            item_obj = sub_shapes.Item(i)

            # 草圖激活
            # noinspection PyBroadException
            try:
                sketch = item_obj.Sketch
                part.Activate(sketch)
            except:
                pass
            part.Activate(item_obj)
    else:
        for i in range(1, howmany + 1):
            item_obj = sub_shapes.Item(i)

            # 草圖激活
            # noinspection PyBroadException
            try:
                sketch = item_obj.Sketch
                part.Activate(sketch)
            except:
                pass
            part.Activate(item_obj)


# 啟動partbody下特徵
def partbodyfeatureactivate(feature):
    catapp = win32.Dispatch('CATIA.Application')
    doc = catapp.ActiveDocument
    part = doc.Part
    bodies = part.Bodies
    body = bodies.Item("PartBody")
    shapes = body.Shapes
    target = feature
    pocket = shapes.Item(target)
    part.Activate(pocket)
    # 草圖激活
    # noinspection PyBroadException
    try:
        sketch = pocket.Sketch
        part.Activate(sketch)
    except:
        pass


# 抑制功能(partbody)
def partdeactivate(feature):
    catapp = win32.Dispatch('CATIA.Application')
    doc = catapp.ActiveDocument
    part = doc.Part
    bodies = part.Bodies
    body = bodies.Item("PartBody")
    shapes = body.Shapes
    target = feature
    pocket = shapes.Item(target)
    # 草圖激活
    # noinspection PyBroadException
    try:
        sketch = pocket.Sketch
        part.Inactivate(sketch)
    except:
        pass
    part.Inactivate(pocket)


# 抑制功能(body)
def bodydeactivate(feature, howmany):
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
    part.Inactivate(remove)
    if howmany == 0:
        for i in range(1, shapes_count + 1):
            item_obj = sub_shapes.Item(i)

            # 草圖激活
            # noinspection PyBroadException
            try:
                sketch = item_obj.Sketch
                part.Inactivate(sketch)
            except:
                pass
            part.Inactivate(item_obj)
    else:
        for i in range(1, howmany + 1):
            item_obj = sub_shapes.Item(i)

            # 草圖激活
            # noinspection PyBroadException
            try:
                sketch = item_obj.Sketch
                part.Inactivate(sketch)
            except:
                pass
            part.Inactivate(item_obj)


# 零件更新
def Update():
    catapp = win32.Dispatch('CATIA.Application')
    doc = catapp.ActiveDocument
    part = doc.Part
    part.Update()


def close_file(file_name):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    file_name = file_name + '.CATPart'
    partDocument = document.Item(file_name)
    partDocument.Close()


def close_window():
    catapp = win32.Dispatch('CATIA.Application')
    specsAndGeomWindow1 = catapp.ActiveWindow
    specsAndGeomWindow1.Close()