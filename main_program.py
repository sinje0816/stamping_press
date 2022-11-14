import os
import win32com.client as win32
import datetime


# 開啟CATIA(由學長提供之函式)
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


def folder_file_name():  # 抓取檔案名稱
    path = os.listdir('C:\\Users\\User\\Desktop\\stamping_press')
    x = []
    for file_name in path:
        x.append(file_name.split('.')[0])
    print(x)


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
    part.Update()


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


# 開啟零件檔
# def open_part():
#     catapp = win32.Dispatch('CATIA.Application')
#     documents1 = catapp.Documents
#     partDocument1 = documents1.Add("Part")

# 零件檔修改
def import_part(path, file_name):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(path + "\\" + file_name + ".CATPart")

# 儲存零件檔
def save_file(path, file_name):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    partDocument1.SaveAs(path + '\\' + file_name + "CATPrat")
    partDocument1.Close()

# 新增資料夾
def new_Folder():
    time = datetime.datetime.now()
    print(time.day, time.hour, time.minute, time.second)
    dir = 'stamping_press' + '_' + str(time.month) + '_' + str(time.day) + '_' + str(time.hour) + '_' + str(
        time.minute) + '_' + str(time.second)
    path = 'C:\\Users\\USER\\Desktop' + '\\' + dir
    os.mkdir(path)
    return path, dir


# catapp = win32.Dispatch('CATIA.Application')
# productDocument1 = catapp.ActiveDocument
# product1 = productDocument1.Product
# products1 = product1.Products
# products1Variant = products1
# combination_file = []#組合檔
# combination_file.append(path + '\\' + file_name + '.CATProduct')
# print(file_name)
# products1Variant.AddComponentsFromFiles(combination_file, "All")

# def combined_dimension():
#     catapp = win32.Dispatch("CATIA.Application")
#     productDocument = catapp.ActiveDocument
#     product = productDocument.Product
#     constraints = product.Connections("CATIAConstraints")
#     constraint = constraints.Item("Offset.173")
#     length = constraint.Dimension
#     length.Value = 1480
#     constraint.Orientation = 1
#     product.Update()

# 偏移組合檔
def add_offset_assembly(element1, element2, dist, relation,
                        binding_conditions):  # (組合檔1, 零件1, 組合檔2, 零件2, 距離, 結合面, 拘束條件)
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
    # product.Update()
    return True


def base_lock(element1, element2):  # 定海神針, 固定基準零件
    catapp = win32.Dispatch('CATIA.Application')
    productDocument = catapp.ActiveDocument
    product = productDocument.Product
    constraints = product.Connections("CATIAConstraints")
    ref = product.CreateReferenceFromName("Product1/%s/!Product1/%s/" % (element1, element2))
    constraint = constraints.AddMonoEltCst(0, ref)
    product.Update()
    return True


def update():
    catapp = win32.Dispatch('CATIA.Application')
    productDocument = catapp.ActiveDocument
    product = productDocument.Product
    product.Update()
    specsAndGeomWindow = catapp.ActiveWindow
    viewer3D = specsAndGeomWindow.ActiveViewer
    viewer3D.Reframe()


def close_plane():
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("PartBody")
    hybridShapes1 = body1.HybridShapes
    hybridShapePlaneOffset1 = hybridShapes1.Item("XZ.PLANE")
    hybridShapes1 = hybridShapePlaneOffset1.Parent
    selection1.Add(hybridShapePlaneOffset1)
    hybridShapePlaneAngle1 = hybridShapes1.Item("YZ.PLANE")
    hybridShapes1 = hybridShapePlaneAngle1.Parent
    selection1.Add(hybridShapePlaneAngle1)
    hybridShapePlaneAngle2 = hybridShapes1.Item("XY.PLANE")
    hybridShapes1 = hybridShapePlaneAngle2.Parent
    selection1.Add(hybridShapePlaneAngle2)
    visPropertySet1 = visPropertySet1.Parent
    visPropertySet1.SetShow(1)
    selection1.Clear()

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