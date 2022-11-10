import os
import win32com.client as win32


def assembly_create():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    productDocument1 = documents1.Add("Product")


def import_file(path, file_name):#(資料夾路徑，檔案名稱)
    catapp = win32.Dispatch('CATIA.Application')
    productDocument1 = catapp.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    products1Variant = products1
    combination_file = []#組合檔

    print(combination_file.append(path + '\\' + file_name + '.CATProduct'))
    products1Variant.AddComponentsFromFiles(combination_file, "All")

def folder_file_name():#抓取檔案名稱
    path = os.listdir('C:\\Users\\User\\Desktop\\stamping_press')
    x = []
    for file_name in path:
        x.append(file_name.split('.')[0])
    print(x)

def product_assemble():#結合組合檔
    catapp = win32.Dispatch('CATIA.Application')
    productDocument1 = catapp.ActiveDocument
    product1 = productDocument1.Product
    product1 = product1.ReferenceProduct
    constraints1 = product1.Connections("CATIAConstraints")
    reference1 = product1.CreateReferenceFromName("Product1/SLIDE_UNIT.1/BOLSTER3/!PartBody/Plane.4")
    reference2 = product1.CreateReferenceFromName("Product1/FRAME.1/BOLSTER1.1/!PartBody/FRAME_XY")
    constraint1 = constraints1.AddBiEltCst(1, reference1, reference2)
    length1 = constraint1.Dimension
    length1.Value = 0
    constraint1.Orientation = 0
    constraints1 = product1.Connections("CATIAConstraints")
    reference3 = product1.CreateReferenceFromName("Product1/SLIDE_UNIT.1/BOLSTER3/!PartBody/Plane.5")
    reference4 = product1.CreateReferenceFromName("Product1/FRAME.1/BOLSTER1.1/!PartBody/FRAME_YZ")
    constraint2 = constraints1.AddBiEltCst(1, reference3, reference4)
    length2 = constraint2.Dimension
    length2.Value = 0
    constraint2.Orientation = 0
    constraints1 = product1.Connections("CATIAConstraints")
    reference5 = product1.CreateReferenceFromName("Product1/FRAME.1/BOLSTER1.1/!PartBody/FRAME_XZ")
    reference6 = product1.CreateReferenceFromName("Product1/SLIDE_UNIT.1/BOLSTER3/!PartBody/Plane.6")
    constraint3 = constraints1.AddBiEltCst(1, reference5, reference6)
    length3 = constraint3.Dimension
    length3.Value = 0
    constraint3.Orientation = 0
    product1.Update()

file_name_list = ['SLIDE_UNIT', 'FRAME', 'CRANK_SHAFT', 'CLUCTH_ASSEMBLY', 'BALANCER', 'BALANCER']

assembly_create()
for x in file_name_list:
    # print(x)
    import_file('C:\\Users\\User\\Desktop\\stamping_press', x)


