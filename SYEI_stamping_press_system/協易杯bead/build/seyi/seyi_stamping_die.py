import win32com.client as win32
import main_program as mprog
import os

def holder_extract1():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("holder.CATPart")
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Extrude4")
    part1.InWorkObject = body1
    hybridShapeFactory1 = part1.HybridShapeFactory
    shapes1 = body1.Shapes
    slot1 = shapes1.Item("Slot.3")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Extrude4;7);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        slot1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = False
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()


def holder_extract2():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("holder.CATPart")
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Extrude4")
    part1.InWorkObject = body1
    hybridShapeFactory1 = part1.HybridShapeFactory
    shapes1 = body1.Shapes
    slot1 = shapes1.Item("Slot.3")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Slot.1;0:(Brp:(Sketch.1;54);Brp:(Sketch.1;58);Brp:(Sketch.1;59);Brp:(Sketch.1;60);Brp:(Sketch.1;61);Brp:(Sketch.1;62);Brp:(Sketch.1;63);Brp:(Sketch.2;3);Brp:(Sketch.1;51)));None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        slot1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = False
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()


def holder_extract3():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("holder.CATPart")
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Extrude4")
    part1.InWorkObject = body1
    hybridShapeFactory1 = part1.HybridShapeFactory
    shapes1 = body1.Shapes
    slot1 = shapes1.Item("Slot.3")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Slot.3;0:(Brp:(Sketch.3;43);Brp:(Sketch.3;50);Brp:(Sketch.3;51);Brp:(Sketch.3;52);Brp:(Sketch.3;53);Brp:(Sketch.3;54);Brp:(Sketch.3;55);Brp:(Sketch.5;3);Brp:(Sketch.3;46)));None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        slot1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = False
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()


def holder_extract4():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("holder.CATPart")
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Extrude4")
    part1.InWorkObject = body1
    hybridShapeFactory1 = part1.HybridShapeFactory
    shapes1 = body1.Shapes
    slot1 = shapes1.Item("Slot.3")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Slot.3;2);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        slot1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = False
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()


def holder_extract5():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("holder.CATPart")
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Extrude4")
    part1.InWorkObject = body1
    hybridShapeFactory1 = part1.HybridShapeFactory
    shapes1 = body1.Shapes
    slot1 = shapes1.Item("Slot.3")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Slot.3;1);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        slot1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = False
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()


def holder_extract6():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("holder.CATPart")
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Extrude4")
    part1.InWorkObject = body1
    hybridShapeFactory1 = part1.HybridShapeFactory
    shapes1 = body1.Shapes
    slot1 = shapes1.Item("Slot.3")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Slot.1;1);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        slot1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = False
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()


def holder_extract7():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("holder.CATPart")
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Extrude4")
    part1.InWorkObject = body1
    hybridShapeFactory1 = part1.HybridShapeFactory
    shapes1 = body1.Shapes
    slot1 = shapes1.Item("Slot.3")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Slot.1;2);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        slot1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = False
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()


def holder_joinextract():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("holder.CATPart")
    part1 = partDocument1.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Extrude4")
    hybridShapes1 = body1.HybridShapes
    hybridShapeExtract1 = hybridShapes1.Item("Extract.2")
    reference1 = part1.CreateReferenceFromObject(hybridShapeExtract1)
    hybridShapeExtract2 = hybridShapes1.Item("Extract.3")
    reference2 = part1.CreateReferenceFromObject(hybridShapeExtract2)
    hybridShapeAssemble1 = hybridShapeFactory1.AddNewJoin(reference1, reference2)
    hybridShapeExtract3 = hybridShapes1.Item("Extract.4")
    reference3 = part1.CreateReferenceFromObject(hybridShapeExtract3)
    hybridShapeAssemble1.AddElement(reference3)
    hybridShapeExtract4 = hybridShapes1.Item("Extract.5")
    reference4 = part1.CreateReferenceFromObject(hybridShapeExtract4)
    hybridShapeAssemble1.AddElement(reference4)
    hybridShapeExtract5 = hybridShapes1.Item("Extract.6")
    reference5 = part1.CreateReferenceFromObject(hybridShapeExtract5)
    hybridShapeAssemble1.AddElement(reference5)
    hybridShapeExtract6 = hybridShapes1.Item("Extract.7")
    reference6 = part1.CreateReferenceFromObject(hybridShapeExtract6)
    hybridShapeAssemble1.AddElement(reference6)
    hybridShapeExtract7 = hybridShapes1.Item("Extract.8")
    reference7 = part1.CreateReferenceFromObject(hybridShapeExtract7)
    hybridShapeAssemble1.AddElement(reference7)
    hybridShapeAssemble1.SetConnex(1)
    hybridShapeAssemble1.SetManifold(1)
    hybridShapeAssemble1.SetSimplify(0)
    hybridShapeAssemble1.SetSuppressMode(0)
    hybridShapeAssemble1.SetDeviation(0.001)
    hybridShapeAssemble1.SetAngularToleranceMode(0)
    hybridShapeAssemble1.SetAngularTolerance(0.5)
    hybridShapeAssemble1.SetFederationPropagation(0)
    body1.InsertHybridShape(hybridShapeAssemble1)
    part1.InWorkObject = hybridShapeAssemble1
    part1.Update()


def hidebody(name):
    catapp = win32.Dispatch('CATIA.Application')
    productDocument1 = catapp.ActiveDocument
    selection1 = productDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("holder.CATPart")
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Extrude4")
    hybridShapes1 = body1.HybridShapes
    hybridShapeExtract1 = hybridShapes1.Item(name)
    hybridShapes1 = hybridShapeExtract1.Parent
    bSTR1 = hybridShapeExtract1.Name
    selection1.Add(hybridShapeExtract1)
    visPropertySet1 = visPropertySet1.Parent
    bSTR2 = visPropertySet1.Name
    bSTR3 = visPropertySet1.Name
    visPropertySet1.SetShow(1)
    selection1.Clear()


def hidebody2():
    catapp = win32.Dispatch('CATIA.Application')
    productDocument1 = catapp.ActiveDocument
    selection1 = productDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("holder.CATPart")
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Extrude4")
    shapes1 = body1.Shapes
    solid1 = shapes1.Item("Extrude4")
    shapes1 = solid1.Parent
    bSTR1 = solid1.Name
    selection1.Add(solid1)
    visPropertySet1 = visPropertySet1.Parent
    bSTR2 = visPropertySet1.Name
    bSTR3 = visPropertySet1.Name
    visPropertySet1.SetShow(1)
    selection1.Clear()



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


def die_extract1():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("die-2.CATPart")
    part1 = partDocument1.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Fillet2")
    shapes1 = body1.Shapes
    rib1 = shapes1.Item("Rib.4")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Fillet2;10);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        rib1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = True
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()


def die_extract2():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("die-2.CATPart")
    part1 = partDocument1.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Fillet2")
    shapes1 = body1.Shapes
    rib1 = shapes1.Item("Rib.4")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Rib.4;0:(Brp:(Sketch.13;2);Brp:(Sketch.11;29)));None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        rib1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = True
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()


def die_extract3():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("die-2.CATPart")
    part1 = partDocument1.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Fillet2")
    shapes1 = body1.Shapes
    rib1 = shapes1.Item("Rib.4")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Rib.3;0:(Brp:(Sketch.9;2);Brp:(Sketch.7;48)));None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        rib1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = True
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()

def die_extract4():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("die-2.CATPart")
    part1 = partDocument1.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Fillet2")
    shapes1 = body1.Shapes
    rib1 = shapes1.Item("Rib.4")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Rib.3;2);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        rib1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = True
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()

def die_extract5():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("die-2.CATPart")
    part1 = partDocument1.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Fillet2")
    shapes1 = body1.Shapes
    rib1 = shapes1.Item("Rib.4")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Rib.3;1);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        rib1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = True
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()

def die_extract6():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("die-2.CATPart")
    part1 = partDocument1.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Fillet2")
    shapes1 = body1.Shapes
    rib1 = shapes1.Item("Rib.4")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Rib.4;2);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        rib1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = True
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()

def die_extract7():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("die-2.CATPart")
    part1 = partDocument1.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Fillet2")
    shapes1 = body1.Shapes
    rib1 = shapes1.Item("Rib.4")
    reference1 = part1.CreateReferenceFromBRepName(
        "RSur:(Face:(Brp:(Rib.4;1);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)",
        rib1)
    hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)
    hybridShapeExtract1.PropagationType = 2
    hybridShapeExtract1.ComplementaryExtract = False
    hybridShapeExtract1.IsFederated = True
    body1.InsertHybridShape(hybridShapeExtract1)
    part1.InWorkObject = hybridShapeExtract1
    part1.Update()

def joindie():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("die-2.CATPart")
    part1 = partDocument1.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Fillet2")
    hybridShapes1 = body1.HybridShapes
    hybridShapeExtract1 = hybridShapes1.Item("Extract.2")
    reference1 = part1.CreateReferenceFromObject(hybridShapeExtract1)
    hybridShapeExtract2 = hybridShapes1.Item("Extract.3")
    reference2 = part1.CreateReferenceFromObject(hybridShapeExtract2)
    hybridShapeAssemble1 = hybridShapeFactory1.AddNewJoin(reference1, reference2)
    hybridShapeExtract3 = hybridShapes1.Item("Extract.4")
    reference3 = part1.CreateReferenceFromObject(hybridShapeExtract3)
    hybridShapeAssemble1.AddElement(reference3)
    hybridShapeExtract4 = hybridShapes1.Item("Extract.5")
    reference4 = part1.CreateReferenceFromObject(hybridShapeExtract4)
    hybridShapeAssemble1.AddElement(reference4)
    hybridShapeExtract5 = hybridShapes1.Item("Extract.6")
    reference5 = part1.CreateReferenceFromObject(hybridShapeExtract5)
    hybridShapeAssemble1.AddElement(reference5)
    hybridShapeExtract6 = hybridShapes1.Item("Extract.7")
    reference6 = part1.CreateReferenceFromObject(hybridShapeExtract6)
    hybridShapeAssemble1.AddElement(reference6)
    hybridShapeExtract7 = hybridShapes1.Item("Extract.8")
    reference7 = part1.CreateReferenceFromObject(hybridShapeExtract7)
    hybridShapeAssemble1.AddElement(reference7)
    hybridShapeAssemble1.SetConnex(1)
    hybridShapeAssemble1.SetManifold(1)
    hybridShapeAssemble1.SetSimplify(0)
    hybridShapeAssemble1.SetSuppressMode(0)
    hybridShapeAssemble1.SetDeviation(0.001)
    hybridShapeAssemble1.SetAngularToleranceMode(0)
    hybridShapeAssemble1.SetAngularTolerance(0.5)
    hybridShapeAssemble1.SetFederationPropagation(0)
    body1.InsertHybridShape(hybridShapeAssemble1)
    part1.InWorkObject = hybridShapeAssemble1
    part1.Update()

def hidebodydie(name):
    catapp = win32.Dispatch('CATIA.Application')
    productDocument1 = catapp.ActiveDocument
    selection1 = productDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("die-2.CATPart")
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Fillet2")
    hybridShapes1 = body1.HybridShapes
    hybridShapeExtract1 = hybridShapes1.Item(str(name))
    hybridShapes1 = hybridShapeExtract1.Parent
    bSTR1 = hybridShapeExtract1.Name
    selection1.Add(hybridShapeExtract1)
    visPropertySet1 = visPropertySet1.Parent
    bSTR2 = visPropertySet1.Name
    bSTR3 = visPropertySet1.Name
    visPropertySet1.SetShow(1)
    selection1.Clear()


def hidebodydie2():
    catapp = win32.Dispatch('CATIA.Application')
    productDocument1 = catapp.ActiveDocument
    selection1 = productDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    documents1 = catapp.Documents
    partDocument1 = documents1.Item("die-2.CATPart")
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Fillet2")
    shapes1 = body1.Shapes
    solid1 = shapes1.Item("Fillet2")
    shapes1 = solid1.Parent
    bSTR1 = solid1.Name
    selection1.Add(solid1)
    visPropertySet1 = visPropertySet1.Parent
    bSTR2 = visPropertySet1.Name
    bSTR3 = visPropertySet1.Name
    visPropertySet1.SetShow(1)
    selection1.Clear()


def save_file_product(path):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    file_name = 'Product1.CATProudct'
    file_name_Product = 'Product1.CATProduct'
    partDocument1 = document.Item(file_name_Product)
    # print(path + '\\' + file_name)
    partDocument1.SaveAs(path + '\\' + file_name)
    # partDocument1.Close()

def saveas(save_dir, target, data_type, data_type2):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    # saveas = document.Item('Product1.CATProduct')
    try:
        saveas = document.Item('%s%s' % (target, data_type))
        saveas.SaveAs('%s\%s%s' % (save_dir, target, data_type2))
    except:
        saveas = catapp.ActiveDocument
        saveas.SaveAs('%s\%s%s' % (save_dir, target, data_type2))
    finally:
        saveas.Save()

def save_file_igs(path):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    file_name = 'Product1.igs'
    file_name_Product = 'Product1.CATProduct'
    partDocument1 = document.Item(file_name_Product)
    # print(path + '\\' + file_name)
    partDocument1.SaveAs(path + '\\' + file_name)

def save_file_part(path, file_name):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    file_name = file_name + '.CATPart'
    partDocument1 = document.Item(file_name)
    # print(path + '\\' + file_name)
    partDocument1.SaveAs(path + '\\' + file_name)
    partDocument1.Close()



def import_product(path, file_name):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(path + "\\" + file_name + ".CATProduct")

def import_part(path, file_name):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(path + "\\" + file_name + ".CATPart")

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
    product.Update()
    productDocument.Save()
    return True

def base_lock(element1, element2, name):  # 定海神針, 固定基準零件
    catapp = win32.Dispatch('CATIA.Application')
    productDocument = catapp.ActiveDocument
    product = productDocument.Product
    constraints = product.Connections("CATIAConstraints")
    ref = product.CreateReferenceFromName("Product1/%s/!Product1/%s/" % (element1, element2))
    constraint = constraints.AddMonoEltCst(0, ref)
    constraint.Name = name
    product.Update()
    productDocument.Save()
    return True

def add_offset():
    base_lock('PUNCH.1', 'PUNCH.1', 0)
    add_offset_assembly('blank.1', 'holder.1', -2, 'zx plane', 0, 'offset1')
    add_offset_assembly('die-2.1', 'blank.1', -0.9, 'zx plane', 0, 'offset2')
    add_offset_assembly('holder.1', 'PUNCH.1', -100, 'zx plane', 0, 'offset3')
    add_offset_assembly('blank.1', 'PUNCH.1', -35.179, 'yz plane', 1, 'offset4')
    add_offset_assembly('blank.1', 'PUNCH.1', 28.215, 'xy plane', 1, 'offset5')

def assembly_create():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    productDocument1 = documents1.Add("Product")

def import_file_Part(path, file_name):  # (資料夾路徑，檔案名稱)
    catapp = win32.Dispatch('CATIA.Application')
    productDocument1 = catapp.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    products1Variant = products1
    combination_file = []  # 組合檔
    combination_file.append(path + '\\' + file_name + '.CATPart')
    products1Variant.AddComponentsFromFiles(combination_file, "All")

def Update():
    catapp = win32.Dispatch('CATIA.Application')
    doc = catapp.ActiveDocument
    part = doc.Part
    part.Update()

def create_new(H1, R1, R2, new_path):
    system_root = os.path.dirname(os.path.realpath(__file__))
    for name in ['holder', 'die-2', 'blank', 'punch']:
        import_part(system_root + '\\seyi_stamping_die', name)
        if name == 'holder':
            param_change(name, 'R2', R2)
            holder_extract1()
            holder_extract2()
            holder_extract3()
            holder_extract4()
            holder_extract5()
            holder_extract6()
            holder_extract7()
            holder_joinextract()
            for i in range(3, 9):
                hidebody("Extract." + str(i))
            hidebody2()
        elif name == 'die-2':
            # die曲面投影
            param_change('die-2', 'R1', R1)
            param_change('die-2', 'H1', H1)
            die_extract1()
            die_extract2()
            die_extract3()
            die_extract4()
            die_extract5()
            die_extract6()
            die_extract7()
            joindie()
            for i in range(3, 9):
                hidebodydie("Extract." + str(i))
            hidebodydie2()
        save_file_part(new_path, name)

    assembly_create()
    saveas(new_path, 'Product1', '.CATProduct', '.CATProduct')

    for name in ['holder', 'die-2', 'blank', 'punch']:
        import_file_Part(new_path, name)

    add_offset()

def close_window():
    catapp = win32.Dispatch('CATIA.Application')
    specsAndGeomWindow1 = catapp.ActiveWindow
    specsAndGeomWindow1.Close()

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