import win32com.client as win32
import parameter as par
import drafting as draft

i = 5

drafting_down_Coordinate_Position, drafting_up_Coordinate_Position, scale = draft.drafting_welding_view_parameter_calculation(par.A[i], par.B[i], par.H[i])
print(drafting_down_Coordinate_Position, drafting_up_Coordinate_Position, scale)
draft.change_Drawing_scale(1 / scale)

def Section():
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.ActiveSheet
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.ActiveView
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    drawingView2 = drawingViews1.Add("AutomaticNaming")
    drawingViewGenerativeBehavior2 = drawingView2.GenerativeBehavior
    drawingView2.X = 759.942657
    drawingView2.Y = 295
    drawingView2.Scale = 0.125
    drawingView2.Angle = 0
    arrayOfVariantOfDouble1 = [0, 543, 0, -2622]
    drawingViewGenerativeBehavior2Variant = drawingViewGenerativeBehavior2
    drawingViewGenerativeBehavior2Variant.DefineSectionView(arrayOfVariantOfDouble1, "SectionView", "Offset", 0, drawingViewGenerativeBehavior1)
    drawingViewGenerativeLinks1 = drawingView2.GenerativeLinks
    drawingViewGenerativeBehavior2 = drawingView2.GenerativeBehavior
    drawingViewGenerativeBehavior2.Update()
    drawingView2.ReferenceView = drawingView1
    drawingView2.AlignedWithReferenceView()

# 剖面圖座標
Section_Coordinate = {'A-A': [[0, -2900, 0, 600], 1],
                      'B-B': [[par.B[i] + 100, 520, par.B[i] + 100, -1200], 0],
                      'C-C': [[-195, -1460, 2286, par.B[i] + 100], 0],
                      'D-D': [[-par.A[i] / 2 - 100, -par.H[i] / 2, par.A[i] / 2 + 100, -par.H[i] / 2], 0],
                      'E-E': [[-100, -(par.S[i] + par.Z[i]) + 100, par.B[i] / 3, -(par.S[i] + par.Z[i]) + 100], 1],
                      'F-F': [[-272, - 176 / 2, -272, -176 * 1.5], 0],
                      'G-G': [[par.B[i] + 100, -(par.S[i] + par.Z[i]) + 150, par.B[i] + 100, -(par.S[i] + par.Z[i]) - 150], 1]}

draft.Section('Front view', drafting_down_Coordinate_Position['section A-A'][0],  drafting_down_Coordinate_Position['section A-A'][1], scale, Section_Coordinate['A-A'][0], 0)
# Section('Front view', drafting_down_Coordinate_Position['section A-A'][0],  drafting_down_Coordinate_Position['section A-A'][1], scale, Section_Coordinate['A-A'][0], 0)