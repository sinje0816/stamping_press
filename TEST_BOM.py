import win32com.client as win32

BOM_list = {'1': ['01 ', '50', '1680x2910', 'SS400', '1', '1628.6', '如圖'],
            '2': ['02 ', '50', '1680x2910', 'SS400', '1', '1628.6', '如圖'],
            '3': ['03 ', '40', '850x1027', 'SS400', '1', '1626.6', '如圖'],
            '4': ['04 ', '32', '850x1050', 'SS400', '1', '207.0', '如圖'],
            '5': ['05 ', '22', '850x590', 'SS400', '1', '84.2', '如圖'],
            '6': ['06 ', '40', '80x182', 'SS400', '2', '3.9', '如圖'],
            '7': ['07 ', '50', '850x710', 'SS400', '1', '127.0', '如圖'],
            '8': ['08 ', '22', '670x476', 'SS400', '1', '53.8', '如圖'],
            '9': ['8-1', '22', '670x150', 'SS400', '2', '17.4', '如圖'],
            '10': [' 9 ', ' 9 ', '協中:250x90x850(待加)L', 'SS400', '1', '28.5', '槽型鋼'],
            '11': [' 9 ', ' 9 ', '協台:250x90x850(待加)L', 'SS400', '1', '29.4', '槽型鋼'],
            '12': ['10 ', '45', '255x300', 'SS400', '2', '24.5', '如圖'],
            '13': ['11 ', '125', 'φ270xφ170', 'SS400', '1', '33.9', '如圖'],
            '14': ['12 ', '19', '670x429', 'SS400', '1', '42.0', '如圖'],
            '15': ['13 ', '19', '48x74', 'SS400', '2', '0.4', '如圖'],
            '16': ['14 ', '40', '105x150', 'SS400', '2', '4.9', '如圖'],
            '17': ['15A ', '50', '40x150', 'SS400', '2', '2.1', '如圖'],
            '18': ['15 ', '50', '40x150', 'SS400', '2', '2.3', '如圖'],
            '19': ['16 ', '40', '105x150', 'SS400', '2', '4.8', '如圖'],
            '20': ['18 ', '19', '145x290', 'SS400', '1', '6.2', '如圖'],
            '21': ['19 ', '90', '960x1588', 'SS400', '2', '609.4', '如圖'],
            '22': ['20 ', '19', '145x300', 'SS400', '1', '5.8', '如圖'],
            '23': ['21 ', '53', '75x75', 'SS400', '1', '2.4', '如圖'],
            '24': ['22 ', '50', '850x710', 'SS400', '1', '156.4', '如圖'],
            '25': ['24 ', '60', '85x420', 'SS400', '2', '15.9', '如圖'],
            '26': ['25 ', '75', '75x75', 'SS400', '5', '2.4', '如圖'],
            '27': ['26 ', '35', '50x50', 'SS400', '2', '0.7', ''],
            '28': ['27 ', '25', '460x280', 'SS400', '2', '25.2', '如圖'],
            '29': ['29 ', '19', '140x80', 'SS400', '4', '1.2', '如圖'],
            '30': ['30 ', '45', '243x300', 'SS400', '2', '0.9', ''],
            '31': ['33 ', '6', 'L65x30x850', 'SS400', '1', '3.6', '如圖'],
            '32': ['34 ', '60', 'φ50', 'SS400', '2', '0.9', ''],
            '33': ['35 ', '50', '850x680', 'SS400', '1', '197.2', '如圖'],
            '34': ['40 ', '22', '850x165', 'SS400', '1', '24.2', ''],
            '35': ['41 ', '6','角鐵:50x50x180L', 'SS400', '1', '197.2', '如圖'],
            '36': ['43 ', '32', '32x85', 'SS400', '1', '0.7', '如圖'],
            '37': ['44 ', '62', 'φ30', 'SS400', '1', '0.3', '如圖']}

print(BOM_list['10'][1])

BOM_list_name = []
for i in BOM_list:
    BOM_list_name.append(i)

BOM_position = [918, 114]
BOM_position_x = [865 + 99, 865 + 116, 865 + 130, 865 + 212.5, 865 + 242, 865 + 255, 865 + 280]
BOM_position_Y = 12

def bom_text_create():
    catapp = win32.Dispatch('CATIA.Application')
    drawingdocument = catapp.ActiveDocument
    drawingsheets = drawingdocument.Sheets
    drawingsheet = drawingsheets.Item('Sheet.1')
    drawingview = drawingsheet.Views.Item('Background View')
    drawingview.Activate()
    drawingview = drawingview.Texts
    # drawingview.Activate()
    for i in BOM_list_name:
        for j in range(0, 7):
            DrawText = drawingview.Add(BOM_list[i][j], BOM_position_x[j], 120 + 12 * int(i)) # 線段長度
            DrawText.SetFontName(0, 0, 'Arial Unicode MS (TrueType)')
            DrawText.SetFontSize(0, 0, 5)  # 調整字體位置和大小(x, y, 字體大小).
            DrawText.AnchorPosition = 2  # 文字框對其位置(TopL=1,MidL=2,BottomL=3,TopM=4,MidM=5,BottomM=6,TopR=7,MidR=8,BottomR=9)

# bom_text_create()

def create_center_line(x_value_1, y_value_1, x_value_2, y_value_2):
    catapp = win32.Dispatch('CATIA.Application')
    drawingdocument = catapp.ActiveDocument
    drawingsheets = drawingdocument.Sheets
    drawingsheet = drawingsheets.Item('Sheet.1')
    drawingviews = drawingsheet.Views
    drawingview = drawingviews.Item('Background View')
    drawingview.Activate()
    factory2d = drawingview.Factory2D
    dotted_line = factory2d.CreateLine(x_value_1, y_value_1, x_value_2, y_value_2)
    dotted_line.Name = 'chain_line'
    object = catapp.ActiveDocument.Selection
    # object.Search('Name=chain_line,all')
    vis = catapp.ActiveDocument.Selection.VisProperties
    vis.SetRealLineType(2, 0.3)
    vis.SetRealWidth(1, 0.13)

for i in BOM_list_name:
    create_center_line(960, 102 + 12 * int(i), 1169, 102 + 12 * int(i))
