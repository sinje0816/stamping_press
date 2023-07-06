import zipfile
import main_program as mprog
import parameter as par
import drafting as draft
from PyQt5 import QtCore, QtGui, QtWidgets
import win32com.client as win32
# i = 5
alpha = 20
# mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -(par.H[i] -40 - 32) / 2 - 40, 'XY.PLANE', 1, 17)
# mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', (par.H[i] - 32 - 40) / 2 + 40, 'XY.PLANE', 0, 14)
# mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', -par.H[i] + 32 - alpha, 'XY.PLANE', 0, 41)
# mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', -par.FRAME20_H[i] + alpha, 'XY.PLANE', 0, 47)
# mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -9, 'XY.PLANE', 0, 44)
# mprog.add_offset_assembly('FRAME20.1', 'FRAME44.1', 9, 'XY.PLANE', 0, 214)
# mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1',
#                                               -(par.H[i] - par.Z[i] - par.S[i] - 34), 'XY.PLANE',
#                                               0, 204)

# f = zipfile.ZipFile('C:\\Users\\USER\\Desktop\\text.zip','w',zipfile.ZIP_DEFLATED)
# dirpath = "C:\\Users\\USER\\Desktop\\CATIA"
# for path, dirnames, filenames in os.walk(dirpath):
#     # 去掉目標根路徑，只對目標資料夾裡面的檔案進行壓縮
#     fpath= path.replace(dirpath,'')
#     for filename in filenames:
#         f.write(os.path.join(path, filename), os.path.join(fpath, filename))


catapp = win32.Dispatch('CATIA.Application')
partdocuments = catapp.Documents
part = partdocuments.part
selection = partdocuments.selection

