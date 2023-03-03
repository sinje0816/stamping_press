import win32com.client as win32

def import_part(path, file_name):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(path + "\\" + file_name + ".stp")

import_part("C:\\Users\\User\\Desktop\\佳馨學姊分析\\2_0_in\\SP\\", "SP_2_0_in")

