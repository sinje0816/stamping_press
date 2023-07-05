import openpyxl as xl


def excel_read():
    wb = xl.load_workbook('尺寸整理表.xlsx')
    parameter = []
    value = []
    worksheet = workbook['Sheet1']
    for a in range(sheet.nrows):
        cells = sheet.row_values(a)
        parameter_name = cells[1]
        parameter_value = cells[2]
        parameter.append(parameter_name)
        value.append(parameter_value)
    return parameter, value

a = excel_read()
print(a[0][0])