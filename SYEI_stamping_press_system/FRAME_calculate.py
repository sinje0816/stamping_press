import DEMO
import excel_parameter_change as epc
import math

# 機架客製化尺寸確認
def frame_calculate(type, alpha, beta, zeta, delta):
    excel = epc.ExcelOp('標準資料')
    type_name, travel_value, close_working_height_value, specifications_travel_min_value, specifications_travel_max_value, specifications_close_working_height_min_value, specifications_close_working_height_max_value = excel.get_standard_parts(1)
    # 行程確認
    if alpha <= specifications_travel_min_value or alpha > specifications_travel_max_value:
        print("行程超出範圍")
        return "alpha error"
    else:
        print("行程正常")
    # 關閉工作高度確認
    if beta <= specifications_close_working_height_min_value or beta > specifications_close_working_height_max_value:
        print("閉合工作高度超出範圍")
        return "beta error"
    else:
        print("閉合工作高度正常")

    # 機架喉部拉高量計算

    def frame_calculate(self, i, specifications_travel_value, specifications_close_working_height_value, zeta, delta):
        # Form = QtWidgets.QWidget()
        # Form.setWindowTitle('warning.studio')
        # Form.resize(400, 300)
        # mbox = QtWidgets.QMessageBox(Form)

        excel = epc.ExcelOp('標準資料')
        type_name, travel_value, close_working_height_value, specifications_travel_min_value, specifications_travel_max_value, specifications_close_working_height_min_value, specifications_close_working_height_max_value = excel.get_standard_parts(i)

        # 驗證行程
        stv = specifications_travel_value
        try:
            if stv < specifications_travel_min_value or stv > specifications_travel_max_value:
                mbox.warning(Form, '警告', '行程超出範圍')