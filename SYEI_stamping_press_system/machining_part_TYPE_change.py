import main_program as mprog
import excel_parameter_change as epc
import time

# 母檔變數變換
def change_parameter(name, i):
    # 已改
    if name == 'FRAME1':
        excel = epc.ExcelOp('FRAME1')
        try:
            FRAME1_parameter_name, FRAME1_parameter_value = excel.part_parameter('FRAME1', i)
            print('FRAME1 Parameter change success')
            mprog.param_change('FRAME1', 'alpha', 10)
            mprog.param_change('FRAME1', 'beta', 10)
            mprog.param_change('FRAME1', 'gamma', 10)
            mprog.param_change('FRAME1', 'delta', 10)
            mprog.param_change('FRAME1', 'zeta', 10)


        except:
            print('FRAME1 Parameter change error')
        if i == 0:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('machining_DD')
            mprog.partbodyfeatureactivate('machining_E上方研磨')
            mprog.activatefeature('110通孔', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('6-M5X16L(配管)', 5)
            mprog.activatefeature('5-M5X10L(配管用)', 2)
            mprog.activatefeature('兩點組合', 0)
            mprog.activatefeature('machining_2x8_12通孔', 0)
            mprog.activatefeature('machining_aaaaa2', 0)
            mprog.activatefeature('machining_aaaaa4', 0)
        elif i == 1:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('machining_DD')
            mprog.partbodyfeatureactivate('machining_E上方研磨')
            mprog.activatefeature('110通孔', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('6-M5X16L(配管)', 5)
            mprog.activatefeature('5-M5X10L(配管用)', 3)
            mprog.activatefeature('兩點組合', 0)
            mprog.activatefeature('machining_2x8_12通孔', 0)
            mprog.activatefeature('machining_aaaaa2', 0)
            mprog.activatefeature('machining_aaaaa4', 0)
        elif i == 2:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('machining_DD')
            mprog.partbodyfeatureactivate('machining_E上方研磨')
            mprog.activatefeature('machining_2x8_12通孔', 0)
            mprog.activatefeature('machining_aaaaa2', 0)
            mprog.activatefeature('machining_aaaaa4', 0)
            mprog.activatefeature('110通孔', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('6-M5X16L(配管)', 5)
            mprog.activatefeature('5-M5X10L(配管用)', 2)
            mprog.activatefeature('兩點組合', 0)
        elif i == 3:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('machining_DD')
            mprog.partbodyfeatureactivate('machining_E上方研磨')
            mprog.activatefeature('machining_2x8_12通孔', 0)
            mprog.activatefeature('machining_aaaaa2', 0)
            mprog.activatefeature('machining_aaaaa4', 0)
            mprog.activatefeature('110通孔', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('6-M5X16L(配管)', 5)
            mprog.activatefeature('5-M5X10L(配管用)', 2)
            mprog.activatefeature('兩點組合', 0)
        elif i == 4:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('machining_DD')
            mprog.partbodyfeatureactivate('machining_E上方研磨')
            mprog.activatefeature('machining_2x8_12通孔', 0)
            mprog.activatefeature('machining_aaaaa2', 0)
            mprog.activatefeature('machining_aaaaa4', 0)
            mprog.activatefeature('1-M5X10L(配管用)', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('6-M5X16L(配管)', 5)
            mprog.activatefeature('5-M5X10L(配管用)', 3)
            mprog.activatefeature('兩點組合', 0)
            mprog.activatefeature('3-M5X10L(配管用)(r1)', 1)

        elif i == 5:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('machining_DD')
            mprog.partbodyfeatureactivate('machining_E上方研磨')
            mprog.activatefeature('machining_2x8_12通孔', 0)
            mprog.activatefeature('machining_aaaaa2', 0)
            mprog.activatefeature('machining_aaaaa4', 0)
            mprog.activatefeature('1-M5X10L(配管用)', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('6-M5X16L(配管)', 0)
            mprog.activatefeature('5-M5X10L(配管用)', 3)
            mprog.activatefeature('兩點組合', 0)
            mprog.activatefeature('3-M5X10L(配管用)(r1)', 2)

        elif i == 6:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('machining_DD')
            mprog.partbodyfeatureactivate('machining_E上方研磨')
            mprog.activatefeature('machining_2x8_12通孔', 0)
            mprog.activatefeature('machining_aaaaa2', 0)
            mprog.activatefeature('machining_aaaaa4', 0)
            mprog.activatefeature('1-M5X10L(配管用)', 0)
            mprog.activatefeature('3-M5X10L(配管用)(r1)', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('6-M5X16L(配管)', 0)
            mprog.activatefeature('5-M5X10L(配管用)', 5)
            mprog.activatefeature('兩點組合', 0)
        elif i == 7:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('machining_DD')
            mprog.partbodyfeatureactivate('machining_E上方研磨')
            mprog.activatefeature('machining_2x8_12通孔', 0)
            mprog.activatefeature('machining_aaaaa2', 0)
            mprog.activatefeature('machining_aaaaa4', 0)
            mprog.activatefeature('1-M5X10L(配管用)', 0)
            mprog.activatefeature('3-M5X10L(配管用)(r1)', 2)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('6-M5X16L(配管)', 0)
            mprog.activatefeature('5-M5X10L(配管用)', 5)
            mprog.activatefeature('兩點組合', 0)
        elif i == 8:
            mprog.partbodyfeatureactivate('T250')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('machining_DD_250')
            mprog.partbodyfeatureactivate('machining_E上方研磨')
            mprog.activatefeature('machining_2x8_12通孔', 0)
            mprog.activatefeature('machining_aaaaa2', 0)
            mprog.activatefeature('machining_aaaaa4', 0)
            mprog.activatefeature('1-M5X10L(配管用)', 0)
            mprog.activatefeature('3-M5X10L(配管用)(r1)', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('6-M5X16L(配管)', 0)
            mprog.activatefeature('5-M5X10L(配管用)', 0)
            mprog.activatefeature('兩點組合', 0)
        try:
            mprog.update()
            print('FRAME1 update success')
        except:
            print('FRAME1 update error')
    elif name == 'FRAME2':
        start = time.time()
        excel = epc.ExcelOp('FRAME2')
        try:
            FRAME2_parameter_name, FRAME2_parameter_value = excel.part_parameter('FRAME2', i)
            mprog.param_change('FRAME2', 'alpha', 10)
            mprog.param_change('FRAME2', 'beta', 10)
            mprog.param_change('FRAME2', 'gamma', 10)
            mprog.param_change('FRAME2', 'delta', 10)
            mprog.param_change('FRAME2', 'zeta', 10)
            print('FRAME2 Parameter change success')
        except:
            print('FRAME2 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('machining_DD')
                mprog.partbodyfeatureactivate('machining_E上方研磨')
                mprog.activatefeature('machining_2x8_12通孔', 0)
                mprog.activatefeature('machining_aaaaa2', 0)
                mprog.activatefeature('machining_aaaaa4', 0)
                mprog.partbodyfeatureactivate('Z1')
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 4)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 1)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('解角器裝置線用', 0)
                mprog.activatefeature('解角器用', 0)
                mprog.activatefeature('變頻器用', 0)
                mprog.activatefeature('解角器護蓋用', 0)
                mprog.activatefeature('M5X10L(配管用)', 0)
            elif i == 1:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('machining_DD')
                mprog.partbodyfeatureactivate('machining_E上方研磨')
                mprog.activatefeature('machining_2x8_12通孔', 0)
                mprog.activatefeature('machining_aaaaa2', 0)
                mprog.activatefeature('machining_aaaaa4', 0)
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 4)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 1)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('解角器裝置線用', 0)
                mprog.activatefeature('解角器用', 0)
                mprog.activatefeature('變頻器用', 0)
                mprog.activatefeature('解角器護蓋用', 0)
                mprog.activatefeature('M5X10L(配管用)', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('1-M5X10L底孔20L配管用', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('machining_DD')
                mprog.partbodyfeatureactivate('machining_E上方研磨')
                mprog.activatefeature('machining_2x8_12通孔', 0)
                mprog.activatefeature('machining_aaaaa2', 0)
                mprog.activatefeature('machining_aaaaa4', 0)
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 4)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 2)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('解角器裝置線用', 0)
                mprog.activatefeature('解角器用', 0)
                mprog.activatefeature('變頻器用', 0)
                mprog.activatefeature('解角器護蓋用', 0)
                mprog.activatefeature('M5X10L(配管用)', 0)
                mprog.activatefeature('1-M5X10L底孔20L配管用', 0)
            elif i == 3:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('machining_DD')
                mprog.partbodyfeatureactivate('machining_E上方研磨')
                mprog.activatefeature('machining_2x8_12通孔', 0)
                mprog.activatefeature('machining_aaaaa2', 0)
                mprog.activatefeature('machining_aaaaa4', 0)
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 4)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 2)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('解角器裝置線用', 0)
                mprog.activatefeature('解角器用', 0)
                mprog.activatefeature('變頻器用', 0)
                mprog.activatefeature('解角器護蓋用', 0)
                mprog.activatefeature('M5X10L(配管用)', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
            elif i == 4:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('machining_DD')
                mprog.partbodyfeatureactivate('machining_E上方研磨')
                mprog.partbodyfeatureactivate('jjj4')
                mprog.partbodyfeatureactivate('mmm4')
                mprog.activatefeature('machining_2x8_12通孔', 0)
                mprog.activatefeature('machining_aaaaa2', 0)
                mprog.activatefeature('machining_aaaaa4', 0)
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 4)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 2)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 2)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('解角器裝置線用', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('4-M8X16L', 0)
                mprog.activatefeature('2-M5(配管用)', 0)
                mprog.activatefeature('台達變頻器用10HP', 0)
                mprog.activatefeature('東元變頻器用10HP', 0)
                mprog.activatefeature('電氣箱連桿', 2)
            elif i == 5:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('machining_DD')
                mprog.partbodyfeatureactivate('machining_E上方研磨')
                mprog.partbodyfeatureactivate('dd4')
                mprog.partbodyfeatureactivate('x4')
                mprog.activatefeature('machining_2x8_12通孔', 0)
                mprog.activatefeature('machining_aaaaa2', 0)
                mprog.activatefeature('machining_aaaaa4', 0)
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 5)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 2)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('4-M8X16L', 0)
                mprog.activatefeature('2-M5(配管用)', 0)
                mprog.activatefeature('電氣箱連桿', 0)
                mprog.activatefeature('東元變頻器10,15HP440V用', 0)
                mprog.activatefeature('台達變頻器10,15HP', 0)
                mprog.activatefeature('東元變頻器,15HP220V', 0)
            elif i == 6:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('machining_DD')
                mprog.partbodyfeatureactivate('machining_E上方研磨')
                mprog.partbodyfeatureactivate('ggg4')
                mprog.partbodyfeatureactivate('xx3')
                mprog.partbodyfeatureactivate('iii4')
                mprog.partbodyfeatureactivate('fff4')
                mprog.activatefeature('machining_2x8_12通孔', 0)
                mprog.activatefeature('machining_aaaaa2', 0)
                mprog.activatefeature('machining_aaaaa4', 0)
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 5)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 2)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('4-M8X16L', 0)
                mprog.activatefeature('2-M5(配管用)', 0)
                mprog.activatefeature('電氣箱連桿', 0)
                mprog.activatefeature('東元變頻器440V,15,20HP', 0)
                mprog.activatefeature('台達變頻器440V,15,20HP；220V15HP', 0)
                mprog.activatefeature('台達變頻器220V,20HP', 0)
                mprog.activatefeature('東元變頻器220V,15,20HP', 0)
            elif i == 7:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('machining_DD')
                mprog.partbodyfeatureactivate('machining_E上方研磨')
                mprog.partbodyfeatureactivate('iii4')
                mprog.partbodyfeatureactivate('ssss4')
                mprog.partbodyfeatureactivate('sss4')
                mprog.activatefeature('machining_2x8_12通孔', 0)
                mprog.activatefeature('machining_aaaaa2', 0)
                mprog.activatefeature('machining_aaaaa4', 0)
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 5)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('4-M8X16L', 0)
                mprog.activatefeature('2-M5(配管用)', 0)
                mprog.activatefeature('電氣箱連桿', 0)
                mprog.activatefeature('台達變頻器220V,20HP', 0)
                mprog.activatefeature('東元變頻器220V/440V,20HP', 0)
                mprog.activatefeature('台達變頻器440V,20HP', 0)
            elif i == 8:
                mprog.partbodyfeatureactivate('T250')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('machining_DD_250')
                mprog.partbodyfeatureactivate('machining_E上方研磨')
                mprog.activatefeature('machining_2x8_12通孔', 0)
                mprog.activatefeature('machining_aaaaa2', 0)
                mprog.activatefeature('machining_aaaaa4', 0)
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 5)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 2)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('4-M8X16L', 0)
                mprog.activatefeature('2-M5(配管用)', 0)
                mprog.activatefeature('電氣箱連桿', 0)
                mprog.activatefeature('東元變頻器25HP', 0)
                mprog.activatefeature('台達變頻器25HP', 0)
        except:
            print('FRAME2 Part activate error')
        finally:
            mprog.update()
            print('FRAME2 update success')
    elif name == 'FRAME5':
        excel = epc.ExcelOp('FRAME5')
        try:
            FRAME5_parameter_name, FRAME5_parameter_value = excel.part_parameter('FRAME5', i)
            print('FRAME5 Parameter change success')
        except:
            print('FRAME5 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('SN1-25_X')
                mprog.partbodyfeatureactivate('before80_I')
                mprog.partbodyfeatureactivate('before80_R')
                mprog.partbodyfeatureactivate('Y')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
                mprog.activatefeature('PT1/2', 0)
            elif i == 1:
                mprog.partbodyfeatureactivate('before80_I')
                mprog.partbodyfeatureactivate('before80_R')
                mprog.partbodyfeatureactivate('Y')
                mprog.partbodyfeatureactivate('AA')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
                mprog.activatefeature('PT1/2', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('before80_I')
                mprog.partbodyfeatureactivate('before80_R')
                mprog.partbodyfeatureactivate('Y')
                mprog.partbodyfeatureactivate('AA')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
                mprog.activatefeature('PT1/2', 0)
            elif i == 3:
                mprog.partbodyfeatureactivate('before80_I')
                mprog.partbodyfeatureactivate('before80_R')
                mprog.partbodyfeatureactivate('S')
                mprog.partbodyfeatureactivate('Rx2')
                mprog.partbodyfeatureactivate('SN1-60_R')
                mprog.partbodyfeatureactivate('Chamfer.2')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 4:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 5:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.partbodyfeatureactivate('6-M6通C5')
                mprog.activatefeature('6-M6通', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 6:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.partbodyfeatureactivate('3-M8通_C5')
                mprog.activatefeature('3-M8通', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 7:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.partbodyfeatureactivate('5-M8通_C5')
                mprog.activatefeature('5-M8通', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 8:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.partbodyfeatureactivate('5-M8通_C5')
                mprog.activatefeature('5-M8通', 0)
                mprog.activatefeature('M6通孔LED', 0)
        except:
            print('FRAME5 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME5 Update success')
            except:
                print('FRAME5 Update error')
    elif name == 'FRAME8':
        excel = epc.ExcelOp('FRAME8')
        # try:
        FRAME8_parameter_name, FRAME8_parameter_value = excel.part_parameter('FRAME8', i)
        #     print('FRAME8 Parameter change success')
        # except:
        #     print('FRAME8 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('上半倒角')
            elif i == 1:
                mprog.partbodyfeatureactivate('上半倒角')
                # mprog.activatefeature('38孔通', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('上半倒角')
            elif i == 3:
                mprog.partbodyfeatureactivate('上半倒角')
            elif i == 4:
                mprog.partbodyfeatureactivate('上半倒角80')
                mprog.partbodyfeatureactivate('R40特徵')
                mprog.partbodyfeatureactivate('下半開槽')
                mprog.activatefeature('38孔通', 0)
            elif i == 5:
                mprog.partbodyfeatureactivate('上半倒角')
                mprog.partbodyfeatureactivate('下半倒角')
                mprog.partbodyfeatureactivate('下半開槽')
                mprog.partbodyfeatureactivate('R40特徵')
                mprog.activatefeature('38孔通', 0)
            elif i == 6:
                mprog.partbodyfeatureactivate('上半倒角')
                mprog.partbodyfeatureactivate('下半倒角')
                mprog.partbodyfeatureactivate('下半開槽')
                mprog.partbodyfeatureactivate('R40特徵')
                mprog.activatefeature('38孔通', 0)
            elif i == 7:
                mprog.partbodyfeatureactivate('上半倒角')
                mprog.partbodyfeatureactivate('下半開槽')
                mprog.partbodyfeatureactivate('R40特徵')
                mprog.activatefeature('38孔通', 0)
            elif i == 8:
                mprog.partbodyfeatureactivate('上半倒角')
                mprog.partbodyfeatureactivate('R40特徵')
                mprog.activatefeature('38孔通', 0)
        except:
            print('FRAME8 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME8 Update success')
            except:
                print('FRAME8 Update error')
    elif name == 'FRAME11':
        start = time.time()
        excel = epc.ExcelOp('FRAME11')
        try:
            FRAME11_parameter_name, FRAME11_parameter_value = excel.part_parameter('FRAME11', i)
            print('FRAME11 Parameter change success')
        except:
            print('FRAME11 Parameter change error')
        if i == 0:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 1:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 2:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 3:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 4:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 5:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 6:
            mprog.activatefeature('挖孔', 0)
        elif i == 7:
            pass
        elif i == 8:
            pass
        try:
            mprog.Update()
            print('FRAME11 Update success')
        except:
            print('FRAME11 Update error')
    elif name == 'FRAME24':
        excel = epc.ExcelOp('FRAME24')
        try:
            FRAME24_parameter_name, FRAME24_parameter_value = excel.part_parameter('FRAME24', i)
            print('FRAME24 Parameter change success')
        except:
            print('FRAME24 Parameter change error')
        if i == 0:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 1:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 2:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 3:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 4:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 5:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 6:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 7:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 8:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        try:
            mprog.Update()
            print('FRAME24 Update success')
        except:
            print('FRAME24 Update error')
    elif name == 'FRAME29':
        excel = epc.ExcelOp('FRAME29')
        try:
            FRAME29_parameter_name, FRAME29_parameter_value = excel.part_parameter('FRAME29', i)
            print('FRAME29 Parameter change success')
        except:
            print('FRAME29 Parameter change error')
        try:
            mprog.partbodyfeatureactivate('Hole.1')
        finally:
            try:
                mprog.Update()
                print('FRAME29 Update success')
            except:
                print('FRAME29 Update error')
    elif name == 'FRAME42':
        excel = epc.ExcelOp('FRAME42')
        try:
            FRAME42_parameter_name, FRAME42_parameter_value = excel.part_parameter('FRAME42', i)
            print('FRAME42 Parameter change success')
        except:
            print('FRAME42 Parameter change error')
        try:
            mprog.Update()
            print('FRAME42 Update success')
        except:
            print('FRAME42 Update error')
    elif name == 'FRAME43':
        excel = epc.ExcelOp('FRAME43')
        try:
            FRAME43_parameter_name, FRAME43_parameter_value = excel.part_parameter('FRAME43', i)
            print('FRAME43 Parameter change success')
        except:
            print('FRAME43 Parameter change error')
        try:
            if i == 4:
                mprog.partbodyfeatureactivate('machining_Pocket')
            else:
                pass
        except:
            print('FRAME43 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME43 Update success')
            except:
                print('FRAME43 Update error')
    elif name == 'FRAME53':
        excel = epc.ExcelOp('FRAME53')
        try:
            FRAME53_parameter_name, FRAME53_parameter_value = excel.part_parameter('FRAME53', i)
            print('FRAME53 Parameter change success')
        except:
            print('FRAME53 Parameter change error')
        try:
            mprog.Update()
            print('FRAME53 Update success')
        except:
            print('FRAME53 Update error')

    # 未改
    elif name == 'FRAME17':
        excel = epc.ExcelOp('FRAME17')
        try:
            FRAME17_parameter_name, FRAME17_parameter_value = excel.part_parameter('FRAME17', i)
            print('FRAME17 Parameter change success')
        except:
            print('FRAME17 Parameter change error')
    elif name == 'FRAME34':
        excel = epc.ExcelOp('FRAME34')
        try:
            FRAME34_parameter_name, FRAME34_parameter_value = excel.part_parameter('FRAME34', i)
            print('FRAME34 Parameter change success')
        except:
            print('FRAME34 Parameter change error')
        try:
            mprog.Update()
            print('FRAME34 Update success')
        except:
            print('FRAME34 Update error')
    elif name == 'FRAME41':
        excel = epc.ExcelOp('FRAME41')
        try:
            FRAME41_parameter_name, FRAME41_parameter_value = excel.part_parameter('FRAME41', i)
            print('FRAME41 Parameter change success')
        except:
            print('FRAME41 Parameter change error')
        try:
            mprog.Update()
            print('FRAME41 Update success')
        except:
            print('FRAME41 Update error')
    elif name == 'FRAME42_old':
        excel = epc.ExcelOp('FRAME42')
        try:
            FRAME42_parameter_name, FRAME42_parameter_value = excel.part_parameter('FRAME42', i)
            print('FRAME42 Parameter change success')
        except:
            print('FRAME42 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('Pad.1')
            elif i == 1:
                mprog.partbodyfeatureactivate('Pad.1')
            elif i == 2:
                mprog.partbodyfeatureactivate('Pad.1')
            elif i == 3:
                mprog.partbodyfeatureactivate('Pad.1')
            elif i == 4:
                pass
            elif i == 5:
                pass
            elif i == 6:
                mprog.activatefeature('角鐵', 0)
            elif i == 7:
                mprog.activatefeature('角鐵', 0)
            elif i == 8:
                mprog.activatefeature('角鐵', 0)
        except:
            print('FRAME42 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME42 Update success')
            except:
                print('FRAME42 Update error')
    elif name == 'FRAME45':
        excel = epc.ExcelOp('FRAME45')
        try:
            FRAME45_parameter_name, FRAME45_parameter_value = excel.part_parameter('FRAME45', i)
            print('FRAME45 Parameter change success')
        except:
            print('FRAME45 Parameter change error')
        if i == 0:
            mprog.activatefeature('25', 0)
        elif i == 1:
            pass
        elif i == 2:
            mprog.partbodyfeatureactivate('Pad.1')
            mprog.partbodyfeatureactivate('通孔B')
        elif i == 3:
            mprog.partbodyfeatureactivate('Pad.1')
            mprog.partbodyfeatureactivate('通孔B')
        elif i == 4:
            pass
        elif i == 5:
            pass
        elif i == 6:
            mprog.partbodyfeatureactivate('Pad.1')
        elif i == 7:
            mprog.partbodyfeatureactivate('Pad.1')
        elif i == 8:
            mprog.partbodyfeatureactivate('Pad.1')
        try:
            mprog.Update()
            print('FRAME45 Update success')
        except:
            print('FRAME45 Update error')
    elif name == 'FRAME46':
        excel = epc.ExcelOp('FRAME46')
        try:
            FRAME46_parameter_name, FRAME46_parameter_value = excel.part_parameter('FRAME46', i)
            print('FRAME46 Parameter change success')
        except:
            print('FRAME46 Parameter change error')
        if i == 0:
            mprog.partbodyfeatureactivate('Pad.1')
        elif i == 1:
            mprog.partbodyfeatureactivate('Pad.1')
        elif i == 2:
            mprog.partbodyfeatureactivate('Pad.1')
        elif i == 3:
            mprog.partbodyfeatureactivate('Pad.1')
        try:
            mprog.Update()
            print('FRAME46 Update success')
        except:
            print('FRAME46 Update error')
    elif name == 'FRAME47':
        excel = epc.ExcelOp('FRAME47')
        try:
            FRAME47_parameter_name, FRAME47_parameter_value = excel.part_parameter('FRAME47', i)
            print('FRAME47 Parameter change success')
        except:
            print('FRAME47 Parameter change error')
        if i == 0:
            mprog.partbodyfeatureactivate('Pad.1')
        elif i == 1:
            mprog.partbodyfeatureactivate('Pad.1')
        elif i == 2:
            mprog.partbodyfeatureactivate('Pad.1')
        elif i == 3:
            mprog.activatefeature('60', 0)
        try:
            mprog.Update()
            print('FRAME47 Update success')
        except:
            print('FRAME47 Update error')

change_parameter('FRAME29', 1)