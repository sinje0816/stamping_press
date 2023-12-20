import main_program as mprog
import excel_parameter_change as epc
import sys


# 母檔變數變換
def change_welding_feature(name, i):
    try:
        # 零件特徵開啟
        if name == 'FRAME1':
            if i == 0:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('AA2')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('VV2')
                mprog.partbodyfeatureactivate('VV4')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角填料_1')
                mprog.partbodyfeatureactivate('R倒角填料_2')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('小噸數E倒角_B')
                mprog.partbodyfeatureactivate('小噸數Y倒角_B1')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_B')
                mprog.partbodyfeatureactivate('喉口倒角_B_圓角')
                mprog.activatefeature('110通孔', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('光電裝置電線固定孔', 5)
                mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 2)
                mprog.activatefeature('兩點組合', 0)
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
                mprog.partbodyfeatureactivate("welding_喉口")
            elif i == 1:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('VV2')
                mprog.partbodyfeatureactivate('VV4')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角填料_1')
                mprog.partbodyfeatureactivate('R倒角填料_2')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('小噸數E倒角_B')
                mprog.partbodyfeatureactivate('小噸數Y倒角_B1')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_B')
                mprog.partbodyfeatureactivate('喉口倒角_B_圓角')
                mprog.activatefeature('110通孔', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('光電裝置電線固定孔', 5)
                mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 3)
                mprog.activatefeature('兩點組合', 0)
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
                mprog.partbodyfeatureactivate("welding_喉口")
            elif i == 2:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('VV2')
                mprog.partbodyfeatureactivate('VV4')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角填料_1')
                mprog.partbodyfeatureactivate('R倒角填料_2')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('小噸數E倒角_B')
                mprog.partbodyfeatureactivate('小噸數Y倒角_B1')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_B')
                mprog.partbodyfeatureactivate('喉口倒角_B_圓角')
                mprog.activatefeature('110通孔', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('光電裝置電線固定孔', 5)
                mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 2)
                mprog.activatefeature('兩點組合', 0)
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
                mprog.partbodyfeatureactivate("welding_喉口")
            elif i == 3:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角填料_1')
                mprog.partbodyfeatureactivate('R倒角填料_2')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('喉口倒角_60')
                mprog.partbodyfeatureactivate('喉口倒圓角_60')
                mprog.partbodyfeatureactivate('喉口倒角補料_60')
                mprog.partbodyfeatureactivate('E_倒角')
                mprog.activatefeature('110通孔', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('光電裝置電線固定孔', 5)
                mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 2)
                mprog.activatefeature('兩點組合', 0)
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
                mprog.partbodyfeatureactivate("welding_喉口")
            elif i == 4:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒角_B_圓角')
                mprog.partbodyfeatureactivate('B倒角_E段')
                mprog.partbodyfeatureactivate('喉口倒角補料_80')
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('光電裝置電線固定孔', 5)
                mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 3)
                mprog.activatefeature('兩點組合', 0)
                mprog.activatefeature('3-M5X10L(配管用)(r1)', 1)
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
                mprog.partbodyfeatureactivate("welding_喉口")
            elif i == 5:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_110')
                mprog.partbodyfeatureactivate('E_倒角')
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('光電裝置電線固定孔', 0)
                mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 3)
                mprog.activatefeature('兩點組合', 0)
                mprog.activatefeature('3-M5X10L(配管用)(r1)', 2)
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
                mprog.partbodyfeatureactivate("welding_喉口")
            elif i == 6:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_160')
                mprog.partbodyfeatureactivate('E_倒角')
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('3-M5X10L(配管用)(r1)', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('光電裝置電線固定孔', 0)
                mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 5)
                mprog.activatefeature('兩點組合', 0)
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
                mprog.partbodyfeatureactivate("welding_喉口")
            elif i == 7:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_200')
                mprog.partbodyfeatureactivate('E_倒角')
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('3-M5X10L(配管用)(r1)', 2)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('光電裝置電線固定孔', 0)
                mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 5)
                mprog.activatefeature('兩點組合', 0)
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
                mprog.partbodyfeatureactivate("welding_喉口")
            elif i == 8:
                mprog.partbodyfeatureactivate('T250')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1_250')
                mprog.partbodyfeatureactivate('V3_250')
                mprog.partbodyfeatureactivate('V5_250')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('喉口倒角_250')
                mprog.partbodyfeatureactivate('喉口倒圓角_250')
                mprog.partbodyfeatureactivate('喉口倒角補料_250')
                mprog.partbodyfeatureactivate('E_倒角')
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('3-M5X10L(配管用)(r1)', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('光電裝置電線固定孔', 0)
                mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 0)
                mprog.activatefeature('兩點組合', 0)
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
                mprog.partbodyfeatureactivate("T250_FF")
            mprog.update()
        elif name == 'FRAME2':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('AA2')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('VV2')
                    mprog.partbodyfeatureactivate('VV4')
                    mprog.partbodyfeatureactivate('喉口倒角_B')
                    mprog.partbodyfeatureactivate('喉口倒圓角_B')
                    mprog.partbodyfeatureactivate('喉口倒角補料_45')
                    mprog.partbodyfeatureactivate('R倒角')
                    mprog.partbodyfeatureactivate('R倒角_qqqq')
                    mprog.partbodyfeatureactivate('R倒角_G1')
                    mprog.partbodyfeatureactivate('Z倒角')
                    mprog.partbodyfeatureactivate('E段_B1倒角')
                    mprog.partbodyfeatureactivate('E段_B倒角_Y')
                    mprog.partbodyfeatureactivate('pad1左&右下倒角')
                    mprog.partdeactivate('machining_DD')
                    mprog.partdeactivate('machining_E上方研磨')
                    mprog.bodydeactivate('machining_2x8_12通孔', 0)
                    mprog.bodydeactivate('machining_aaaaa2', 0)
                    mprog.bodydeactivate('machining_aaaaa4', 0)
                    mprog.partbodyfeatureactivate("welding_喉口")
                    mprog.bodydeactivate('解角器固定架用', 0)
                    mprog.bodydeactivate('解角器固定架用_2', 0)
                    mprog.bodydeactivate('6_M10*12L', 0)
                    mprog.partdeactivate('JJ')
                elif i == 1:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('AA2')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('VV2')
                    mprog.partbodyfeatureactivate('VV4')
                    mprog.partbodyfeatureactivate('喉口倒角_B')
                    mprog.partbodyfeatureactivate('喉口倒圓角_B')
                    mprog.partbodyfeatureactivate('喉口倒角補料_35')
                    mprog.partbodyfeatureactivate('R倒角')
                    mprog.partbodyfeatureactivate('R倒角_qqqq')
                    mprog.partbodyfeatureactivate('R倒角_G1')
                    mprog.partbodyfeatureactivate('Z倒角')
                    mprog.partbodyfeatureactivate('E段_B1倒角')
                    mprog.partbodyfeatureactivate('E段_B倒角_Y')
                    mprog.partbodyfeatureactivate('pad1左&右下倒角')
                    mprog.partdeactivate('machining_DD')
                    mprog.partdeactivate('machining_E上方研磨')
                    mprog.bodydeactivate('machining_2x8_12通孔', 0)
                    mprog.bodydeactivate('machining_aaaaa2', 0)
                    mprog.bodydeactivate('machining_aaaaa4', 0)
                    mprog.partbodyfeatureactivate("welding_喉口")
                    mprog.bodydeactivate('解角器固定架用', 0)
                    mprog.bodydeactivate('解角器固定架用_2', 0)
                    mprog.bodydeactivate('6_M10*12L', 0)
                    mprog.partdeactivate('JJ')
                elif i == 2:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('AA2')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('VV2')
                    mprog.partbodyfeatureactivate('VV4')
                    mprog.partbodyfeatureactivate('喉口倒角_B')
                    mprog.partbodyfeatureactivate('喉口倒圓角_B')
                    mprog.partbodyfeatureactivate('喉口倒角補料_45')
                    mprog.partbodyfeatureactivate('R倒角')
                    mprog.partbodyfeatureactivate('R倒角_qqqq')
                    mprog.partbodyfeatureactivate('R倒角_G1')
                    mprog.partbodyfeatureactivate('Z倒角')
                    mprog.partbodyfeatureactivate('E段_B1倒角')
                    mprog.partbodyfeatureactivate('E段_B倒角_Y')
                    mprog.partbodyfeatureactivate('pad1左&右下倒角')
                    mprog.partdeactivate('machining_DD')
                    mprog.partdeactivate('machining_E上方研磨')
                    mprog.bodydeactivate('machining_2x8_12通孔', 0)
                    mprog.bodydeactivate('machining_aaaaa2', 0)
                    mprog.bodydeactivate('machining_aaaaa4', 0)
                    mprog.partbodyfeatureactivate("welding_喉口")
                    mprog.bodydeactivate('解角器固定架用', 0)
                    mprog.bodydeactivate('解角器固定架用_2', 0)
                    mprog.bodydeactivate('6_M10*12L', 0)
                    mprog.partdeactivate('JJ')
                elif i == 3:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('AA2')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('V1')
                    mprog.partbodyfeatureactivate('V3')
                    mprog.partbodyfeatureactivate('V5')
                    mprog.partbodyfeatureactivate('喉口倒角_非B')
                    mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                    mprog.partbodyfeatureactivate('喉口倒角補料_60')
                    mprog.partbodyfeatureactivate('R倒角_無qqqq')
                    mprog.partbodyfeatureactivate('R倒角_G1')
                    mprog.partbodyfeatureactivate('Z倒角')
                    mprog.partbodyfeatureactivate('E段倒角_非B')
                    mprog.partbodyfeatureactivate('pad1左&右下倒角')
                    mprog.partdeactivate('machining_DD')
                    mprog.partdeactivate('machining_E上方研磨')
                    mprog.bodydeactivate('machining_2x8_12通孔', 0)
                    mprog.bodydeactivate('machining_aaaaa2', 0)
                    mprog.bodydeactivate('machining_aaaaa4', 0)
                    mprog.partbodyfeatureactivate("welding_喉口")
                    mprog.bodydeactivate('解角器固定架用', 0)
                    mprog.bodydeactivate('解角器固定架用_2', 0)
                    mprog.bodydeactivate('6_M10*12L', 0)
                    mprog.partdeactivate('JJ')
                elif i == 4:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('V1')
                    mprog.partbodyfeatureactivate('V3')
                    mprog.partbodyfeatureactivate('V5')
                    mprog.partbodyfeatureactivate('喉口倒角_B')
                    mprog.partbodyfeatureactivate('喉口倒圓角_B')
                    mprog.partbodyfeatureactivate('喉口倒角補料_80')
                    mprog.partbodyfeatureactivate('Z倒角')
                    mprog.partbodyfeatureactivate('E段倒角_B')
                    mprog.partbodyfeatureactivate('jjj4')
                    mprog.partbodyfeatureactivate('mmm4')
                    mprog.partbodyfeatureactivate('pad1左&右下倒角')
                    mprog.partdeactivate('machining_DD')
                    mprog.partdeactivate('machining_E上方研磨')
                    mprog.bodydeactivate('machining_2x8_12通孔', 0)
                    mprog.bodydeactivate('machining_aaaaa2', 0)
                    mprog.bodydeactivate('machining_aaaaa4', 0)
                    mprog.partbodyfeatureactivate("welding_喉口")
                elif i == 5:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('V1')
                    mprog.partbodyfeatureactivate('V3')
                    mprog.partbodyfeatureactivate('V5')
                    mprog.partbodyfeatureactivate('喉口倒角_非B')
                    mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                    mprog.partbodyfeatureactivate('喉口倒角補料_110')
                    mprog.partbodyfeatureactivate('Z倒角')
                    mprog.partbodyfeatureactivate('E段倒角_非B')
                    mprog.partbodyfeatureactivate('dd4')
                    mprog.partbodyfeatureactivate('x4')
                    mprog.partbodyfeatureactivate('pad1左&右下倒角')
                    mprog.partdeactivate('machining_DD')
                    mprog.partdeactivate('machining_E上方研磨')
                    mprog.bodydeactivate('machining_2x8_12通孔', 0)
                    mprog.bodydeactivate('machining_aaaaa2', 0)
                    mprog.bodydeactivate('machining_aaaaa4', 0)
                    mprog.partbodyfeatureactivate("welding_喉口")
                elif i == 6:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('V1')
                    mprog.partbodyfeatureactivate('V3')
                    mprog.partbodyfeatureactivate('V5')
                    mprog.partbodyfeatureactivate('喉口倒角_非B')
                    mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                    mprog.partbodyfeatureactivate('喉口倒角補料_160')
                    mprog.partbodyfeatureactivate('Z倒角')
                    mprog.partbodyfeatureactivate('E段倒角_非B')
                    mprog.partbodyfeatureactivate('ggg4')
                    mprog.partbodyfeatureactivate('xx3')
                    mprog.partbodyfeatureactivate('iii4')
                    mprog.partbodyfeatureactivate('fff4')
                    mprog.partbodyfeatureactivate('pad1左&右下倒角')
                    mprog.partdeactivate('machining_DD')
                    mprog.partdeactivate('machining_E上方研磨')
                    mprog.bodydeactivate('machining_2x8_12通孔', 0)
                    mprog.bodydeactivate('machining_aaaaa2', 0)
                    mprog.bodydeactivate('machining_aaaaa4', 0)
                    mprog.partbodyfeatureactivate("welding_喉口")
                elif i == 7:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('V1')
                    mprog.partbodyfeatureactivate('V3')
                    mprog.partbodyfeatureactivate('V5')
                    mprog.partbodyfeatureactivate('喉口倒角_非B')
                    mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                    mprog.partbodyfeatureactivate('喉口倒角補料_200')
                    mprog.partbodyfeatureactivate('Z倒角')
                    mprog.partbodyfeatureactivate('E段倒角_非B')
                    mprog.partbodyfeatureactivate('iii4')
                    mprog.partbodyfeatureactivate('ssss4')
                    mprog.partbodyfeatureactivate('sss4')
                    mprog.partbodyfeatureactivate('pad1左&右下倒角')
                    mprog.partdeactivate('machining_DD')
                    mprog.partdeactivate('machining_E上方研磨')
                    mprog.bodydeactivate('machining_2x8_12通孔', 0)
                    mprog.bodydeactivate('machining_aaaaa2', 0)
                    mprog.bodydeactivate('machining_aaaaa4', 0)
                    mprog.partbodyfeatureactivate("welding_喉口")
                elif i == 8:
                    mprog.partbodyfeatureactivate('T250')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('V1_250')
                    mprog.partbodyfeatureactivate('V3_250')
                    mprog.partbodyfeatureactivate('V5_250')
                    mprog.partbodyfeatureactivate('喉口倒角_250')
                    mprog.partbodyfeatureactivate('喉口倒圓角_250')
                    mprog.partbodyfeatureactivate('喉口倒角補料_250')
                    mprog.partbodyfeatureactivate('Z倒角')
                    mprog.partbodyfeatureactivate('E段倒角_非B')
                    mprog.partbodyfeatureactivate('pad1左&右下倒角')
                    mprog.partdeactivate('machining_DD')
                    mprog.partdeactivate('machining_E上方研磨')
                    mprog.bodydeactivate('machining_2x8_12通孔', 0)
                    mprog.bodydeactivate('machining_aaaaa2', 0)
                    mprog.bodydeactivate('machining_aaaaa4', 0)
                    mprog.partbodyfeatureactivate("T250_FF")
            except BaseException:
                print('FRAME2 Part activate error')
            finally:
                mprog.update()
        elif name == 'FRAME3':
            try:
                if i == 0:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('SN1_25250_M', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 1)
                    mprog.partdeactivate('FRAME_SN1_25250_PQ')
                    mprog.partdeactivate('FRAME_SN1_2560_S')
                    mprog.partdeactivate('FRAME_SN1_2560_U')
                    mprog.partdeactivate('machining_AA')
                    mprog.partdeactivate('machining_L')
                    mprog.partdeactivate('machining_bbbbb')
                    # mprog.partdeactivate("machining_bbbbb_2")
                    mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                elif i == 1:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('SN1_25250_M', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 1)
                    mprog.partdeactivate('FRAME_SN1_25250_PQ')
                    mprog.partdeactivate('FRAME_SN1_2560_S')
                    mprog.partdeactivate('FRAME_SN1_2560_U')
                    mprog.partdeactivate('machining_AA')
                    mprog.partdeactivate('machining_L')
                    mprog.partdeactivate('machining_bbbbb')
                    # mprog.partdeactivate("machining_bbbbb_2")
                    mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                elif i == 2:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('SN1_25250_M', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 1)
                    mprog.partdeactivate('FRAME_SN1_25250_PQ')
                    mprog.partdeactivate('FRAME_SN1_2560_S')
                    mprog.partdeactivate('FRAME_SN1_2560_U')
                    mprog.partdeactivate('machining_AA')
                    mprog.partdeactivate('machining_L')
                    mprog.partdeactivate('machining_bbbbb')
                    # mprog.partdeactivate("machining_bbbbb_2")
                    mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                elif i == 3:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 1)
                    mprog.partdeactivate('FRAME_SN1_25250_PQ')
                    mprog.partdeactivate('FRAME_SN1_2560_S')
                    mprog.partdeactivate('FRAME_SN1_2560_U')
                    mprog.partdeactivate('machining_AA')
                    mprog.partdeactivate('machining_bbbbb')
                    # mprog.partdeactivate("machining_bbbbb_2")
                    mprog.partdeactivate('machining_L')
                    mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                elif i == 4:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('SN1_25250_M', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 1)
                    mprog.partdeactivate('FRAME_SN1_25250_PQ')
                    mprog.partdeactivate('FRAME_SN1_80250_R')
                    mprog.partdeactivate('machining_AA')
                    mprog.partdeactivate('machining_L')
                    mprog.partdeactivate('machining_bbbbb')
                    # mprog.partdeactivate("machining_bbbbb_2")
                    mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                elif i == 5:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 1)
                    mprog.partdeactivate('FRAME_SN1_25250_PQ')
                    mprog.partdeactivate('FRAME_SN1_80250_R')
                    mprog.partdeactivate('machining_AA')
                    mprog.partdeactivate('machining_L')
                    mprog.partdeactivate('machining_bbbbb')
                    # mprog.partdeactivate("machining_bbbbb_2")
                    mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                elif i == 6:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('SN1_25250_M', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 1)
                    mprog.partdeactivate('FRAME_SN1_25250_PQ')
                    mprog.partdeactivate('FRAME_SN1_80250_R')
                    mprog.partdeactivate('machining_AA')
                    mprog.partdeactivate('machining_L')
                    mprog.partdeactivate('machining_bbbbb')
                    # mprog.partdeactivate("machining_bbbbb_2")
                    mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                elif i == 7:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('SN1_25250_M', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 1)
                    mprog.partdeactivate('FRAME_SN1_25250_PQ')
                    mprog.partdeactivate('FRAME_SN1_80250_R')
                    mprog.partdeactivate('machining_AA')
                    mprog.partdeactivate('machining_L')
                    mprog.partdeactivate('machining_bbbbb')
                    # mprog.partdeactivate("machining_bbbbb_2")
                    mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                elif i == 8:
                    mprog.activatefeature('SN1_25250_Body', 4)
                    mprog.activatefeature('SN1_25250_M', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 1)
                    mprog.partdeactivate('FRAME_SN1_25250_PQ')
                    mprog.partdeactivate('FRAME_SN1_80250_R')
                    mprog.partdeactivate('machining_AA')
                    mprog.partdeactivate('machining_L')
                    mprog.partdeactivate('machining_bbbbb')
                    # mprog.partdeactivate("machining_bbbbb_2")
                    mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
            except:
                print('FRAME3 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME3 Update error')
        elif name == 'FRAME4':  # 已更改
            try:
                if i == 0:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('Hole_1', 3)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.partdeactivate("machining_P")
                elif i == 1:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('Hole_1', 2)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.partdeactivate("machining_P")
                elif i == 2:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.partdeactivate("machining_P")
                elif i == 3:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.activatefeature('Hole_3', 2)
                    mprog.partdeactivate("machining_N")
                    mprog.partdeactivate("machining_O")
                    mprog.partdeactivate("machining_M")
                    mprog.partdeactivate("machining_P")
                elif i == 4:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.activatefeature('Hole_3', 2)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partdeactivate("machining_N")
                    mprog.partdeactivate("machining_O")
                    mprog.partdeactivate("machining_M")
                    mprog.partdeactivate("machining_P")
                elif i == 5:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.activatefeature('Hole_3', 2)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partdeactivate("machining_N")
                    mprog.partdeactivate("machining_O")
                    mprog.partdeactivate("machining_M")
                    mprog.partdeactivate("machining_P")
                elif i == 6:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.activatefeature('Hole_3', 2)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partdeactivate("machining_N")
                    mprog.partdeactivate("machining_O")
                    mprog.partdeactivate("machining_M")
                    mprog.partdeactivate("machining_P")
                elif i == 7:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.activatefeature('Hole_3', 2)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partdeactivate("machining_N")
                    mprog.partdeactivate("machining_O")
                    mprog.partdeactivate("machining_M")
                    mprog.partdeactivate("machining_P")
                elif i == 8:
                    mprog.activatefeature('SN1_25250_Body', 0)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.activatefeature('Hole_3', 4)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partdeactivate("machining_N")
                    mprog.partdeactivate("machining_O")
                    mprog.partdeactivate("machining_M")
                    mprog.partdeactivate("machining_P")
            except:
                print('FRAME4 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME4 Update error')
        elif name == 'FRAME5':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('SN1-25_X')
                    mprog.partbodyfeatureactivate('CD')
                    mprog.partbodyfeatureactivate("I")
                    mprog.partbodyfeatureactivate('SN1_25_H')
                    mprog.activatefeature('2-M8通孔', 0)
                    mprog.activatefeature('TY牌電磁閥用', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                    mprog.activatefeature('PT1/2', 0)
                elif i == 1:
                    mprog.partbodyfeatureactivate('CD')
                    mprog.partbodyfeatureactivate("I")
                    mprog.partbodyfeatureactivate("SN1_3545_H")
                    mprog.activatefeature('2-M8通孔', 0)
                    mprog.activatefeature('TY牌電磁閥用', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                    mprog.activatefeature('PT1/2', 0)
                elif i == 2:
                    mprog.partbodyfeatureactivate('CD')
                    mprog.partbodyfeatureactivate("I")
                    mprog.partbodyfeatureactivate("SN1_3545_H")
                    mprog.activatefeature('2-M8通孔', 0)
                    mprog.activatefeature('TY牌電磁閥用', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                    mprog.activatefeature('PT1/2', 0)
                elif i == 3:
                    mprog.partbodyfeatureactivate('CD')
                    mprog.partbodyfeatureactivate("I")
                    mprog.partbodyfeatureactivate('NM')
                    mprog.partbodyfeatureactivate('SN1_60_R')
                    mprog.activatefeature('2-M8通孔', 0)
                    mprog.activatefeature('TY牌電磁閥用', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                elif i == 4:
                    mprog.partbodyfeatureactivate('SN1_80250_I')
                    mprog.partbodyfeatureactivate('SN1_80250_OP')
                    mprog.partbodyfeatureactivate('SN1_80250_Q')
                    mprog.activatefeature('2-M8通孔', 0)
                    mprog.activatefeature('TY牌電磁閥用', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                elif i == 5:
                    mprog.partbodyfeatureactivate('SN1_80250_I')
                    mprog.partbodyfeatureactivate('SN1_80250_OP')
                    mprog.partbodyfeatureactivate('SN1_80250_Q')
                    mprog.partbodyfeatureactivate('6-M6通C5')
                    mprog.activatefeature('6-M6通', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                elif i == 6:
                    mprog.partbodyfeatureactivate('SN1_80250_I')
                    mprog.partbodyfeatureactivate('SN1_80250_OP')
                    mprog.partbodyfeatureactivate('SN1_80250_Q')
                    mprog.partbodyfeatureactivate('3-M8通_C5')
                    mprog.activatefeature('3-M8通', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                elif i == 7:
                    mprog.partbodyfeatureactivate('SN1_80250_I')
                    mprog.partbodyfeatureactivate('SN1_80250_OP')
                    mprog.partbodyfeatureactivate('SN1_80250_Q')
                    mprog.partbodyfeatureactivate('5-M8通_C5')
                    mprog.activatefeature('5-M8通', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                elif i == 8:
                    mprog.partbodyfeatureactivate('SN1_80250_I')
                    mprog.partbodyfeatureactivate('SN1_80250_OP')
                    mprog.partbodyfeatureactivate('SN1_80250_Q')
                    mprog.partbodyfeatureactivate('5-M8通_C5')
                    mprog.activatefeature('5-M8通', 0)
                    mprog.activatefeature('M6通孔LED', 0)
            except BaseException:
                print('FRAME5 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME6':
            try:
                if i == 0:
                    mprog.bodydeactivate('FRAME_SN1_2545_Body', 0)
                    mprog.bodydeactivate('M16_i', 0)
                    mprog.activatefeature('SN1_2545_Body', 0)
                elif i == 1:
                    mprog.bodydeactivate('FRAME_SN1_2545_Body', 0)
                    mprog.bodydeactivate('M16_i', 0)
                    mprog.activatefeature('SN1_2545_Body', 0)
                elif i == 2:
                    mprog.bodydeactivate('FRAME_SN1_2545_Body', 0)
                    mprog.bodydeactivate('M16_i', 0)
                    mprog.activatefeature('SN1_2545_Body', 0)
                elif i == 3:
                    mprog.bodydeactivate('FRAME_SN1_60_Body', 0)
                    mprog.bodydeactivate('M16_h FRAME9_L', 0)
                    mprog.activatefeature('SN1_60_Body', 0)
                elif i == 4:
                    mprog.bodydeactivate('FRAME_SN1_80110_Body', 0)
                    mprog.bodydeactivate('M16_h FRAME9_L', 0)
                    mprog.activatefeature('SN1_80110_Body', 0)
                elif i == 5:
                    mprog.bodydeactivate('FRAME_SN1_80110_Body', 0)
                    mprog.bodydeactivate('M16_h FRAME9_L', 0)
                    mprog.activatefeature('SN1_80110_Body', 0)
                elif i == 7:
                    mprog.bodydeactivate('FRAME_SN1_200_Body', 0)
                    mprog.bodydeactivate('M16_h FRAME9_L', 0)
                    mprog.activatefeature('SN1_200_Body', 0)
            except:
                print('FRAME6 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME6 Update error')
        elif name == 'FRAME7':
            try:
                if i == 0:
                    mprog.activatefeature('SN1_2560_Body', 0)
                    mprog.partbodyfeatureactivate("SN1_2545_C")
                    mprog.partbodyfeatureactivate("SN1_2545_D")
                    mprog.partbodyfeatureactivate("SN1_2545_E")
                    mprog.partdeactivate('SN1_25_VWX')
                    mprog.partdeactivate('DD')
                elif i == 1:
                    mprog.activatefeature('SN1_2560_Body', 0)
                    mprog.partbodyfeatureactivate("SN1_2545_C")
                    mprog.partbodyfeatureactivate("SN1_2545_D")
                    mprog.partbodyfeatureactivate("SN1_2545_E")
                    mprog.partdeactivate('DD')
                elif i == 2:
                    mprog.activatefeature('SN1_2560_Body', 0)
                    mprog.partbodyfeatureactivate("SN1_2545_C")
                    mprog.partbodyfeatureactivate("SN1_2545_D")
                    mprog.partbodyfeatureactivate("SN1_2545_E")
                    mprog.partdeactivate('DD')
                elif i == 3:
                    mprog.activatefeature('SN1_2560_Body', 1)
                    mprog.partbodyfeatureactivate("SN1_60_C")
                    mprog.partbodyfeatureactivate("SN1_60_D")
                    mprog.partbodyfeatureactivate("SN1_60_E")
                    mprog.partdeactivate('DD')
                elif i == 4:
                    mprog.activatefeature('SN1_80250_Body', 7)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partdeactivate("Y")
                    mprog.partdeactivate('DD')
                elif i == 5:
                    mprog.partdeactivate("Y")
                    mprog.activatefeature('SN1_80250_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partdeactivate('DD')
                elif i == 6:
                    mprog.partdeactivate("Y")
                    mprog.activatefeature('SN1_80250_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partdeactivate('DD')
                elif i == 7:
                    mprog.partdeactivate("Y")
                    mprog.activatefeature('SN1_80250_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partdeactivate('DD')
                elif i == 8:
                    mprog.partdeactivate("Y")
                    mprog.activatefeature('SN1_80250_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partdeactivate('DD')
            except:
                print('FRAME7 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME7 Update error')
        elif name == 'FRAME8':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('上半倒角_25')
                    mprog.partbodyfeatureactivate('上半倒角_2')
                elif i == 1:
                    mprog.partbodyfeatureactivate('上半倒角_1')
                    mprog.partbodyfeatureactivate('上半倒角_2')
                    # mprog.activatefeature('38孔通', 0)
                elif i == 2:
                    mprog.partbodyfeatureactivate('上半倒角_1')
                    mprog.partbodyfeatureactivate('上半倒角_2')
                elif i == 3:
                    mprog.partbodyfeatureactivate('上半倒角_1')
                    mprog.partbodyfeatureactivate('上半倒角_2')
                elif i == 4:
                    mprog.partbodyfeatureactivate('上半倒角80')
                    mprog.partbodyfeatureactivate('R40特徵')
                    mprog.partbodyfeatureactivate('下半開槽')
                    mprog.activatefeature('38孔通', 0)
                elif i == 5:
                    mprog.partbodyfeatureactivate('上半倒角_1')
                    mprog.partbodyfeatureactivate('上半倒角_2')
                    mprog.partbodyfeatureactivate('下半倒角')
                    mprog.partbodyfeatureactivate('下半開槽')
                    mprog.partbodyfeatureactivate('R40特徵')
                    mprog.activatefeature('38孔通', 0)
                elif i == 6:
                    mprog.partbodyfeatureactivate('上半倒角_1')
                    mprog.partbodyfeatureactivate('上半倒角_2')
                    mprog.partbodyfeatureactivate('下半倒角')
                    mprog.partbodyfeatureactivate('下半開槽')
                    mprog.partbodyfeatureactivate('R40特徵')
                    mprog.activatefeature('38孔通', 0)
                elif i == 7:
                    mprog.partbodyfeatureactivate('上半倒角_1')
                    mprog.partbodyfeatureactivate('上半倒角_2')
                    mprog.partbodyfeatureactivate('下半開槽')
                    mprog.partbodyfeatureactivate('R40特徵')
                    mprog.activatefeature('38孔通', 0)
                elif i == 8:
                    mprog.partbodyfeatureactivate('上半倒角_1')
                    mprog.partbodyfeatureactivate('上半倒角_2')
                    mprog.partbodyfeatureactivate('R40特徵')
                    mprog.activatefeature('38孔通', 0)
            except BaseException:
                print('FRAME8 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME9':
            try:
                if i == 0:
                    mprog.activatefeature('SN1_2580_Body', 0)
                elif i == 1:
                    mprog.activatefeature('SN1_2580_Body', 0)
                elif i == 2:
                    mprog.activatefeature('SN1_2580_Body', 0)
                elif i == 3:
                    mprog.activatefeature('SN1_2580_Body', 0)
                elif i == 4:
                    mprog.activatefeature('channel steel', 0)
                elif i == 5:
                    mprog.activatefeature('channel steel', 0)
                elif i == 6:
                    mprog.activatefeature('channel steel', 0)
                elif i == 7:
                    mprog.activatefeature('channel steel', 0)
                elif i == 8:
                    mprog.activatefeature('channel steel', 0)
            except:
                print('FRAME9 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME9 Update error')
        elif name == 'FRAME10':
            try:
                if i == 0:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partdeactivate('SN1_25250_K')
                elif i == 1:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partdeactivate('SN1_25250_K')
                elif i == 2:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partdeactivate('SN1_25250_K')
                elif i == 3:
                    mprog.activatefeature('SN1_60_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partdeactivate('SN1_25250_K')
                elif i == 4:
                    mprog.activatefeature('SN1_80110_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partdeactivate('SN1_25250_K')
                elif i == 5:
                    mprog.activatefeature('SN1_80110_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partdeactivate('SN1_25250_K')
                elif i == 6:
                    mprog.activatefeature('SN1_160250_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partdeactivate('SN1_25250_K')
                elif i == 7:
                    mprog.activatefeature('SN1_160250_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partdeactivate('SN1_25250_K')
                elif i == 8:
                    mprog.activatefeature('SN1_160250_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partdeactivate('SN1_25250_K')
            except:
                print('FRAME10 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME11':
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
            mprog.Update()
        elif name == 'FRAME12':
            pass
        elif name == 'FRAME13':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_25_45_CD')
                    mprog.partbodyfeatureactivate('SN1_25_45_G')
                    mprog.partbodyfeatureactivate('F')
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.bodydeactivate('Hole_4', 0)
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 1:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('F')
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.bodydeactivate('Hole_4', 0)
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 2:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_25_45_CD')
                    mprog.partbodyfeatureactivate('SN1_25_45_G')
                    mprog.partbodyfeatureactivate('F')
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.bodydeactivate('Hole_4', 0)
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 3:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('F')
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.bodydeactivate('Hole_4', 0)
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 4:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('F')
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.bodydeactivate('Hole_4', 0)
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 5:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('F')
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.bodydeactivate('Hole_4', 0)
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 6:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('F')
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.bodydeactivate('Hole_4', 0)
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 7:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('F')
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.bodydeactivate('Hole_4', 0)
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 8:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('F')
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.bodydeactivate('Hole_4', 0)
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
            except:
                print('FRAME13 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME13 Update error')
        elif name == 'FRAME14':
            try:
                if i == 0:  # 25N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 1:  # 35N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 2:  # 45N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 3:  # 60N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 4:  # 80N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('outside_hole_E', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 5:  # 110N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 6:  # 160N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 7:  # 200N
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('outside_hole_E', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 8:  # 250N
                    mprog.partbodyfeatureactivate('250N')
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                print('FRAME14 machining success')
            except:
                print('FRAME14 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME15':
            try:
                if i == 0:  # 25N
                    mprog.activatefeature('Hole', 0)
                    mprog.partdeactivate('F')
                elif i == 1:  # 35N
                    mprog.activatefeature('Hole', 0)
                    mprog.partdeactivate('F')
                elif i == 2:  # 45N
                    mprog.activatefeature('Hole', 0)
                    mprog.partdeactivate('F')
                elif i == 3:  # 60N
                    mprog.activatefeature('Hole', 0)
                    mprog.partdeactivate('F')
                elif i == 4:  # 80N
                    mprog.activatefeature('Hole', 0)
                    mprog.partdeactivate('F')
                elif i == 5:  # 110N
                    mprog.activatefeature('Hole', 0)
                    mprog.partdeactivate('F')
                elif i == 6:  # 160N
                    mprog.activatefeature('Hole', 0)
                    mprog.partdeactivate('F')
                elif i == 7:  # 200N
                    mprog.partbodyfeatureactivate('SN1_200250')
                    mprog.activatefeature('Hole', 0)
                    mprog.partdeactivate('F')
                elif i == 8:  # 250N
                    mprog.partbodyfeatureactivate('SN1_200250')
                    mprog.activatefeature('Hole', 0)
                    mprog.partdeactivate('F')
            except:
                print('FRAME15 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME15 Update error')
        elif name == 'FRAME17':
            try:
                if i == 0:  # 25N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 1:  # 35N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 2:  # 45N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 3:  # 60N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 4:  # 80N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('outside_hole_E', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 5:  # 110N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 6:  # 160N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 7:  # 200N
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('outside_hole_E', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                elif i == 8:  # 250N
                    mprog.partbodyfeatureactivate('250N')
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.bodydeactivate('processing_h', 0)
                    mprog.bodydeactivate('processing_i', 0)
                    mprog.bodydeactivate('processing_m', 0)
                    mprog.partdeactivate("H")
                    mprog.partdeactivate("I")
                print('FRAME17 machining success')
            except:
                print('FRAME17 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME18':
            pass
        elif name == 'FRAME19':
            try:
                if i == 0:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.partbodyfeatureactivate('SN1_25_X')
                    mprog.partbodyfeatureactivate('SN1_2535_Y')
                    mprog.partdeactivate("machining_AI")
                elif i == 1:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.partbodyfeatureactivate('SN1_2535_Y')
                    mprog.partbodyfeatureactivate("SN1_3545_AD")
                    mprog.partdeactivate("machining_AI")
                elif i == 2:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.partbodyfeatureactivate('SN1_45_Y')
                    mprog.partbodyfeatureactivate('SN1_45_X')
                    mprog.partbodyfeatureactivate("SN1_3545_AD")
                    mprog.partdeactivate("machining_AI")
                elif i == 3:
                    mprog.activatefeature('SN1_60_Body', 0)
                    mprog.partdeactivate("machining_AI")
                elif i == 4:
                    mprog.activatefeature('SN1_80110_Body', 0)
                    mprog.partdeactivate("machining_AI")
                elif i == 5:
                    mprog.activatefeature('SN1_80110_Body', 0)
                    mprog.partdeactivate("machining_AI")
                elif i == 6:
                    mprog.activatefeature('SN1_160_Body', 0)
                    mprog.bodydeactivate('FRMAE_SN1_160_Body', 0)
                    mprog.bodydeactivate('FRAME_34_Hole_1', 0)
                    mprog.partdeactivate("machining_AI")
                elif i == 7:
                    mprog.activatefeature('SN1_200_Body', 0)
                    mprog.partdeactivate("machining_AI")
                elif i == 8:
                    mprog.activatefeature('SN1_250_Body', 0)
                    mprog.partdeactivate("machining_AI")
            except:
                print('FRAME19 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME19 Update error')
        elif name == 'FRAME20':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
                elif i == 1:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
                elif i == 2:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
                elif i == 3:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_60250_C')
                    mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
                elif i == 4:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_60250_C')
                    mprog.partbodyfeatureactivate('SN1_80250_DEF')
                    mprog.partbodyfeatureactivate('SN1_80250_G')
                    mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
                elif i == 5:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_60250_C')
                    mprog.partbodyfeatureactivate('SN1_80250_DEF')
                    mprog.partbodyfeatureactivate('SN1_80250_G')
                    mprog.partbodyfeatureactivate('SN1_110160_ijk')
                    mprog.activatefeature('Hole_2', 0)
                elif i == 6:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_60250_C')
                    mprog.partbodyfeatureactivate('SN1_80250_DEF')
                    mprog.partbodyfeatureactivate('SN1_80250_G')
                    mprog.partbodyfeatureactivate('SN1_110160_ijk')
                    mprog.activatefeature('Hole_2', 0)
                elif i == 7:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_60250_C')
                    mprog.partbodyfeatureactivate('SN1_80250_DEF')
                    mprog.partbodyfeatureactivate('SN1_80250_G')
                    mprog.partbodyfeatureactivate('SN1_200250_hij')
                    mprog.activatefeature('Hole_2', 1)
                elif i == 8:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_60250_C')
                    mprog.partbodyfeatureactivate('SN1_80250_DEF')
                    mprog.partbodyfeatureactivate('SN1_80250_G')
                    mprog.partbodyfeatureactivate('SN1_200250_hij')
                    mprog.activatefeature('Hole_2', 1)
            except:
                print('FRAME20 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME20 Update error')
        elif name == 'FRAME21':
            try:
                mprog.partdeactivate("H")
                if i == 0:  # 25
                    mprog.activatefeature('Hole', 0)
                elif i == 1:  # 35
                    mprog.activatefeature('Hole', 0)
            except:
                print('FRAME21 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME22':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('A(1)D')
                    mprog.partbodyfeatureactivate('SN1-2560_BE')
                    mprog.partbodyfeatureactivate('SN1_2545_I')
                    mprog.partbodyfeatureactivate('SN1_2545_GH')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.bodydeactivate('SN1_2560_6_M8X16L_底孔27L', 0)
                elif i == 1:
                    mprog.partbodyfeatureactivate('A(1)D')
                    mprog.partbodyfeatureactivate('SN1-2560_BE')
                    mprog.partbodyfeatureactivate('SN1_2545_I')
                    mprog.partbodyfeatureactivate('SN1_2545_GH')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.bodydeactivate('SN1_2560_6_M8X16L_底孔27L', 0)
                elif i == 2:
                    mprog.partbodyfeatureactivate('A(1)D')
                    mprog.partbodyfeatureactivate('SN1-2560_BE')
                    mprog.partbodyfeatureactivate('SN1_2545_I')
                    mprog.partbodyfeatureactivate('SN1_2545_GH')
                    mprog.partbodyfeatureactivate('SN1_45_G(45)H(45)')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.bodydeactivate('SN1_2560_6_M8X16L_底孔27L', 0)
                elif i == 3:
                    mprog.partbodyfeatureactivate('A(1)D')
                    mprog.partbodyfeatureactivate('SN1-2560_BE')
                    mprog.partbodyfeatureactivate('SN1_60_I')
                    mprog.partbodyfeatureactivate('SN1_60_JK')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.partdeactivate('A(2)')
                    mprog.activatefeature('Hole_3', 0)
                    mprog.bodydeactivate('SN1_2560_6_M8X16L_底孔27L', 0)
                elif i == 4:
                    mprog.partdeactivate('A(2)')
                    mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                    mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                    mprog.partbodyfeatureactivate('SN1_80250_JK')
                    mprog.partbodyfeatureactivate('SN1_80250_GH')
                    mprog.partbodyfeatureactivate('SN1_80250_IN')
                    mprog.partbodyfeatureactivate('SN1_80250_LM1')
                    mprog.partbodyfeatureactivate('SN1_80250_LM2')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                    mprog.activatefeature('Hole_3', 0)
                elif i == 5:
                    mprog.partdeactivate('A(2)')
                    mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                    mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                    mprog.partbodyfeatureactivate('SN1_80250_JK')
                    mprog.partbodyfeatureactivate('SN1_80250_GH')
                    mprog.partbodyfeatureactivate('SN1_80250_IN')
                    mprog.partbodyfeatureactivate('SN1_80250_LM1')
                    mprog.partbodyfeatureactivate('SN1_80250_LM2')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                    mprog.activatefeature('Hole_3', 0)
                elif i == 6:
                    mprog.partdeactivate('A(2)')
                    mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                    mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                    mprog.partbodyfeatureactivate('SN1_80250_JK')
                    mprog.partbodyfeatureactivate('SN1_80250_GH')
                    mprog.partbodyfeatureactivate('SN1_80250_IN')
                    mprog.partbodyfeatureactivate('SN1_80250_LM1')
                    mprog.partbodyfeatureactivate('SN1_80250_LM2')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                    mprog.activatefeature('Hole_3', 0)
                elif i == 7:
                    mprog.partdeactivate('A(2)')
                    mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                    mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                    mprog.partbodyfeatureactivate('SN1_80250_JK')
                    mprog.partbodyfeatureactivate('SN1_80250_GH')
                    mprog.partbodyfeatureactivate('SN1_80250_IN')
                    mprog.partbodyfeatureactivate('SN1_80250_LM1')
                    mprog.partbodyfeatureactivate('SN1_80250_LM2')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                    mprog.activatefeature('Hole_3', 0)
                elif i == 8:
                    mprog.partdeactivate('A(2)')
                    mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                    mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                    mprog.partbodyfeatureactivate('SN1_80250_JK')
                    mprog.partbodyfeatureactivate('SN1_80250_GH')
                    mprog.partbodyfeatureactivate('SN1_80250_IN')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                    mprog.activatefeature('Hole_3', 0)
            except:
                print('FRAME22 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME22 Update error')
        elif name == 'FRAME23':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif i == 1:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif i == 2:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif i == 3:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif i == 4:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('80至250除料')
                    mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
                elif i == 5:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('80至250除料')
                    mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
                elif i == 6:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('80至250除料')
                elif i == 7:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('80至250除料')
                elif i == 8:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('80至250除料')
                    mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
            except:
                print('FRAME23 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME23_1':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif i == 1:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif i == 2:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif i == 3:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif i == 4:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('80至250除料')
                    mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
                elif i == 5:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('80至250除料')
                    mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
                elif i == 6:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('80至250除料')
                elif i == 7:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('80至250除料')
                elif i == 8:
                    mprog.partbodyfeatureactivate('D1')
                    mprog.partbodyfeatureactivate('D2')
                    mprog.partbodyfeatureactivate('J1')
                    mprog.partbodyfeatureactivate('J2')
                    mprog.partbodyfeatureactivate('80至250除料')
                    mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
            except:
                print('FRAME23_1 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME24':
            if i == 0:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 1:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 2:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 3:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 4:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 5:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 6:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 7:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 8:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            mprog.Update()
        elif name == 'FRAME24_1':
            if i == 0:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 1:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 2:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 3:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 4:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 5:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 6:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 7:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            elif i == 8:
                mprog.partdeactivate('Pad.2')
                mprog.partbodyfeatureactivate('Pad.1')
                mprog.partbodyfeatureactivate('倒角C')
                mprog.activatefeature('M5', 0)
                mprog.partdeactivate('machining_bbbbb')
                mprog.partdeactivate('machining_bbbbb_2')
                mprog.bodydeactivate('machining_Hole_1', 0)
            mprog.Update()
        elif name == 'FRAME25':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('3_M8通', 0)
                elif i == 1:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('3_M8通', 0)
                elif i == 2:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('3_M8通', 0)
                elif i == 3:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('3_M8通', 0)
            except:
                print('FRAME25 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME25 Update error')
        elif name == 'FRAME26':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif i == 1:
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif i == 2:
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif i == 3:
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif i == 4:
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif i == 5:
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.partbodyfeatureactivate('F')
                elif i == 6:
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif i == 7:
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif i == 8:
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.partbodyfeatureactivate('250除料')
            except:
                print('FRAME26 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME27':
            try:
                if i == 0:
                    mprog.partdeactivate('machining_aaaaa2')
                    mprog.partbodyfeatureactivate("35N前_FF")
                elif i == 1:
                    mprog.partdeactivate('machining_aaaaa2')
                    mprog.partbodyfeatureactivate("35N前_FF")
                elif i == 2:
                    mprog.partdeactivate('machining_aaaaa2')
                    mprog.bodydeactivate('machining_ddddd', 0)
                    mprog.partbodyfeatureactivate("35N前_FF")
                elif i == 3:
                    mprog.partdeactivate('machining_aaaaa2')
                    mprog.bodydeactivate('machining_ddddd', 0)
                    mprog.partbodyfeatureactivate("FF")
            except BaseException:
                print('FRAME27 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME27_1':
            try:
                if i == 0:
                    mprog.partdeactivate('machining_aaaaa2')
                    mprog.partbodyfeatureactivate('35N前_FF')
                elif i == 1:
                    mprog.partdeactivate('machining_aaaaa2')
                    mprog.partbodyfeatureactivate('35N前_FF')
                elif i == 2:
                    mprog.partdeactivate('machining_aaaaa2')
                    mprog.bodydeactivate('machining_ddddd', 0)
                    mprog.partbodyfeatureactivate('35N前_FF')
                elif i == 3:
                    mprog.partdeactivate('machining_aaaaa2')
                    mprog.bodydeactivate('machining_ddddd', 0)
                    mprog.partbodyfeatureactivate('FF')
            except BaseException:
                print('FRAME27_1 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME28':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif i == 1:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif i == 2:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif i == 3:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif i == 4:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif i == 5:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif i == 6:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif i == 7:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif i == 8:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
            except:
                print('FRAME28 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME28 Update error')
        elif name == 'FRAME29':
            try:
                mprog.partdeactivate('Hole.1')
                mprog.partdeactivate("F")
            finally:
                    mprog.Update()
        elif name == 'FRAME30':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partdeactivate('SN1_25250_L')
                    mprog.partdeactivate('SN1_25250_HIJ')
                    mprog.partbodyfeatureactivate('SN1_25250_CD')
                    mprog.partbodyfeatureactivate('SN1_25250_FG')
                    mprog.activatefeature('CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.bodydeactivate('4_M14通', 0)
                elif i == 1:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partdeactivate('SN1_25250_L')
                    mprog.partdeactivate('SN1_25250_HIJ')
                    mprog.partbodyfeatureactivate('SN1_25250_CD')
                    mprog.partbodyfeatureactivate('SN1_25250_FG')
                    mprog.activatefeature('CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.bodydeactivate('4_M14通', 0)
                elif i == 2:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partdeactivate('SN1_25250_L')
                    mprog.partdeactivate('SN1_25250_HIJ')
                    mprog.partbodyfeatureactivate('SN1_25250_CD')
                    mprog.partbodyfeatureactivate('SN1_25250_FG')
                    mprog.activatefeature('CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.bodydeactivate('4_M14通', 0)
                elif i == 3:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partdeactivate('SN1_25250_L')
                    mprog.partdeactivate('SN1_25250_HIJ')
                    mprog.partbodyfeatureactivate('SN1_25250_CD')
                    mprog.partbodyfeatureactivate('SN1_25250_FG')
                    mprog.activatefeature('CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.bodydeactivate('4_M14通', 0)
                elif i == 4:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partdeactivate('SN1_25250_L')
                    mprog.partdeactivate('SN1_25250_HIJ')
                    mprog.partbodyfeatureactivate('SN1_25250_CD')
                    mprog.partbodyfeatureactivate('SN1_25250_FG')
                    mprog.activatefeature('CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.bodydeactivate('4_M14通', 0)
                elif i == 5:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partdeactivate('SN1_25250_L')
                    mprog.partdeactivate('SN1_25250_HIJ')
                    mprog.partbodyfeatureactivate('SN1_25250_CD')
                    mprog.partbodyfeatureactivate('SN1_25250_FG')
                    mprog.activatefeature('CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.bodydeactivate('4_M14通', 0)
                elif i == 6:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partdeactivate('SN1_25250_L')
                    mprog.partdeactivate('SN1_25250_HIJ')
                    mprog.partbodyfeatureactivate('SN1_25250_CD')
                    mprog.partbodyfeatureactivate('SN1_25250_FG')
                    mprog.activatefeature('CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.bodydeactivate('4_M14通', 0)
                elif i == 7:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partdeactivate('SN1_25250_L')
                    mprog.partdeactivate('SN1_25250_HIJ')
                    mprog.partbodyfeatureactivate('SN1_25250_CD')
                    mprog.partbodyfeatureactivate('SN1_25250_FG')
                    mprog.activatefeature('CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.bodydeactivate('4_M14通', 0)
                elif i == 8:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partdeactivate('SN1_25250_L')
                    mprog.partdeactivate('SN1_25250_HIJ')
                    mprog.partbodyfeatureactivate('SN1_25250_CD')
                    mprog.partbodyfeatureactivate('SN1_25250_FG')
                    mprog.activatefeature('CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.bodydeactivate('4_M14通', 0)


            except:
                print('FRAME30 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME30 Update error')
        elif name == 'FRAME31':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.bodydeactivate('FRAME_Body', 0)
                elif i == 1:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.bodydeactivate('FRAME_Body', 0)
                elif i == 2:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.bodydeactivate('FRAME_Body', 0)
            except:
                print('FRAME31 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME31 Update error')
        elif name == 'FRAME32':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('Body')
                elif i == 1:
                    mprog.partbodyfeatureactivate('Body')
                elif i == 2:
                    mprog.partbodyfeatureactivate('Body')
                elif i == 3:
                    mprog.partbodyfeatureactivate('Body')
                elif i == 4:
                    mprog.partbodyfeatureactivate('Body')
                elif i == 5:
                    mprog.partbodyfeatureactivate('Body')
                elif i == 6:
                    mprog.partbodyfeatureactivate('Body')
                elif i == 7:
                    mprog.partbodyfeatureactivate('Body')
                elif i == 8:
                    mprog.partbodyfeatureactivate('Body')
            except:
                print('FRAME32 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME32 Update error')
        elif name == "FRAME33":
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_2545_CEF')
                    mprog.partbodyfeatureactivate('SN1_2545_J')
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partdeactivate('FRAME_2545_LMN')
                    mprog.partdeactivate("O")
                elif i == 1:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_2545_CEF')
                    mprog.partbodyfeatureactivate('SN1_2545_J')
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partdeactivate('FRAME_2545_LMN')
                    mprog.partdeactivate("O")
                elif i == 2:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_2545_CEF')
                    mprog.partbodyfeatureactivate('SN1_2545_J')
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partdeactivate('FRAME_2545_LMN')
                    mprog.partdeactivate("O")
                elif i == 3:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partdeactivate("O")
            except:
                print('FRAME33 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME34':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_2545_CD')
                    mprog.partbodyfeatureactivate('SN1_2545_L')
                    mprog.partbodyfeatureactivate('SN1_2535_J')
                    mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partdeactivate('FRAME_1_2545_PQR')
                    mprog.partdeactivate("S")
                elif i == 1:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_2545_CD')
                    mprog.partbodyfeatureactivate('SN1_2545_L')
                    mprog.partbodyfeatureactivate('SN1_2535_J')
                    mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partdeactivate('FRAME_1_2545_PQR')
                    mprog.partdeactivate("S")
                elif i == 2:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_2545_CD')
                    mprog.partbodyfeatureactivate('SN1_2545_L')
                    mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partdeactivate('FRAME_1_2545_PQR')
                    mprog.partdeactivate("S")
                elif i == 3:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_60_NM')
                    mprog.partbodyfeatureactivate('SN1_60_L')
                    mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partdeactivate("S")

            except:
                print('FRAME34 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME35':
            try:
                if i == 0:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partdeactivate('SN1_25250_L')
                elif i == 1:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partdeactivate('SN1_25250_L')
                elif i == 2:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partdeactivate('SN1_25250_L')
                elif i == 3:
                    mprog.activatefeature('SN1_60_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partdeactivate('SN1_25250_L')
                elif i == 4:
                    mprog.activatefeature('SN1_80200_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partdeactivate('SN1_25250_L')
                elif i == 5:
                    mprog.activatefeature('SN1_80200_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partdeactivate('SN1_25250_L')
                elif i == 6:
                    mprog.activatefeature('SN1_80200_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partdeactivate('SN1_25250_L')
                elif i == 7:
                    mprog.activatefeature('SN1_80200_Body', 6)
                    mprog.activatefeature('Hole', 2)
                    mprog.partdeactivate('SN1_25250_L')
                elif i == 8:
                    mprog.activatefeature('SN1_250_Body', 0)
                    mprog.activatefeature('Hole', 3)
                    mprog.partdeactivate('SN1_25250_L')

            except:
                print('FRAME35 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME35 Update error')
        elif name == 'FRAME36':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('SN1_2535_Body')
                    mprog.partdeactivate('FRAME_SN1_25_Body')
                elif i == 1:
                    mprog.partbodyfeatureactivate('SN1_2535_Body')
                    mprog.partdeactivate('FRAME_SN1_3560_Body')
                elif i == 2:
                    mprog.partbodyfeatureactivate('SN1_4560_Body')
                    mprog.partdeactivate('FRAME_SN1_3560_Body')
                elif i == 3:
                    mprog.partbodyfeatureactivate('SN1_4560_Body')
                    mprog.partdeactivate('FRAME_SN1_3560_Body')
            except:
                print('FRAME36 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME36 Update error')
        elif name == 'FRAME37':  # 已更改
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partdeactivate('FRAME_Body')
                elif i == 1:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partdeactivate('FRAME_Body')
                elif i == 2:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partdeactivate('FRAME_Body')
                elif i == 3:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partdeactivate('FRAME_Body')
            except:
                print('FRAME37 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME38':
            try:
                if i == 0:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif i == 1:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif i == 2:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif i == 3:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif i == 4:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif i == 5:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif i == 6:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif i == 7:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif i == 8:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            except:
                print('FRAME38 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME38 Update error')
        elif name == 'FRAME41':
            try:
                if i == 1:  # 35
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.partbodyfeatureactivate('Chamfer.2')
                    mprog.partbodyfeatureactivate('Chamfer.3')
            except:
                print('FRAME41 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME42':
            try:
                if i == 3:  # 60
                    mprog.partbodyfeatureactivate('Chamfer60/80')
                elif i == 4:  # 80
                    mprog.partbodyfeatureactivate('Chamfer60/80')
                    mprog.partbodyfeatureactivate('ChamferOnly80')
            except:
                print('FRAME42 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME43':
            try:
                if i == 3:
                    mprog.partbodyfeatureactivate('Chamfer.1')
                elif i == 4:
                    mprog.partdeactivate('machining_Pocket')
                    mprog.bodydeactivate("Hole_1", 0)
                elif i == 3:
                    mprog.partdeactivate('machining_Pocket')
                    mprog.bodydeactivate("Hole_1", 0)
                elif i == 5:
                    mprog.partdeactivate('machining_Pocket')
                    mprog.bodydeactivate("Hole_1", 0)
                elif i == 6:
                    mprog.partdeactivate('machining_Pocket')
                    mprog.bodydeactivate("Hole_1", 0)
                elif i == 7:
                    mprog.partdeactivate('machining_Pocket')
                    mprog.bodydeactivate("Hole_1", 0)
                elif i == 8:
                    mprog.partdeactivate('machining_Pocket')
                    mprog.bodydeactivate("Hole_1", 0)
            except BaseException:
                print('FRAME43 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME45':
            try:
                if i == 3:  # 60N
                    mprog.partbodyfeatureactivate('Chamfer.G')
                    mprog.partbodyfeatureactivate('Chamfer.F')
                elif i == 4:  # 80N
                    mprog.partbodyfeatureactivate('Chamfer.G')
                    mprog.partbodyfeatureactivate('Chamfer.F')
                elif i == 5:  # 110N
                    mprog.partbodyfeatureactivate('Chamfer.G')
                    mprog.partbodyfeatureactivate('Chamfer.F')
                elif i == 6:  # 160N
                    mprog.partbodyfeatureactivate('Chamfer.G')
                    mprog.partbodyfeatureactivate('Chamfer.F')
                elif i == 7:  # 200N
                    mprog.partbodyfeatureactivate('Chamfer.G')
                    mprog.partbodyfeatureactivate('Chamfer.F')
                elif i == 8:  # 250N
                    mprog.partbodyfeatureactivate('Chamfer.G')
                    mprog.partbodyfeatureactivate('Chamfer.F')
            except:
                print('FRAME45 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME45 Update error')
        elif name == 'FRAME47':
            try:
                if i == 4:  # 80
                    mprog.partbodyfeatureactivate('上導角')
                elif i == 5:  # 110
                    mprog.partbodyfeatureactivate('上導角')
                    mprog.partbodyfeatureactivate('下導角')
                elif i == 6:  # 160
                    mprog.partbodyfeatureactivate('上導角')
                    mprog.partbodyfeatureactivate('下導角')
                elif i == 7:  # 200
                    mprog.partbodyfeatureactivate('上導角')
                    mprog.partbodyfeatureactivate('下導角')
                elif i == 8:  # 250
                    mprog.partbodyfeatureactivate('250上導角')
            except:
                print('FRAME47 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME48':
            pass
        elif name == 'FRAME49':
            try:
                if i == 4:
                    mprog.activatefeature('body', 0)
                    mprog.partdeactivate('FRAME_SN1_80250_c')
                    mprog.bodydeactivate('M12通', 0)
                elif i == 5:
                    mprog.activatefeature('body', 0)
                    mprog.partdeactivate('FRAME_SN1_80250_c')
                    mprog.bodydeactivate('M12通', 0)
                elif i == 6:
                    mprog.activatefeature('body', 0)
                    mprog.bodydeactivate('M12通', 0)
                elif i == 7:
                    mprog.activatefeature('body', 0)
                    mprog.partdeactivate('FRAME_SN1_80250_c')
                    mprog.bodydeactivate('M12通', 0)
                elif i == 8:
                    mprog.activatefeature('body', 0)
                    mprog.partdeactivate('FRAME_SN1_80250_c')
                    mprog.bodydeactivate('M12通', 0)
            except:
                print('FRAME49 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME49 Update error')
        elif name == 'FRAME50':
            try:
                if i == 4:
                    mprog.activatefeature('Body', 0)
                    mprog.partdeactivate('FRAME_SN1_c')
                    mprog.bodydeactivate('Hole_1', 0)
                elif i == 5:
                    mprog.activatefeature('Body', 0)
                    mprog.partdeactivate('FRAME_SN1_c')
                    mprog.bodydeactivate('Hole_1', 0)
                elif i == 6:
                    mprog.activatefeature('Body', 0)
                    mprog.partdeactivate('FRAME_SN1_c')
                    mprog.bodydeactivate('Hole_1', 0)
                elif i == 7:
                    mprog.activatefeature('Body', 0)
                    mprog.partdeactivate('FRAME_SN1_c')
                    mprog.bodydeactivate('Hole_1', 0)
                elif i == 8:
                    mprog.activatefeature('Body', 0)
                    mprog.partdeactivate('FRAME_SN1_c')
                    mprog.bodydeactivate('Hole_1', 0)
            except:
                print('FRAME50 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME50 Update error')
        elif name == 'FRAME51':
            # 判斷是否需要模墊加工
            try:
                if i == 8:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partbodyfeatureactivate('SN1_250_CD')
                    mprog.partbodyfeatureactivate('Chamfer_FG_R')
                    mprog.partbodyfeatureactivate('SN1_250_CD')
                    mprog.activatefeature('M10*35L', 0)
                else:
                    print('FRAME51 feature activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME51_1':
            # 判斷是否需要模墊加工
            try:
                if i == 8:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partbodyfeatureactivate('SN1_250_CD')
                    mprog.partbodyfeatureactivate('Chamfer_FG_L')
                    mprog.partbodyfeatureactivate('SN1_250_CD')
                    mprog.activatefeature('M10*35L', 0)
                else:
                    print('FRAME51_1 feature activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME52':
            try:
                if i == 5:
                    mprog.partdeactivate('machining_除料')
                    mprog.partdeactivate('machining_通孔')
                    mprog.partbodyfeatureactivate('top')
                    mprog.partbodyfeatureactivate('bottom')
                elif i == 6:
                    mprog.partdeactivate('machining_除料')
                    mprog.partdeactivate('machining_通孔')
                    mprog.partbodyfeatureactivate('top')
                    mprog.partbodyfeatureactivate('bottom')
                elif i == 7:
                    mprog.partdeactivate('machining_除料')
                    mprog.partdeactivate('machining_通孔')
                    mprog.partbodyfeatureactivate('top')
                    mprog.partbodyfeatureactivate('bottom')
                elif i == 8:
                    mprog.partdeactivate('machining_除料')
                    mprog.partdeactivate('machining_通孔')
                    mprog.partbodyfeatureactivate('top')
                    mprog.partbodyfeatureactivate('bottom')
            except:
                print('FRAME52 Parameter activate error')
            finally:
                    mprog.Update()
        elif name == 'FRAME53':
            try:
                if i == 8:
                    mprog.partbodyfeatureactivate('EdgeFillet.1')
                    mprog.partbodyfeatureactivate('Chamfer.1')
            except:
                print('FRAME53 Parameter activate error')
            mprog.Update()
        elif name == 'FRAME53_1':
            try:
                if i == 8:
                    mprog.partbodyfeatureactivate('EdgeFillet.1')
                    mprog.partbodyfeatureactivate('Chamfer.1')
            except:
                print('FRAME53_1 Parameter activate error')
            mprog.Update()
        elif name == 'FRAME54':
            try:
                if i == 6:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partbodyfeatureactivate('Chamfer_E')
                elif i == 7:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partbodyfeatureactivate('Chamfer_E')
                elif i == 8:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partbodyfeatureactivate('Chamfer_E')
            except:
                print('FRAME54 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                except:
                    print('FRAME54 Update error')
    # print出跳錯欄位
    except Exception as e:
        s = sys.exc_info()
        print('報錯行數：{}\n報錯內容：{}'.format(__file__, s[2].tb_lineno, s[1]))