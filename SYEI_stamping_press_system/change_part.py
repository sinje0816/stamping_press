import main_program as mprog
import excel_parameter_change as epc


def change(name, i):
    if name == 'FRAME12':
        excel = epc.ExcelOp('FRAME12')

        try:
            excel.part_parameter('FRAME12', i)
            FRAME12_parameter_name, FRAME12_parameter_value = excel.part_parameter('FRAME12', i)
            print('FRAME12 Parameter change success')
        except:
            print('FRAME12 Parameter change error')
        try:
            if i == 0:  # 25N
                mprog.partbodyfeatureactivate('Chamfer.1', 0)
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)

            elif i == 1:  # 35N
                mprog.partbodyfeatureactivate('Chamfer.1', 0)
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)

            elif i == 2:  # 45N
                mprog.partbodyfeatureactivate('Chamfer.1', 0)
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)

            elif i == 3:  # 60N
                mprog.partbodyfeatureactivate('Chamfer.1', 0)
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)

            elif i == 4:  # 80N
                mprog.partbodyfeatureactivate('Chamfer.1', 0)
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)

            elif i == 5:  # 110N
                mprog.partbodyfeatureactivate('Chamfer.1', 0)
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)

            elif i == 6:  # 160N
                mprog.partbodyfeatureactivate('C8_160', 0)
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)

            elif i == 7:  # 200N
                mprog.partbodyfeatureactivate('Chamfer.1', 0)
                mprog.activatefeature('Pipe_clamp', 0)
                mprog.activatefeature('drain_hole', 0)

            elif i == 8:  # 250N
                mprog.partbodyfeatureactivate('Chamfer.1', 0)
                mprog.activatefeature('Pipe_clamp', 0)
                mprog.activatefeature('drain_hole', 0)
        except:
            print('FRAME12 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME12 Update success')
            except:
                print('FRAME12 Update error')
#######################################################################################################
    if name == 'FRAME14':
        excel = epc.ExcelOp('FRAME14')

        try:
            excel.part_parameter('FRAME14', i)
            FRAME14_parameter_name, FRAME14_parameter_value = excel.part_parameter('FRAME14', i)
            print('FRAME14 Parameter change success')
        except:
            print('FRAME14 Parameter change error')
        try:
            if i == 0:  # 25N
                mprog.activatefeature('inside_hole', 0)

            elif i == 1:  # 35N
                mprog.activatefeature('inside_hole', 0)

            elif i == 2:  # 45N
                mprog.activatefeature('inside_hole', 0)

            elif i == 3:  # 60N
                mprog.activatefeature('inside_hole', 0)

            elif i == 4:  # 80N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('outside_hole_E', 0)

            elif i == 5:  # 110N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)

            elif i == 6:  # 160N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)

            elif i == 7:  # 200N
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('outside_hole_E', 0)

            elif i == 8:  # 250N
                mprog.partbodyfeatureactivate('250N', 0)
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)
        except:
            print('FRAME14 Parameter activate error')
        finally:
            try:
                # mprog.Update()
                print('FRAME14 Update success')
            except:
                print('FRAME14 Update error')

        #######################################################################################################
    if name == 'FRAME15':
        excel = epc.ExcelOp('FRAME15')

        try:
            excel.part_parameter('FRAME15', i)
            FRAME15_parameter_name, FRAME15_parameter_value = excel.part_parameter('FRAME15', i)
            print('FRAME15 Parameter change success')
        except:
            print('FRAME15 Parameter change error')
        try:
            if i == 0:  # 25N
                    mprog.activatefeature('Hole', 0)

            elif i == 1:  # 35N
                    mprog.activatefeature('Hole', 0)

            elif i == 2:  # 45N
                    mprog.activatefeature('Hole', 0)

            elif i == 3:  # 60N
                    mprog.activatefeature('Hole', 0)

            elif i == 4:  # 80N
                    mprog.activatefeature('Hole', 0)

            elif i == 5:  # 110N
                    mprog.activatefeature('Hole', 0)

            elif i == 6:  # 160N
                    mprog.activatefeature('Hole', 0)

            elif i == 7:  # 200N
                    mprog.partbodyfeatureactivate('250N')
                    mprog.activatefeature('Hole', 0)

            elif i == 8:  # 250N
                    mprog.partbodyfeatureactivate('250N')
                    mprog.activatefeature('Hole', 0)
        except:
            print('FRAME15 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME15 Update success')
            except:
                print('FRAME15 Update error')

###############################################################################
    if name == 'FRAME17':
        excel = epc.ExcelOp('FRAME17')

        try:
            excel.part_parameter('FRAME17', i)
            FRAME17_parameter_name, FRAME17_parameter_value = excel.part_parameter('FRAME17', i)
            print('FRAME17 Parameter change success')
        except:
            print('FRAME17 Parameter change error')
        try:
            if i == 0:  # 25N
                mprog.activatefeature('inside_hole', 0)

            elif i == 1:  # 35N
                mprog.activatefeature('inside_hole', 0)

            elif i == 2:  # 45N
                mprog.activatefeature('inside_hole', 0)

            elif i == 3:  # 60N
                mprog.activatefeature('inside_hole', 0)

            elif i == 4:  # 80N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('outside_hole_E', 0)

            elif i == 5:  # 110N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)

            elif i == 6:  # 160N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)

            elif i == 7:  # 200N
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('outside_hole_E', 0)

            elif i == 8:  # 250N
                mprog.partbodyfeatureactivate('250N', 0)
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)
        except:
            print('FRAME17 Parameter activate error')
        finally:
            try:
                # mprog.Update()
                print('FRAME17 Update success')
            except:
                print('FRAME17 Update error')


##############################################################
    if name == 'FRAME27':
        excel = epc.ExcelOp('FRAME27')

        try:
            excel.part_parameter('FRAME27', i)
            FRAME27_parameter_name, FRAME27_parameter_value = excel.part_parameter('FRAME27', i)
            print('FRAME27 Parameter change success')
        except:
            print('FRAME27 Parameter change error')
        try:
            if i == 0:  # 25N
                mprog.partbodyfeatureactivate('Pocket.1', 0)
                mprog.activatefeature('Hole', 0)
                mprog.activatefeature('Piping', 0)

            elif i == 1:  # 35N
                mprog.partbodyfeatureactivate('35N', 0)
                mprog.activatefeature('Hole', 0)
                mprog.activatefeature('Piping', 0)

            elif i == 2:  # 45N
                mprog.partbodyfeatureactivate('Pocket.1', 0)
                mprog.activatefeature('Hole', 0)
                mprog.activatefeature('Piping', 0)

            elif i == 3:  # 60N
                mprog.partbodyfeatureactivate('Pocket.1', 0)
                mprog.activatefeature('Hole', 0)
                mprog.activatefeature('Piping', 0)

        except:
            print('FRAME27 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME27 Update success')
            except:
                print('FRAME27 Update error')

        ###############################################################
    if name == 'FRAME45':
        excel = epc.ExcelOp('FRAME45')

        try:
            excel.part_parameter('FRAME45', i)
            FRAME45_parameter_name, FRAME45_parameter_value = excel.part_parameter('FRAME45', i)
            print('FRAME45 Parameter change success')
        except:
            print('FRAME45 Parameter change error')
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
                print('FRAME45 Update success')
            except:
                print('FRAME45 Update error')

###############################################################
    if name == 'FRAME51':
        excel = epc.ExcelOp('FRAME51')

        try:
            excel.part_parameter('FRAME51', i)
            FRAME51_parameter_name, FRAME51_parameter_value = excel.part_parameter('FRAME51', i)
            print('FRAME51 Parameter change success')
        except:
            print('FRAME51 Parameter change error')
        try:
            if i == 8:  # 250N
                mprog.partbodyfeatureactivate('Chamfer.1', 0)
                mprog.activatefeature('Hole', 0)

        except:
            print('FRAME51 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME51 Update success')
            except:
                print('FRAME51 Update error')
###############################################################################################
    if name == 'FRAME52':
        excel = epc.ExcelOp('FRAME52')

        try:
            excel.part_parameter('FRAME52', i)
            FRAME52_parameter_name, FRAME52_parameter_value = excel.part_parameter('FRAME52', i)
            print('FRAME52 Parameter change success')
        except:
            print('FRAME52 Parameter change error')
        try:
            if i == 5:  # 110N
                mprog.partbodyfeatureactivate('top', 0)
                mprog.partbodyfeatureactivate('bottom', 0)

            if i == 6:  # 160N
                mprog.partbodyfeatureactivate('top', 0)
                mprog.partbodyfeatureactivate('bottom', 0)

            if i == 7:  # 200N
                mprog.partbodyfeatureactivate('top', 0)
                mprog.partbodyfeatureactivate('bottom', 0)

            if i == 8:  # 250N
                mprog.partbodyfeatureactivate('top', 0)
                mprog.partbodyfeatureactivate('bottom', 0)

        except:
            print('FRAME52 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME52 Update success')
            except:
                print('FRAME52 Update error')


change('FRAME45', 6)