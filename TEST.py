import win32com.client as win32
import main_program as mprog

hide_part_name = ['BOLSTER1', 'BOLSTER2', 'BOLSTER3', 'BALANCER_LEFT_All', 'BALANCER_RIGHT_All', 'CRANK_SHAFT_CLOCK',
                  'CLUCTH_ASSEMBLY_All', 'SLIDE_UNIT_All', 'CRANK_SHAFT.1', 'JOINT_All', 'MAIN_GEAR1', 'MAIN_GEAR2',
                  'MAIN_GEAR3', 'MAIN_GEAR4', 'JOINT1', 'GIB1', 'GIB2', 'FRAME35', 'Fixture']

for part_name in hide_part_name:
    mprog.hide_show_part(part_name, 1)

