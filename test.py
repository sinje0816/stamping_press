import os
import win32com.client as win32
import datetime, time
import main_program as mprog

body_draft_area_center_initX = 500
body_draft_area_center_initY = 220

def drafting_parameter_calculation(width, height, depth, scale_p):#電子型錄WHD, 比例
    scale_tmp = scale_p  # temping original scale
    scale = 1 / scale_p  # proportion convert to ratio
    drafting_area_centerX = body_draft_area_center_initX
    drafting_area_centerY = body_draft_area_center_initY
    w_scale = width * scale  # width after scaling
    h_scale = height * scale  # height after scaling
    d_scale = depth * scale
    drafting_area_X_range = h_scale * 3 + d_scale * 2 + body_draft_area_center_initX * 4
    drafting_area_Y_range = d_scale * 2 + w_scale + body_draft_area_center_initY * 2
    drafting_area_extremum = [drafting_area_centerX - drafting_area_X_range / 2,
                              drafting_area_centerX + drafting_area_X_range / 2,
                              drafting_area_centerY - drafting_area_Y_range / 2,
                              drafting_area_centerY + drafting_area_Y_range / 2]  # drafting area extremums, format:[X-min,X-max,Y-min,Y-max]