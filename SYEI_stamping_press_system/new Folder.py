import os
import datetime


def save_Folder():
    time = datetime.datetime.now()
    print(time.day, time.hour, time.minute, time.second)
    dir = 'test' + '_' + str(time.month) + '_' + str(time.day) + '_' + str(time.hour) + '_' + str(
        time.minute) + '_' + str(time.second)
    path = 'C:\\Users\\USER\\Desktop' + '\\' + dir
    os.mkdir(path)
