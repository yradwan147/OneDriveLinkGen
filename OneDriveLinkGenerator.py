import pyautogui
import win32gui
import time
import os
import csv
import math
import traceback
from tkinter import Tk

def check_pdfs(Data, x = None):
    if(x):
        _range = x
    else:
        _range = math.ceil(len(os.listdir(file_path + '\\' + folders_names[file_counter]))/30)
    for i in range(_range):
        try:
            time.sleep(2)
            finds = list(pyautogui.locateAllOnScreen(PATH + '/icon1.png'))
            for x in finds:
                pos = pyautogui.center(x)
                pyautogui.click(pos, button = "right")
                time.sleep(3)
                sharebutton = pyautogui.locateOnScreen(PATH + '/onedrive.png')
                pos1 = pyautogui.center(sharebutton)
                pyautogui.click(pos1)
                time.sleep(3)
                openPNG = pyautogui.locateOnScreen(PATH + '/open.png')
                pos_open = pyautogui.center(openPNG)
                pyautogui.click(pos_open)
                time.sleep(1)
                NU_only = pyautogui.locateOnScreen(PATH + '/NU_only.png')
                pos_NU = pyautogui.center(NU_only)
                pyautogui.click(pos_NU)
                time.sleep(1)
                editing = pyautogui.locateOnScreen(PATH + '/editing.png')
                pos_editing = pyautogui.center(editing)
                pyautogui.click(pos_editing)
                time.sleep(1)
                download = pyautogui.locateOnScreen(PATH + '/download.png')
                pos_download = pyautogui.center(download)
                pyautogui.click(pos_download)
                time.sleep(1)
                apply_button = pyautogui.locateOnScreen(PATH + '/finish.png')
                pos_apply = pyautogui.center(apply_button)
                pyautogui.click(pos_apply)
                time.sleep(2)
                share = pyautogui.locateOnScreen(PATH + '/share1.png')
                pos2 = pyautogui.center(share)
                pyautogui.click(pos2)
                time.sleep(3)
                w=win32gui
                title = w.GetWindowText(w.GetForegroundWindow())
                title = title[6:].strip('"')

                pyautogui.click(960, 540)

                if title not in Data:
                    Data.append(title)
                    with open(PATH + "\Output\Output_PrivacyEnabled.csv", "a+", encoding="utf-8") as outputf:
                        writer = csv.writer(outputf)
                        writer.writerow([folders_names[file_counter], title, Tk().clipboard_get()])

            pyautogui.scroll(-2000)
            print("scrolled")

        except Exception as e:
            print(traceback.format_exc())
            pass

##FILE_NO = input("Number of files: ")
PATH = os.path.dirname(os.path.realpath(__file__))
file_path = r'E:\OneDrive - Nile University\0- Publications Bank'

file_counter = 0
Done = []
folders_names = ['01_CIS', '02_CNT', '03_NISC', '04_WINC', '05_EAS', '06_NU', '07_BUS', '08_CEM', '09_ITCS', '10_MOT', '11_SESC']

os.startfile(file_path)
time.sleep(2)
folders = list(pyautogui.locateAllOnScreen(PATH + '/folder.png'))
folders = folders[1:12]
for v in folders:
    file_pos = pyautogui.center(v)
    pyautogui.click(file_pos, clicks=2)
    pyautogui.moveTo(1080, 540)
    time.sleep(2)
    folders_inner = list(pyautogui.locateAllOnScreen(PATH + '/folder2.png'))
    print(folders_inner)
    for d in folders_inner:
        file_inner_pos = pyautogui.center(d)
        pyautogui.click(file_inner_pos, clicks=2)
        pyautogui.moveTo(1080, 540)
        time.sleep(2)
        check_pdfs(Done, 1)
        back = pyautogui.locateOnScreen(PATH + '/back.png')
        pyautogui.click(pyautogui.center(back))
    time.sleep(2)
    check_pdfs(Done)
    time.sleep(2)
    back = pyautogui.locateOnScreen(PATH + '/back.png')
    pyautogui.click(pyautogui.center(back))

    file_counter += 1







