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
                    with open(PATH + "\Output\Output.csv", "a+", encoding="utf-8") as outputf:
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
folders_names = ['04_WINC']
Done = []
file_counter = 0

time.sleep(4)
check_pdfs(Done, 1)