import pyautogui
import win32gui
import time
import os
import csv
import math
import traceback
from tkinter import Tk
import pandas as pd

PATH = os.path.dirname(os.path.realpath(__file__))
file_path = r'E:\OneDrive - Nile University\0- Publications Bank'
Data = []
Error = []

df = pd.read_csv(PATH + '/last_convert.csv',)
files = []
files.append(df["Codes Need links"])

data_files = files[0].to_numpy()

os.startfile(file_path)
time.sleep(2)


for x in data_files:
    try:
        search = pyautogui.locateOnScreen(PATH + "/search.png")
        pos = pyautogui.center(search)
        pyautogui.click(pos)
        pyautogui.click(search)
        print(str(x))
        time.sleep(5)
        pyautogui.write(str(x))
        pyautogui.press('enter')
        time.sleep(3)
        pdf = pyautogui.locateOnScreen(PATH + "/pdf_final.png")
        pos1 = pyautogui.center(pdf)
        pyautogui.click(pos1, button = "right")
        time.sleep(3)
        sharebutton = pyautogui.locateOnScreen(PATH + '/onedrive.png')
        posx = pyautogui.center(sharebutton)
        pyautogui.click(posx)
        time.sleep(5)
        share = pyautogui.locateOnScreen(PATH + '/share1.png')
        pos2 = pyautogui.center(share)
        pyautogui.click(pos2)
        time.sleep(5)
        w=win32gui
        title = w.GetWindowText(w.GetForegroundWindow())
        title = title[6:].strip('"')

        pyautogui.click(960, 540)
        back = pyautogui.locateOnScreen(PATH + '/back.png')
        pyautogui.click(pyautogui.center(back))
        time.sleep(2)

        if title not in Data:
            Data.append(title)
            with open(PATH + "\Output\Output_missing.csv", "a+", encoding="utf-8") as outputf:
                writer = csv.writer(outputf)
                writer.writerow([str(x), Tk().clipboard_get()])
    except Exception as e:
            print(traceback.format_exc())
            Error.append(str(x))
            back = pyautogui.locateOnScreen(PATH + '/back.png')
            pyautogui.click(pyautogui.center(back))


print("Done: \n")
for x in Data:
    print(x + "\n")
print("Incomplete: \n")
for y in Error:
    print(y + "\n")
