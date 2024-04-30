
import pandas as pd # for reading excel file
import pywhatkit # for sending whatsapp message
import pyautogui
from tkinter import *

total_students = 50 # For the main loop 
df = pd.read_excel("message_main.xlsx") # Name of the excel sheet


for i in range(total_students):

    # Excel Location input 
    name = df.iloc[i,1] 
    message = df.iloc[i,4] 
    number = df.iloc[i,2]
    cell_number = f'+91{number}' # Concatinating with the country code

    # Sending whatsapp message function    
    pywhatkit.sendwhatmsg_instantly(cell_number,message,15,True,5)

    win = Tk()
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()

    pyautogui.moveTo(screen_width*0.694, screen_height*0.964)
    pyautogui.click()
    pyautogui.press('enter')



