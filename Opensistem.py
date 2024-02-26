import os
import subprocess
import pyperclip
import pyautogui
from dotenv import load_dotenv
from storagefunctions import (pressingKey,make_window_visible)
from time import sleep

# Habre el CRM Onyx App
def openCRM():    
    sleep(1)
    pressingKey('win')
    sleep(1)
    pyautogui.write('Nuevo CRM')    
    sleep(1)
    pressingKey('enter')
    make_window_visible('Control de Acceso Unificado CRM')
    
# Conecta el CRM Onyx App    
def connectToCRM():    
    crm_select_app = None
    crm_dashboard = None       
    
    pressingKey('tab',2)
    sleep(1)
    pyautogui.write(os.getenv('USER_CRM'))
    sleep(1)
    pressingKey('tab') 
    pyperclip.copy(os.getenv('PASS_CRM'))
    sleep(0.5)
    pyautogui.hotkey('ctrl','v')
    sleep(1)
    pressingKey('tab')
    sleep(2)
    pressingKey('enter')
    
    while crm_select_app is None:
        print("buscando ventana emergente!")
        crm_select_app = pyautogui.locateOnScreen('C:/Automatizacion/assets/select_app.png', grayscale = True,confidence=0.9)
    print("CRM select box is present!")
    
    pressingKey('enter')
    pressingKey('tab',4)
    pressingKey('enter')
    
    while crm_dashboard is None:
        crm_dashboard = pyautogui.locateOnScreen('C:/Automatizacion/assets/crm_dashboard.png', grayscale = True,confidence=0.85)    
    print("CRM Dashboard GUI is present!")    
    