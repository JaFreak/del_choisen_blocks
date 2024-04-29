import win32com.client
import time
from random import randint
from pythoncom import VT_R8, VT_ARRAY, VT_DISPATCH, VT_BSTR, VT_BYREF, com_error, VT_I2, VT_VARIANT, VT_I4
from time import sleep

app = win32com.client.Dispatch("AutoCAD.Application")
doc = app.ActiveDocument
sets = doc.SelectionSets
scount = sets.Count

def selsetcheck():
    while sets.Count != 0:
        try:
            for i in (sets):
                #print (i.Name)
                n=i.Name
                sets.Item(n).Delete()
        except com_error as seterror:
            if seterror.hresult == -2147417851:
                print ('ошибка на сервере')
                time.sleep (0.02) 
    print ('selsetcheck очищено')       
    doc.Application.Update()

def selsetcheck2():
    while sets.Count != 0:
        try:
            for i in (sets):
                if i.Name != 'ssels2':
                    print (i.Name)
                    n=i.Name
                    sets.Item(n).Delete()
                    print (i.Name, 'очищено')
        except com_error as seterror:
            if seterror.hresult == -2147417851 or seterror.hresult == -2147418111:
                print ('ошибка на сервере')
                time.sleep (0.1)        
    doc.Application.Update()