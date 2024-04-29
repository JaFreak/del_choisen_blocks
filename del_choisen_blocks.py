import SelectionSets_chek
from SelectionSets_chek import selsetcheck
import time
import win32com.client
from random import randint
from pythoncom import VT_R8, VT_ARRAY, VT_DISPATCH, VT_BSTR, VT_BYREF, com_error, VT_I2, VT_VARIANT, VT_I4
from time import sleep
import numpy as np



app = win32com.client.Dispatch("AutoCAD.Application")

doc = app.ActiveDocument
modelsp = doc.ModelSpace

sets = doc.SelectionSets
layers = doc.Layers



def inv_layer_create():
    try:
        layers.Add('0_невидимый')
        layerInv = layers.Item('0_невидимый')
        layers.Item('0_невидимый').LayerOn = False
    except:
        pass

def cpi (fcp, centrMinx, centrMiny, s, centrx, centry,):
    ucpx0 = fcp[0]
    ucpy0 = fcp[1]
    cpm1_0 = np.array([[ucpx0,ucpy0,1]])
    cpm1_1 = np.array([[1,0,0],[0,1,0],[centrMinx,centrMiny,1]])
    cpm1_2 = np.array([[s,0,0],[0,s,0],[0,0,1]])
    cpm1_3 = np.array([[1,0,0],[0,1,0],[centrx,centry,1]])
    cpm1_01 = np.dot(cpm1_0, cpm1_1)
    cpm1_12 = np.dot(cpm1_01, cpm1_2)
    cpm1_23 = np.dot(cpm1_12, cpm1_3)
    cpm1 = list(cpm1_23[0,:])
    ucpx = cpm1[0]
    ucpy = cpm1[1]
    cpi = [ucpx,ucpy,0]
    return [cpi]


def del_choisen_blocks():
    selsetcheck()
    start = int(0)
    finish = int(2)
    while start < finish:
        try:
            FilterType = win32com.client.VARIANT(VT_ARRAY|VT_I2, [  -4, 0, 8, 2, 62, -4 ]) 
            FilterData = win32com.client.VARIANT(VT_ARRAY|VT_VARIANT, ['<AND', type_for_del, lay, name_for_del, baseColor, 'AND>'])       #выбор по типу и слою, определяемым по выбранному объекту
            selset2 = 'ssels2'
            sset = app.ActiveDocument.SelectionSets.Add (selset2)
            sset.Select(5,None,None,FilterType,FilterData)
            f = sets.Item(selset2).Count
            finish = int(f)
            print ('finish',finish, 'f', f)
            for i in enumerate((sets.Item(selset2))):
                s = 1.1     #коэфициент масштабирования
                gb = i[1].GetBoundingBox()
                fcp2 = gb[1]
                fcp4 = gb[0]
                fcp1 = (fcp4[0],fcp2[1],0)
                fcp3 = (fcp2[0],fcp4[1],0)
                centr = i[1].InsertionPoint
                centrx = centr[0]
                centry = centr[1]
                centrMinx = centrx * -1
                centrMiny = centry * -1
                #вверх-влево     
                cp = cpi (fcp1, centrMinx, centrMiny, s, centrx, centry,) 
                cp1 = cp[0]

                #вверх-вправо    
                cp = cpi (fcp2, centrMinx, centrMiny, s, centrx, centry,) 
                cp2 = cp[0]              

                #вниз-вправо  
                cp = cpi (fcp3, centrMinx, centrMiny, s, centrx, centry,)
                cp3 = cp[0]

                #вниз-влево  
                cp = cpi (fcp4, centrMinx, centrMiny, s, centrx, centry,)
                cp4 = cp[0]

                pl = cp1+cp2+cp3+cp4
                plist = win32com.client.VARIANT(VT_ARRAY | VT_R8,pl)

                sset3 = app.ActiveDocument.SelectionSets.Add (selset3)
                FilterType2 = win32com.client.VARIANT(VT_ARRAY|VT_I2, [ -4, -4, 0, 8, 2, 62, -4, -4 ]) 
                FilterData2 = win32com.client.VARIANT(VT_ARRAY|VT_VARIANT, ['<NOT', '<AND', type_for_del, lay, name_for_del, baseColor, 'AND>','NOT>' ])
                sset3.SelectByPolygon (7,plist,FilterType2,FilterData2)
                print("выбрано объектов:", sset3.Count)
                if hasattr(sset3, '__iter__'):
                    if sets.Item(selset3).Count != 0:                
                        i[1].Layer = '0_невидимый'
                i[1].Color = 3
                app.ActiveDocument.SelectionSets.Item(selset3).Delete()
                doc.Application.Update()
                time.sleep (0.01)
                if i[0] >= finish-1:
                    start = finish
                    print('stop')
                    time.sleep (0.1)
                    sset4 = app.ActiveDocument.SelectionSets.Add (selset4)
                    FilterType3 = win32com.client.VARIANT(VT_ARRAY|VT_I2, [ -4, 0, 8, 2, 62, -4 ]) 
                    FilterData3 = win32com.client.VARIANT(VT_ARRAY|VT_VARIANT, ['<AND', type_for_del, lay, name_for_del, '3', 'AND>' ])
                    sset4.Select(5,None,None,FilterType3,FilterData3)
                    print ('ok')
                    if hasattr(sets.Item(selset4), '__iter__'):
                        if sets.Item(selset4).Count != 0:  
                            for bl in sets.Item(selset4):
                                bl.Color = baseColor
                    break
        except com_error as error:
            if error.hresult == -2147418111:
                time.sleep (0.1)
                doc.Application.Update()
                selsetcheck()
                print ('отрабатали отозванный вызов')
            start = 0
        except AttributeError as aterror:
            print ('отрабатываем AttributeError')
            if aterror == 'Add.Count':
                time.sleep (0.1)
                doc.Application.Update()
                selsetcheck()
                print ('отрабатали AttributeError')
                start = 0
            else:
                print(aterror)
                time.sleep (0.1)
                doc.Application.Update()
                selsetcheck()
                print ('отрабатали AttributeError Else')
                start = 0


def inv_layer_empty():
    selsetcheck()
    FilterTypeDelInv = win32com.client.VARIANT(VT_ARRAY|VT_I2, [ 8 ]) 
    FilterDataDelInv = win32com.client.VARIANT(VT_ARRAY|VT_VARIANT, ['0_невидимый']) 
    selsetDelInv = 'sselsDelInv'
    try:
        ssetDelInv = app.ActiveDocument.SelectionSets.Add (selsetDelInv)
        ssetDelInv.Select(5,None,None,FilterTypeDelInv,FilterDataDelInv)
        if hasattr(ssetDelInv, '__iter__'):
            for f in sets.Item(selsetDelInv):
                f.Delete()
                selsetcheck()
                time.sleep (0.01)
    except com_error as error:
        if error.hresult == -2147418111:
            selsetcheck()
            print ('отрабатали отозванный вызов при очистке невидимого слоя')
            time.sleep (0.01)

def inv_layer_del():
    try:
        layers.Item('0_невидимый').Delete()
        time.sleep (0.01)
        print('невидимы слой удален')
    except:
        inv_layer_empty()
        time.sleep (0.01)
        layers.Item('0_невидимый').Delete()

# ниже расставляем функции
selsetcheck()

inv_layer_create()

selset = 'ssels'

selset3 = 'ssels3'
selset4 = 'ssels4'

sset = app.ActiveDocument.SelectionSets.Add (selset)

doc.Utility.Prompt("Выберите объект эталон")


etype = ''
sset.SelectOnScreen ()

for i in sset:
    entype = i.EntityType
    lay = i.Layer
    baseColor = i.Color
    if entype == 23:
        type_for_del = 'LWPolyline'
        name_for_del = 'None'
    elif entype == 24:
        type_for_del = 'LWPolyline'
        name_for_del = 'None'
    elif entype == 8:
        type_for_del = 'Circle'
        name_for_del = 'None'
    elif entype == 22:
        type_for_del = 'Point'
        name_for_del = 'None'
    elif entype == 21:
        type_for_del = 'Mtext'
        name_for_del = 'None'
    elif entype == 32:
        type_for_del = 'Text'
        name_for_del = 'None' 
    elif entype == 32:
        type_for_del = 'Text'
        name_for_del = 'None'
    elif entype == 7:
        type_for_del = 'INSERT'
        name_for_del = i.Name


del_choisen_blocks ()
print ('del_intersectable_blocks - DONE!')

time.sleep(1)



doc.Application.Update()

doc.Utility.Prompt("Готово")







