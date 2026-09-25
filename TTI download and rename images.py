#!/usr/bin/env python
# coding: utf-8

# In[2]:


import tkinter as tk
import requests
import os
import urllib
import xlrd
import xlsxwriter
#from osgeo import ogr
#from osgeo import gdal
#from random import random

def Download ():

    # open shapefile
    dato1=testoinput.get()
    dato2=testoinput2.get()
    #dato3=int(testoinput3.get())
    #dato4=int(testoinput4.get())
    dato = dato1+dato2
    #ds = ogr.Open("C:/Documents/TomTom/Comunicazioni/QTM SPL\Progetti 2021/IND/06-10-2021/Prayagraj/IND-TBTRESTRMSFT-Prayagraj.shp", 0)
    
    workbook = xlrd.open_workbook(dato)

    # Open the worksheet
    worksheet = workbook.sheet_by_index(0)
    
    
    
    #ds = ogr.Open(dato, 0)
    #layer = ds.GetLayer()
               
    # get extent
    #number = layer.GetFeatureCount()
    
    row_count = worksheet.nrows
    print("Total number of JPGs to be downloadeded/remamed: "+str(row_count-1))
    
    for i in range(1, row_count):
    #        R= int((random()*dato4*991)/1000)
             A1= worksheet.cell_value(i, 2)
             urllib.request.urlretrieve(A1, dato1+worksheet.cell_value(i, 3))
             print("Downloading: "+A1)
    #        A1='https://panoserve.sso.maps.az.tt3.com/mml/'+C+'/mono/camera_1/'+str(R)+'/image'
    #        try:
    #            urllib.request.urlretrieve(A1, B+C+'-'+str(R)+'.jpg')
    #            ciccio = C+'-'+str(R)+'.jpg'+'    >>>> OK'
    #            print(ciccio)
    #            OK=OK+1
    #        except Exception:
    #            ciccio = C+'-'+str(R)+'.jpg'+'    **** NOT FOUND ****'
    #            print(ciccio)
    #            continue
    
    
        
    # get a single feature
    #feature = layer.GetFeature(1)

    #B = "C:/Documents/TomTom/Comunicazioni/QTM SPL/Progetti 2021/IND/06-10-2021/Sessions JPGs/"
    #B=dato1+'Sessions JPGs\\'
    #f = B.replace('\\','/')
    #if not os.path.exists(B):os.mkdir(B)
    #for z in os.listdir(f):
    #      os.remove(os.path.join(f, z))

    #
    #namefile = B+'MomaLight sessions download followup '+dato2+'.xlsx'
    #workbook = xlsxwriter.Workbook(namefile)
    #worksheet = workbook.add_worksheet()
    #

    #row = 0
    #col = 0
    #worksheet.write(row, col, 'Session name')
    #col = 1
    #worksheet.write(row, col, 'N° of images downloaded')
    
    
    #else:
    #    os.rmdir(B)
    #    os.mkdir(B)
    
    # loop over all features in a layer
    #J=1
    #row=1
    #for feature in layer:
    #    C= feature.GetField("SessionNam")
    #    i=0
    #    OK=0
    #    worksheet.write(row, 0, C)
                
    #    for i in range(dato3):
    #        R= int((random()*dato4*991)/1000)
    #        A1= 'https://panoserve.sso.maps.az.tt3.com/mml/'+C+'/mono/camera_1/'+str(R)+'/image'
    #        try:
    #            urllib.request.urlretrieve(A1, B+C+'-'+str(R)+'.jpg')
    #            ciccio = C+'-'+str(R)+'.jpg'+'    >>>> OK'
    #            print(ciccio)
    #            OK=OK+1
    #        except Exception:
    #            ciccio = C+'-'+str(R)+'.jpg'+'    **** NOT FOUND ****'
    #            print(ciccio)
    #            continue
         
        
        
    #    worksheet.write(row, 1, str(OK))
    #    
    #    perc=int(J/number*100)
    #    J=J+1
    #    row=row+1
    #    print(str(perc)+'% completed')  
    #workbook.close()
    print("End of downloading")
    
 
 
finestra = tk.Tk()
finestra.geometry("700x500")
finestra.title("JPG images downloader for TTI v1.0 @2022")
finestra.configure(background="white")

Label1= tk.Label(finestra, text= 'This tool downloads and renames TTI JPGs sourced on XLSX file', fg='green', font=('helvetica', 12, 'bold'), background="white")
Label1.grid(row=0, column=0, sticky="N")
Label2= tk.Label(finestra, text= 'Please insert below the XLSX source file path:', fg='green', font=('helvetica', 12, 'bold'), background="white")
Label2.grid(row=1, column=0, sticky="N")
Label3= tk.Label(finestra, text= 'e.g.: C:\Documents\TomTom\TTI\\', fg='green', font=('helvetica', 12, 'bold'), background="white")
Label3.grid(row=2, column=0, sticky="N", padx="100")

Label4= tk.Label(finestra, text='Please insert below XLSX file name (e.g.: TII US.xlsx):', fg='green', font=('helvetica', 12, 'bold'), background="white")
Label4.grid(row=4, column=0, sticky="N", ipadx="100")

#Label5= tk.Label(finestra, text='Please insert the number of images to be downloaded for each session:', fg='green', font=('helvetica', 12, 'bold'), background="white")
#Label5.grid(row=6, column=0, sticky="N", ipadx="10")

#Label6= tk.Label(finestra, text='Please insert the average total number of frames composing the sessions:', fg='green', font=('helvetica', 12, 'bold'), background="white")
#Label6.grid(row=8, column=0, sticky="N", ipadx="10")


testoinput= tk.Entry(finestra)
testoinput.grid(row=3, column=0, sticky="WE", padx="10")

testoinput2= tk.Entry(finestra)
testoinput2.grid(row=5, column=0, sticky="WE", padx="10")                 

#testoinput3= tk.Entry(finestra)
#testoinput3.grid(row=7, column=0, sticky="WE", padx="10")                 

#testoinput4= tk.Entry(finestra)
#testoinput4.grid(row=9, column=0, sticky="WE", padx="10")                 


bottone1= tk.Button(finestra, text='Click here to start downloading',command=Download, bg='red',fg='white')
bottone1.grid(row=12, column=0, sticky="N", ipadx="100")



finestra.mainloop()    


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




