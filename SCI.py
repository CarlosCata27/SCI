import pandas as pd
import qrcode
import os
import shutil
import numpy as np
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from datetime import datetime
from PIL import Image, ImageDraw
from tkinter import filedialog as fd
from tkinter import*

#Conexion to Firebase Storage
import pyrebase

#Interface tkinter
root = Tk()

root.title("Sistema de credenciales")
root.geometry("720x720")
root.iconbitmap("SCI.ico")
root.config(bg="#900C3F",cursor="arrow") #color y tipo de cursor

#DATAFRAME
Nombre = StringVar()
nombre = StringVar()
Nombre2 = StringVar()
Apellido = StringVar()
Apellido2 = StringVar()
Sede = StringVar()
Folio = StringVar()
Registro = StringVar()
Vigencia = StringVar()
Empleado = StringVar()
#TITLE
texto0 = Label(root,text="Sistema de credenciales para la UTEyCV", bd=4,font="arial 16",bg="#900C3F",fg="#fff")
texto0.place(x=20,y=50)
#FORM
texto1 = Label(root,text="Nombre:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
texto1.place(x=20,y=150)
input1 = Entry(root,textvariable=Nombre,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
input1.place(x=150,y=150)
texto2 = Label(root,text="Nombre 2:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
texto2.place(x=20,y=200)
input2 = Entry(root,textvariable=Nombre2,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
input2.place(x=150,y=200)
texto3 = Label(root,text="Apellido:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
texto3.place(x=20,y=250)
input3 = Entry(root,textvariable=Apellido,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
input3.place(x=150,y=250)
texto4 = Label(root,text="Apellido 2:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
texto4.place(x=20,y=300)
input4 = Entry(root,textvariable=Apellido2,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
input4.place(x=150,y=300)
texto5 = Label(root,text="Sede:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
texto5.place(x=20,y=350)
input5 = Entry(root,textvariable=Sede,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
input5.place(x=150,y=350)
texto6 = Label(root,text="Folio:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
texto6.place(x=20,y=400)
input6 = Entry(root,textvariable=Folio,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
input6.place(x=150,y=400)
texto7 = Label(root,text="Registro:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
texto7.place(x=20,y=450)
input7 = Entry(root,textvariable=Registro,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
input7.place(x=150,y=450)
texto8 = Label(root,text="Vigencia:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
texto8.place(x=20,y=500)
input8 = Entry(root,textvariable=Vigencia,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
input8.place(x=150,y=500)
texto9 = Label(root,text="Empleado:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
texto9.place(x=20,y=550)
input9 = Entry(root,textvariable=Empleado,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
input9.place(x=150,y=550)
#SAVED FUNCTION
def save():
    datos = [
        Nombre.get(),
        Nombre2.get(),
        Apellido.get(),
        Apellido2.get(),
        Sede.get(),
        Folio.get(),
        Registro.get(),
        Vigencia.get(),
        Empleado.get()
    ]
    print(datos)
    DFDatosRaw = pd.DataFrame(datos)
    print(DFDatosRaw)
    DFDatosRaw2 = DFDatosRaw.transpose()
    print(DFDatosRaw2)

    saved = Label(root,text="Registro guardado!",bg="yellow")
    saved.pack()
#CLEAR INPUTS FUNCTION
# def clear():
#     input1.delete("1.0","end")

buttonName = Button(root,text="Guardar Datos",bd=3,command = save,bg="#94FF40",font="arial 12",cursor="plus")
buttonName.place(x=20,y=600)

#Cierre de la ventana 
root.mainloop()


firebaseConfig = {
    "apiKey": "AIzaSyA_QQZNU0vBCJhhmedvg-X1xDnBi2w44BA",
    "authDomain": "credenciales-uteycv.firebaseapp.com",
    "projectId": "credenciales-uteycv",
    "storageBucket": "credenciales-uteycv.appspot.com",
    "messagingSenderId": "722175152275",
    "appId": "1:722175152275:web:8ac288aa49373e7259bd66",
    "measurementId": "G-Q6BCTG1MEB"}

firebase = pyrebase.initialize_app(firebaseConfig)

#Config storage from firebase
storage = firebase.storage()

#Get date in a friendly format
def FechaActualCompleta(date):
    months = ("ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
    day = date.day
    month = months[date.month - 1]
    year = date.year
    messsage = "{} DE {} DEL {}".format(day, month, year)
    return messsage

def recuperarExcel():
        nombrearch=fd.askopenfilename(initialdir = "./",title = "Seleccione archivo",filetypes = (("Libro de Excel","*.xlsx"),("Libro de Excel 97 a Excel 2003","*.xls"),("Todos los archivos","*.*")))
        if nombrearch!='':
            if(nombrearch.endswith('.xlsx') or nombrearch.endswith('.xls')):
                sys.exit(1)

# Load excel file to analize, you can select it
#archivo = pd.ExcelFile('Credenciales.xlsx')
archivo = fd.askopenfilename(initialdir = "./",title = "Seleccione archivo",filetypes = (("Todos los archivos","*.*"),("Libro de Excel","*.xlsx")))
print("Opening...")
if archivo!='':
    if(not(archivo.endswith('.xlsx') or archivo.endswith('.xls'))):
        print("opened")
        sys.exit(1)
    else:
        print("No opened")

archivo.getCells().insertRows(2,1)

df = pd.read_excel(archivo,0,usecols='A:I',skiprows=range(1))

#Create DataFrame with our parameters
DFDatos = pd.DataFrame({
    'Nombre':df.iloc[:,0],
    'Nombre 2':df.iloc[:,1],
    'Apellido Paterno':df.iloc[:,2],
    'Apellido Materno':df.iloc[:,3],
    'Sede':df.iloc[:,4],
    'Folio':df.iloc[:,5],
    'Registro Postgrado':df.iloc[:,6],
    'Vigencia':df.iloc[:,7],
    'Numero Empleado':df.iloc[:,8],
})

#Data cleaning and normalization
DFDatos['Nombre'] = DFDatos['Nombre'].replace(pd.NA,0)
DFDatos = DFDatos[DFDatos['Nombre'] != 0]

DFDatos['Nombre 2'] = DFDatos['Nombre 2'].replace(pd.NA,'')

DFDatos['Nombre'] = DFDatos['Nombre'].map(lambda x: x.rstrip(' '))
DFDatos['Nombre 2'] = DFDatos['Nombre 2'].map(lambda x: x.lstrip(' '))
DFDatos['Apellido Paterno'] = DFDatos['Apellido Paterno'].map(lambda x: x.rstrip(' '))
DFDatos['Apellido Materno'] = DFDatos['Apellido Materno'].map(lambda x: x.lstrip(' '))

DFDatos['Folio'] = DFDatos['Folio'].map(lambda x: x.lstrip('FOLIO:'))

DFDatos['Numero Empleado'] = DFDatos['Numero Empleado'].astype(int)

#Creation and validation directorys to save final files 
shutil.rmtree('CodigosQR', ignore_errors=True)
shutil.rmtree('PDFs', ignore_errors=True)
shutil.rmtree('Credenciales', ignore_errors=True)
shutil.rmtree('Recortes', ignore_errors=True)
os.mkdir('PDFs')
os.mkdir('Credenciales')
os.mkdir('CodigosQR')
os.mkdir('Recortes')

#Inicializate how to create each pdf and ID
#for row in range(len(DFDatos.index)):
for row in range(1):
    #Get name foreach file
    nombrePDF = str(DFDatos.iloc[row,5])
    #Create PDF file
    myCanvas = canvas.Canvas(f'./PDFs/{nombrePDF}.pdf', pagesize=letter)
    width, height = letter  #612 792
    
    #Get UPIITA's logo online
    Upiita = 'https://firebasestorage.googleapis.com/v0/b/credenciales-uteycv.appspot.com/o/LogoUPIITA.png?alt=media&token=36f618f6-45e0-4f7f-a26b-355759536a3a'

    #UPIITA's logo localization in PDF file
    Fondo = myCanvas.drawImage(Upiita,x=width/6,y=height/4,height=450,width=450,mask='auto')

    #Phrases with Helveltica font
    myCanvas.setFont("Helvetica", 50)
    Nombre = ''.join([str(DFDatos.iloc[row,0]),' ', str(DFDatos.iloc[row,1])])
    Apellidos = ''.join([str(DFDatos.iloc[row,2]),' ', str(DFDatos.iloc[row,3])])
    Empleado = ''.join(['Empleado: ',str(DFDatos.iloc[row,8])])

    myCanvas.drawCentredString(width/2,700,Nombre)
    myCanvas.drawCentredString(width/2,650,Apellidos)
    myCanvas.drawCentredString(width/2,590,Empleado)

    myCanvas.setFont("Helvetica", 30)
    myCanvas.drawCentredString(width/2,450,'VIGENCIA: ')
    myCanvas.drawCentredString(width/2,350,'FECHA DE EXPEDICIÓN')
    myCanvas.drawCentredString(width/2,250,'LUGAR DE ORIGEN')

    #Phrases with Helveltica-Bold font
    myCanvas.setFont("Helvetica-Bold", 30)
    Folio = 'FOLIO: '+str(DFDatos.iloc[row,5])

    myCanvas.drawCentredString(width/2,75,Folio)

    #Get date from expedition to add
    Now = datetime.now()
    Fechahoy = FechaActualCompleta(Now)

    myCanvas.drawCentredString(width/2,400,FechaActualCompleta(DFDatos.iloc[row,7]))
    myCanvas.drawCentredString(width/2,300,Fechahoy)
    myCanvas.drawCentredString(width/2,200,'UTEyCV')

    #Saving PDF file 
    myCanvas.save()

    #Saving our PDF file in Firebase cloud
    file = f'PDFs/{nombrePDF}.pdf'
    cloudfilename = f'PDFs/{nombrePDF}.pdf'

    storage.child(cloudfilename).put(file)

    #Creating QR Code from url firebase
    url_PDF = storage.child(cloudfilename).get_url(None)
    CodigoQR = qrcode.make(url_PDF)
    CodigoQR.save(f'CodigosQR/{nombrePDF}.png')

    # PDF file with ID

    #Set image background
    myCanvas = canvas.Canvas(f'Credenciales/{nombrePDF}.pdf', pagesize=(3322,2214))
    url = 'https://firebasestorage.googleapis.com/v0/b/credenciales-uteycv.appspot.com/o/1.png?alt=media&token=45f62ed5-23d7-47f2-9de5-67aab19e9d1c'
    Base = myCanvas.drawImage(url,0,0,height=2214,width=3322,mask='auto')

    #Image selection from Fotos ID
    if(os.path.exists(f'Fotos/{nombrePDF}.jpg')):
        img = Image.open(f'./Fotos/{nombrePDF}.jpg').convert("RGB")
        cut_img = img.crop((250,450,950,1150))
        npImage=np.array(cut_img)
        h,w=cut_img.size
        alpha = Image.new('L', cut_img.size,0)
        draw = ImageDraw.Draw(alpha)
        draw.pieslice([0,0,h,w],0,360,fill=255)
        npAlpha=np.array(alpha)
        npImage=np.dstack((npImage,npAlpha))
        Image.fromarray(npImage).save(f'Recortes/{nombrePDF}.png')
    else:
        pass
    
    #If image doesn't exist, dont try to paste in ID pdf file
    if(os.path.exists(f'Recortes/{nombrePDF}.png')):
        pathCrop = f'Recortes/{nombrePDF}.png'
        Base = myCanvas.drawImage(pathCrop,200,370,height=1055,width=1055,mask='auto')
    else:
        pass
    #Phrases with Helveltica-Bold font
    myCanvas.setFont("Helvetica-Bold", 200)
    myCanvas.drawString(1300,1100,Nombre)
    myCanvas.drawString(1300,850,Apellidos)
    
    myCanvas.setFont("Helvetica-Bold", 150)
    myCanvas.drawString(1300,550,DFDatos.iloc[row,4])

    #Phrases with Helveltica font
    myCanvas.setFont("Helvetica", 100)
    myCanvas.drawString(1300,400,Empleado)
    
    #Date in format m/Y
    temp = DFDatos.iloc[row,7].strftime('%m/%Y')
    myCanvas.drawString(2400,125,'Vigencia: '+str(temp))

    #Save page 1
    myCanvas.showPage()
    #Get QR Code saved in a local dir
    pathQR = f'CodigosQR/{nombrePDF}.png'
    #Set image background
    url = 'https://firebasestorage.googleapis.com/v0/b/credenciales-uteycv.appspot.com/o/2.png?alt=media&token=d8badf27-6bd9-4b0a-b541-6a3b974ae7e8'
    Base = myCanvas.drawImage(url,0,0,height=2214,width=3322,mask='auto')
    #Set QR Code in the specified position
    Base = myCanvas.drawImage(pathQR,300,520,height=1050,width=1050,mask='auto')

    #Phrases with Helveltica font
    myCanvas.setFont("Helvetica", 60)

    myCanvas.drawCentredString(825,500,f'{Nombre} {Apellidos}')
    
    #Saving the ID as a PDF file
    myCanvas.save()

#Delete CodigosQR and crop images dir, we dont need it
shutil.rmtree('./CodigosQR', ignore_errors=True)
shutil.rmtree('Recortes', ignore_errors=True)