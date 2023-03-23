import pandas as pd
import qrcode
import os
import shutil
import numpy as np
import sys

#importar UI
from tkinter import *
from tkinter import filedialog as fd
from tkinter import ttk
from tkcalendar import DateEntry
import customtkinter as ctk

#importar PDF gen
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

from datetime import datetime

#importar Pillow
from PIL import Image, ImageDraw
from PIL import Image

#Conexion to Firebase Storage
from Pyrebase.firebaseStorage import FirebaseStorage
#Config storage from firebase
storage = FirebaseStorage()

#Define global DataFrame
DFDatos = pd.DataFrame(columns=['Nombre','Nombre 2','Apellido Paterno','Apellido Materno','Puesto','Area','Folio','Registro Postgrado','Vigencia','Numero Empleado','Ruta Imagen','Encargado'])

print(len(DFDatos.index))

#CLEAR INPUTS FUNCTION
def clear(inNombre,inNombre2,inApellido,inApellido2,inSede,inFolio,inEmpleado,inEncargado):
    inNombre.delete("0","end")
    inNombre2.delete("0","end")
    inApellido.delete("0","end")
    inApellido2.delete("0","end")
    inSede.delete("0","end")
    inFolio.delete("0","end")
    inEmpleado.delete("0","end")
    inEncargado.delete("0","end")

ctk.set_appearance_mode("System")


def Interface():
    #Interface tkinter
    app = ctk.CTk()
    app.title("Sistema de credenciales")
    app.geometry("1280x720")
    app.iconbitmap("SCI.ico")

    Titulos = ctk.CTkFont(family="Montserrat", size=36, weight='bold')

    #DATAFRAME
    Nombre = StringVar()
    Nombre2 = StringVar()
    Apellido = StringVar()
    Apellido2 = StringVar()
    Puesto = StringVar()
    Area = StringVar()
    Folio = StringVar()
    Registro = StringVar()
    Vigencia = StringVar()
    Empleado = StringVar()
    Encargado = StringVar()

    #TITLE
    title = ctk.CTkLabel(master=app,
                        text="Sistema de Credenciales para la UTEyCV",
                        font=Titulos)
    title.place(relx=0.5, rely=0.5)

    #FORM
    textNombre = Label(app,text="Nombre:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
    textNombre.place(x=20,y=150)
    inNombre = Entry(app,textvariable=Nombre,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
    inNombre.place(x=150,y=150)

    textNombre2 = Label(app,text="Segundo nombre:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
    textNombre2.place(x=20,y=200)
    inNombre2 = Entry(app,textvariable=Nombre2,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
    inNombre2.place(x=150,y=200)

    textApellido = Label(app,text="Apellido paterno:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
    textApellido.place(x=20,y=250)
    inApellido = Entry(app,textvariable=Apellido,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
    inApellido.place(x=150,y=250)

    textApellido2 = Label(app,text="Apellido materno:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
    textApellido2.place(x=20,y=300)
    inApellido2 = Entry(app,textvariable=Apellido2,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
    inApellido2.place(x=150,y=300)

    dropdownElements = {
        'Director': ['Director de la Unidad Academica'],
        'Sub-director':['Subdirector Academico',
                        'Subdirector Administrativo',
                        'Subdirector De Servicios Educativos E Integracion Social'],
        'Jefe de Departamento':['Jefe de la Unidad de Informatica',
                                'JEfe de la seccion de estudios de posgrado e investigacion',
                                'Jefe del Departamento de Investigacion',
                                'Jefe del Departamente de Posgrado',
                                'Jefe del Laboratorio Institucional de la Red de Robotica y Mecatronica',
                                'Jefe del Departamento de Ciencias Basicas',
                                'Jefe del Departamento de Tecnologias Avanzadas',
                                'Jefe del Departamento de Formacion Integral e Institucional',
                                'Jefe del Departamento de Ingenieria',
                                'Jefe de la Unidad de Tecnologia Educativa y Campus Virtual',
                                'Jefe del Departamento de Evaluacion y Seguimiento Academico',
                                'Jefe del Departamento de Capital Humano',
                                'Jefe del Departamento de Recursos Financieros',
                                'Jefe del Departamento de Recursos Materiales y Servicios',
                                'Jefe del Departamento de Gestion Escolar',
                                'Jefe del Departamento de Servicios Estudiantiles',
                                'Jefa del Departamento de Extensión Y Apoyos Educativos']
    }

    def seleccion_dropdown(*args):
        opcionSelec = comboPuesto.get()
        subOpciones = dropdownElements[opcionSelec]

        #Actualizar las opciones del segundo ComboBox
        comboArea["values"] = subOpciones
        comboArea.set(subOpciones[0])

    Puesto.set(list(dropdownElements.keys())[0])
    Puesto.trace_add('write', seleccion_dropdown)
    #Creamos el Combobox
    comboPuesto = ttk.Combobox(app, textvariable=Puesto ,values=list(dropdownElements.keys()))
    comboPuesto.current(0)
    comboPuesto.pack()

    opcionesArea = dropdownElements[Puesto.get()]
    comboArea = ttk.Combobox(app, textvariable=Area, values=opcionesArea)
    comboArea.current(0)
    comboArea.pack()

    textFolio = Label(app,text="Folio:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
    textFolio.place(x=20,y=400)
    inFolio = Entry(app,textvariable=Folio,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
    inFolio.place(x=150,y=400)

    def selectDate(event):
        fecha= date_entryRegistro.get()

    #Label de Fecha de Registro
    textRegistro = Label(app,text="Registro:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
    textRegistro.place(x=20,y=450)

    #Se crea el DateEntry
    date_entryRegistro = DateEntry(app, textvariable=Registro, date_pattern='dd/mm/yyyy', width = 12, background = 'blue', foreground = 'white', borderwidth = 2)
    date_entryRegistro.place(x = 150, y=450)

    #Llama a la funcion del evento
    date_entryRegistro.bind("DateEntrySelected", selectDate)

    textVigencia = Label(app,text="Vigencia:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
    textVigencia.place(x=20,y=500)

    def selectDateVigencia(event):
        fecha= date_entryVigencia.get()

    #Se crea el DateEntry
    date_entryVigencia = DateEntry(app, textvariable=Vigencia, date_pattern='dd/mm/yyyy', width = 12, background = 'blue', foreground = 'white', borderwidth = 2)
    date_entryVigencia.place(x = 150, y=500)

    #Llama a la funcion del evento
    date_entryVigencia.bind("DateEntrySelected", selectDate)

    textEmpleado = Label(app,text="Empleado:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
    textEmpleado.place(x=20,y=550)
    inEmpleado = Entry(app,textvariable=Empleado,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
    inEmpleado.place(x=150,y=550)

    textEncargado = Label(app,text="Encargado:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
    textEncargado.place(x=20,y=600)
    inEncargado = Entry(app,textvariable=Encargado,bd=4,font="arial 12",bg="#CFCFCF",fg="#000")
    inEncargado.place(x=150,y=600)

    textImagen = Label(app,text="Imagen:", bd=4,font="arial 12",bg="#900C3F",fg="#fff")
    textImagen.place(x=20,y=650)

    def browsefunc():
        global fileSelect
        fileSelect = StringVar()
        fileSelect = fd.askopenfilename(filetypes=(("Img files",".png .jpg .jpeg .webp"),("All files","*.*")))
        #DFDatos.iloc[row,9]=filePad
        print(fileSelect)
        #ent1.insert(tk.END, filename) # add this

    b1 = Button(app,text="Seleccionar imagen",font=40,command=browsefunc)
    b1.place(x=150,y=700)\

    #SAVED FUNCTION
    def save():
        datos = [
            Nombre.get(),
            Nombre2.get(),
            Apellido.get(),
            Apellido2.get(),
            Puesto.get(),
            Area.get(),
            Folio.get(),
            Registro.get(),
            Vigencia.get(),
            Empleado.get(),
            Encargado.get(),
            fileSelect
        ]
        #Save data inside DataFrame, each iteration in the interface insert data
        DFDatos.loc[len(DFDatos.index)]=datos

        saved = Label(app,text="Registro guardado!",bg="yellow")
        saved.pack()
        clear(inNombre,inNombre2,inApellido,inApellido2,inSede,inFolio,inEmpleado,inEncargado)

    buttonName = Button(app,text="Guardar Datos",bd=3,command = save,bg="#94FF40",font="arial 12",cursor="plus")
    buttonName.place(x=20,y=750)

    #Cierre de la ventana
    app.mainloop()

Frame = Interface()

#Get date in a friendly format
def FechaActualCompleta(date):
    months = ("ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
    day = date.day
    month = months[date.month - 1]
    year = date.year
    messsage = "{} DE {} DEL {}".format(day, month, year)
    return messsage

#Begin process to make ID's


#Data cleaning and normalization
""" DFDatos['Nombre'] = DFDatos['Nombre'].replace(pd.NA,0)
DFDatos = DFDatos[DFDatos['Nombre'] != 0] """

DFDatos['Nombre 2'] = DFDatos['Nombre 2'].replace(pd.NA,'')

""" DFDatos['Nombre'] = DFDatos['Nombre'].map(lambda x: x.rstrip(' '))
DFDatos['Nombre 2'] = DFDatos['Nombre 2'].map(lambda x: x.lstrip(' '))
DFDatos['Apellido Paterno'] = DFDatos['Apellido Paterno'].map(lambda x: x.rstrip(' '))
DFDatos['Apellido Materno'] = DFDatos['Apellido Materno'].map(lambda x: x.lstrip(' '))
"""
DFDatos['Folio'] = DFDatos['Folio'].map(lambda x: x.lstrip('FOLIO:'))

DFDatos['Numero Empleado'] = DFDatos['Numero Empleado'].astype("int64")

DFDatos['Vigencia'] = pd.to_datetime(DFDatos['Vigencia'], format="%d/%m/%Y")


DFDatos.to_excel('CredencialesRealizadas.xlsx')
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
    if(os.path.exists(DFDatos.iloc[row,9])):
        img = Image.open(DFDatos.iloc[row,9]).convert("RGB")
        img.show()

        #Calcula las dimensiones de la imagen
        width, height = img.size

        #Se centra en el centro de la imagen
        square_size = min(width,height)

        # Calcula las coordenadas del cuadrado de recorte
        left = (width - square_size) / 2
        top = (height - square_size) / 2
        right = (width + square_size) / 2
        bottom = (height + square_size) / 2

        cut_img = img.crop((left,top,right,bottom))
        cut_img.show()
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
