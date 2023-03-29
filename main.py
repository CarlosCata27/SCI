import pandas as pd
import qrcode
import os
import sys
import shutil
import numpy as np
import sys
import PyPDF2
import openpyxl

import gui

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

from datetime import datetime,date

#importar Pillow
from PIL import Image, ImageDraw
from PIL import Image

#Conexion to Firebase Storage
from Pyrebase.firebaseStorage import FirebaseStorage
#Config storage from firebase
storage = FirebaseStorage()

#Nombre del archivo de base de datos?
fileName = 'CredencialesRealizadas.xlsx'
archivoNuevo = False

if(os.path.exists(fileName)):
    archivoNuevo = False
    DFExistentes = pd.read_excel(fileName)
else:
    archivoNuevo = True
    newExcel = openpyxl.Workbook()
    newExcel.save(fileName)


DFDatos = pd.DataFrame(columns=['Folio',
                                'Nombre','Nombre 2',
                                'Apellido Paterno',
                                'Apellido Materno',
                                'Puesto','Area',
                                'Registro',
                                'Vigencia',
                                'Numero Empleado',
                                'Encargado',
                                'Ruta Imagen'])


def receive_data(app_instance, data):
    print("Datos recibidos", data)
    DFDatos.loc[len(DFDatos.index)]=data

app = gui.AppGui(data_callback=receive_data)
app.mainloop()

#Get date in a friendly format
def FechaActualCompleta(date):
    months = ("ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
    day = date.day
    month = months[date.month - 1]
    year = date.year
    messsage = "{} DE {} DEL {}".format(day, month, year)
    return messsage

DFDatos['Nombre 2'] = DFDatos['Nombre 2'].replace(pd.NA,'')

#DFDatos['Folio'] = DFDatos['Folio'].map(lambda x: x.lstrip('FOLIO:'))

#DFDatos['Numero Empleado'] = DFDatos['Numero Empleado'].astype("int64")

DFDatos['Vigencia'] = pd.to_datetime(DFDatos['Vigencia'], format="%d/%m/%Y")

def consecutivoFolio(Dataframe):
    #SCI, año YY, ##
    Anio = str(datetime.now().year)
    Anio = Anio[-2:]
    padding=4

    DFReal = Dataframe[Dataframe['Folio'] != 0]
    DFFolios = Dataframe[Dataframe['Folio'] == 0]
    
    if(not(DFReal.empty)):
        registrosExistentes = len(DFReal.index)
        numero = int(DFReal.iloc[registrosExistentes-1,0][-4:])+1
    else:
        numero = 1

    print(DFFolios)
    for i in range(len(DFFolios.index)):
        Consecutivo = str(numero).zfill(padding)
        print(numero)
        Folio = "SCI"+Anio+Consecutivo
        print(Folio)
        DFFolios.iloc[i,0] = Folio

        numero+=1
        
    DFAux = pd.concat([DFReal,DFFolios],ignore_index=True)
    return DFAux

if (archivoNuevo == True):
    DFAux = consecutivoFolio(DFDatos)
    Dataframe = DFAux
else:
    #Concatenacion de los registros hechos desde la GUI a los registros que obtuvimos del excel
    Dataframe = pd.concat([DFExistentes,DFDatos],ignore_index=True)

    Dataframe = consecutivoFolio(Dataframe)



print(Dataframe)


#with pd.ExcelWriter(fileName, mode='w', engine='openpyxl') as writer:
Dataframe.to_excel(fileName, index=False)

#DFDatos.to_excel('CredencialesRealizadas.xlsx')
#Creation and validation directorys to save final files
shutil.rmtree('CodigosQR', ignore_errors=True)
shutil.rmtree('Recortes', ignore_errors=True)
os.mkdir('CodigosQR')
os.mkdir('Recortes')

registrosCreados = len(DFDatos.index)
DFDatos = Dataframe.tail(registrosCreados)

#Inicializate how to create each pdf and ID
#for row in range(len(DFDatos.index)):
for row in range(1):
    #Get name foreach file
    nombrePDF = str(DFDatos.iloc[row, 3]) + str(DFDatos.iloc[row, 4]) + str(DFDatos.iloc[row, 1]) + str(DFDatos.iloc[row, 2]) + str(DFDatos.iloc[row,0])
    #Create PDF file
    myCanvas = canvas.Canvas(f'./PDFs/{nombrePDF}.pdf', pagesize=letter)
    width, height = letter  #612 792

    #Get UPIITA's logo online
    Upiita = 'https://firebasestorage.googleapis.com/v0/b/credenciales-uteycv.appspot.com/o/LogoUPIITA.png?alt=media&token=36f618f6-45e0-4f7f-a26b-355759536a3a'

    #UPIITA's logo localization in PDF file
    Fondo = myCanvas.drawImage(Upiita,x=width/6,y=height/4,height=450,width=450,mask='auto')

    #Phrases with Helveltica font
    myCanvas.setFont("Helvetica", 50)
    Nombre = ''.join([str(DFDatos.iloc[row,1]),' ', str(DFDatos.iloc[row,2])])
    Apellidos = ''.join([str(DFDatos.iloc[row,3]),' ', str(DFDatos.iloc[row,4])])
    Empleado = ''.join(['No Empleado o RFC: ',str(DFDatos.iloc[row,9])])

    myCanvas.drawCentredString(width/2,700,Nombre)
    myCanvas.drawCentredString(width/2,650,Apellidos)
    myCanvas.drawCentredString(width/2,590,Empleado)

    myCanvas.setFont("Helvetica", 30)
    myCanvas.drawCentredString(width/2,450,'VIGENCIA: ')
    myCanvas.drawCentredString(width/2,350,'FECHA DE EXPEDICIÓN')
    myCanvas.drawCentredString(width/2,250,'LUGAR DE ORIGEN')

    #Phrases with Helveltica-Bold font
    myCanvas.setFont("Helvetica-Bold", 30)
    Folio = 'FOLIO: '+str(DFDatos.iloc[row,0])

    myCanvas.drawCentredString(width/2,75,Folio)

    #Get date from expedition to add
    Now = datetime.now()
    Fechahoy = FechaActualCompleta(Now)

    myCanvas.drawCentredString(width/2,400,FechaActualCompleta(DFDatos.iloc[row,8]))
    myCanvas.drawCentredString(width/2,300,Fechahoy)
    myCanvas.drawCentredString(width/2,200,'UTEyCV')

    #Saving PDF file
    myCanvas.save()

    #Saving our PDF file in Firebase cloud
    file = f'PDFs/{nombrePDF}.pdf'
    cloudfilename = f'PDFs/{nombrePDF}.pdf'

    storage.upload_file(file, cloudfilename)

    #Creating QR Code from url firebase
    url_PDF = storage.get_file_url(cloud_filename=cloudfilename)
    CodigoQR = qrcode.make(url_PDF)
    nombreCred = "Cred"+nombrePDF
    CodigoQR.save(f'CodigosQR/{nombrePDF}.png')

    # PDF file with ID

    #Set image background
    myCanvas = canvas.Canvas(f'Credenciales/{nombreCred}.pdf', pagesize=(3322,2214))
    url = 'https://firebasestorage.googleapis.com/v0/b/credenciales-uteycv.appspot.com/o/1.png?alt=media&token=45f62ed5-23d7-47f2-9de5-67aab19e9d1c'
    Base = myCanvas.drawImage(url,0,0,height=2214,width=3322,mask='auto')

    #Image selection from Fotos ID
    if(os.path.exists(DFDatos.iloc[row,11])):
        img = Image.open(DFDatos.iloc[row,11]).convert("RGB")
        #img.show()

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
        #cut_img.show()
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
    myCanvas.drawString(1300,550,DFDatos.iloc[row,6])

    #Phrases with Helveltica font
    myCanvas.setFont("Helvetica", 100)
    myCanvas.drawString(1300,400,Empleado)
    #Date in format m/Y

    temp = DFDatos.iloc[row,8].strftime('%m/%Y')
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

input_folder = "./Credenciales"
filename = nombreCred+".pdf"

input_file = os.path.join(input_folder,filename)
output_file = os.path.join(input_folder,filename)


# Abre el archivo PDF original
with open(input_file, "rb") as file:
    # Crea un objeto PDFReader
    reader = PyPDF2.PdfReader(file)

    # Crea un objeto PDFWriter para escribir el PDF modificado
    writer = PyPDF2.PdfWriter()

    # Itera sobre las páginas del PDF
    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]

        # Gira la página 180 grados en sentido horario
        if page_num == 1:  # Cambia el índice de la página aquí
            page.rotate(180)

        # Agrega la página al objeto PDFWriter
        writer.add_page(page)

    # Guarda el PDF modificado
    with open(output_file, "wb") as output:
        writer.write(output)

python = sys.executable
os.execl(python, python, * sys.argv)