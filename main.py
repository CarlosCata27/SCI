import pandas as pd
import qrcode
import os
import sys
import shutil
import numpy as np
import PyPDF2
import openpyxl
import gui
import fitz
import glob
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, mm
from reportlab.pdfbase import pdfdoc
from datetime import datetime
from itertools import groupby
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
        Folio = "SCI"+Anio+Consecutivo
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

#with pd.ExcelWriter(fileName, mode='w', engine='openpyxl') as writer:
Dataframe.to_excel(fileName, index=False)

registrosCreados = len(DFDatos.index)
DFDatos = Dataframe.tail(registrosCreados)

if(DFDatos.empty):
    sys.exit()

#DFDatos.to_excel('CredencialesRealizadas.xlsx')
#Creation and validation directorys to save final files
shutil.rmtree('CodigosQR', ignore_errors=True)
shutil.rmtree('Recortes', ignore_errors=True)
os.mkdir('CodigosQR')
os.mkdir('Recortes')
if(not(os.path.exists('PDFs'))):
    os.mkdir('PDFs')
if(not(os.path.exists('Credenciales'))):
    os.mkdir('Credenciales')
if(not(os.path.exists('./Credenciales/Redimensionadas'))):
    os.mkdir('./Credenciales/Redimensionadas')
if(not(os.path.exists('./Credenciales/Redimensionadas/img'))):
    os.mkdir('./Credenciales/Redimensionadas/img')



#Inicializate how to create each pdf and ID
for row in range(len(DFDatos.index)):
#for row in range(1):
    #Get name foreach file
    nombrePDF = str(DFDatos.iloc[row, 3]) + str(DFDatos.iloc[row, 4]) + str(DFDatos.iloc[row, 1]) + str(DFDatos.iloc[row, 2]) + str(DFDatos.iloc[row,0])
    #Create PDF file
    myCanvas = canvas.Canvas(f'./PDFs/{nombrePDF}.pdf', pagesize=letter)
    width, height = letter  #612 792

    #Get UPIITA's logo online
    Upiita = 'https://firebasestorage.googleapis.com/v0/b/credenciales-uteycv.appspot.com/o/Hoja%20credencial.png?alt=media&token=5e9008f8-f44c-4f24-ab4b-326bab8bdfd9'

    #UPIITA's logo localization in PDF file
    Fondo = myCanvas.drawImage(Upiita,x=0,y=0,height=letter[1],width=letter[0],mask='auto')

    #Phrases with Helveltica font
    myCanvas.setFont("Helvetica", 50)
    Nombre = ''.join([str(DFDatos.iloc[row,1]),' ', str(DFDatos.iloc[row,2])])
    Apellidos = ''.join([str(DFDatos.iloc[row,3]),' ', str(DFDatos.iloc[row,4])])
    Empleado = ''.join([str(DFDatos.iloc[row,9])])

    myCanvas.drawCentredString(width/2,700,Nombre)
    myCanvas.drawCentredString(width/2,650,Apellidos)
    myCanvas.drawCentredString(width/2,590,'No EMPLEADO o RFC: ')
    myCanvas.drawCentredString(width/2,540,Empleado)

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

    #If image doesn't exist, dont try to paste in ID pdf file
    if(os.path.exists(f'Recortes/{nombrePDF}.png')):
        pathCrop = f'Recortes/{nombrePDF}.png'
        Base = myCanvas.drawImage(pathCrop,200,370,height=1055,width=1055,mask='auto')

    #Phrases with Helveltica-Bold font
    myCanvas.setFont("Helvetica-Bold", 200)
    myCanvas.drawString(1300,1100,Nombre)
    myCanvas.drawString(1300,850,Apellidos)

    myCanvas.setFont("Helvetica-Bold", 150)
    myCanvas.drawString(1300,600,DFDatos.iloc[row,5])
    myCanvas.drawString(1300,450,DFDatos.iloc[row,6])

    #Phrases with Helveltica font
    myCanvas.setFont("Helvetica", 100)
    myCanvas.drawString(1300,250,Empleado)
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

        
    input_folder = "./Credenciales"
    output_folder = "./Credenciales/Redimensionadas"
    filename = nombreCred+".pdf"

    input_file = os.path.join(input_folder,filename)
    output_file = os.path.join(output_folder,nombreCred+"_resized.pdf")

    target_width, target_height = 85.60 * mm, 53.98 * mm
    # Abre el archivo PDF original
    with open(input_file, "rb") as file:
        # Crea un objeto PDFReader
        reader = PyPDF2.PdfReader(file)
        # Crea un objeto PDFWriter para escribir el PDF modificado
        writer = PyPDF2.PdfWriter()
        # Itera sobre las páginas del PDF
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]

            # Redimensionar la página al tamaño de la identificación
            page.scale_to(target_width, target_height)

            # Rotar la segunda página 180 grados en sentido horario
            if page_num == 1:
                page.rotate(180)

            # Agregar la página al objeto PDFWriter
            writer.add_page(page)

        # Guarda el PDF modificado
        with open(output_file, "wb") as output:
            writer.write(output)
    os.remove(input_file)


# Combinar los PDF redimensionados en un nuevo PDF tamaño "letter"


# Obtener todos los archivos PDF redimensionados en la carpeta "Credenciales/Redimensionados"
resized_files = glob.glob("./Credenciales/Redimensionadas/*_resized.pdf")

output_folder = "./Credenciales/Redimensionadas/img"

# Iterar sobre los archivos PDF redimensionados
for resized_file in resized_files:
    input_pdf = resized_file  # Ruta del archivo PDF
    

    # Dimensiones deseadas en puntos (dpi)
    """ target_width_pt = 96.36
    target_height_pt = 95.817 """
    target_width_pt = 300
    target_height_pt = 300

    # Abrir el archivo PDF
    pdf = fitz.open(input_pdf)

    # Crear una lista para almacenar las imágenes de cada página
    images = []

    # Abrir el archivo PDF
    pdf = fitz.open(input_pdf)
    name_without_extension = os.path.splitext(os.path.basename(input_pdf))[0]

    # Iterar sobre las páginas del PDF
    for page_num in range(len(pdf)):
        page = pdf.load_page(page_num)

        # Renderizar la página como imagen
        pix = page.get_pixmap(matrix=fitz.Matrix(target_width_pt/72, target_height_pt/72))

        # Guardar la imagen como archivo PNG
        output_image = os.path.join(output_folder,name_without_extension+f"{page_num}.png")  # Ruta de salida para la imagen
        pix.save(output_image, "png")

    # Cerrar el archivo PDF
    pdf.close()


# Obtener la fecha y hora actual
now = datetime.now()
# Formatear la fecha actual en formato "YY-MM-DD-HH-MM"
date_string = now.strftime("%y-%m-%d--%H-%M")
input_folder = "./Credenciales"

# Agregar la fecha al nombre del archivo
output_combined = os.path.join(input_folder, f"Credenciales_{date_string}.pdf")

cred_per_page = 4
margin_x = 30
margin_y = 600

# Crear nuevo PDF tamaño "letter"
newPDF = canvas.Canvas(output_combined, pagesize=letter)

# Inicializar las coordenadas x e y
x = margin_x
y = margin_y

# Inicializar el contador de credenciales
cred_count = 0
frontal=0
horizontal=1

target_width, target_height = 85.60 * mm, 53.98 * mm

# Obtener todos los archivos PDF redimensionados en la carpeta "Credenciales"
resized_files = glob.glob("./Credenciales/Redimensionadas/img/*.png")

# Ordenar los archivos por nombre base
resized_files.sort(key=lambda x: os.path.splitext(os.path.basename(x))[0])
#y1=600,y2=447,x=30
#y1=600,y2=447,x=320

#y1=300,y2=147,x=30
#y1=300,y2=147,x=320

# Iterar sobre las archivos PNG redimensionados
for resized_file in resized_files:
    

    # Agregar la página al nuevo PDF tamaño "letter" en las coordenadas especificadas
    #newPDF.setPageSize(letter)
    if frontal ==0:
        newPDF.drawImage(resized_file,x=x,y=y,height=target_height,width=target_width,mask='auto')
        frontal+=1
    elif frontal==1:
        newPDF.drawImage(resized_file,x=x,y=y-153,height=target_height,width=target_width,mask='auto')
        frontal=0
        if horizontal==1:
            horizontal+=1
            x=320
        else:
            horizontal=1
            x=margin_x
            y = 250


    # Actualizar el contador de credenciales
    cred_count += 1

    # Verificar si se alcanzó el límite de credenciales por página
    if cred_count == (2*cred_per_page):
        # Reiniciar el contador de credenciales
        cred_count = 0

        # Agregar salto de página en el nuevo PDF tamaño "letter"
        newPDF.showPage()
        x = margin_x
        y = margin_y
        

# Guardar y cerrar el nuevo PDF tamaño "letter"
newPDF.save()


if(not(os.path.exists('./Credenciales/Realizadas'))):
    os.mkdir('./Credenciales/Realizadas')

resized_files = glob.glob("./Credenciales/Redimensionadas/*.pdf")
for file in resized_files:
    shutil.move(file, './Credenciales/Realizadas')


#Delete CodigosQR and crop images dir, we dont need it
shutil.rmtree('./CodigosQR', ignore_errors=True)
shutil.rmtree('Recortes', ignore_errors=True)
shutil.rmtree('./Credenciales/Redimensionadas/img', ignore_errors=True)
shutil.rmtree('./Credenciales/Redimensionadas', ignore_errors=True)
