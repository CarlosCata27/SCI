import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import qrcode
from datetime import datetime

#Funcion que recupera la fecha actual en formato amigable
def FechaActual(date):
    months = ("ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
    day = date.day
    month = months[date.month - 1]
    year = date.year
    messsage = "{} DE {} DEL {}".format(day, month, year)
    return messsage


# Load the Excel file
archivo = pd.ExcelFile('Credenciales.xlsx')
df = pd.read_excel(archivo,'Hoja 1',usecols='A:G',skiprows=range(1))

DFDatos = pd.DataFrame({
    'Nombre':df.iloc[:,0],
    'Nombre 2':df.iloc[:,1],
    'Apellido Paterno':df.iloc[:,2],
    'Apellido Materno':df.iloc[:,3],
    'Sede':df.iloc[:,4],
    'Folio':df.iloc[:,5],
    'Registro Postgrado':df.iloc[:,6]
})

DFDatos['Nombre'] = DFDatos['Nombre'].replace(pd.NA,0)
DFDatos = DFDatos[DFDatos['Nombre'] != 0]

DFDatos['Nombre 2'] = DFDatos['Nombre 2'].replace(pd.NA,'')

DFDatos['Nombre'] = DFDatos['Nombre'].map(lambda x: x.rstrip(' '))
DFDatos['Nombre 2'] = DFDatos['Nombre 2'].map(lambda x: x.lstrip(' '))
DFDatos['Apellido Paterno'] = DFDatos['Apellido Paterno'].map(lambda x: x.rstrip(' '))
DFDatos['Apellido Materno'] = DFDatos['Apellido Materno'].map(lambda x: x.lstrip(' '))

DFDatos['Folio'] = DFDatos['Folio'].map(lambda x: x.lstrip('FOLIO:'))
#print(DFDatos)
#for row in range(len(DFDatos.index)):
for row in range(1):
    nombrePDF = str(DFDatos.iloc[row,5])
    myCanvas = canvas.Canvas(f'{nombrePDF}.pdf', pagesize=letter)
    width, height = letter  #keep for later
    print(width, height)
    
    Upiita = 'logo_upiita_ipn_varios_colores_logo_upiita_gris.png'

    Fondo = myCanvas.drawImage(Upiita,x=width/6,y=height/4,height=450,width=450,mask=[0,0,0,0,0,0])

    #Textos con la fuente normal
    myCanvas.setFont("Helvetica", 50)
    Nombre = ''.join([str(DFDatos.iloc[row,0]),' ', str(DFDatos.iloc[row,1])])
    Apellidos = ''.join([str(DFDatos.iloc[row,2]),' ', str(DFDatos.iloc[row,3])])

    myCanvas.drawCentredString(width/2,700,Nombre)
    myCanvas.drawCentredString(width/2,650,Apellidos)

    myCanvas.setFont("Helvetica", 30)
    myCanvas.drawCentredString(width/2,350,'FECHA DE EXPEDICIÓN')
    myCanvas.drawCentredString(width/2,250,'LUGAR DE ORIGEN')

    #Textos con fuente negrita
    myCanvas.setFont("Helvetica-Bold", 30)
    Folio = 'FOLIO: '+str(DFDatos.iloc[row,5])

    myCanvas.drawCentredString(width/2,75,Folio)

    Now = datetime.now()
    Fechahoy = FechaActual(Now)


    myCanvas.drawCentredString(width/2,300,Fechahoy)
    myCanvas.drawCentredString(width/2,200,'UTEYCV')


    
    CodigoQR = qrcode.make(f'{nombrePDF}.pdf')
    f = open(f'{nombrePDF}.png', "wb")
    CodigoQR.save(f)
    f.close()
    myCanvas.save()