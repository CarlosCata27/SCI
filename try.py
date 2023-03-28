from datetime import datetime 
#SCI, año YY, ##
Anio = str(datetime.now().year)
Anio = Anio[-2:]
print(Anio)
Folio = "SCI"+Anio
registrosExistentes = 2
for i in range(2):
    print(Folio)