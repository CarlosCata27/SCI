# SCI
Sistema de Credencialización Interno

Install libraries
 - pandas
 - reportlab
 - qrcode
 - pyrebase4
 - pillow
 - customtkinter
 - pyPDF2

If you want an executable file, install the following librarie
 - pyinstaller

After that go to the same path as .py files, insert in console the following command
```
 pyinstaller --windowed --onefile .\SCI.py
```
Also customtkinter need to be included as a --add-data so we put this command in terminal
```
 pyinstaller --windowed --onefile --add-data "C:/Users/USER/AppData/Local/Programs/Python/Python3##/Lib/site-packages/customtkinter;customtkinter/" .\main.py
```
Change to ur user

Finally in main.spec you need to add in hiddenimports
```
hiddenimports=['babel.numbers'],
```

Now we can execute our pyinstaller with
```
pyinstaller .\main.spec
```