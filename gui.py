import tkinter as tk
import customtkinter as ctk
import sys, os

from tkcalendar import DateEntry
from opcionesCombo import dropdownElements
from tkinter import filedialog as fd

class ToplevelWindow(ctk.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("400x100")
        self.title("Registro Exitoso")
        self.grab_set()

        Subtitulo = ctk.CTkFont(family="Montserrat", size=24, weight='normal')

        self.label = ctk.CTkLabel(self, text="Registro realizado exitosamente!", font=Subtitulo)
        self.label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        #buttonName = ctk.CTkButton(self, text="Realizar un nuevo Registro", command=restart_program, font=Boton, cursor="plus")
        #buttonName.place(relx=0.3, rely=0.3, anchor=tk.CENTER)

class FormFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        for i in range(9):
            self.grid_rowconfigure(i, weight=1)
        for i in range(5):
            self.grid_columnconfigure(i, weight=1)

        Subtitulo = ctk.CTkFont(family="Montserrat", size=24, weight='normal')
        NormalFont = ctk.CTkFont(family="Assistant", size=16, weight='normal')

        self.subtitle= ctk.CTkLabel(self, text='Registro para Credencial', font=Subtitulo)
        self.subtitle.grid(row=0, padx=20, pady=20, columnspan=5)

        global nombre
        nombre = tk.StringVar()
        labelNombre = ctk.CTkLabel(self, text="Nombre:", font=NormalFont)
        labelNombre.grid(row=2, column=0, padx=20, pady=5, sticky=tk.E)
        inputNombre = ctk.CTkEntry(self, textvariable=nombre, border_width=2, corner_radius=10)
        inputNombre.grid(row=2, column=1, padx=20, pady = 5)

        global nombre2
        nombre2 = tk.StringVar()
        labelNombre2 = ctk.CTkLabel(self, text="Segundo Nombre:", font=NormalFont)
        labelNombre2.grid(row=2, column=2, padx=20, pady=5, sticky=tk.E)
        inputNombre2 = ctk.CTkEntry(self, textvariable=nombre2, border_width=2, corner_radius=10)
        inputNombre2.grid(row=2, column=3, padx=20, pady = 5)

        global apellidoP
        apellidoP = tk.StringVar()
        labelApellidoP = ctk.CTkLabel(self, text="Apellido Paterno:", font=NormalFont)
        labelApellidoP.grid(row=3, column=0, padx=20, pady=5, sticky=tk.E)
        inputApellidoP = ctk.CTkEntry(self, textvariable=apellidoP, border_width=2, corner_radius=10)
        inputApellidoP.grid(row=3, column=1, padx=20, pady=5, sticky=tk.E)

        global apellidoM
        apellidoM = tk.StringVar()
        labelApellidoM = ctk.CTkLabel(self, text="Apellido Materno:", font=NormalFont)
        labelApellidoM.grid(row=3, column=2, padx=20, pady=5, sticky=tk.E)
        inputApellidoM = ctk.CTkEntry(self, textvariable=apellidoM, border_width=2, corner_radius=10)
        inputApellidoM.grid(row=3, column=3, padx=20, pady=5, sticky=tk.E)

        global puesto
        puesto = tk.StringVar()
        def seleccion_dropdown(*args):
            opcionSelec = puesto.get()
            if opcionSelec in dropdownElements:
                subOpciones = dropdownElements[opcionSelec]

                #Actualizar las opciones del segundo ComboBox
                comboArea.configure(values=subOpciones)
                comboArea.set(subOpciones[0])
                print("aqui 1")
            else:
                print("aqui 2")
                comboArea.set('')
                comboArea["values"] = []

        puesto.set(list(dropdownElements.keys())[0])
        puesto.trace_add('write', seleccion_dropdown)

        labelPuesto = ctk.CTkLabel(self, text="Puesto:", font=NormalFont)
        labelPuesto.grid(row=4, column=0, padx=20, pady=5, sticky=tk.E)
        comboPuesto = ctk.CTkComboBox(self, variable=puesto ,values=list(dropdownElements.keys()), font=NormalFont)
        comboPuesto.grid(row=4, column=1, padx=20, pady=5, sticky=tk.E)

        global area
        area = tk.StringVar()
        opcionesArea = dropdownElements[puesto.get()]
        print(opcionesArea)
        labelArea = ctk.CTkLabel(self, text="Area:", font=NormalFont)
        labelArea.grid(row=4, column=2, padx=20, pady=5, sticky=tk.E)
        comboArea = ctk.CTkComboBox(self, variable=area, values=opcionesArea)
        comboArea.grid(row=4, column=3, padx=20, pady=5, sticky=tk.E)

        global registro, vigencia
        registro = tk.StringVar()
        vigencia = tk.StringVar()
        def selectDate(event):
            fecha= date_entryRegistro.get()

        labelRegistro = ctk.CTkLabel(self, text="Fecha Registro:", font=NormalFont)
        labelRegistro.grid(row=5, column=0, padx=20, pady=5, sticky=tk.E)
        #Se crea el DateEntry
        date_entryRegistro = DateEntry(self, textvariable=registro, date_pattern='dd/mm/yyyy', width = 14, background = 'grey', foreground = 'white', borderwidth = 3)
        date_entryRegistro.grid(row=5, column=1, padx=20, pady=5, sticky=tk.E)

        #Llama a la funcion del evento
        date_entryRegistro.bind("DateEntrySelected", selectDate)

        def selectDateVigencia(event):
            fecha= date_entryVigencia.get()

        labelRegistro = ctk.CTkLabel(self, text="Fecha Expiracion:", font=NormalFont)
        labelRegistro.grid(row=5, column=2, padx=20, pady=5, sticky=tk.W)

        date_entryVigencia = DateEntry(self, textvariable=vigencia, date_pattern='dd/mm/yyyy', width = 14, background = 'grey', foreground = 'white', borderwidth = 3)
        date_entryVigencia.grid(row=5, column=3, padx=20, pady=5, sticky=tk.E)

        global empleado
        empleado = tk.StringVar()
        labelEmpleado = ctk.CTkLabel(self, text="No. Empleado o RFC:", font=NormalFont)
        labelEmpleado.grid(row=6, column=0, padx=20, pady=5, sticky=tk.W)
        inputEmpleado = ctk.CTkEntry(self, textvariable=empleado, border_width=2, corner_radius=10)
        inputEmpleado.grid(row=6, column=1, padx=20, pady=5, sticky=tk.E)

        global responsable
        responsable = tk.StringVar()
        labelResponsable = ctk.CTkLabel(self, text="Resposable de Credencial:", font=NormalFont)
        labelResponsable.grid(row=7, column=0, padx=20, pady=5, sticky=tk.E)
        inputResponsable = ctk.CTkEntry(self, textvariable=responsable, border_width=2, corner_radius=10)
        inputResponsable.grid(row=7, column=1, padx=20, pady=5, sticky=tk.E)

        textImagen = ctk.CTkLabel(self, text="Imagen para Credencial:", font=NormalFont)
        textImagen.grid(row=8, column=0, padx=20, pady=5, sticky=tk.E)

        global fileSelect
        fileSelect = tk.StringVar()

        def browsefunc():
            file = fd.askopenfilename(filetypes=(("Img files",".png .jpg .jpeg .webp"),("All files","*.*")))
            fileSelect.set(file)
            #DFDatos.iloc[row,9]=filePad
            print(fileSelect)
            textURLimg.configure(text=file)
            #ent1.insert(tk.END, filename) # add this

        textURLimg = ctk.CTkLabel(self, text=fileSelect, font=NormalFont)
        textURLimg.grid(row=8, column=2, columnspan=3, padx=20, pady=5, sticky=tk.E)

        botonImagen = ctk.CTkButton(self,text="Seleccionar imagen", font=NormalFont ,command=browsefunc)
        botonImagen.grid(row=8, column=1, padx=20, pady=5, sticky=tk.E)

class AppGui(ctk.CTk):
    def __init__(self, data_callback=None, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.data_callback = data_callback
        self.geometry("1280x720")
        self.title("Sistema de credenciales")
        # Intenta establecer el icono de la aplicación
        try:
            self.iconbitmap("SCI.ico")
        except tk.TclError:
            pass  # Omitimos el icono si no es compatible
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.resizable(width=False, height=False)

        self.grab_set()


        Titulos = ctk.CTkFont(family="Montserrat", size=32, weight='bold')
        Boton = ctk.CTkFont(family="Montserrat", size=24, weight='bold')

        title = ctk.CTkLabel(master=self, text="Sistema de Credenciales para la UTEyCV", font=Titulos)
        title.grid(row=0, column=0, padx = 20, pady = 20)

        self.formFrame = FormFrame(master=self)
        self.formFrame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        buttonName = ctk.CTkButton(self, text="Guardar Datos", command=self.save, font=Boton, cursor="plus")
        buttonName.grid(row=2, column=0, padx=20, pady=20)

        self.toplevel_window = None
        

    def save(self, event=None):
        datos = [
            0,
            nombre.get(),
            nombre2.get(),
            apellidoP.get(),
            apellidoM.get(),
            puesto.get(),
            area.get(),
            registro.get(),
            vigencia.get(),
            empleado.get(),
            responsable.get(),
            fileSelect.get()
            ]

        if self.data_callback:
            self.data_callback(self, datos)

        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = ToplevelWindow(self)
        else:
            self.toplevel_window.focus()