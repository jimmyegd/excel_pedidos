import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.ttk import Notebook
import os
import excel
import exceldiff


ws = tk.Tk()



w = 790 # width for the Tk root
h = 380 # height for the Tk root

# get screen width and height
wscreen = ws.winfo_screenwidth() # width of the screen
hscreen = ws.winfo_screenheight() # height of the screen

# calculate x and y coordinates for the Tk root window
x = (wscreen/2) - (w/2)
y = (hscreen/2) - (h/2)

# set the dimensions of the screen 
# and where it is placed
ws.geometry('%dx%d+%d+%d' % (w, h, x, y))



ws.title('Pedidos')


notebook = Notebook(ws)
notebook.pack(pady=10, expand=True)


frame1 = tk.Frame(notebook, width=780, height=340)
frame2 = tk.Frame(notebook, width=780, height=340)

def select_file():
    filetypes = (
        ('Excel files', '*.xlsx'),
        # ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    if filename:
        filepath = os.path.abspath(filename)
        print(filepath)
        l4.config(text=filepath)


def select_file2():
    filetypes = (
        ('Excel files', '*.xlsx'),
        # ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    if filename:
        filepath = os.path.abspath(filename)
        print(filepath)
        l4_frame2.config(text=filepath)

def select_file3():
    filetypes = (
        ('Excel files', '*.xlsx'),
        # ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    if filename:
        filepath = os.path.abspath(filename)
        print(filepath)
        l8_frame2.config(text=filepath)


def limpiar():
    l4.config(text='Ruta Archivo')
    l7.config(text='Ruta Carpeta')
    dias_alcance_var.set('')
    dias_analisis_var.set('')
    nombre_archivo_var.set('')


def select_folder():
    foldername = fd.askdirectory()
    if foldername:
        folderpath = os.path.abspath(foldername)
        l7.config(text=folderpath)


def select_folder2():
    foldername = fd.askdirectory()
    if foldername:
        folderpath = os.path.abspath(foldername)
        l76_frame2.config(text=folderpath)


def procesar_doc():
    excel.alcance(l4.cget("text"), dias_analisis_var.get(), dias_alcance_var.get(),
                  l7.cget("text") + "//" + nombre_archivo_var.get())
    # l72.get()#diasaanalisis
    #  l92.get()#diasalcance
    # l4.get()  #ruta archivo
    #  l7.get() #foldername
    #  l112.get() #nombrearchivo


def cruzar_datos():
    exceldiff.cruzarDatos(l4_frame2.cget("text"), l8_frame2.cget("text"), l76_frame2.cget("text") + "//"+ nombre_archivo_var_frame2.get())

def limpiar_frame2():
    l4_frame2.config(text='Ruta Archivo')
    l76_frame2.config(text='Ruta Carpeta')
    l8_frame2.config(text='Ruta Archivo')
    nombre_archivo_var_frame2.set('')


open_button2 = ttk.Button(
    frame2,
    text='Seleccionar Archivo...',
    command=select_file2
)

open_button3 = ttk.Button(
    frame2,
    text='Seleccionar Archivo...',
    command=select_file3
)

open_button3_output = ttk.Button(
    frame2,
    text='Seleccionar Carpeta...',
    command=select_folder2
)

open_button = ttk.Button(
    frame1,
    text='Seleccionar Archivo...',
    command=select_file
)

open_buttonOutPut = ttk.Button(
    frame1,
    text='Seleccionar Carpeta...',
    command=select_folder
)

procesar = ttk.Button(
    frame1,
    text='Procesar',
    command=procesar_doc
)

limpiar = ttk.Button(
    frame1,
    text='Limpiar',
    command=limpiar
)

procesar_frame2 = ttk.Button(
    frame2,
    text='Procesar',
    command=cruzar_datos
)

limpiar_frame2 = ttk.Button(
    frame2,
    text='Limpiar',
    command=limpiar_frame2
)

#frame 1--- pestaña archivo alcance
l1 = tk.Label(frame1, text='')
l1.grid(row=1,column=1)

l2 = tk.Label(frame1, text='')
l2.grid(row=1,column=2)

l3 = tk.Label(frame1, text='Archivo a procesar:')
l3.grid(row=2,column=1, ipadx=40)

l4 = tk.Label(frame1, text='Ruta Archivo', width=55)
l4.grid(row=2,column=2)

open_button.grid(row=2,column=3)

l66 = tk.Label(frame1, text='')
l66.grid(row=3,column=1)

l77 = tk.Label(frame1, text='')
l77.grid(row=3,column=2)

l6 = tk.Label(frame1, text='Archivo de Salida:')
l6.grid(row=4,column=1 , ipadx=40)

l7 = tk.Label(frame1, text='Ruta Carpeta')
l7.grid(row=4,column=2)

l51 = tk.Label(frame1, text='')
l51.grid(row=5,column=1)

l52 = tk.Label(frame1, text='')
l52.grid(row=5,column=2)

l71 = tk.Label(frame1, text='Días Analisis:')
l71.grid(row=7,column=1 , ipadx=40)

dias_analisis_var=tk.StringVar()
l72 = tk.Entry(frame1, textvariable=dias_analisis_var)
l72.grid(row=7,column=2)

l81 = tk.Label(frame1, text='')
l81.grid(row=8,column=1)

l82 = tk.Label(frame1, text='')
l82.grid(row=8,column=2)

l91 = tk.Label(frame1, text='Días de Alcance:')
l91.grid(row=9,column=1 , ipadx=40)

dias_alcance_var=tk.StringVar()
l92 = tk.Entry(frame1, textvariable=dias_alcance_var)
l92.grid(row=9,column=2)

l101 = tk.Label(frame1, text='')
l101.grid(row=10,column=1)

l102 = tk.Label(frame1, text='')
l102.grid(row=10,column=2)

l111 = tk.Label(frame1, text='Nombre Archivo de Salida:')
l111.grid(row=11,column=1 , ipadx=40)

nombre_archivo_var=tk.StringVar()
l112 = tk.Entry(frame1, textvariable=nombre_archivo_var)
l112.grid(row=11,column=2)

l121 = tk.Label(frame1, text='')
l121.grid(row=12,column=1)

l122 = tk.Label(frame1, text='')
l122.grid(row=12,column=2)

procesar.grid(row=13,column=2)


l121 = tk.Label(frame1, text='')
l121.grid(row=14,column=1)

l122 = tk.Label(frame1, text='')
l122.grid(row=14,column=2)

limpiar.grid(row=15,column=2)


open_buttonOutPut.grid(row=4, column=3)

#frame 1  -- archivo alcance fin
#
#
#
#
#
#frame 2 -- cruzar datos analista

l1_frame2 = tk.Label(frame2, text='')
l1_frame2.grid(row=1,column=1)

l2_frame2 = tk.Label(frame2, text='')
l2_frame2.grid(row=1,column=2)

l3_frame2 = tk.Label(frame2, text='Archivo Alcance:')
l3_frame2.grid(row=2,column=1, ipadx=40)

l4_frame2 = tk.Label(frame2, text='Ruta Archivo', width=55)
l4_frame2.grid(row=2,column=2)

open_button2.grid(row=2,column=3)

l5_frame2 = tk.Label(frame2, text='')
l5_frame2.grid(row=3,column=1)

l6_frame2 = tk.Label(frame2, text='')
l6_frame2.grid(row=3,column=2)

l7_frame2 = tk.Label(frame2, text='Archivo Resta:')
l7_frame2.grid(row=4,column=1, ipadx=40)

l8_frame2 = tk.Label(frame2, text='Ruta Archivo', width=55)
l8_frame2.grid(row=4,column=2)

open_button3.grid(row=4,column=3)

l55_frame2 = tk.Label(frame2, text='')
l55_frame2.grid(row=5,column=1)

l65_frame2 = tk.Label(frame2, text='')
l65_frame2.grid(row=5,column=2)

l6_frame2 = tk.Label(frame2, text='Archivo de Salida:')
l6_frame2.grid(row=6,column=1 , ipadx=40)

l76_frame2 = tk.Label(frame2, text='Ruta Carpeta')
l76_frame2.grid(row=6,column=2)
open_button3_output.grid(row=6,column=3)

l57_frame2 = tk.Label(frame2, text='')
l57_frame2.grid(row=7,column=1)

l67_frame2 = tk.Label(frame2, text='')
l67_frame2.grid(row=7,column=2)

l68_frame2 = tk.Label(frame2, text='Nombre Archivo:')
l68_frame2.grid(row=8,column=1 , ipadx=40)

nombre_archivo_var_frame2=tk.StringVar()
l112_frame2 = tk.Entry(frame2, textvariable=nombre_archivo_var_frame2)
l112_frame2.grid(row=8,column=2)

l59_frame2 = tk.Label(frame2, text='')
l59_frame2.grid(row=9,column=1)

l69_frame2 = tk.Label(frame2, text='')
l69_frame2.grid(row=9,column=2)

procesar_frame2.grid(row=10, column=2)

l511_frame2 = tk.Label(frame2, text='')
l511_frame2.grid(row=11,column=1)

l611_frame2 = tk.Label(frame2, text='')
l611_frame2.grid(row=11,column=2)


limpiar_frame2.grid(row=12,column=2)

frame1.pack(fill="both", expand=True)
frame2.pack(fill="both", expand=True)



notebook.add(frame1, text="Archivo Alcance")
notebook.add(frame2, text="Cruzar Datos")


ws.mainloop()