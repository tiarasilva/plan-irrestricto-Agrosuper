from constants import *
from tkinter import *
from tkinter import ttk, messagebox
import tkinter as tk
from PIL import Image, ImageTk

def messageBox(dict_datos_modificados, dict_otro_tipo):
  app = Tk()
  app.configure(bg='#ffffff')

  # Info Image
  information_image = Image.open("Notice.png")
  information_image = information_image.resize((50, 50))
  information_image = ImageTk.PhotoImage(information_image)
  logo_label = tk.Label(image = information_image, borderwidth=0, bg = 'white')
  logo_label.image = information_image
  logo_label.grid(column = 0, row = 1, padx = 10, pady = 10)

  # Instrucciones
  instructions = tk.Label(app, text = "Alerta falta de información", bg="white", justify=LEFT)
  instructions.configure(font=("bold", 16))
  instructions.grid(column = 1, row = 1)

  # Description
  messageboxText = "En las siguientes columnas no se encontró la información en los excel de Colaboraciones. Estos fueron rellanaron con los siguientes datos: Para el volumen de contenedor promedio fue con 24.000 ton, para el porcentaje de utilización con un 35%. \n\n*En el excel fueron marcadas con rojo para mayor detalle"
  text = Text(app, bg = '#ffffff', bd = 0, borderwidth=0, highlightthickness=0, height=5, width=70, padx = 15, pady = 5)
  text.insert(INSERT, messageboxText)
  text.config(state=DISABLED, font="Calibrí")
  text.grid(column = 0, row = 2, columnspan = 12)

  # Table
  tree = ttk.Treeview(app, columns=(1, 2, 3, 4, 5), show="headings", height="8", padding = (20, 20, 20, 20))
  # vsb = Scrollbar(app, orient="vertical", command=tree.yview)
  # vsb.place(x=30+200+2, y=95, height=200+20)
  # tree.configure(yscrollcommand=vsb.set)

  tree.column(1, anchor=CENTER, stretch=NO, width=70)
  tree.heading(1, text = 'Fila')
  tree.column(2, anchor=CENTER, stretch=NO, width=200)
  tree.heading(2, text = 'Llave')
  tree.column(3, anchor=CENTER, stretch=NO, width=180)
  tree.heading(3, text = 'Nombre')
  tree.column(4, anchor=CENTER, stretch=NO, width=100)
  tree.heading(4, text = 'Valor original')
  tree.column(5, anchor=CENTER, stretch=NO, width=100)
  tree.heading(5, text = 'Cambiado a')

  row = 0
  for key, value in dict_datos_modificados.items():
    tree.insert('', 'end', values = (key, value['llave'], value['name'], value['original_value'], value['change_value']))
    row += 1
  tree.grid(column = 0, row = 3, columnspan = 12)

  # Arbol de datos de otro tipo de venta
  messageboxText2 = f"Además, se eliminaron los siguientes pedidos produccidos que no corresponden al tipo de venta {seleccion_tipo_venta}."
  text = Text(app, bg = '#ffffff', bd = 0, borderwidth=0, highlightthickness=0, height=2, width=70, padx = 5, pady = 5)
  text.insert(INSERT, messageboxText2)
  text.config(state=DISABLED, font="Calibrí")
  text.grid(column = 0, row = 4, columnspan = 12)

  tree_otra_tipo_venta = ttk.Treeview(app, columns=(1, 2, 3, 4), show="headings", height="8", padding=(10, 10, 10, 10))
  tree_otra_tipo_venta.column(1, anchor=CENTER, stretch=NO, width=70)
  tree_otra_tipo_venta.heading(1, text = 'SKU')
  tree_otra_tipo_venta.column(2, anchor=CENTER, stretch=NO, width=180)
  tree_otra_tipo_venta.heading(2, text = 'Oficina')
  tree_otra_tipo_venta.column(3, anchor=CENTER, stretch=NO, width=300)
  tree_otra_tipo_venta.heading(3, text = 'Descripción')
  tree_otra_tipo_venta.column(4, anchor=CENTER, stretch=NO, width=100)
  tree_otra_tipo_venta.heading(4, text = 'Producción Mes')

  for key, value in dict_otro_tipo.items():
    tree_otra_tipo_venta.insert('', 'end', values = (value['sku'], value['Oficina'], value['Descripción'], value['Producción mes']))
  tree_otra_tipo_venta.grid(column = 0, row = 5, columnspan = 12)

  # Exit button
  boton = Button(app, text = "Salir", command = app.destroy, width=8, highlightbackground='#ffffff')
  boton.grid(column = 10, row = 6, pady = 20)


  app.title('Alerta falta de datos')
  app.eval('tk::PlaceWindow . center')
  app.mainloop()