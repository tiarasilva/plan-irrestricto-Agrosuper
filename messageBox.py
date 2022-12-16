from tkinter import *

app = Tk()
app.title('Alerta falta de datos')

part_text = StringVar()
label = Label(app, text="En las siguientes columnas no se encontró la información y se rellanaron con datos. Están coloreadas rojo para mayor detalle", pady=20)
label.grid(row = 0, column = 0)
app.geometry('700x300')
app.mainloop()