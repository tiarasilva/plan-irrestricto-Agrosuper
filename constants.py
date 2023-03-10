#-------------------------- CONSTANTES ---------------------------#
year = 2022

seleccion_tipo_venta = "Venta Directa" # Venta Directa o Venta Local

filename = 'Plan de ventas irrestricto.xlsx'
filename_asignacion = 'Colaboraciones plan de ventas/Asignación venta.xlsx'
filename_util = 'Colaboraciones plan de ventas/Fechas de zarpe - Logística.xlsx'
filename_por_despachar = 'Colaboraciones plan de ventas/Proyeccion Plan de Venta.xlsx'
filename_stock = 'Colaboraciones plan de ventas/Pedidos Stock.xlsx'
filename_por_producir = 'Colaboraciones plan de ventas/Producción.xlsx'
filename_vol_cont = 'Colaboraciones plan de ventas/Volumen por contenedor.xlsx'
filename_puerto = 'Colaboraciones plan de ventas/Pedidos Planta-Puerto-Embarcado.xlsx'
filename_maestro_materiales = 'Colaboraciones plan de ventas/Maestro de materiales.xlsx'
filename_parametros = 'Colaboraciones plan de ventas/Parametros.xlsx'
path_img = "Img/Notice.png"

PO = 'Pollo'
PV = 'Pavo'
GO = 'Cerdo'
GA = 'Pollo'                                            # o Gallina

# Definición de colores
grey = 'aaabac'
lightlightBlue = 'dbe5f1'
lightBlue = '8ba9d7'
lightPale = 'FFDFCC'
lightOrange = 'f8c9ad'
orange = 'c14811'
blue = '2f5496'
white = 'ffffff'
lightRed = 'fec7cd'
red = 'FF0000'
darkRed = 'b0292f'
yellow = 'ffeb9c'
lightGreen = 'c6eecd'
green = '006100'

tamano = {
  'Llave': 23,
  'Sector': 8,                   
  'Oficina': 16,             
  'Material': 9,
  'Descripción': 27,
  'Nivel Jer. 2': 15,
  'Nivel Jer. 3': 17,
  'RV Producción mes N + 1': 14,
  'RV Venta mes N + 1': 14,
  '% Uti. producción': 7,
  'Producción disponible': 14,
  'Stock al día': 14,
  'Por producir mes N': 14,
  'Producción por despachar mes N': 14,
  'Total disponible': 14,
  'Delay': 14,
  'Vol. prom. Por contenedor': 14,
  'Atraso a facturar': 14,
  'Ajuste atraso': 14,
  'Facturación atraso': 14,
  'Producción para venta nueva': 14,
  'Venta del mes': 14,
  'Ajuste venta nueva': 14,
  'Facturación Venta nueva': 14,
  'Saldo Volumen disponible': 14,
  'Disponible stock sin venta': 14,
  'Ajuste stock sin venta': 14,
  'Facturación stock': 14,
  'En puerto a facturar': 14,
  'Plan Irrestricto': 14,
  'Plan Ajustado': 14,
  'Motivo Ajuste': 30
}

month_translate_CL_EN = {
  'enero': 'january',
  'febrero': 'february',
  'marzo': 'march',
  'abril': 'april', 
  'mayo': 'may',
  'junio': 'june',
  'julio': 'july', 
  'agosto': 'august',
  'septiembre': 'september',
  'octubre': 'october',
  'noviembre': 'november',
  'diciembre': 'december'
}

month_translate_EN_CL = {
  'January': 'Enero',
  'February': 'Febrero',
  'March': 'Marzo',
  'April': 'Abril', 
  'May': 'Mayo',
  'June': 'Junio',
  'July': 'Julio', 
  'August': 'Agosto',
  'September': 'Septiembre',
  'October': 'Octubre',
  'November': 'Noviembre',
  'December': 'Diciembre'
}

week_translate_EN_CL = {
  'Monday': 'Lunes',
  'Tuesday': 'Martes',
  'Wednesday': 'Miércoles',
  'Thursday': 'Jueves',
  'Friday': 'Viernes',
  'Saturday': 'Sábado',
  'Sunday': 'Domingo'
}
month_number = {
  'enero': 1,
  'febrero': 2,
  'marzo': 3,
  'abril': 4, 
  'mayo': 5,
  'junio': 6,
  'julio': 7, 
  'agosto': 8,
  'septiembre': 9,
  'octubre': 10,
  'noviembre': 11,
  'diciembre': 12   
}