# -*- coding: utf-8 -*-

##############################################################################
#
#

import calendar
import pprint
import xlsxwriter

totalHoras = 11
cantidadLabs = 4
meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
dias = ["lunes","martes","miercoles","jueves","viernes","sabado","domingo"]
cantidadColumnas=totalHoras+1
last_col=(len(dias)+1)*cantidadColumnas
mes=9

def escribirHoras(fila = 2):
    inicio = 0
    for i in range(len(dias)):
        for j in range(cantidadColumnas):
            if j==0:
                worksheet.write(fila, inicio+j, "DN", tituloDia)
            else:
                # Write some numbers, with row/column notation.
                worksheet.write(fila, inicio+j, j, tituloDia)
        inicio=inicio+cantidadColumnas 

def escribirDias(fila=2):
    #merge_range(first_row, first_col, last_row, last_col, cell_format)
    cont=0
    for i in range(cantidadColumnas,last_col,cantidadColumnas):
        # Write some numbers, with row/column notation.
        worksheet.merge_range(fila, i-cantidadColumnas,fila,i-1, str(dias[cont]),tituloTabla)
        cont=cont+1

def escribirNumeroDia(fila=2, semanas=4, laboratorios=4, diaSemana=[1,2,3,4,5]):
    cont=0
    for i in range(cantidadColumnas,last_col,cantidadColumnas):
        worksheet.merge_range(fila, i-cantidadColumnas,fila+laboratorios-1,i-cantidadColumnas, str(diaSemana[cont]),contenidoTabla)
        cont=cont+1
    return i
    

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('calendario.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 20)

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

arrayContenido = {           
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'} 
contenidoTabla = workbook.add_format(arrayContenido)
tituloTabla = workbook.add_format({           
            'border': 2,
            'align': 'center',
            'valign': 'vcenter',
            'bold': True})
tituloDia = workbook.add_format({           
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bold': True})

#year = input("Año? ")
#month = input ("Mes ")

#mes=month-1

#------------------------------------------

filaMes = 0
for i in range(len(meses)):
    
    #diaNumero=calendar.monthcalendar(year, month)
    diaNumero=calendar.monthcalendar(2014, i+1)
    worksheet.merge_range(filaMes,0,filaMes,last_col-13, meses[i], tituloTabla)
    escribirDias(filaMes+1)
    escribirHoras(filaMes+2)
    for j in range(len(diaNumero)):
        #merge_range(first_row, first_col, last_row, last_col, cell_format)
        escribirNumeroDia(filaMes+3,len(diaNumero),cantidadLabs, diaNumero[j])
        filaMes=filaMes+cantidadLabs
    filaMes=filaMes+len(diaNumero)
workbook.close()

# Insert an image.
#worksheet.insert_image('B5', 'logo.png')
   

    
