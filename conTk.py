# -*- coding: utf-8 -*- 
# Simple calendario con tkinter 
import calendar 
import Tkinter as tk 
import datetime 
# Obtenemos los valores del año y mes a mostrar 

year = datetime.date.today().year 
month = datetime.date.today().month 

# Asignamos el año y mes al calendario 

str1 = calendar.month(year, month) 
root = tk.Tk() 
root.title("Calendario") 

print str1

# Lo posicionamos en un label 

label1 = tk.Label(root, text=str1, font=('courier', 14, 'bold'), bg='white') 
label1.pack(padx=3, pady=5) 

print "Funco?"
# ejecutamos el evento loop 

root.mainloop()
