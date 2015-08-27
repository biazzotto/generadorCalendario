# -*- coding: utf-8 -*-
import calendar
import pprint
year = input("Año?")
month = input ("Mes")
meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
#print (calendar.month(year,month,w=2,l=1))
#print "-------------------------------------"
mes=month-1
print mes
print meses[mes]
pprint.pprint(calendar.monthcalendar(year, month))

