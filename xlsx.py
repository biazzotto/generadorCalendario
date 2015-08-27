# -*- coding: utf-8 -*-
#Para el excel
import xlsxwriter


try:
    import cStringIO as StringIO
except ImportError:
    import StringIO

#Descarga Libro iva
def printIva(request, idIva):  
    iva = RegistroIva.objects.get(id=idIva)
    # Create the HttpResponse object with the appropriate PDF headers.   

     # create a workbook in memory
    output = StringIO.StringIO()
   
    workbook = xlsxwriter.Workbook(output)   
    arrayContenido = {           
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'}
    arrayMoney = {           
            'border': 1,
            'align': 'rigth',
            'valign': 'vcenter',
            'num_format': '[$$-2C0A]   #.#0'}   
    contenidoTabla = workbook.add_format(arrayContenido)
    money = workbook.add_format(arrayMoney)
   
    def addHoja(worksheet, tipoLibro):
        negrita = workbook.add_format()
        negrita.set_bold()   
       
               
        worksheet.set_column('A:C', 15)
        worksheet.set_column('D:D', 40)
        worksheet.set_column('E:S', 15)
        worksheet.write('A1', 'IMPRESORA DEL CENTRO S.R.L.', negrita)
        worksheet.write('A2', u'DOMICILIO: JULIO CESAR LASTRA 2220 - Bº SANTA ISABEL 1º SECCIÓN - CÓRDOBA', negrita)
        worksheet.write('A3', 'CUIT: 30-71103466-4', negrita)
        worksheet.write('A4', 'IVA RESPONSABLE INSCRIPTO', negrita)   
        worksheet.write('E4', 'IVA %s' % tipoLibro, negrita)
        worksheet.write('E6', 'PERIODO: ', negrita)
        worksheet.write('F6', '%s' % iva.periodo(), negrita)
        ##CREANDO TITULOS TABLA
        tituloTabla = workbook.add_format({           
            'border': 2,
            'align': 'center',
            'valign': 'vcenter'})
        worksheet.merge_range('A8:A9', 'FECHA', tituloTabla)
        worksheet.merge_range('B8:C8', 'COMPROBANTE', tituloTabla)
        worksheet.write('B9', 'TIPO',tituloTabla)
        worksheet.write('C9', u'NÚMERO',tituloTabla)
        worksheet.merge_range('D8:D9', u'NOMBRE Y APELLIDO O RAZÓN SOCIAL', tituloTabla)
        worksheet.merge_range('E8:E9', u'C.U.I.T.', tituloTabla)
        if tipoLibro == 'COMPRAS':  
            worksheet.merge_range('F8:F9', u'TOTAL\nFACTURADO', tituloTabla)
            worksheet.merge_range('G8:J8', u'NETO GRAVADO', tituloTabla)
            worksheet.write('G9', '21%',tituloTabla)
            worksheet.write('H9', '27%',tituloTabla)
            worksheet.write('I9', '17,355%',tituloTabla)
            worksheet.write('J9', '10,50%',tituloTabla)
            worksheet.merge_range('K8:N8', u'IVA LOQUIDADO', tituloTabla)
            worksheet.write('K9', '21%',tituloTabla)
            worksheet.write('L9', '27%',tituloTabla)
            worksheet.write('M9', '17,355%',tituloTabla)
            worksheet.write('N9', '10,50%',tituloTabla)
            worksheet.merge_range('O8:O9', u'COMPRAS\nFACT. C/B', tituloTabla)
            worksheet.merge_range('P8:P9', u'CONCEPTO\nNO GRAV.', tituloTabla)
            worksheet.merge_range('Q8:Q9', u'RETENCIÓN\nIVA', tituloTabla)
            worksheet.merge_range('R8:R9', u'RETENCIÓN\nGANANCIAS', tituloTabla)
            worksheet.merge_range('S8:S9', u'IMP. CTA', tituloTabla)
        else:
            worksheet.merge_range('F8:F9', u'COND', tituloTabla)
            worksheet.merge_range('G8:G9', u'TOTAL\nFACTURA', tituloTabla)
            worksheet.merge_range('H8:I8', u'NETO GRAVADO', tituloTabla)
            worksheet.write('H9', '21%',tituloTabla)
            worksheet.write('I9', '10,5%',tituloTabla)
            worksheet.merge_range('J8:K8', u'IVA LIQUIDADO', tituloTabla)
            worksheet.write('J9', '21%',tituloTabla)
            worksheet.write('K9', '10,5%',tituloTabla)
            worksheet.merge_range('L8:L9', u'EXENTOS', tituloTabla)
            worksheet.merge_range('M8:M9', u'RETEN.', tituloTabla)
           
        return worksheet
       
   
  
    #CARGO LIBRO COMPRAS
    compras = addHoja(workbook.add_worksheet('LIBRO IVA COMPRAS'), 'COMPRAS')
    count = 10   
   
    for fc in iva.facturasCompra():
        compras.write('A%d' % count, str(fc.fecha.strftime('%d/%m/%Y')),contenidoTabla)
        compras.write('B%d' % count, str(fc.letra),contenidoTabla)
        compras.write('C%d' % count, str(fc.numero),contenidoTabla)
        compras.write('D%d' % count, str(fc.proveedor.nombre),contenidoTabla)
        compras.write('E%d' % count, str(fc.proveedor.cuit),contenidoTabla)
        compras.write('F%d' % count, fc.total(),money)
        if (fc.iva=='21'):
            compras.write('G%d' % count, fc.subtotal(),money)
        else:
            compras.write('G%d' % count, '',contenidoTabla)
        if (fc.iva=='27'):
            compras.write('H%d' % count, fc.subtotal(),money)
        else:
            compras.write('H%d' % count, '',contenidoTabla)
        if (fc.iva=='17.355'):
            compras.write('I%d' % count, fc.subtotal(),money)
        else:
            compras.write('I%d' % count, '',contenidoTabla)
        if (fc.iva=='10.5'):
            compras.write('J%d' % count, fc.subtotal(),money)
        else:
            compras.write('J%d' % count, '',contenidoTabla)
       
        if (fc.iva=='21' and fc.letra=='A'):
            compras.write('K%d' % count, fc.subtotal(),money)
        else:
            compras.write('K%d' % count, '',contenidoTabla)
        if (fc.iva=='27' and fc.letra=='A'):
            compras.write('L%d' % count, fc.subtotal(),money)
        else:
            compras.write('L%d' % count, '',contenidoTabla)
        if (fc.iva=='17.355' and fc.letra=='A'):
            compras.write('M%d' % count, fc.subtotal(),money)
        else:
            compras.write('M%d' % count, '',contenidoTabla)
        if (fc.iva=='10.5' and fc.letra=='A'):
            compras.write('N%d' % count, fc.subtotal(),money)
        else:
            compras.write('N%d' % count, '',contenidoTabla)       
        if (fc.letra=='B' or fc.letra=='C'):
            compras.write('O%d' % count, fc.total(),money)
        else:
            compras.write('O%d' % count, '',contenidoTabla)
           
        if (fc.noGravado>0):
            compras.write('P%d' % count, fc.noGravado,money)
        else:
            compras.write('P%d' % count, '',contenidoTabla)
           
        if (fc.retIva>0):
            compras.write('Q%d' % count, fc.retIva,money)
        else:
            compras.write('Q%d' % count, '',contenidoTabla)
       
        if (fc.retGanancias>0):
            compras.write('R%d' % count, fc.retGanancias,money)
        else:
            compras.write('R%d' % count, '',contenidoTabla)
       
        if (fc.retImpCta>0):
            compras.write('S%d' % count, fc.retImpCta,money)
        else:
            compras.write('S%d' % count, '',contenidoTabla)
               
        count = count + 1
       
    #CARGO LIBRO VENTAS
    ventas = addHoja(workbook.add_worksheet('LIBRO IVA VENTAS'), 'VENTAS')   
    factVentas = iva.facturasVenta()       
   
   
    workbook.close()
    #Creando El response
    output.seek(0)
   
    response = HttpResponse(output.read(), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response['Content-Disposition'] = "attachment; filename=RegistroIva%s.xlsx" % (iva.periodo())
    
    print response
   
    return response
