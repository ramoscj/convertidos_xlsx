import xlsxwriter
import datetime

def crearArchivoXlsx(nombreArchivo, valoresArchivo:[]):
    workbook = xlsxwriter.Workbook('{0}.xlsx'.format(nombreArchivo))

    columnas = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z')

    bold2 = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'locked': True, 'border': 2, 'bg_color': '#FFA500'})

    format_data = workbook.add_format({'font_name': 'Arial', 'font_size': 9, 'align': 'center', 'locked': True, 'border': 1})

    try:
        for valoresEntrada in valoresArchivo:
            nombreHoja = valoresEntrada[0]
            encabezado = valoresEntrada[1]
            data =  valoresEntrada[2]
            worksheet = workbook.add_worksheet(nombreHoja)
            row = 0
            col = 0
            for valor in encabezado:
                ajustar = '%s%s:%s%s' % (columnas[col], 1, columnas[col], 1)
                worksheet.set_column(ajustar, 30)
                worksheet.write(row, col, valor, bold2)
                col += 1
            row += 1
            for key, dataValor in data.items():
                col = 0
                for llave, valor in dataValor.items():
                    if type(valor) is datetime.date:
                        valor = valor.strftime("%d/%m/%Y")
                    worksheet.write(row, col, valor, format_data)
                    col += 1
                row += 1
        workbook.close()
        return True
    except Exception as e:
        raise Exception('Error en crearArchivoXlsx: %s' % e)
