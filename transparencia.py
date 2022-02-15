import os 
from xlrd import open_workbook
from openpyxl import Workbook, load_workbook

path = "/Users/eduardomurrieta/repositories/github/valenzuela/excels"

os.chdir(path)
# wb = Workbook()
#Abrimos el archivo para leerlo
transparencia_excel = load_workbook("{}/{}".format(path,'transparencia.xlsx'))
sheet_transparencia_excel = transparencia_excel["Sheet"]
sheet_transparencia_excel_2 = transparencia_excel["Sheet2"]

# sheets = transparencia_excel.sheetnames
# grab the active worksheet

# ws = wb.active
# wb.save("transparencia.xlsx")
# wb.create_sheet('Sheet2')
# wb.save('transparencia.xlsx')

def insertar(row_count, cur_col,maxima_row,sheet, salto_row, que_columna, sheet2, is_number, file):
    for cur_row in range(salto_row, row_count):
        cell = sheet.cell(cur_row, cur_col)                    
        if (cur_row-1) == salto_row:
            continue
        if cell.value:
            if(sheet2):
                transparencia_excel.active = 1
                if (is_number):
                    sheet_transparencia_excel_2.cell(row=(maxima_row+cur_row-salto_row-1), column=que_columna, value=float(cell.value.strip()))
                else:
                    sheet_transparencia_excel_2.cell(row=(maxima_row+cur_row-salto_row-1), column=que_columna, value=cell.value.strip())                        
            else:
                transparencia_excel.active = 0
                if (isinstance(cell.value,int) or isinstance(cell.value,float)):
                    sheet_transparencia_excel.cell(row=(maxima_row+cur_row-salto_row-1), column=que_columna, value=cell.value)
                else:
                    sheet_transparencia_excel.cell(row=(maxima_row+cur_row-salto_row-1), column=que_columna, value=cell.value.strip())
                sheet_transparencia_excel.cell(row=(maxima_row+cur_row-salto_row-1), column=46, value=file)
def iiter_another_table(book,name_table):
    sh = book.sheet_by_name(name_table)
    row_count = sh.nrows
    print('hoja secundaria')
    print(row_count)
    col_count = sh.ncols
    maxima_row = sheet_transparencia_excel_2.max_row
    iter_columnas(col_count , sh, maxima_row, row_count, 2,1, book, True, '')

def iter_columnas(col_count , sheet, maxima_row, row_count, row, salto_inicio, book, sheet2, file):
    for cur_col in range(0, col_count):
        transparencia_excel.active = 0
        cell = sheet.cell(row, cur_col)
        if ('ID' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 1, sheet2, True, file)
        if (' Ejercicio' == cell.value or 'EJERCICIO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 2, sheet2, True, file)
        if (' Tipo de Procedimiento (catálogo)' == cell.value or ' Tipo de Procedimiento' == cell.value or 'TIPO DE PROCEDIMIENTO (CATÁLOGO)' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 3, sheet2, False, file)
        if (' Materia' == cell.value or ' Materia O Tipo de Contratación (catálogo)' == cell.value or ' Categoría:' == cell.value or 'MATERIA (CATÁLOGO)' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 4, sheet2, False, file)
        if (' Número de Expediente, Folio O Nomenclatura' == cell.value or 'NÚMERO DE EXPEDIENTE  FOLIO O NOMENCLATURA QUE LO IDENTIFIQUE' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 5, sheet2, False, file)
        if (' Nombre(s) Del Adjudicado' == cell.value or ' Nombre(s) Del Contratista O Proveedor' ==  cell.value or ' Nombre(s)' == cell.value or 'NOMBRE(S) DEL ADJUDICADO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 6, sheet2, False, file)
        if (' Primer Apellido Del Adjudicado' == cell.value or ' Primer Apellido' == cell.value or ' Primer Apellido Del Contratista O Proveedor' == cell.value or 'PRIMER APELLIDO DEL ADJUDICADO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 7, sheet2, False, file)
        if (' Segundo Apellido Del Adjudicado' == cell.value or ' Segundo Apellido' == cell.value or ' Segundo Apellido Del Contratista O Proveedor' == cell.value or 'SEGUNDO APELLIDO DEL ADJUDICADO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 8, sheet2, False, file)
        if (' Nombre O Razón Social Del Adjudicado (Tabla' in cell.value):
            transparencia_excel.save("transparencia.xlsx")
            transparencia_excel.active = 1
            tabla = cell.value.split('Adjudicado (', 1)[1].replace(")","").strip().upper().replace("_","")
            iiter_another_table(book, tabla)
        if (' Nombre Completo Del O Los Contratista(s) Elegidos (Tabla' in cell.value):
            transparencia_excel.save("transparencia.xlsx")
            transparencia_excel.active = 1
            tabla = cell.value.split('Elegidos (', 1)[1].replace(")","").strip().upper().replace("_","")
            iiter_another_table(book, tabla)
        if (' Origen de Los Recursos Públicos (Tabla' in cell.value):
            transparencia_excel.save("transparencia.xlsx")
            transparencia_excel.active = 1
            tabla = cell.value.split('Públicos (', 1)[1].replace(")","").strip().upper().replace("_","")
            iiter_another_table(book, tabla)
        if (' Denominación O Razón Social' == cell.value or
            ' Denominación O Razón Social Del Contratista' == cell.value or
            ' Razón Social' == cell.value or
            ' Razón Social Del Contratista O Proveedor' == cell.value or
            ' Nombre O Razón Social Del Adjudicado' in cell.value or
            'RAZÓN SOCIAL DEL ADJUDICADO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 9, sheet2, False, file)
        if (' Rfc de La Persona Física O Moral Contratista O Proveedor' == cell.value or ' Registro Federal de Contribuyentes (rfc) de La Persona Física O Moral Adjudicada' == cell.value or 'REGISTRO FEDERAL DE CONTRIBUYENTES (RFC) DE LA PERSONA FÍSICA O MORAL ADJUDICADA' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 10, sheet2, False, file)
        if (' Área(s) Solicitante' == cell.value or 'ÁREA(S) SOLICITANTE(S)' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 11, sheet2, False, file)
        if (' Área(s) Contratante(s)' == cell.value or ' Unidad Administrativa Contratante' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 12, sheet2, False, file)
        if (' Número Que Identifique Al Contrato' == cell.value or 'NÚMERO QUE IDENTIFIQUE AL CONTRATO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 13, sheet2, True, file)
        if (' Monto Del Contrato sin Impuestos (en Pesos Mex.)' == cell.value or ' Monto Del Contrato sin Impuestos (en Mxn)' == cell.value or
            'MONTO DEL CONTRATO SIN IMPUESTOS INCLUIDOS' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 14, sheet2, True, file)
        if (' Monto Total Del Contrato con Impuestos Incluidos' == cell.value or ' Monto Total Del Contrato con Impuestos Incluidos (mxn)' == cell.value or
            'MONTO TOTAL DEL CONTRATO CON IMPUESTOS INCLUIDOS (EXPRESADO EN PESOS MEXICANOS)' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 15, sheet2, True, file)
        if (' Tipo de Moneda' == cell.value or 'TIPO DE MONEDA' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 16, sheet2, False, file)
        if (' Tipo de Cambio de Referencia, en su Caso' == cell.value or 'TIPO DE CAMBIO DE REFERENCIA  EN SU CASO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 17, sheet2, False, file)
        if (' Forma de Pago' == cell.value or 'FORMA DE PAGO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 18, sheet2, False, file)
        if (' Origen de Los Recursos Públicos (catálogo)' == cell.value or ' Origen de Los Recursos Públicos' == cell.value or 'ORIGEN DE LOS RECURSOS PÚBLICOS' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 19, sheet2, False, file)
        if (' Fuente de Financiamiento' == cell.value or 'FUENTES DE FINANCIAMIENTO' cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 20, sheet2, False, file)
        if (' Lugar Donde Se Realizará La Obra Pública, en su Caso' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 21, sheet2, False, file)
        if (' Descripción de Las Obras, Bienes O Servicios' == cell.value or 'DESCRIPCIÓN DE OBRAS  BIENES O SERVICIOS' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 22, sheet2, False, file)
        if (' Fecha de La Convocatoria O Invitación' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 23, sheet2, False, file)
        if (' Fecha Del Contrato' == cell.value or 'FECHA DEL CONTRATO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 24, sheet2, False, file)
        if (' Hipervínculo Al Comunicado de Suspensión, en su Caso' == cell.value or 'HIPERVÍNCULO AL COMUNICADO DE SUSPENSIÓN  RESCISIÓN O TERMINACIÓN ANTICIPADA DEL CONTRATO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 25, sheet2, False, file)
        if (' Hipervínculo Al (los) Dictámenes, en su Caso' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 26, sheet2, False, file)
        if (' Hipervínculo a La Convocatoria O Invitaciones' == cell.value or ' Hipervínculo a La Convocatoria O Invitaciones Emitidas' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 27, sheet2, False, file)
        if (' Hipervínculo Al Fallo de La Junta de Aclaraciones O Al Documento Correspondiente' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 28, sheet2, False, file)
        if (' Hipervínculo Al Documento Donde Conste La Presentación Las Propuestas' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 29, sheet2, False, file)
        # #obras públicas campo aun faltantes que no tenia
        if (' Descripción de Las Razones Que Justifican su Elección' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 30, sheet2, False, file)
        if (' Área(s) Responsable de su Ejecución' == cell.value or 'ÁREA(S) RESPONSABLE(S) DE LA EJECUCIÓN DEL CONTRATO' == cell.value or
            'ÁREA(S) RESPONSABLE(S) QUE GENERA(N)  POSEE(N)  PUBLICA(N) Y ACTUALIZAN LA INFORMACIÓN' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 31, sheet2, False, file)
        if (' Objeto Del Contrato' == cell.value or 'OBJETO DEL CONTRATO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 32, sheet2, False, file)
        if (' Fecha de Inicio Del Plazo de Entrega O Ejecución' == cell.value or 'FECHA DE INICIO DEL PLAZO DE ENTREGA O EJECUCIÓN DE SERVICIOS CONTRATADOS U OBRA PÚBLICA' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 33, sheet2, False, file)
        if (' Fecha de Término Del Plazo de Entrega O Ejecución' == cell.value or 'FECHA DE TÉRMINO DEL PLAZO DE ENTREGA O EJECUCIÓN DE SERVICIOS U OBRA PÚBLICA' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 34, sheet2, False, file)
        if (' Hipervínculo Al Documento Del Contrato Y Anexos, en Versión Pública, en su Caso' == cell.value or
            'HIPERVÍNCULO AL DOCUMENTO DEL CONTRATO Y ANEXOS  VERSIÓN PÚBLICA SI ASÍ CORRESPONDE' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 35, sheet2, False, file)
        if ('ESTATUS CONTRATO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 36, sheet2, False, file)
        if (' Tipo de Fondo de Participación O Aportación Respectiva' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 37, sheet2, False, file)
        if (' Breve Descripción de La Obra Pública, en su Caso' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 38, sheet2, False, file)
        if (' Etapa de La Obra Pública Y/o Servicio de La Misma (catálogo)' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 39, sheet2, False, file)
        if (' Mecanismos de Vigilancia Y Supervisión de La Ejecución, en su Caso' == cell.value or
            'MECANISMOS DE VIGILANCIA Y SUPERVISIÓN CONTRATOS' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 40, sheet2, False, file)
        if (' Hipervínculo a Los Informes de Avance Financiero, en su Caso' == cell.value or
            'HIPERVÍNCULO A LOS INFORMES DE AVANCE FINANCIERO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 41, sheet2, False, file)
        if (' Monto Total de Garantías Y/o Contragarantías, en Caso de Que Se Otorgaran durante El Procedimiento' == cell.value or
            'MONTO TOTAL DE GARANTÍAS Y/O CONTRAGARANTÍAS  EN CASO DE QUE SE OTORGARAN DURANTE EL PROCEDIMIENTO' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 42, sheet2, False, file)
        if ('Institucion' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 43, sheet2, False, file)
        if ('Municipio/Estado' == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 44, sheet2, False, file)
        if (" Carácter Del Procedimiento (catálogo)" == cell.value):
            insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 45, sheet2, False, file)
        insertar(row_count, cur_col, maxima_row, sheet, salto_inicio, 46, sheet2, False, file)

def parse_home():
    for file in os.listdir():
        if file in ['.DS_Store']:
            continue 
        if file in ['transparencia.xlsx']:
            continue
        print("{}/{}".format(path,file))
        try:
            book = open_workbook("{}/{}".format(path,file),on_demand=True)
        except:
            print('este no se pudo abrir {}'.format(file))
            continue
        # data = open_file.values
        sheet = book.sheet_by_index(0)

        row_count = sheet.nrows
        print(row_count)
        col_count = sheet.ncols
        # maxima_columna = sheet_transparencia_excel.max_column
        maxima_row = sheet_transparencia_excel.max_row
        transparencia_excel.save("transparencia.xlsx")
        if (sheet.cell_value(rowx=3, colx=0) == 'Periodos:'):
            iter_columnas(col_count , sheet, maxima_row, row_count, 5, 4, book, False, file)
        else:
            iter_columnas(col_count , sheet, maxima_row, row_count, 4, 3, book, False, file)

def run():
    parse_home()
if __name__ == '__main__':
    run()