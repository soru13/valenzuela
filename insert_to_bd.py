from bd import conexion
import csv, ipdb, psycopg2, os
from openpyxl import load_workbook
path = "/app/excels"
os.chdir(path)
transparencia_excel = load_workbook("{}/{}".format(path,'transparencia.xlsx'))
sheet_transparencia_excel = transparencia_excel["Sheet"]
sheet_transparencia_excel_2 = transparencia_excel["Sheet2"]

def consultarId(tabla, propiedad, valor):
    try:
        with conexion.cursor() as cursor:
            consulta_id = "select id from {} where {}= '{}';".format(tabla,propiedad,valor)
            print(consulta_id)
            cursor.execute(consulta_id)

            # Con fetchall traemos todas las filas
            retorno = cursor.fetchall()
            print(retorno)
            return retorno[0][0]
    except psycopg2.Error as e:
        print("Ocurrió un error al consultar con where: ", e)
        conexion.close()
    else:
        conexion.close()

def insertar(tabla, valor):
    try:
        with conexion.cursor() as cursor:
            insert = "INSERT INTO %s (name, usuario_registro_id, created, modified) VALUES (%s, 1, current_timestamp, current_timestamp);"
            # Podemos llamar muchas veces a .execute con datos distintos
            cursor.execute(insert, (tabla, valor))
        conexion.commit()  # Si no haces commit, los cambios no se guardan
    except psycopg2.Error as e:
        print("Ocurrió un error al insertar: ", e)
    finally:
        conexion.close()

def insertarTransparencia(tabla, valor):
    try:
        with conexion.cursor() as cursor:
            insert = "INSERT INTO %s (name, usuario_registro_id, created, modified) VALUES (%s, 1, current_timestamp, current_timestamp);"
            # Podemos llamar muchas veces a .execute con datos distintos
            cursor.execute(consulta, (tabla, valor))
        conexion.commit()  # Si no haces commit, los cambios no se guardan
    except psycopg2.Error as e:
        print("Ocurrió un error al insertar: ", e)
    finally:
        conexion.close()
def parse_home():
    with open('./excels/hoja1.csv', newline='') as File:  
        reader = csv.reader(File)
        # print(reader)
        line_count = 0
        for row in reader:
            # print(row)
            if line_count == 0:
                # print(f'Column names are {", ".join(row)}')
                line_count += 1
            # if line_count == 1:
            #     print(f'Column names are {", ".join(row)}')
            #     line_count += 1
            #     insertar('ejercicio',row[line_count])
            # if line_count == 2:
            #     print(f'Column names are {", ".join(row)}')
            #     line_count += 1
            #     insertar('tipo_procedimiento',row[line_count])
            else:
                # print(f'\t{row[0]} works in the {row[1]} department, and was born in {row[2]}.')
                line_count += 1
                # print(range(len(row)))
                for i in range(len(row)):

                    # print(row[i])
                    if i == 0:
                        print(row[i])
                        code_id = row[i]
                    if i == 1:
                        # print(row[i])
                        # ejercicio = row[i]
                        print('ejercicio')

                    if i == 2:
                        print('tipo_procedimiento')
                        # if ("ADJUDICACION DIRECTA" == row[i] or
                        #     "ADJUDICACION  DIRECTA" == row[i] or
                        #     "Adjudicación directa" == row[i] or
                        #     "ADJUDIACION DIRECTA" == row[i] or
                        #     "Adjudicación Directa" == row[i] or
                        #     "Adjudicacion Directa" == row[i] or
                        #     "adjudicación directa" == row[i] or
                        #     "ADJUDICACIÓN DIRECTA" == row[i] or
                        #     "ADJUDICACION DIRECTA CON RECURSOS PROPIOS" == row[i]):
                        #     # tipo_procedimiento = consultarId('tipo_procedimiento', 'name','Adjudicación directa')
                        # elif ("COMPRA DIRECTA" == row[i]):
                        #     pass
                        #     # tipo_procedimiento = consultarId('tipo_procedimiento', 'name','compra directa')
                        # else:
                        #     pass
                        #     # print(row[i])
                        #     # tipo_procedimiento = consultarId('tipo_procedimiento', 'name',row[i])
                        #     # print(tipo_procedimiento)
                    if i ==3:
                        print('materia_tipo')
                        # if("Adquisición" == row[i] or
                        #     "Adquisiciones" == row[i]):
                        #     print(row[i])
                        #     # materia_tipo = consultarId('materia_tipo', 'name','Adquisiciones')
                        # else:
                        #     pass
                        #     # materia_tipo = consultarId('materia_tipo', 'name',row[i])
                    if i ==4:
                        print('folio')
                        folio = row[i]
                    if i == 5:
                        print('nombre_id')
                        nombre_id = ''
                        print(row[i])
                        print('nombre_id')
                        print(type(row[i]))


                        if (isinstance(row[i],int) or isinstance(row[i],float)):
                            print('nombre id')
                            print(row[i])
                            print('nombre id')
                            nombre_id = consultarId('full_name', 'code_id',row[i])
                        else:
                            nombres = row[i]

                    # if i ==4:

                    # insertarTransparencia()
                #     break
                # for item in row:
                #     print(item)
                    # insertar(item)
            # ipdb.set_trace()

def run():
    print(sheet_transparencia_excel.max_row)
    for col in range(0, sheet_transparencia_excel.max_column):
        for row in range(1, sheet_transparencia_excel.max_row):
            fila = sheet_transparencia_excel.cell(row= row, column = col+1)
            print(fila.value)
    # parse_home()
if __name__ == '__main__':
    run()