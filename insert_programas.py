from curses import beep
from gettext import NullTranslations
from bd import conexion
import psycopg2, os
import pandas as pd
from datetime import datetime

path = "/app/excels/programas_sociales/csv"
os.chdir(path)

def consultarIdBeneficiario(tabla, nombres, apellido_paterno, apellido_materno):
    consulta_id = ''
    try:
        with conexion.cursor() as cursor:
            consulta_id = "select id from {} where names ='{}' and last_name ='{}' and surname='{}';".format(tabla,nombres,apellido_paterno,apellido_materno)
            cursor.execute(consulta_id)
            # Con fetchall traemos todas las filas
            retorno = cursor.fetchall()
            # print(retorno)
            if len(retorno)>0:
                return retorno[0][0]
            else:
                return None
    except psycopg2.Error as e:
        print(consulta_id)
        print("Ocurrió un error al consultar con where: ", e)
        conexion.close()
    else:
        conexion.close()
def consultarId(tabla, propiedad, valor):
    consulta_id = ''
    try:
        with conexion.cursor() as cursor:
            if valor != None:
                consulta_id = "select id from {} where {}='{}';".format(tabla,propiedad,valor)
            else:
                consulta_id = "select id from {} where {}='{}';".format(tabla,propiedad,'')
            cursor.execute(consulta_id)
            # Con fetchall traemos todas las filas
            retorno = cursor.fetchall()
            # print(retorno)
            if len(retorno)>0:
                return retorno[0][0]
            else:
                return None
    except psycopg2.Error as e:
        print(consulta_id)
        print("Ocurrió un error al consultar con where: ", e)
        conexion.close()
    else:
        conexion.close()
def insertar(insert, cerrar_conexion):
    try:
        with conexion.cursor() as cursor:
            cursor.execute(insert)
        conexion.commit()  # Si no haces commit, los cambios no se guardan
    except psycopg2.Error as e:
        print(insert)
        print("Ocurrió un error al insertar: ", e)
        conexion.close()
    finally:
        if cerrar_conexion:
            print('cerro conexion')
            conexion.close()

def run():
    duplicados =[]
    for file in os.listdir():
        if file in ['.DS_Store']:
            continue
        #nombre del archivo
        print("{}/{}".format(path,file))
        # Read the file
        data = pd.read_csv("{}/{}".format(path,file), usecols=['Periodo', 'Año','Municipio','Unidad Responsable','Programa','Año de inicio del programa','Prioritario','Tipo de programa','Primer Apellido', 'Segundo Apellido', 'Nombre del beneficiario','Importe beneficiario'], low_memory=False)
        # Output the number of rows
        print("Total rows: {0}".format(len(data)))
        # See which headers are available
        print(list(data))
        for index in range(0,len(data)):
            # print("{} {} {}".format(data['Primer Apellido'][index], data['Segundo Apellido'][index], data['Nombre del beneficiario'][index]))
            # if data['Nombre del beneficiario'][index]=='CLAUDIA SUGEIH':
            #     break
            # duplicado = consultarIdBeneficiario('full_name', data['Nombre del beneficiario'][index],data['Primer Apellido'][index],data['Segundo Apellido'][index])
            # if duplicado:
            #     continue
            full_name_id=ejercicio_id=fecha_inicio_periodo=fecha_fin_periodo=municipio_id=dependencia_responsable_id=ano_inicio_programa=programa_id=tipo_programa_id=importe=prioritario=""
            ejercicio = data['Año'][index]
            if ('Enero - Febrero' == data['Periodo'][index]):
                fecha_inicio_periodo = datetime.strptime('01/01/{}'.format(ejercicio), '%d/%m/%Y').strftime('%Y-%m-%d')
                fecha_fin_periodo = datetime.strptime('28/02/{}'.format(ejercicio), '%d/%m/%Y').strftime('%Y-%m-%d')
            if ('Marzo - Abril' == data['Periodo'][index]):
                fecha_inicio_periodo = datetime.strptime('01/03/{}'.format(ejercicio), '%d/%m/%Y').strftime('%Y-%m-%d')
                fecha_fin_periodo = datetime.strptime('30/04/{}'.format(ejercicio), '%d/%m/%Y').strftime('%Y-%m-%d')
            if ('Mayo - Junio' == data['Periodo'][index]):
                fecha_inicio_periodo = datetime.strptime('01/05/{}'.format(ejercicio), '%d/%m/%Y').strftime('%Y-%m-%d')
                fecha_fin_periodo = datetime.strptime('30/06/{}'.format(ejercicio), '%d/%m/%Y').strftime('%Y-%m-%d')
            if ('Julio - Agosto' == data['Periodo'][index]):
                fecha_inicio_periodo = datetime.strptime('01/07/{}'.format(ejercicio), '%d/%m/%Y').strftime('%Y-%m-%d')
                fecha_fin_periodo = datetime.strptime('30/08/{}'.format(ejercicio), '%d/%m/%Y').strftime('%Y-%m-%d')
            if ('Septiembre - Octubre' == data['Periodo'][index]):
                fecha_inicio_periodo = datetime.strptime('01/09/{}'.format(ejercicio), '%d/%m/%Y').strftime('%Y-%m-%d')
                fecha_fin_periodo = datetime.strptime('30/10/{}'.format(ejercicio), '%d/%m/%Y').strftime('%Y-%m-%d')
            if ('Noviembre - Diciembre' == data['Periodo'][index]):
                fecha_inicio_periodo = datetime.strptime('01/11/{}'.format(ejercicio), '%d/%m/%Y').strftime('%Y-%m-%d')
                fecha_fin_periodo = datetime.strptime('30/12/{}'.format(ejercicio), '%d/%m/%Y').strftime('%Y-%m-%d')
            if (data['Año'][index] == None):
                ejercicio_id = consultarId('ejercicio', 'name',1970)
            else:
                ejercicio_id = consultarId('ejercicio', 'name',data['Año'][index])
            if (data['Municipio'][index] == None):
                municipio_id = 'NULL'
            else:
                municipio_id = consultarId('municipality', 'name',data['Municipio'][index])
                if (municipio_id == None):
                    municipio_id = 'NULL'
            if (data['Unidad Responsable'][index] == None):
                dependencia_responsable_id = ''
            else:
                dependencia_responsable_id = consultarId('dependencia_responsable', 'name',data['Unidad Responsable'][index])
                if (dependencia_responsable_id == None):
                    dependencia_responsable_id = 'NULL'
            if (data['Programa'][index] == None):
                programa_id = 'NULL'
            else:
                programa_id = consultarId('programa_social', 'name',data['Programa'][index])
                if (programa_id == None):
                    programa_id = 'NULL'
            if (data['Año de inicio del programa'][index] == None):
                ano_inicio_programa=''
            else:
                ano_inicio_programa = data['Año de inicio del programa'][index]
            if ('Sí' == data['Prioritario'][index]):
                prioritario=True
            else:
                prioritario = False
            if (data['Tipo de programa'][index] == None):
                tipo_programa_id = ''
            else:
                tipo_programa_id = consultarId('tipo_programa', 'name',data['Tipo de programa'][index])
                if (tipo_programa_id == None):
                    tipo_programa_id = 'NULL'
            full_name_id = consultarIdBeneficiario('full_name', data['Nombre del beneficiario'][index],data['Primer Apellido'][index],data['Segundo Apellido'][index])
            if full_name_id == None:
                full_name_id = consultarId('full_name', 'code_id',0)
            if (data['Importe beneficiario'][index] == None):
                importe=0
            else:
                importe = data['Importe beneficiario'][index]
            insert = "insert into beneficiarios (fehca_inicio_periodo,fehca_fin_periodo,ejercicio_id,state_id,municipio_id,dependencia_id,programa_social_id,fecha_inicio_programa,prioritario,tipo_programa_id,full_name_id,importe,usuario_registro_id, created, modified ) VALUES ('{}','{}',{},{},{},{},{},{},{},{},{},{},1, current_timestamp, current_timestamp);".format(
                fecha_inicio_periodo,
                fecha_fin_periodo,
                ejercicio_id,
                25,
                municipio_id,
                dependencia_responsable_id,
                programa_id,
                ano_inicio_programa,
                prioritario,
                tipo_programa_id,
                full_name_id,
                importe,
                )
            if(index != len(data)-1):
                insertar(insert, False)
            if(index == len(data)-1):
                print(insert)
                insertar(insert, False)

if __name__ == '__main__':
    run()