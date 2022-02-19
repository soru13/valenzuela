import psycopg2

conexion = psycopg2.connect(host="transparencia_postgres_1", database="transparencia", user="postgres", password="DucEasGMJjkQY2AthnC")

# with open("credenciales.json") as archivo_credenciales:
#     credenciales = json.load(archivo_credenciales)
# try:
#     conexion = psycopg2.connect(**credenciales)
# except psycopg2.Error as e:
#     print("Ocurrió un error al conectar a PostgreSQL: ", e)