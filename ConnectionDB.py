import psycopg2

connection = psycopg2.connect(
        host="localhost",
        port="5432",
        database="Agents",
        user="postgres",
        password="12345"
    )

cur = connection.cursor()
