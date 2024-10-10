import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()

class Database:

    # SQL Запрос на все данные для отображения состава апелляционной комиссии в таблице
    select_commission_members = """
        SELECT 
            CONCAT(ROW_NUMBER() OVER (ORDER BY e.employee_id), '.') AS No,
            CONCAT(SUBSTRING(e.first_name, 1, 1), '.', SUBSTRING(e.surname, 1, 1), '. ', e.last_name) AS "Ф.И.О Эксперта",
            r.role AS Должность
        FROM 
            CommissionMembers cm
        JOIN 
            Employees e ON cm.employee_id = e.employee_id
        JOIN 
            roles r ON e.role_id = r.role_id;
    """

    # Логика подключения к БД
    _connection = None

    @classmethod
    def get_connection(cls):
        if cls._connection is None:
            try:
                cls._connection = psycopg2.connect(
                    dbname=os.getenv("DB_NAME"),
                    user=os.getenv("DB_USER"),
                    password=os.getenv("DB_PASSWORD"),
                    host=os.getenv("DB_HOST"),
                    port=os.getenv("DB_PORT"),
                )
                print("Соединение с PostgreSQL установлено")
            except Exception as error:
                print("Ошибка при подключении к PostgreSQL", error)
                cls._connection = None
        return cls._connection

    @classmethod
    def close_connection(cls):
        if cls._connection:
            cls._connection.close()
            print("Соединение с PostgreSQL закрыто")
            cls._connection = None

    # Метод получения состава 
    @classmethod
    def get_commission_members(cls):
        with Database.get_connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(cls.select_commission_members)
                results = cursor.fetchall()
                return results