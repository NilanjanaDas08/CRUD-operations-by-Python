import pyodbc

def database_connection():
    try:
        conn = pyodbc.connect(
            'DRIVER={ODBC Driver 18 for SQL Server};'
            'SERVER=DESKTOP-DU3RE5I\SQLEXPRESS;'
            'DATABASE=Office;'
            'Trusted_Connection=yes;'
            'TrustServerCertificate=yes;'
        )
        return conn
    except pyodbc.Error as e:
        print("Error occurred while connecting to the database:", e)
        return None

def insert_employees_from_dict(employees_list):
    conn = database_connection()
    if conn:
        try:
            cursor = conn.cursor()
            insert_query = '''
            INSERT INTO Employee (EmployeeID, First_Name, Last_Name, Age)
            VALUES (?, ?, ?, ?)
            '''
            
            # Convert list of dictionaries to a list of tuples
            values_list = [
                (
                    emp['EmployeeID'],
                    emp['First_Name'],
                    emp['Last_Name'],
                    emp['Age']
                )
                for emp in employees_list
            ]
            
            cursor.executemany(insert_query, values_list)
            conn.commit()
            print("Employees inserted successfully!")
        
        except pyodbc.Error as e:
            print("Error occurred:", e)
        
        finally:
            cursor.close()
            conn.close()

def delete_employees():
    conn = database_connection()
    if conn:
        try:
            cursor = conn.cursor()
            
            delete_query = "DELETE FROM Employee WHERE EmployeeID IN(?)"
            
            cursor.execute(delete_query,(4,))
            conn.commit()
            print("Employees deleted successfully")
        
        except pyodbc.Error as e:
            print("Error occurred:", e)
        
        finally:
            cursor.close()
            conn.close()

def get_all_employees():
    conn = database_connection()
    if conn:
        try:
            cursor = conn.cursor()
            
            select_query = "SELECT * FROM Employee"
            cursor.execute(select_query)
            
            rows = cursor.fetchall()
            
            for row in rows:
                print(row)
        
        except pyodbc.Error as e:
            print("Error occurred:", e)
        
        finally:
            cursor.close()
            conn.close()

# Example usage:
employees = [
    {'EmployeeID': 1, 'First_Name': 'Jishan', 'Last_Name': 'Doe', 'Age': 30},
    {'EmployeeID': 2, 'First_Name': 'Zayn', 'Last_Name': 'Smith', 'Age': 25},
    {'EmployeeID': 3, 'First_Name': 'Anil', 'Last_Name': 'Parker', 'Age': 28},
    {'EmployeeID': 4, 'First_Name': 'Nilanjana', 'Last_Name': 'Das', 'Age': 23}

]
insert_employees_from_dict(employees)

# Delete  employees
delete_employees()

# Retrieve and display all employees
get_all_employees()
