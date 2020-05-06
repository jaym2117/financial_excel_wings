import pypyodbc as pyodbc
import pandas as pd

class SQL_Config:
    def __init__(self, db_host, db_name, db_user, db_password):
        self.db_host = db_host
        self.db_name = db_name
        self.db_user = db_user
        self.db_password = db_password
        self.connection_string = 'Driver={SQL Server};Server=' + db_host + ';Database=' + db_name + ';UID=' + db_user + ';PWD=' + db_password + ';'
   
    def sqlStatement(self, statement):
        db = pyodbc.connect(self.connection_string)
        return pd.read_sql(statement, db)

    def writeToSQL(self, statement):
        db = pyodbc.connect(self.connection_string)
        cursor = db.cursor()
        cursor.execute(statement)
        db.commit()
        db.close()