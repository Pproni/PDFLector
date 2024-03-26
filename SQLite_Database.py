import sqlite3 as sql

def Create_DB():
    conn = sql.connect("Prueba.db")
    conn.commit()
    conn.close()

def CreateTable():
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    cursor.execute(
        """CREATE TABLE Jobs_changes(
            Initial_Date text, Initial_JobName text,
            Final_JobName text
        )"""
    )
    conn.commit()
    conn.close()

def inserRow(Table,Initial_Date,cambiosList):
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = f"INSERT INTO {Table} VALUES ('{Initial_Date}', ?, ?)"
    cursor.execute(instruccion, cambiosList)
    conn.commit()
    conn.close()

def readRows():
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = f"SELECT *  FROM Cambios"
    cursor.execute(instruccion)
    datos = cursor.fetchall()
    conn.commit()
    conn.close()
    print(datos)

def insertRows(Table,Initial_Date,cambiosList):
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = f"INSERT INTO {Table} VALUES ('{Initial_Date}', ?, ?)"
    cursor.executemany(instruccion, cambiosList)
    conn.commit()
    conn.close()    

def readOrdered(field):
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = f"SELECT *  FROM Cambios ORDER BY {field}" 
    cursor.execute(instruccion)
    datos = cursor.fetchall()
    conn.commit()
    conn.close()
    print(datos)

def search(Table,Column,Element):
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = f"SELECT *  FROM {Table} WHERE {Column} like '{Element}%'"
    cursor.execute(instruccion)
    datos = cursor.fetchall()
    conn.commit()
    conn.close()
    return datos

def updatefields(Table, Oldname, Newname):
    if Table == 'Jobs_changes':
        column_finalname = 'Final_JobName'
        column_initallname = 'Initial_JobName'
    elif Table == 'Names_changes':
        column_finalname = 'Final_Name'
        column_initallname = 'Initial_Name'
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = f"UPDATE {Table} SET {column_finalname}='{Newname}' WHERE {column_initallname}='{Oldname}'"
    cursor.execute(instruccion)
    conn.commit()
    conn.close()

def deleteRow(Table, Column, Rowvalue):
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = f"DELETE FROM {Table} WHERE {Column}='{Rowvalue}'"
    cursor.execute(instruccion)
    conn.commit()
    conn.close()

def updateColumnName(oldname, newname):
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = f'''ALTER TABLE Cambios RENAME COLUMN {oldname} TO {newname}'''
    cursor.execute(instruccion)
    conn.commit()
    conn.close()

def removeColumn():
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = '''ALTER TABLE Cambios DROP COLUMN Try'''
    cursor.execute(instruccion)
    conn.commit()
    conn.close()

def addColumn():
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = '''ALTER TABLE Cambios ADD COLUMN Try'''
    cursor.execute(instruccion)
    conn.commit()
    conn.close()

def getTableColumns(table_name):
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = f"PRAGMA table_info({table_name})"
    cursor.execute(instruccion)
    columns = [column[1] for column in cursor.fetchall()]
    conn.commit()
    conn.close()
    print(columns)
    
def updatemultipleValues():
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = """UPDATE Cambios 
            SET Initial_Name = 'Toño', Final_Name = 'Tigre'
            WHERE Initial_Name = 'UWU' """
    cursor.execute(instruccion)
    conn.commit()
    conn.close()

def renameTable():
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = '''ALTER TABLE Cambios RENAME TO Names_changes'''
    cursor.execute(instruccion)
    conn.commit()
    conn.close()

def getValuesList(Table,Column,InitialDate):
    conn = sql.connect("Prueba.db")
    cursor = conn.cursor()
    instruccion = f"SELECT *  FROM {Table} WHERE {Column} like '{InitialDate}%'"
    cursor.execute(instruccion)
    datos = cursor.fetchall()
    cambios = [list(fila[1:]) for fila in datos]
    conn.commit()
    conn.close()
    return cambios

if __name__ == '__main__':
    check = False
    #testlist = [['1','2'],['3','4'],['3','6'],['3','5']]
    testlist = ['1','2']
    Initial_date = '2024-01-26'
    #insertRows('Jobs_changes',Initial_date,testlist)
    #updatefields('Jobs_changes','3','7')
    #cambios = getValuesList('Names_changes','Initial_Date',Initial_date)
    #print(cambios)
    #inserRow('Names_changes',Initial_date,testlist)
    if not getValuesList('Jobs_changes','Initial_Date','2024-01-19'):
        check = False
    else:
        check = True
        
    print(check)
    deleteRow('Jobs_changes','Initial_Date','2024-01-26')