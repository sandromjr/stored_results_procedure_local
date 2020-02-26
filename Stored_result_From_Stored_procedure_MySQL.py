import datetime
import mysql.connector
from mysql.connector import Error

data = (datetime.date(2020, 1, 31))

try:
    connection = mysql.connector.connect(host='127.0.0.1',
                                         database='***',
                                         user='***',
                                         password='***')

    cursor = connection.cursor()
    cursor.callproc('MJ_posicao_carteira', (data,))

    print("\nPrinting each laptop record")
    for j in cursor.stored_results():
        for row in j:
            print("Data = ", row[0])
            print("Cliente  = ", row[1])
            print("Descricao = ", row[2])
            print("Quantidade = ", row[3])
            print("Preco Unitario = ", row[4])
            print("Valor de Mercado  = ", row[5])
#        print(row.fetchall())

except Error as e:
    print("Error reading data from MySQL table", e)
finally:
    if (connection.is_connected()):
        connection.close()
        cursor.close()
        print("MySQL connection is closed")
        
# ------------------------------------------------------------------------------------------------
# COMEÇA UM NOVO CODIGO
# ------------------------------------------------------------------------------------------------
#
#from openpyxl import Workbook
#arquivo_excel = Workbook()
#planilha1 = arquivo_excel.active
#valores = [
#    ("Categoria", "Valor"),
#    ("Restaurante", 45.99),
#    ("Transporte", 208.45),
#    ("Viagem", 558.54)
#]
#
#for linha in valores:
#    planilha1.append(linha)
#
#arquivo_excel.save("C:/Users/***/teste.xlsx")
#