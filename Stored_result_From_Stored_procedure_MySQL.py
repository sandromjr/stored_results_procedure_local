#Cria a função que conecta ao banco de dados, realiza a consulta a partir de _
#uma Stored Procedure e o resultado é armazenado em uma planilha excel. _
#O arquivo será salvo no diretório escolhido

def posicao_carteira_xl():
    import datetime
    import mysql.connector
    from openpyxl import Workbook
    from mysql.connector import Error

#Declara a data que será realizada a pesquisa
    data = (datetime.date(2020, 1, 31))
#Armazena o horário e data que ao início da função
    time = datetime.datetime.now()

#Controle em Run
    print(f"Date: {data} SELECTED")

#Declara o cabeçalho da planilha
    header = [
        ("Data", "Cliente", "Descricao", "Quantidade", "Preco_Unitario", "Valor_de_Mercado")
    ]

#Inicia o arquivo excel
    arquivo_excel = Workbook()
    planilha1 = arquivo_excel.active

#Controle em Run
    print("Excel File: OPENED")

#Inicia a conexão
    try:
        connection = mysql.connector.connect(host='127.0.0.1',
                                         database='lts1',
                                         user='sandro',
                                         password='Smj091050#')
#Controle em Run
        print("Database: CONNECTED")
        cursor = connection.cursor()

#Chama a Stored Procedure (Com args)
        cursor.callproc('MJ_posicao_carteira', (data,))

#Controle em Run
        print("Stored Procedure: CALLED SUCCESSFULLY")

#Copia as informações do cabeçalho para a primeira linha do Excel
        for linha in header:
            planilha1.append(linha)

#Copia todos os dados da consulta no Excel, abaixo do cabeçalho
        for j in cursor.stored_results():
            for row in j:
                planilha1.append(row)

#Insere os dados de atualização ao final da tabela
        planilha1.append([("")])
        planilha1.append([(time.strftime("UPDATE %d/%m/%Y %H:%M:%S"))])

#Controle em Run
        print("--- Downloading Datas ---")


    #except Error as e:
    #    print("Error reading data from MySQL table", e)
    finally:
#Fecha a conexão com o banco de dados
        if (connection.is_connected()):
            connection.close()
            cursor.close()
#Controle em Run
            print("MySQL connection is closed")

#Salva a planilha com os dados no diretório
    arquivo_excel.save("C:/Users/smarchini/Documents/Posicao_Carteira.xlsx")

#Controle em Run
    print("File Posicao_Carteira.xlsx CREATED")

#Executa a função
posicao_carteira_xl()
