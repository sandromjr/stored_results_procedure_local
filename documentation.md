Este script é feito para se comunicar com o banco de dados (local) e
realizar a chamada de uma Stored Procedure no MySQL. O resultado é 
armazenado e copiado em uma planilha Excel que será salva no local
escolhido.

--------------------------------------------------------------------------

FUNÇÃO: posicao_carteira_xl()

Define-se uma data para a consulta (var data), o cabeçalho (tuple header).
Inicia um workbook e então faz-se a conexão com o banco de dados. Chama-se
então a Stored Procedure e armazena os dados do resultado. A planilha
então começará a  ser preenchida pelo cabeçalho, em seguida pelos dados
armazenados pela consulta e ao final contendo os dados sobre a última
atualização da própria planilha. A conexão com o banco de dados será
encerrada e a planilha será salva no caminho especificado.

--------------------------------------------------------------------------

LIMITAÇÕES: O banco de dados local só possibilita a conexão do host do 
servidor. Portanto o código em Python só retornará resultado ao ser
executado pela máquina específica.
