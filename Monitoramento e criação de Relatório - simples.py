from datetime import date
from datetime import datetime
import time
import pandas as pd
import requests as requests
import pyodbc

while True: ## Fazendo um loop para infinito para ficar fazendo o monitoramento constantemente ###
 requisiçao = requests.get(
    'https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL') ### Fazendo a requisição para o site, e trazendo no formato Json, pois é um dado estruturado ###
 Requisiçao = requisiçao.json() 

#### Requisições e as variaveis para passar no banco ###

 Cotacao_euro_compra = Requisiçao["EURBRL"]["bid"]
 Cotacao_euro_venda = Requisiçao["EURBRL"]["ask"]
 Cotacao_dolar_compra = Requisiçao["USDBRL"]["bid"]
 Cotacao_dolar_venda = Requisiçao["USDBRL"]["ask"]
 Cotacao_btc_compra = Requisiçao["BTCBRL"]["bid"]
 Cotacao_btc_venda = Requisiçao["BTCBRL"]["ask"]

### Variaveis que serão executadas no banco de dados, para armazenar os dados ddas requisições ###
 Data_Horario_Requisicao = datetime.now()
 Euro = f"""'Euro','{Cotacao_euro_compra}','{Cotacao_euro_venda}','{Data_Horario_Requisicao}'"""
 Dolar = f"""'Dolar',{Cotacao_dolar_compra},{Cotacao_dolar_venda},'{Data_Horario_Requisicao}'"""
 Bitcoin = f"""'Bitcoin','{Cotacao_btc_compra}','{Cotacao_btc_venda}','{Data_Horario_Requisicao}'"""

### querys para inserir os dados no banco de dados ###
 query = f"""INSERT INTO COTACAO.dbo.Monitoramento VALUES({Dolar})
            INSERT INTO COTACAO.dbo.Monitoramento VALUES({Euro})
            INSERT INTO COTACAO.dbo.Monitoramento VALUES({Bitcoin})
 """
### Conectando no banco atraves da biblioteca pyodbc ###
 conexão = pyodbc.connect(
    "Driver={ODBC Driver 17 for SQL Server}; Server= DESKTOP-MICJEJC\SQLEXPRESS; Database = COTACAO; Trusted_connection=yes;")
 cursor = conexão.cursor()
 cursor.execute(query)
 cursor.commit() ### NECESSARIO POIS ESTÁ SE ALTERANDO O BANCO DE DADOS ###

### CONECTANDO NOVAMENTE PARA CRIAR UMA TABELA RELATORIO ###
### Criando o range para consulta no banco de dados e criação do relatório ###
### Observação, sei que poderia ter passado apenas o date today, porém optei por dividir pois vai que no futuro tenha-se que utilizar alguma data relativa ###
 
 Dia_atual = date.today().strftime("%d")
 Mes_atual = date.today().strftime("%m")
 Ano_atual = date.today().strftime("%Y")
 Data_query = Ano_atual + "-" + Mes_atual + "-" + Dia_atual
 ### Passando a variavel da query e delimitando pelo tempo que quero que sejam retirados essas informações ###
 query1 = f"""select * from COTACAO.dbo.Monitoramento with(nolock)
	where CAST(COTACAO.dbo.Monitoramento.Datahoracadastro as datetimeoffset) 
	between '{Data_query} 00:01:00' and '{Data_query} 23:59:00'"""
 conexao1 = pyodbc.connect(
    "Driver={ODBC Driver 17 for SQL Server}; Server= DESKTOP-MICJEJC\SQLEXPRESS; Database = COTACAO; Trusted_connection=yes;")
 cursor = conexão.cursor()
 cursor.execute(query1)    
 from_db=[] ### Criando uma lista vazia para inserir os dados retirados do banco de dados ###
 for row in cursor:
    result = list(row)
    from_db.append(result)
 
#### Criando o Relatorio Excel em uma pasta para alimentar o BI Optei por crirar no formato excel, mas da para escolher outros formatos ### 

 Datalatina = date.today().strftime("%d-%m-%Y")
 colunas = ['Moeda','Compra','Venda','Datacotacao']
 Tabela = pd.DataFrame(from_db, columns=colunas)
 path = "C:/Teste"
 Nome_arquivo = f'Cotação-{Datalatina}.xlsx'
 Tabela.to_excel(f"{path}/{Nome_arquivo}",index=False)
 print("Código executado com sucesso")

 time.sleep(1800.01) ### Colocando um timer de 30 minutos entre cada execução do códgio ###
