import pyodbc
import pandas as pd 
from datetime import date
from datetime import datetime
from datetime import timedelta
import os

### Data utilizada para a consulta no SQL ###

data_consulta_atual = date.today()

### Datas e Horas utilizadas para nomeção do arquivo e criação de pastas ###
hora_geracao = datetime.now().strftime('%H')
ano_geracao = date.today().strftime('%Y')
mes_geracao = date.today().strftime('%m')
dia_geracao = date.today().strftime('%d')
data_geracao_anterior = datetime.now() - (timedelta(days=1))
data_geracao_historico = datetime.now() - (timedelta(days=8))
data_geracao_anterior = data_geracao_anterior.strftime('%Y-%m-%d')
data_geracao_historico = data_geracao_historico.strftime('%Y-%m-%d')
from_db1 = []
from_db2 = []
from_db3 = []
tabela = ''
tabela2=''
tabela3 = ''

### Endereço do diretório de Criação ###
endereco_pasta = '//adldas2k1226/producao$/Relatórios/Relatório Diário Edição'

def consulta_sql_1_geracao():
    """
    Nesta parte passaremos as consultas, os conectores, e a criação do arquivo excel
    - A Range Desta consulta consulta do dia anterior as 14 horas, até o dia atual as 08:00
    """
    
    ### Consulta Utilizada Para criação
    
    query = f"""
    select distinct d.codDIario, d.Nome, cd.codCaderno, cd.Nome, bl.id as buscaLote, tb.edicaoDiario, bl.DataPublicacao, bl.DataDivulgacao, bl.DataHoraInclusao, count(tb.codPublicacao) as QtdePublicacao
    from publicacao.PublicacaoBuscaLote bl (nolock)
    join publicacao.ControleMigracao cm (nolock) on cm.idPublicacaoBuscaLote = bl.Id	
    join publicacao.Diario d (nolock) on d.codDIario = cm.codDiario 
    join publicacao.CadernoDiario cd (nolock) on cd.codCaderno = cm.codCaderno and cd.Id = cm.IdCadernoDiario
    join dbo.tbPublicacao tb (nolock) on tb.buscaLote = bl.Id
    where bl.idArquivoPublicFornecedor is not null
    and bl.DataHoraInclusao between '{data_geracao_anterior} 14:00:00' and '{data_consulta_atual} 08:00:00'
    group by
    d.codDIario, d.Nome, cd.codCaderno, cd.Nome, bl.id, tb.edicaoDiario, bl.DataPublicacao, bl.DataDivulgacao, bl.DataHoraInclusao
    order by bl.DataHoraInclusao   """
    
    
    conecao = pyodbc.connect("Driver={SQL Server Native Client 11.0}; Server=srvbanco035; Database=AdvisePublicacao_Producao; Trusted_Connection=yes;")
    cursor = conecao.cursor()
    cursor.execute(query)

    for row in cursor:
        global from_db1
        global tabela
        result = list(row)
        from_db1.append(result)

        
        """
        Criando o DataFrame e Fazendo um tratamento prévio dos dados para melhor visualização
        """

        columns = [
            'Código Diário',
            'Nome Caderno',
            'Código Caderno',
            'Nome Diário',
            'Busca Lote',
            'Edição',
            'Data Publicacao',
            'Data Divulgacao',
            'Data Inclusao',
            'Quantidade Pub'
            ] 
        tabela = pd.DataFrame(from_db1, columns=columns)
        tabela['Data Publicacao'] = tabela['Data Publicacao'].dt.strftime('%d/%m/%Y')
        tabela['Data Divulgacao'] = tabela['Data Divulgacao'].dt.strftime('%d/%m/%Y')
        tabela['Data Inclusao'] = tabela['Data Inclusao'].dt.strftime('%d/%m/%Y %H:%M')
        
        tabela = tabela[[
            'Código Diário',
            'Nome Caderno',
            'Nome Diário',
            'Código Caderno',
            'Busca Lote',
            'Edição',
            'Data Divulgacao',
            'Data Publicacao',
            'Data Inclusao',
            'Quantidade Pub'
            ]]

    return tabela



def consulta_sql_2_geracao():
    """
    Nesta parte passaremos as consultas, os conectores, e a criação do arquivo excel
    -- A Range desta consulta, consulta somente o Dia Atual
    """
    
    ### Consulta Utilizada Para criação
    
    query = f"""
    select distinct d.codDIario, d.Nome, cd.codCaderno, cd.Nome, bl.id as buscaLote, tb.edicaoDiario, bl.DataPublicacao, bl.DataDivulgacao, bl.DataHoraInclusao, count(tb.codPublicacao) as QtdePublicacao
    from publicacao.PublicacaoBuscaLote bl (nolock)
    join publicacao.ControleMigracao cm (nolock) on cm.idPublicacaoBuscaLote = bl.Id	
    join publicacao.Diario d (nolock) on d.codDIario = cm.codDiario 
    join publicacao.CadernoDiario cd (nolock) on cd.codCaderno = cm.codCaderno and cd.Id = cm.IdCadernoDiario
    join dbo.tbPublicacao tb (nolock) on tb.buscaLote = bl.Id
    where bl.idArquivoPublicFornecedor is not null
    and bl.DataHoraInclusao between '{data_consulta_atual} 08:00:00' and '{data_consulta_atual} 14:00:00'
    group by
    d.codDIario, d.Nome, cd.codCaderno, cd.Nome, bl.id, tb.edicaoDiario, bl.DataPublicacao, bl.DataDivulgacao, bl.DataHoraInclusao
    order by bl.DataHoraInclusao   """
    
    
    conecao = pyodbc.connect("Driver={SQL Server Native Client 11.0}; Server=srvbanco035; Database=AdvisePublicacao_Producao; Trusted_Connection=yes;")
    cursor = conecao.cursor()
    cursor.execute(query)

   
    for row in cursor:
        global from_db2
        global tabela3    
        result = list(row)
        from_db2.append(result)
        """
        Criando o DataFrame e Fazendo um tratamento prévio dos dados para melhor visualização
        """
        columns = [
            'Código Diário',
            'Nome Caderno',
            'Código Caderno',
            'Nome Diário',
            'Busca Lote',
            'Edição',
            'Data Publicacao',
            'Data Divulgacao',
            'Data Inclusao',
            'Quantidade Pub'
             ] 
        tabela3 = pd.DataFrame(from_db2, columns=columns)
        tabela3['Data Publicacao'] = tabela3['Data Publicacao'].dt.strftime('%d/%m/%Y')
        tabela3['Data Divulgacao'] = tabela3['Data Divulgacao'].dt.strftime('%d/%m/%Y')
        tabela3['Data Inclusao'] = tabela3['Data Inclusao'].dt.strftime('%d/%m/%Y %H:%M')
        
        tabela3 = tabela3[[
                'Código Diário',
                'Nome Caderno',
                'Nome Diário',
                'Código Caderno',
                'Busca Lote',
                'Edição',
                'Data Divulgacao',
                'Data Publicacao',
                'Data Inclusao',
                'Quantidade Pub'
                ]]
    
    return tabela3

    


def consulta_sql_geracao_continua_historico():
    """
    Nesta parte passaremos as consultas, os conectores, e a criação do arquivo excel
    - A Range Desta consulta consulta do dia anterior as 14 horas, até o dia atual as 08:00
    """
    
    ### Consulta Utilizada Para criação
    
    query = f"""
    select distinct d.codDIario, d.Nome, cd.codCaderno, cd.Nome, bl.id as buscaLote, tb.edicaoDiario, bl.DataPublicacao, bl.DataDivulgacao, bl.DataHoraInclusao, count(tb.codPublicacao) as QtdePublicacao
    from publicacao.PublicacaoBuscaLote bl (nolock)
    join publicacao.ControleMigracao cm (nolock) on cm.idPublicacaoBuscaLote = bl.Id	
    join publicacao.Diario d (nolock) on d.codDIario = cm.codDiario 
    join publicacao.CadernoDiario cd (nolock) on cd.codCaderno = cm.codCaderno and cd.Id = cm.IdCadernoDiario
    join dbo.tbPublicacao tb (nolock) on tb.buscaLote = bl.Id
    where bl.idArquivoPublicFornecedor is not null
    and bl.DataHoraInclusao between '{data_geracao_historico} 00:00:00' and '{data_consulta_atual} 23:59:59'
    group by
    d.codDIario, d.Nome, cd.codCaderno, cd.Nome, bl.id, tb.edicaoDiario, bl.DataPublicacao, bl.DataDivulgacao, bl.DataHoraInclusao
    order by bl.DataHoraInclusao   """
    
    
    conecao = pyodbc.connect("Driver={SQL Server Native Client 11.0}; Server=srvbanco035; Database=AdvisePublicacao_Producao; Trusted_Connection=yes;")
    cursor = conecao.cursor()
    cursor.execute(query)


    for row in cursor:
        global from_db3
        global tabela2
        result = list(row)
        from_db3.append(result)
        
        """
        Criando o DataFrame e Fazendo um tratamento prévio dos dados para melhor visualização
        """

        columns = [
            'Código Diário',
            'Nome Caderno',
            'Código Caderno',
            'Nome Diário',
            'Busca Lote',
            'Edição',
            'Data Publicacao',
            'Data Divulgacao',
            'Data Inclusao',
            'Quantidade Pub'
            ] 
        tabela2 = pd.DataFrame(from_db3, columns=columns)
        tabela2['Data Publicacao'] = tabela2['Data Publicacao'].dt.strftime('%d/%m/%Y')
        tabela2['Data Divulgacao'] = tabela2['Data Divulgacao'].dt.strftime('%d/%m/%Y')
        tabela2['Data Inclusao'] = tabela2['Data Inclusao'].dt.strftime('%d/%m/%Y %H:%M')
        
        tabela2 = tabela2[[
            'Código Diário',
            'Nome Caderno',
            'Nome Diário',
            'Código Caderno',
            'Busca Lote',
            'Edição',
            'Data Divulgacao',
            'Data Publicacao',
            'Data Inclusao',
            'Quantidade Pub'
            ]]

    return tabela2      


def criacao_pastas():
        """   
        Definindo primeiramente a criação de diretórios
        """
        pasta_ano = os.path.normpath((os.path.join(endereco_pasta,ano_geracao)))
        pasta_mes = os.path.normpath(os.path.join(endereco_pasta,pasta_ano,mes_geracao))
        pasta_dia = os.path.normpath(os.path.join(endereco_pasta,pasta_mes,dia_geracao))
        """ 
        Construção do nome da pasta
        """
        nome = str(hora_geracao)
        nome_arquivo = 'Relatório de Lotes Gerados' + ' - ' + nome + ' H' + '.xlsx'
        nome_txt = 'Não existem novos arquivos' + ' - ' + nome + ' H' '.txt'

        if datetime.now().strftime('%Y-%m-%d %H:%M') > datetime.now().strftime('%Y-%m-%d 13:59'):
            tabela_consulta_14 = consulta_sql_2_geracao()
            tabela_historico = consulta_sql_geracao_continua_historico()
            if os.path.isdir(pasta_dia) == True:
                if len(tabela_consulta_14) == 0:
                    open(f'{pasta_dia}/{nome_txt}',"x")
                else:
                    writer = pd.ExcelWriter(f'{pasta_dia}/{nome_arquivo}',engine='xlsxwriter')
                    tabela_consulta_14.to_excel(writer, sheet_name='Diario-Lote',index=False)
                    tabela_historico.to_excel(writer,sheet_name = 'Histórico',index=False)
                    writer.save()
            else:
                if len(tabela_consulta_14) == 0:
                    os.makedirs(pasta_dia)
                    open(f'{pasta_dia}/{nome_txt}',"x")
                else:
                    os.makedirs(pasta_dia)
                    writer = pd.ExcelWriter(f'{pasta_dia}/{nome_arquivo}',engine='xlsxwriter')
                    tabela_consulta_14.to_excel(writer, sheet_name='Diario-Lote',index=False)
                    tabela_historico.to_excel(writer,sheet_name = 'Histórico',index=False)
                    writer.save()
        else:
            tabela_consulta_8 = consulta_sql_1_geracao()
            tabela_historico = consulta_sql_geracao_continua_historico()
            if os.path.isdir(pasta_dia) == True:
                if len(tabela_consulta_8) == 0:
                    open(f'{pasta_dia}/{nome_txt}',"x")
                else:
                    writer = pd.ExcelWriter(f'{pasta_dia}/{nome_arquivo}',engine='xlsxwriter')
                    tabela_consulta_8.to_excel(writer, sheet_name='Diario-Lote',index=False)
                    tabela_historico.to_excel(writer,sheet_name = 'Histórico',index=False)
                    writer.save()
            else:
                if len(tabela_consulta_8) == 0:
                    os.makedirs(pasta_dia)
                    open(f'{pasta_dia}/{nome_txt}',"x")
                else:
                    os.makedirs(pasta_dia)
                    writer = pd.ExcelWriter(f'{pasta_dia}/{nome_arquivo}',engine='xlsxwriter')
                    tabela_consulta_8.to_excel(writer, sheet_name='Diario-Lote',index=False)
                    tabela_historico.to_excel(writer,sheet_name = 'Histórico',index=False)
                    writer.save()
        print('{}'.format('Relatorio Gerado'))

    
criacao_pastas()
