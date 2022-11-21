import requests
from bs4 import BeautifulSoup
import regex as re
import pandas as pd

def ultima_pagina_apartamento():
    
    requisicação1 = 'https://imobiliariaperez.com.br/comprar/apartamento-a-venda?page=1'
    headers = {'User-Agente':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36 OPR/92.0.0.0'}

    site = requests.get(requisicação1,headers=headers)
    soup =BeautifulSoup(site.content,'html.parser')
    qnt_paginas = soup.find('div',class_='pagination')
    lista = ['']
    for x in qnt_paginas:
         listas = list(x)
         lista.append(listas)

    lista = lista[14]
    return lista

def ultima_pagina_casa():
    
    requisicação1 = 'https://imobiliariaperez.com.br/comprar/casa-a-venda?page=1'
    headers = {'User-Agente':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36 OPR/92.0.0.0'}

    site = requests.get(requisicação1,headers=headers)
    soup =BeautifulSoup(site.content,'html.parser')
    qnt_paginas = soup.find('div',class_='pagination')
    lista = ['']
    for x in qnt_paginas:
         listas = list(x)
         lista.append(listas)

    lista = lista[14]
    return lista

def scraping():
    pagina = ultima_pagina_apartamento()
    pagina_vazia  = [int(val) for val in pagina]
    pagina_vazia=str(pagina_vazia).strip('[]')
    dic_informações = {'Bairro':[],'Preço':[],'Código':[],'Home':[],'Link':[]}
    for x in range(1,int(pagina_vazia)+1):
        link = f'https://imobiliariaperez.com.br/comprar/apartamento-a-venda?page={x}'
        headers = {'User-Agente':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36 OPR/92.0.0.0'}
        site = requests.get(link,headers=headers)
        soup =BeautifulSoup(site.content,'html.parser')
        informações = soup.find_all('div', class_=re.compile('list-property-info'))
        for informação in informações:
            carac = informação.find('div',class_=re.compile('list-title')).get_text().strip()
            Preço = informação.find('div',class_=re.compile('list-pric')).get_text().strip()
            Código = informação.find('div',class_=re.compile('list-code')).get_text().strip()
            home  = informação.find('div',class_=re.compile('slide-home-itens')).get_text().strip()
            link2 = informação.find('a',href=re.compile('https'))
            dic_informações['Bairro'].append(carac)
            dic_informações['Preço'].append(Preço)
            dic_informações['Código'].append(Código)
            dic_informações['Home'].append(home)
            dic_informações['Link'].append(link2)
    
    pagina = ultima_pagina_casa()
    pagina_vazia  = [int(val) for val in pagina]
    pagina_vazia=str(pagina_vazia).strip('[]')
    for x in range(1,int(pagina_vazia)+1):
        link3 = f'https://imobiliariaperez.com.br/comprar/casa-a-venda?page={x}'
        headers = {'User-Agente':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36 OPR/92.0.0.0'}
        site = requests.get(link3,headers=headers)
        soup =BeautifulSoup(site.content,'html.parser')
        informações = soup.find_all('div', class_=re.compile('list-property-info'))
        for informação in informações:
            carac = informação.find('div',class_=re.compile('list-title')).get_text().strip()
            Preço = informação.find('div',class_=re.compile('list-pric')).get_text().strip()
            Código = informação.find('div',class_=re.compile('list-code')).get_text().strip()
            home  = informação.find('div',class_=re.compile('slide-home-itens')).get_text().strip()
            link2 = informação.find('a',href=re.compile('https'))
            dic_informações['Bairro'].append(carac)
            dic_informações['Preço'].append(Preço)
            dic_informações['Código'].append(Código)
            dic_informações['Home'].append(home)
            dic_informações['Link'].append(link2)
    
    
    
    return dic_informações

def Tratamento_excel():
    dic= scraping()
    tabela=pd.DataFrame(dic)

    tabela['Home'] = tabela['Home'].str.replace('\n','')

    tabela = tabela.astype({
        'Link' : 'string',
        'Bairro': 'string',
        'Código': 'string',
        'Home': 'string',
        'Preço': 'string'
        })

    tabela.to_excel(f'\\Users\Gustavo\Desktop\Limpeza\semtratamento.xlsx',index=False)
    
    regex_1= tabela['Bairro'].str.findall("[venda]{5}.{1,100}[-]\s.{1,100}[/][A-Z]{2}")
    regex_2 = tabela['Código'].str.findall('\d{1,10}')
    regex_3 = tabela['Home'].str.findall('\d\s[Dorm]{4}')
    regex_4 = tabela['Home'].str.findall('\d\s[Banh]{4}')
    regex_5 = tabela['Home'].str.findall('\d\s[Suítes]{6}')
    regex_6 = tabela['Home'].str.findall('\d\s[Vagas]{5}')
    regex_7 = tabela['Home'].str.findall('\d{2,4}[m]{1}.')
    regex_8 = tabela['Link'].str.findall('\w{5}[:].{1,200}\d{2,6}')
    regex_9 = tabela['Preço'].str.findall('\d{1,30}\W\d{1,30}\W\d{1,30}|\d{1,30}\W\d{1,30}')
    regex_10 = tabela['Bairro'].str.findall('^\w{1,20}') 
    regex_1, regex_2,regex_3,regex_4,regex_5,regex_6,regex_7,regex_8,regex_9,regex_10 = regex_1.astype(str),regex_2.astype(str),regex_3.astype(str),regex_4.astype(str),regex_5.astype(str),regex_6.astype(str),regex_7.astype(str),regex_8.astype(str),regex_9.astype(str),regex_10.astype(str)
    tabela = tabela.assign(regex_1 = regex_1,regex_2 = regex_2,Dormitórios = regex_3,Banheiros=regex_4,Suites = regex_5,vagas = regex_6,M2 = regex_7,link = regex_8,Preço = regex_9,Imovel = regex_10)
    regex = tabela['regex_1'].str.findall('\s[-].{1,100}[/][A-Z]{2}')
    rege = tabela['regex_1'].str.findall('[no]{2}.{1,100}[-].{1}')
    regex, rege = regex.astype(str),rege.astype(str)
    tabela = tabela.assign(Cidade_Bairro = regex,B = rege)
    tabela['Preço'] = tabela['Preço'].str.replace('\[\]','')
    tabela[['Cidade','Bairro2']] = tabela['Cidade_Bairro'].str.split('/',expand=True)


    tabela['regex_2'],tabela['Dormitórios'],tabela['Banheiros'],tabela['Suites'],tabela['vagas'],tabela['M2'],tabela['link'],tabela['Preço'],tabela['Imovel'],tabela['Cidade'],tabela['Bairro2'],tabela['B'] = tabela['regex_2'].str.replace('\D',''),tabela['Dormitórios'].str.replace('\D',''),tabela['Banheiros'].str.replace('\D',''),tabela['Suites'].str.replace('\D',''),tabela['vagas'].str.replace('\D',''),tabela['M2'].str.replace('\D',''),tabela['link'].str.replace("\D[']|[']\D",""),tabela['Preço'].str.replace("\D[']|[']\D",""),tabela['Imovel'].str.replace("\W",""),tabela['Cidade'].str.replace('.{1,10}[-]\s',''),tabela['Bairro2'].str.replace('\W',''),tabela['B'].str.replace(".{1,10}['][no]{2}\s","")
    tabela['B'] = tabela['B'].str.replace("\s[-].{1,10}['].{1,5}","")
    tabela.to_excel(f'\\Users\Gustavo\Desktop\Limpeza\TESTE.xlsx',index=False)
    tabela = tabela.drop(columns=['Bairro',
                                  'Código'
                                  ,'Home',
                                  'Link',
                                  'regex_1',
                                  'Cidade_Bairro'])
    tabela = tabela.rename(columns={'regex_2' :'Código do Imóvel',
                                    'vagas':'Vagas',
                                    'B':'Bairro',
                                    'Bairro2':'Estado',
                                    'link':'Hyperlink',
                                    'Preço':'Valor'})




    tabela['Banheiros'] = tabela['Banheiros'].replace({
        '':'NULL'
        })
    tabela['Suites'] = tabela['Suites'].replace({
        '':'NULL'
        })
    tabela['Vagas'] = tabela['Vagas'].replace({
        '':'NULL'
     })
    tabela['M2'] = tabela['M2'].replace({
        '':'NULL'
        })
    tabela['Valor'] = tabela['Valor'].replace({
        '':'NULL'
        })
    tabela['Código do Imóvel'] = tabela['Código do Imóvel'].replace({
     '':'NULL'
        })
    tabela['Hyperlink'] = tabela['Hyperlink'].replace({
        '':'NULL'
        })
    tabela['Imovel'] = tabela['Imovel'].replace({
        '':'NULL'
        })
    tabela['Bairro'] = tabela['Bairro'].replace({
        '':'NULL'
        })
    tabela['Cidade'] = tabela['Cidade'].replace({
        '':'NULL'
     })
    tabela['Estado'] = tabela['Estado'].replace({
        '':'NULL'
        })

    tabela['Dormitórios'] = tabela['Dormitórios'].replace({
        '':'NULL'
        })
    
    
    tabela = tabela[['Código do Imóvel',
                 'Imovel',
                 'Valor',
                 'Dormitórios',
                 'Suites',
                 'Banheiros',
                 'Vagas',
                 'M2',
                 'Bairro',
                 'Cidade',
                 'Estado',
                 'Hyperlink']]

    tabela.to_excel(f'\\Users\Gustavo\Desktop\Limpeza\Tabela Tratada.xlsx',index=False)



Tratamento_excel()

# scraping_apartamento()

