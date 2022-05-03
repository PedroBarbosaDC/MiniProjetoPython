from itertools import groupby
import pandas as pd

# importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx') 

# visualizar a base de dados
pd.set_option('display.max_columns', None) # mostrar todas as colunas da tabela
#print(tabela_vendas[['ID Loja','Valor Final']]) # filtrar algumas colunas da tabela

print('-'*50)
# faturamento por loja
# agrupar coluna somando  valores
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('-'*50)
# quantidades de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-'*50)
# ticket médio de vendas
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame() # faz a divisão e salva resultado numa tabela(to_frame)
print(ticket_medio)

# enviar email com as tabelas
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'To address'
mail.Subject = 'Message subject'
mail.HTMLBody = '<h2>HTML Message body</h2>'

mail.Send()








