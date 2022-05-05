import email
from itertools import groupby
import pandas as pd
import win32com.client as win32

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
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar email com as tabelas
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'pedro_barbosa_tec@outlook.com'
#mail.CC = 'pedaocds@gmail.com'  # envia email com copia
mail.Attachments.Add('E:\Vendas.xlsx')  #enviar um anexo
mail.Subject = 'Relatório de Vendas por Loja' #titulo do email
mail.HTMLBody = f''' 
<p>Prezados,</p> 

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p> 
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>
<p>Att.,</p>
<p>Pedro</p>
'''

mail.Send()
print('#'*50)
print('Email enviado')






