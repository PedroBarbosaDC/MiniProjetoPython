from itertools import groupby
import pandas as pd

# importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx') 

# visualizar a base de dados
pd.set_option('display.max_columns', None) # mostrar todas as colunas da tabela
#print(tabela_vendas[['ID Loja','Valor Final']]) # filtrar algumas colunas da tabela

# faturamento por loja
# agrupar coluna somando  valores
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)













