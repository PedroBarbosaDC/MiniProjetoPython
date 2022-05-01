import pandas as pd

# importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx') 

# visualizar a base de dados
pd.set_option('display.max_columns', None) # mostrar todas as colunas da tabela
print(tabela_vendas)