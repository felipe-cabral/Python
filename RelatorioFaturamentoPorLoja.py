import pandas as pd

# Importar a base de dados
tabela_vendas = pd.read_excel("Vendas.xlsx")

# Visualizar a base de dados
print(tabela_vendas)
pd.set_option('display.max_columns', None)

# Faturamento por Loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-'*50)

# Quantidade por Loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-'*50)

# Ticket médio por produto por Loja
ticket_medio = (faturamento['Valor Final']/quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0:'Ticket Médio'})
print(ticket_medio)
print('-'*50)

# Enviar e-mail com o relatório

import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'felipecabralaugusto@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''<p>Prezados</p>

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}


<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição!</p>
'''


mail.Send()

print('Email Enviado!')
