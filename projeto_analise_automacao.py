# importar biblioteca "pandas"
import pandas as pd
import win32com.client as win32

# passo 1 - importar a base de dados
tabela_vendas = pd.read_excel('vendas.xlsx')

# passo 2 - visualizar a base de dados
# mostrar todas as colunas
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# passo 3 - faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# passo 4 - quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# print para separar um passo do outro com traços
print('-' * 50)

# passo 5 - ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# passo 6 - enviar um email com o relatório
# código de outlook para enviar email (padrão)
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'lessacontabilconsult@gmail.com'
mail.Subject = 'Teste envio tabela'
mail.HTMLBody = f'''
<p>Prezados,</p> 

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida, estou à disposição.</p>

<p>Att</p>
'''

mail.Send()

print('Email Enviado')
