# > Importar a base de dados -> instalar o pandas e o openpyxl
import pandas as pd
import win32com.client as win32

# > Visualizar a base de dados -> utilizar o comando set_option para o pandas não limitar a quantidade de colunas
tabela_vendas = pd.read_excel("Vendas.xlsx")
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# > Faturamento por loja -> Criar uma lista com ID Loja e Valor final e atribuir a uma nova variável e depois agrupar
# exibidas lojas pelo ID e somar os valos finais
faturamento_lojas = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento_lojas)

# > Quantidade de produtos vendidos por loja
quantidade_produtos_vendidos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade_produtos_vendidos)

# > Ticket médio por produto vendidos por loja
ticket_produtos_vendidos = tabela_vendas[['ID Loja', 'Quantidade', 'Valor Final']].groupby('ID Loja').sum()
ticket_produtos_vendidos['Ticket Médio'] = round(ticket_produtos_vendidos['Valor Final'] / ticket_produtos_vendidos['Quantidade'], 1)
print(ticket_produtos_vendidos)

# > Enviar um email com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'worlorns@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f''''
<h2>Prezados,</h2>

<p>Segue o relatório de Vendas por cada Loja.</p>

<h3>Relatório:</h3>
{ticket_produtos_vendidos.to_html(formatters={'Valor Final': 'R$ {:,.2f}'.format, 'Ticket Médio': 'R$ {}'.format})}

<p>att, Leonardo Sanders.</p>
'''

mail.Send()

print('E-mail enviado!')