import pandas as pd
import win32com.client as win32


tabela_vendas = pd.read_excel('Vendas.xlsx')

faturamento_loja = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

ticket = (faturamento_loja['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket = ticket.rename(columns = {0: 'Ticket Médio'})

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'To address' #Endereço a quem se deve enviar 
mail.Subject = 'Relatório de vendas' #Assunto da mensagem 

#Conteúdo do email:
mail.HTMLBody = f''' 
<p>Segue relátorio de vendas por loja:</p>

<p>Faturamento por loja:</p>
{faturamento_loja.to_html(formatters ={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade de produtos vendidos por loja:</p>
{quantidade.to_html()}

<p>Ticket médio dos produtos em cada loja:</p>
{ticket.to_html(formatters = {'Ticket Médio' : 'R${:,.2f}'.format})}

<p>Caso reste alguma dúvida, entre em contato com o administrador.</p>
<p>Att</p>
<p>Administrador</p>
'''

mail.Send()
print('=' * 50)
print('Processo Concluído')