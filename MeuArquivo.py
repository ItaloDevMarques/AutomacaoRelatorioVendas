import pandas as pd
import win32com.client as win32

#importação da base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

#visualização da base de dados
pd.set_option('display.max_columns', None)

#faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum(numeric_only=True)
print(faturamento)
print('-' *50)

#produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum(numeric_only=True)
print(quantidade)
print('-' *50)

#valor médio por produto em cada loja
valor_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
valor_medio = valor_medio.rename(columns={0: 'Valor Médio'})
print(valor_medio)

#configuração para envio automatico de relatório a um ou mais emails
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'Aqui vai o(os) email(s) destinatário(s)'
mail.Subject = 'Aqui vai o assunto/titulo do relatório'
#abaixo segue um exemplo dos dados tratados, e formatados em tabelas para serem enviados pelo email
#neste campo abaixo é tudo o que será enviado no corpo do email, precisa ser utilizado formatação em html
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Valor Médio dos Produtos em cada Loja:</p>
{valor_medio.to_html(formatters={'Valor Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>[Insira o seu nome aqui] </p>
'''
mail.Send()

print('Email Enviado com Sucesso!')


