import pandas as pd #vai ler o arquivo excele  evai armazena-lo


#importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')


#visualização da base de dados
pd.set_option('display.max_columns', None)#mostrar as colunas sem um limite máximo definido
print(tabela_vendas)


#faturamento (valor final)
faturamento = tabela_vendas [['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)


#quantidade de produtos
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)


print('-' * 50)
#ticket medio por produto
ticket_medio = (faturamento['Valor Final'] / quantidade ['Quantidade']).to_frame()




#enviar gmail com relatório
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreatItem(0)
mail.To = 'meduardarodrigues@gmail.com'
mail.Subject = 'Relatório de Vendas'
mail.HTMLBody= f''' 
<p>Prezadas,
Segue o Relatório de vendas por cada loja!
    
    
    Faturamento: {faturamento.to_html()}
    
    
    Quantidade Vendida: {quantidade.to_html()}
    
    
    Ticket Médio dos Produtos em cada Loja: {ticket_medio.to_html()}
    
    
    Qualquer duvida estarei á disposição!!
    Att.,
    Maria</p>
    '''


mail.Send()


print('Email Enviado!')


