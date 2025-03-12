import pandas as pd #vai ler o arquivo excel  e vai armazena-lo

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
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})

#envio pelo gmail
import smtplib
import email.message

def enviar_email():
    corpo_email = f'''
    <p>Olá!</p>
    <p>Segue o arquivo</p>
      Faturamento: {faturamento.to_html(formatter={'Valor Final': 'R${:,2.f}'.format})}
    
    
    Quantidade Vendida: {quantidade.to_html()}
    
    
    Ticket Médio dos Produtos em cada Loja: {ticket_medio.to_html(formatters={'Tiket Médio': 'R${:,.2f'.format})}
   
    <p> Qualquer duvida estarei á disposição!!
    Att.,
    Maria</p>
    '''

    msg = email.message.Message()
    msg['Subject'] = "Tabela de Vendas"
    msg['From'] = 'yyduds91@gmail.com' #coloque seu gmail!
    msg['To'] = 'yyduds91@gmail.com' #coloque seu gmail!
    password = 'j c e c a b z p p a w n g h s o' #esta senha é específica para o uso correto do código, por favor não mude<3!
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credenciais
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')



