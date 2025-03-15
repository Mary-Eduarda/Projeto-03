import pandas as pd
import smtplib
import email.message

# Importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualização da base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# Faturamento (valor final)
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produtos
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
# Ticket médio por produto
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})


# Envio pelo Gmail
def enviar_email():
    corpo_email = f'''
    <p>Olá!</p>
    <p>Segue o arquivo</p>
    <p><strong>Faturamento:</strong></p>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

    <p><strong>Quantidade Vendida:</strong></p>
    {quantidade.to_html()}

    <p><strong>Ticket Médio dos Produtos em cada Loja:</strong></p>
    {ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

    <p>Qualquer dúvida, estarei à disposição!</p>
    <p>Atenciosamente,</p>
    <p>Maria</p>
    '''

    msg = email.message.Message()
    msg['Subject'] = "Tabela de Vendas"
    msg['From'] = 'yyduds91@gmail.com'  # Coloque seu Gmail!
    msg['To'] = 'yyduds91@gmail.com'  # Coloque o destinatário do email!
    password = 'jcec abzpp awnghso'  # Senha de aplicativo gerada para o uso
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = None #inicializando a variavel s com None

    try:
        s = smtplib.SMTP('smtp.gmail.com:587')
        s.starttls()
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        print('Email enviado com sucesso!')

    except smtplib.SMTPAuthenticationError as e:
        print(f"Falha na autenticação: {e}")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")

    finally:
     if s:
        s.quit()
enviar_email()