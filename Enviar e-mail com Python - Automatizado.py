import win32com.client as win32

# integração com o outlook
outlook = win32.Dispatch('outlook.application')

# Email novo
email = outlook.CreateItem(0)

faturamento = 1500
qtde_produtos = 10
ticket_medio = faturamento / qtde_produtos

# Informações do seu e-mail _ meu pessoal
email.To = "pjonathas972@gmail.com"
email.Subject = "Status faturamento"
email.HTMLBody = f"""
<p>Olá Cliente, aqui é o status de faturamento mensal, caso quiser gráfico solicite!</p>

<p>O faturamento da loja foi de R${faturamento}</p>
<p>Vendemos {qtde_produtos} produtos</p>
<p>O ticket Médio foi de R${ticket_medio}</p>

<p>Abs,</p>
<p>'EMPRESA FORNECIDA DO SERVIÇO'</p>
"""

# anexo = "C://Users/jonathas/Downloads/arquivo.xlsx"
# email.Attachments.Add(anexo)

email.Send()
print("Email Enviado")
