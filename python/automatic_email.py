#criar biblioteca para usar o windows32 como cliente win32com.client
import win32com.client as win32
from datetime import date
#criando hoje com datetime e importando date

today = date.today()

today_format = today.strftime("%d/%m/%Y")

#integrar python com outlook para despachar pela aplicação do outlook
outlook = win32.Dispatch('outlook.application')

#criar o objeto do email
email = outlook.CreateItem(0)

#secretaria01@grupounus.com.br; dados03@grupounus.com.br 

email.To = 'dados13@grupounus.com.br'
          
email.Subject = f'[CLIENTES ATIVOS POR EMPRESA] - {today_format}'

email.HTMLBody = f"""
<p>Prezados,</p>

<p>Segue em anexo a quantidade de clientes que possuem conjuntos ativos por cooperativa.</p>

<p>Atenciosamente,</p>

<p>Equipe de Análise de Dados.</p>

<p><i>(Esse é um e-mail automático, por favor não responda)</i></p>""" 

email.Send()