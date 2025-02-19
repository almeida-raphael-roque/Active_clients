#criar biblioteca para usar o windows32 como cliente win32com.client
import win32com.client as win32
from datetime import date
import pandas as pd

#criando hoje com datetime e importando date
today = date.today()
today_format = today.strftime("%d/%m/%Y")

#criando dataframe a partir da planilha

file_path = r"C:\Users\raphael.almeida\Grupo Unus\analise de dados - Arquivos em excel\CAMPANHA_RANKING_ATIVACOES.xlsx"
df = pd.read_excel(file_path, sheet_name = "ATIVAÇÕES")
df_viavante = df[df['empresa']=='Viavante'].drop_duplicates(subset=['cliente'], keep = 'first')
df_stcoop = df[df['empresa']=='Stcoop'].drop_duplicates(subset=['cliente'], keep = 'first')
df_segtruck = df[df['empresa']=='Segtruck'].drop_duplicates(subset=['cliente'], keep = 'first')

clientes_seg = len(df_segtruck)
clientes_st = len(df_stcoop)
clientes_via = len(df_viavante)
clientes_geral = clientes_seg + clientes_st + clientes_via

def enviar_email():

    #integrar python com outlook para despachar pela aplicação do outlook
    outlook = win32.Dispatch('outlook.application')

    #criar o objeto do email
    email = outlook.CreateItem(0)

    #secretaria01@grupounus.com.br; dados03@grupounus.com.br 

    email.To = 'dados13@grupounus.com.br; secretaria01@grupounus.com.br; dados03@grupounus.com.br'
            
    email.Subject = f'[CLIENTES ATIVOS POR EMPRESA] - {today_format}'

    email.HTMLBody = f"""
    <p>Prezados,</p>

    <p>Segue em anexo a quantidade de clientes que possuem conjuntos ativos, por cooperativa, do dia {today_format}.</p>

   <b>SEGTRUCK: {clientes_seg}</b><br>
    <b>STCOOP: {clientes_st}</b><br>
    <b>VIAVANTE: {clientes_via}</b><br>
    <b>TOTAL: {clientes_geral}</b>

    <p>Atenciosamente,</p>

    <p><b>Equipe de Análise de Dados - Grupo Unus</p></b>

    <p><i>(Esse é um e-mail automático, por favor não responda)</i></p>""" 

    email.Send()



enviar_email()
