#Importando bibliotecas
import win32com.client
import datetime
import time
import sys
from dateutil.relativedelta import relativedelta

hoje = datetime.datetime.now()
ano = hoje.year
mes = hoje.month
dia = hoje.day
hora = hoje.hour
minuto = hoje.minute

inicio_mes = str(ano)+'-'+str(mes)+'-'+'01'
#------------------------------INFORMACOES DO USUARIO-------------------------------------------------------------------
pergunta = input('Digite s caso deseje enviar o faturamento para a Renata Gava e Thiago Pereira. Para digitar novos destinatários, digite qualquer outra informação.\n')
if pergunta == 's':
    emails = "renata.gava@hkm.ind.br;thiago.pereira@hkm.ind.br"
else:
    emails = ''
    while True:
        novo_email = input('Digite os logins dos destinatários (ex: carlos.junior) separados por enter. Para encerrar, pressione enter com o campo vazio.')
        if not novo_email:
            break
        emails += novo_email + '@hkm.ind.br;'
print('Por favor, aguarde.')
#------------------------------EXCEL-------------------------------------------------------------------
xlapp = win32com.client.DispatchEx("Excel.Application")
xlapp.Visible = 0
wb = xlapp.workbooks.open("C:\\Users\\carlos.junior\\Desktop\\FaturamentoDiario\\Faturamento_Diario.xlsm")
xlapp.Application.Run("FiltrosDatasPivots")
wb.RefreshAll()
xlapp.CalculateUntilAsyncQueriesDone()
xlapp.Application.Run("exportpic")
wb.Save()
wb.Close()
xlapp.Quit()
path_to_jpg = "C:\\Users\\carlos.junior\\Desktop\\FaturamentoDiario\\Faturamento.jpg"
#-------------------------------PANDAS-----------------------------------------------------------------
import pandas as pd
from babel.numbers import format_currency
#from pretty_html_table import build_table

faturamento = pd.read_excel(
    "C:\\Users\\carlos.junior\\Desktop\\FaturamentoDiario\\Faturamento_Diario.xlsm", sheet_name="vwFaturamentoOS")

filtro_data = faturamento[(faturamento['Data de Emissão da NF'] >= inicio_mes)]
filtro_data = filtro_data[['OS', 'NF', 'Valor Médio da OS', 'Valor Restante a Faturar', 'Data de Emissão da NF', 'Cliente', 'Carteirista']]
filtro_data = filtro_data.rename(columns={'Valor Médio da OS': 'Valor da OS'})
filtro_data['NF'] = filtro_data['NF'].astype('int')
faturamento_mensal = filtro_data['Valor da OS'].sum()
faturamento_mensal = "R${:,.2f}".format(faturamento_mensal)
filtro_data["Valor da OS"] = filtro_data["Valor da OS"].apply(lambda x: format_currency(x, currency="BRL", locale="pt_BR"))
filtro_data["Valor Restante a Faturar"] = filtro_data["Valor Restante a Faturar"].apply(lambda x: format_currency(x, currency="BRL", locale="pt_BR"))
df_filtro_data = pd.DataFrame(filtro_data)
df_filtro_data
body_filtro_data = '<html><body>' + df_filtro_data.to_html() + '</body></html>'

# -----------------------------OUTLOOK----------------------------------------------------------------
outlook = win32com.client.Dispatch("Outlook.Application")
Msg = outlook.CreateItem(0)
Msg.To = emails
Msg.Subject = f"Faturamento Diário {dia}-{mes}-{ano}"
Msg.HTMLBody = f'''
 Bom dia,<br><br>
 No presente momento, o faturamento do mês está em {faturamento_mensal}.<br>
 Abaixo, a lista detalhada:<br>
 
 {body_filtro_data}

 Em caso de dúvidas ou sugestões, favor entrar em contato.<br>
 Este é um e-mail automático, mas sinta-se livre para respondê-lo.
 '''
Msg.Attachments.Add(path_to_jpg)
Msg.Send()

fim = input('Fim do script. Pressione enter para finalizar.')