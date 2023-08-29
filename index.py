import openpyxl
import datetime

# Preparando de forma dinâmica para pegar o nome das planilhas
data_atual = datetime.datetime.now()
mes_atual = data_atual.month
ano_atual = str(data_atual.year)[-2:]
nome_meses = ['Jan','Fev','Mar','Abr','Maio','Jun','Jul','Ago', 'Set','Out','Nov','Dez']
nome_colaborador = 'Davi Ghiggino'

hora_extra = False

nome_planilha = f'Timesheet - {nome_colaborador} - {nome_meses[mes_atual-1]} {ano_atual}.xlsx'
workbook = openpyxl.load_workbook(nome_planilha)
sheet = workbook.active

# Preenchendo a planilha em todos os dias úteis
sheet['D2'] = nome_colaborador
for index in range(9,40):
  if(sheet[f'K{index}'].value is None or hora_extra):
    sheet[f'C{index}'] = "09:00:00"  
    sheet[f'D{index}'] = "18:00:00"  
    sheet[f'E{index}'] = "01:00:00"  
  
    

workbook.save(nome_planilha)
print("Dados escritos com sucesso na planilha!")

