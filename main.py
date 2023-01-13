import openpyxl as op
import json
from PySimpleGUI import PySimpleGUI as sg

def createSpreadSheetAndRecords(name = 'teste', values = []):
    book = op.Workbook()

    book.create_sheet(name)

    page = book[name]

    page.append(['client', 'method', 'currency', 'value'])

    for value in values:
        page.append(value)
    
    book.save( name + '.xlsx' )

def getAmount(amount):
    amount_result = []
    amount_result.append(amount['currency'])
    amount_result.append(amount['value'])
    return amount_result
    


def generateDatas(records):
    datas = []
    for record in records:
        datas.append([
                record['client'],
                record['method'],
                *getAmount(amount=record['amount'])
        ])
    return datas

def fileGenerate(file_name):
    with open('values.json', encoding='utf-8') as data:
        records = json.load(data)

    result = generateDatas(records=records)

    print(result)

    createSpreadSheetAndRecords(name=file_name, values = result)


sg.theme('Reddit')

layout = [
    [sg.Text('nome do arquivo'), sg.Input(key='file_name')],
    [sg.Button('Gerar Arquivo')]
]

page = sg.Window('Gerar arquivo', layout)

while True:
    events, values = page.read()

    if events == sg.WINDOW_CLOSED:
        break

    if events == 'Gerar Arquivo':
        print(values)
        if values['file_name'] == '' or values['file_name'] is None:
            sg.popup_error('Ih rapaz... tem que adicionar um nome para o arquivo ai :sad_pepe:')
        else:
            fileGenerate(file_name=values['file_name'])
            break