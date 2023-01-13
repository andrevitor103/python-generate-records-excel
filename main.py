import openpyxl as op
import json

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

with open('values.json', encoding='utf-8') as values:
    records = json.load(values)

result = generateDatas(records=records)

print(result)

createSpreadSheetAndRecords(values = result)