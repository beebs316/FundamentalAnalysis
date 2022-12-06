from polygon import RESTClient
import numpy as np
#import talib as ta
import json
import requests
import pprint
import datetime
from datetime import timedelta
import pandas as pd
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
from xlwings import Range


def writeDataAnnual(earningsDataAnnual,col):
    joe = earningsDataAnnual
    start_date_Annual_1 = joe['start_date']
    end_date_Annual_1 = joe['end_date']
    fiscal_period_Annual_1 = joe['fiscal_period']
    fiscal_year_Annual_1 = joe['fiscal_year']
    revenues_Annual_1 = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues_Annual_1 = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin_Annual_1 = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent_Annual_1 = (gross_margin_Annual_1 / revenues_Annual_1) * 100
    opperating_expenses_Annual_1 = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss_Annual_1 = joe['financials']['income_statement']['net_income_loss']['value']
    #diluted_earnings_per_share_Annual_1 = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period_Annual_1, fiscal_year_Annual_1)
    print('Time Period:', start_date_Annual_1, '-', end_date_Annual_1)
    print('Revenues =', revenues_Annual_1)
    print('Cost of Revenues =', cost_of_revenues_Annual_1)
    print('Gross Margin =', gross_margin_Annual_1)
    print('Gross Margin Mercent =', gross_margin_percent_Annual_1, '%')
    print('Opperating Expenses =', opperating_expenses_Annual_1)
    print('Net Income =', net_income_loss_Annual_1)
    #print('Diluted EPS =', diluted_earnings_per_share_Annual_1)

    ws1.range(col + '36').value = end_date_Annual_1
    ws1.range(col + '37').value = fiscal_period_Annual_1 + " " + fiscal_year_Annual_1
    ws1.range(col + '38').value = (revenues_Annual_1 / 1000000)
    ws1.range(col + '39').value = (cost_of_revenues_Annual_1 / 1000000)
    ws1.range(col + '40').value = (gross_margin_Annual_1 / 1000000)
    ws1.range(col + '41').value = (net_income_loss_Annual_1 / 1000000)
    ws1.range(col + '42').value = (opperating_expenses_Annual_1 / 1000000)
    ws1.range(col + '43').value = gross_margin_percent_Annual_1
    #ws1.range(col + '44').value = diluted_earnings_per_share_Annual_1

key = ""
client = RESTClient(key)

wbTest = xw.Book('earningsData.xlsx')
ws1 = wbTest.sheets['Sheet1']

#wbTest.
# ws1.charts[3].delete()
# ws1.charts[2].delete()
# ws1.charts[1].delete()
# ws1.charts[0].delete()
#print(ws1.charts.delete)

existingCharts = len(ws1.charts)
print(existingCharts)

while existingCharts > 0:
    deleteValue = existingCharts - 1
    ws1.charts[deleteValue].delete()
    existingCharts = len(ws1.charts)

ws1.clear()




#ws1.charts[0].delete()
#Range('A1:T14').api.Delete(DeleteShiftDirection.xlShiftToLeft)  # or xlShiftUp

tiker = "TSLA"
url = 'https://api.polygon.io/vX/reference/financials?ticker=' + tiker + '&timeframe=quarterly&limit=9&apiKey=' + key
urlAnnual = 'https://api.polygon.io/vX/reference/financials?ticker=' + tiker + '&timeframe=annual&limit=16&apiKey=_1q67Oh4iSpjA1oGCnQ31PYh9cOIK6Rp'
print(url)
response = requests.get(url).json()
data = list(response.items())
joe = (data[0])

responseAnnual = requests.get(urlAnnual).json()
dataAnnual = list(responseAnnual.items())
joeAnnual = (dataAnnual[0])
# print(joe[0])
# print(type(joe[1]))
# print(joe[2])
# print(type(data[0]))
joe = joe[1]
joeAnnual = joeAnnual[1]
# print(joe[0])
# print(joe[1])
sizeAnnual = len(joeAnnual)
print(len(joeAnnual))
earningsData_1 = joe[0]
earningsData_2 = joe[1]
earningsData_3 = joe[2]
earningsData_4 = joe[3]
earningsData_5 = joe[6]

earningsDataTotal_1 = joe[2]
earningsDataTotal_2 = joe[3]
earningsDataTotal_3 = joe[4]

earningsDataTotal_1_Q3 = joe[3]
earningsDataTotal_2_Q3 = joe[4]
earningsDataTotal_3_Q3 = joe[5]

earningsDataAnnual_1 = joeAnnual[0]
earningsDataAnnual_2 = joeAnnual[0]
earningsDataAnnual_3 = joeAnnual[0]
earningsDataAnnual_4 = joeAnnual[0]
earningsDataAnnual_5 = joeAnnual[0]
earningsDataAnnual_6 = joeAnnual[0]
earningsDataAnnual_7 = joeAnnual[0]
earningsDataAnnual_8 = joeAnnual[0]
earningsDataAnnual_9 = joeAnnual[0]
earningsDataAnnual_10 = joeAnnual[0]
graphStart = ''
if (sizeAnnual > 0):
    earningsDataAnnual_1 = joeAnnual[0]
if sizeAnnual > 1:
    earningsDataAnnual_2 = joeAnnual[1]
if sizeAnnual > 2:
    earningsDataAnnual_3 = joeAnnual[2]
if sizeAnnual > 3:
    earningsDataAnnual_4 = joeAnnual[3]
if sizeAnnual > 4:
    earningsDataAnnual_5 = joeAnnual[4]
if sizeAnnual > 5:
    earningsDataAnnual_6 = joeAnnual[5]
if sizeAnnual > 6:
    earningsDataAnnual_7 = joeAnnual[6]
if sizeAnnual > 7:
    earningsDataAnnual_8 = joeAnnual[7]
if sizeAnnual > 8:
    earningsDataAnnual_9 = joeAnnual[8]
if sizeAnnual > 9:
    earningsDataAnnual_10 = joeAnnual[9]

earningsDataYearByQuarter_1_1 = joe[0]
earningsDataYearByQuarter_1_2 = joe[3]
earningsDataYearByQuarter_1_3 = joe[6]

earningsDataYearByQuarter_2_1 = joe[1]
earningsDataYearByQuarter_2_2 = joe[4]
earningsDataYearByQuarter_2_3 = joe[7]

earningsDataYearByQuarter_3_1 = joe[2]
earningsDataYearByQuarter_3_2 = joe[5]
earningsDataYearByQuarter_3_3 = joe[8]

joe = earningsDataAnnual_1
start_date_Annual_1 = joe['start_date']
end_date_Annual_1 = joe['end_date']
fiscal_period_Annual_1 = joe['fiscal_period']
fiscal_year_Annual_1 = joe['fiscal_year']
revenues_Annual_1 = joe['financials']['income_statement']['revenues']['value']
cost_of_revenues_Annual_1 = joe['financials']['income_statement']['cost_of_revenue']['value']
gross_margin_Annual_1 = joe['financials']['income_statement']['gross_profit']['value']
gross_margin_percent_Annual_1 = (gross_margin_Annual_1 / revenues_Annual_1) * 100
opperating_expenses_Annual_1 = joe['financials']['income_statement']['operating_expenses']['value']
net_income_loss_Annual_1 = joe['financials']['income_statement']['net_income_loss']['value']
#diluted_earnings_per_share_Annual_1 = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
print('')
print('Previous Quarter')
print('Fiscal Period and Year =', fiscal_period_Annual_1, fiscal_year_Annual_1)
print('Time Period:', start_date_Annual_1, '-', end_date_Annual_1)
print('Revenues =', revenues_Annual_1)
print('Cost of Revenues =', cost_of_revenues_Annual_1)
print('Gross Margin =', gross_margin_Annual_1)
print('Gross Margin Mercent =', gross_margin_percent_Annual_1, '%')
print('Opperating Expenses =', opperating_expenses_Annual_1)
print('Net Income =', net_income_loss_Annual_1)
#print('Diluted EPS =', diluted_earnings_per_share_Annual_1)


if sizeAnnual >= 10:
    joe = earningsDataAnnual_1
    writeDataAnnual(joe, 'K')
    joe = earningsDataAnnual_2
    writeDataAnnual(joe, 'J')
    joe = earningsDataAnnual_3
    writeDataAnnual(joe, 'I')
    joe = earningsDataAnnual_4
    writeDataAnnual(joe, 'H')
    joe = earningsDataAnnual_5
    writeDataAnnual(joe, 'G')
    joe = earningsDataAnnual_6
    writeDataAnnual(joe, 'F')
    joe = earningsDataAnnual_7
    writeDataAnnual(joe, 'E')
    joe = earningsDataAnnual_8
    writeDataAnnual(joe, 'D')
    joe = earningsDataAnnual_9
    writeDataAnnual(joe, 'C')
    joe = earningsDataAnnual_10
    writeDataAnnual(joe, 'B')
if sizeAnnual == 9:
    joe = earningsDataAnnual_1
    writeDataAnnual(joe, 'J')
    joe = earningsDataAnnual_2
    writeDataAnnual(joe, 'I')
    joe = earningsDataAnnual_3
    writeDataAnnual(joe, 'H')
    joe = earningsDataAnnual_4
    writeDataAnnual(joe, 'G')
    joe = earningsDataAnnual_5
    writeDataAnnual(joe, 'F')
    joe = earningsDataAnnual_6
    writeDataAnnual(joe, 'E')
    joe = earningsDataAnnual_7
    writeDataAnnual(joe, 'D')
    joe = earningsDataAnnual_8
    writeDataAnnual(joe, 'C')
    joe = earningsDataAnnual_9
    writeDataAnnual(joe, 'B')
if sizeAnnual == 8:
    joe = earningsDataAnnual_1
    writeDataAnnual(joe, 'I')
    joe = earningsDataAnnual_2
    writeDataAnnual(joe, 'H')
    joe = earningsDataAnnual_3
    writeDataAnnual(joe, 'G')
    joe = earningsDataAnnual_4
    writeDataAnnual(joe, 'F')
    joe = earningsDataAnnual_5
    writeDataAnnual(joe, 'E')
    joe = earningsDataAnnual_6
    writeDataAnnual(joe, 'D')
    joe = earningsDataAnnual_7
    writeDataAnnual(joe, 'C')
    joe = earningsDataAnnual_8
    writeDataAnnual(joe, 'B')
if sizeAnnual == 7:
    joe = earningsDataAnnual_1
    writeDataAnnual(joe, 'H')
    joe = earningsDataAnnual_2
    writeDataAnnual(joe, 'G')
    joe = earningsDataAnnual_3
    writeDataAnnual(joe, 'F')
    joe = earningsDataAnnual_4
    writeDataAnnual(joe, 'E')
    joe = earningsDataAnnual_5
    writeDataAnnual(joe, 'D')
    joe = earningsDataAnnual_6
    writeDataAnnual(joe, 'C')
    joe = earningsDataAnnual_7
    writeDataAnnual(joe, 'B')
if sizeAnnual == 6:
    joe = earningsDataAnnual_1
    writeDataAnnual(joe, 'G')
    joe = earningsDataAnnual_2
    writeDataAnnual(joe, 'F')
    joe = earningsDataAnnual_3
    writeDataAnnual(joe, 'E')
    joe = earningsDataAnnual_4
    writeDataAnnual(joe, 'D')
    joe = earningsDataAnnual_5
    writeDataAnnual(joe, 'C')
    joe = earningsDataAnnual_6
    writeDataAnnual(joe, 'B')
if sizeAnnual == 5:
    joe = earningsDataAnnual_1
    writeDataAnnual(joe, 'F')
    joe = earningsDataAnnual_2
    writeDataAnnual(joe, 'E')
    joe = earningsDataAnnual_3
    writeDataAnnual(joe, 'D')
    joe = earningsDataAnnual_4
    writeDataAnnual(joe, 'C')
    joe = earningsDataAnnual_5
    writeDataAnnual(joe, 'B')
if sizeAnnual == 4:
    joe = earningsDataAnnual_1
    writeDataAnnual(joe, 'E')
    joe = earningsDataAnnual_2
    writeDataAnnual(joe, 'D')
    joe = earningsDataAnnual_3
    writeDataAnnual(joe, 'C')
    joe = earningsDataAnnual_4
    writeDataAnnual(joe, 'B')
if sizeAnnual == 3:
    joe = earningsDataAnnual_1
    writeDataAnnual(joe, 'D')
    joe = earningsDataAnnual_2
    writeDataAnnual(joe, 'C')
    joe = earningsDataAnnual_3
    writeDataAnnual(joe, 'B')
if sizeAnnual == 3:
    joe = earningsDataAnnual_1
    writeDataAnnual(joe, 'C')
    joe = earningsDataAnnual_2
    writeDataAnnual(joe, 'B')




joe = earningsData_1
start_date_1 = joe['start_date']
end_date_1 = joe['end_date']
fiscal_period_1 = joe['fiscal_period']
fiscal_year_1 = joe['fiscal_year']
revenues_1 = joe['financials']['income_statement']['revenues']['value']
cost_of_revenues_1 = joe['financials']['income_statement']['cost_of_revenue']['value']
gross_margin_1 = joe['financials']['income_statement']['gross_profit']['value']
gross_margin_percent_1 = (gross_margin_1 / revenues_1) * 100
opperating_expenses_1 = joe['financials']['income_statement']['operating_expenses']['value']
net_income_loss_1 = joe['financials']['income_statement']['net_income_loss']['value']
diluted_earnings_per_share_1 = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
print('')
print(tiker, 'Earnings Data:')
print('')
print('Current quarter')
print('Fiscal Period and Year =', fiscal_period_1, fiscal_year_1)
print('Fiscal Period =', fiscal_period_1)
quarter = 0
if fiscal_period_1 == 'Q1':
    print('Q1 Confirmed')
    quarter = 1

if fiscal_period_1 == 'Q2':
    print('Q2 Confirmed')
    quarter = 2

if fiscal_period_1 == 'Q3':
    print('Q3 Confirmed')
    quarter = 3

print(type(fiscal_period_1))
print('Time Period:', start_date_1, '-', end_date_1)
print('Revenues =', revenues_1)
print('Cost of Revenues =', cost_of_revenues_1)
print('Gross Margin =', gross_margin_1)
print('Gross Margin Percent =', gross_margin_percent_1, '%')
print('Operating Expenses =', opperating_expenses_1)
print('Net Income =', net_income_loss_1)
print('Diluted EPS =', diluted_earnings_per_share_1)

ws1.range('A1').value = 'Ticker'
ws1.range('B1').value = tiker
ws1.range('A2').value = 'Date of earnings release '
ws1.range('B2').value = end_date_1

ws1.range('A3').value = 'Operating Results:'
ws1.range('A4').value = 'In Millions (except EPS)'

ws1.range('B5').value = 'Quarter by Quarter'
ws1.range('B35').value = 'Year over Year (Annual)'

ws1.range('A8').value = 'Revenues       '
ws1.range('A9').value = 'Cost of Revenues       '
ws1.range('A10').value = 'Gross Margin      '
ws1.range('A11').value = 'Net Income        '
ws1.range('A12').value = 'Operating Expenses        '
ws1.range('A13').value = 'Gross Margin Percent'
ws1.range('A14').value = 'Diluted EPS'

ws1.range('A38').value = 'Revenues               '
ws1.range('A39').value = 'Cost of Revenues               '
ws1.range('A40').value = 'Gross Margin              '
ws1.range('A41').value = 'Net Income                '
ws1.range('A42').value = 'Operating Expenses                '
ws1.range('A43').value = 'Gross Margin Percent'
ws1.range('A44').value = 'Diluted EPS'



ws1.range('B6').value = end_date_1
ws1.range('B7').value = fiscal_period_1 + " " + fiscal_year_1 + " (Current)"
ws1.range('B8').value =  (revenues_1 / 1000000)
ws1.range('B9').value =  (cost_of_revenues_1 / 1000000)
ws1.range('B10').value =  (gross_margin_1 / 1000000)
ws1.range('B11').value =  (net_income_loss_1 / 1000000)
ws1.range('B12').value =  (opperating_expenses_1 / 1000000)
ws1.range('B13').value = gross_margin_percent_1
ws1.range('B14').value = diluted_earnings_per_share_1


if quarter == 1:
    print('calc Q4 here')

joe = earningsData_2
start_date_2 = joe['start_date']
end_date_2 = joe['end_date']
fiscal_period_2 = joe['fiscal_period']
fiscal_year_2 = joe['fiscal_year']
revenues_2 = joe['financials']['income_statement']['revenues']['value']
cost_of_revenues_2 = joe['financials']['income_statement']['cost_of_revenue']['value']
gross_margin_2 = joe['financials']['income_statement']['gross_profit']['value']
gross_margin_percent_2 = (gross_margin_2 / revenues_2) * 100
opperating_expenses_2 = joe['financials']['income_statement']['operating_expenses']['value']
net_income_loss_2 = joe['financials']['income_statement']['net_income_loss']['value']
diluted_earnings_per_share_2 = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
print('')
print('Previous Quarter')
print('Fiscal Period and Year =', fiscal_period_2, fiscal_year_2)
print('Time Period:', start_date_2, '-', end_date_2)
print('Revenues =', revenues_2)
print('Cost of Revenues =', cost_of_revenues_2)
print('Gross Margin =', gross_margin_2)
print('Gross Margin Mercent =', gross_margin_percent_2, '%')
print('Opperating Expenses =', opperating_expenses_2)
print('Net Income =', net_income_loss_2)
print('Diluted EPS =', diluted_earnings_per_share_2)

ws1.range('C6').value = end_date_2
ws1.range('C7').value = fiscal_period_2 + " " + fiscal_year_2
ws1.range('C8').value =  (revenues_2 / 1000000)
ws1.range('C9').value =  (cost_of_revenues_2 / 1000000)
ws1.range('C10').value =  (gross_margin_2 / 1000000)
ws1.range('C11').value =  (net_income_loss_2 / 1000000)
ws1.range('C12').value =  (opperating_expenses_2 / 1000000)
ws1.range('C13').value = gross_margin_percent_2
ws1.range('C14').value = diluted_earnings_per_share_2

if quarter == 2:
    ws1.range('G8').value = 'Revenues               '
    ws1.range('G9').value = 'Cost of Revenues               '
    ws1.range('G10').value = 'Gross Margin              '
    ws1.range('G11').value = 'Net Income                '
    ws1.range('G12').value = 'Operating Expenses                '
    ws1.range('G13').value = 'Gross Margin Percent'
    ws1.range('G14').value = 'Diluted EPS'

    ws1.range('L8').value = 'Revenues               '
    ws1.range('L9').value = 'Cost of Revenues               '
    ws1.range('L10').value = 'Gross Margin              '
    ws1.range('L11').value = 'Net Income                '
    ws1.range('L12').value = 'Operating Expenses                '
    ws1.range('L13').value = 'Gross Margin Percent'
    ws1.range('L14').value = 'Diluted EPS'

    print('calc Q4 here')
    print('')
    print('total 1')
    joe = earningsDataTotal_1
    start_date_Total_1 = joe['start_date']
    end_date_Total_1 = joe['end_date']
    fiscal_period_Total_1 = joe['fiscal_period']
    fiscal_year_Total_1 = joe['fiscal_year']
    revenues_Total_1 = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues_Total_1 = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin_Total_1 = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_Total_1_percent = (gross_margin_Total_1 / revenues_Total_1) * 100
    opperating_expenses_Total_1 = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss_Total_1 = joe['financials']['income_statement']['net_income_loss']['value']
#    diluted_earnings_per_share_Total_1 = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print(tiker, 'Earnings Data:')
    print('')
    print('Current quarter')
    print('Fiscal Period and Year =', fiscal_period_Total_1, fiscal_year_Total_1)
    print('Fiscal Period =', fiscal_period_Total_1)

    print('')
    print('total 2')
    joe = earningsDataTotal_2
    start_date_Total_2 = joe['start_date']
    end_date_Total_2 = joe['end_date']
    fiscal_period_Total_2 = joe['fiscal_period']
    fiscal_year_Total_2 = joe['fiscal_year']
    revenues_Total_2 = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues_Total_2 = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin_Total_2 = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_Total_1_percent = (gross_margin_Total_2 / revenues_Total_2) * 100
    opperating_expenses_Total_2 = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss_Total_2 = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share_Total_2 = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print(tiker, 'Earnings Data:')
    print('')
    print('Current quarter')
    print('Fiscal Period and Year =', fiscal_period_Total_2, fiscal_year_Total_2)
    print('Fiscal Period =', fiscal_period_Total_2)

    print('')
    print('total 3')
    joe = earningsDataTotal_3
    start_date_Total_3 = joe['start_date']
    end_date_Total_3 = joe['end_date']
    fiscal_period_Total_3 = joe['fiscal_period']
    fiscal_year_Total_3 = joe['fiscal_year']
    revenues_Total_3 = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues_Total_3 = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin_Total_3 = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_Total_3_percent = (gross_margin_Total_3 / revenues_Total_3) * 100
    opperating_expenses_Total_3 = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss_Total_3 = joe['financials']['income_statement']['net_income_loss']['value']
#    diluted_earnings_per_share_Total_3 = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print(tiker, 'Earnings Data:')
    print('')
    print('Current quarter')
    print('Fiscal Period and Year =', fiscal_period_Total_3, fiscal_year_Total_3)
    print('Fiscal Period =', fiscal_period_Total_3)

    Q4revenue = revenues_Annual_1 - (revenues_Total_1 + revenues_Total_2 + revenues_Total_3)
    print('Q4 revenue =', Q4revenue)

    Q4cost_of_revenues = cost_of_revenues_Annual_1 - (cost_of_revenues_Total_1 + cost_of_revenues_Total_2 + cost_of_revenues_Total_3)
    print('Q4 cost_of_revenues =', Q4cost_of_revenues)

    Q4gross_margin = gross_margin_Annual_1 - (gross_margin_Total_1 + gross_margin_Total_2 + gross_margin_Total_3)
    print('Q4 gross_margin =', Q4gross_margin)

    Q1_Q3net_income_loss = (net_income_loss_Total_1 + net_income_loss_Total_2 + net_income_loss_Total_3)

    Q4net_income_loss = net_income_loss_Annual_1 - (net_income_loss_Total_1 + net_income_loss_Total_2 + net_income_loss_Total_3)


    print('Q3 net_income_loss =', net_income_loss_Total_1/ 1000000)
    print('Q2 net_income_loss =', net_income_loss_Total_2/ 1000000)
    print('Q1 net_income_loss =', net_income_loss_Total_3/ 1000000)
    print('Q1-Q3 net_income_loss =', Q1_Q3net_income_loss/ 1000000)
    print('2020 net_income_loss =', net_income_loss_Annual_1/ 1000000)
    print('Q4 CALCULATED net_income_loss =', (net_income_loss_Annual_1 - Q1_Q3net_income_loss)/ 1000000)
    print('Q4 net_income_loss =', Q4net_income_loss/ 1000000)

    Q4opperating_expenses = opperating_expenses_Annual_1 - (opperating_expenses_Total_1 + opperating_expenses_Total_2 + opperating_expenses_Total_3)
    print('Q4 opperating_expenses =', Q4opperating_expenses)

    ws1.range('D6').value = end_date_Annual_1
    ws1.range('D7').value = "Q4" + " " + fiscal_year_Total_3
    ws1.range('D8').value = (Q4revenue / 1000000)
    ws1.range('D9').value = (Q4cost_of_revenues / 1000000)
    ws1.range('D10').value = (Q4gross_margin / 1000000)
    ws1.range('D11').value = (Q4net_income_loss / 1000000)
    ws1.range('D12').value = (Q4opperating_expenses / 1000000)
    # ws1.range('D13').value = gross_margin_percent
    # ws1.range('D14').value = diluted_earnings_per_share

    #Q3
    joe = earningsData_3
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
#    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    #print('Diluted EPS =', diluted_earnings_per_share)

    ws1.range('E6').value = end_date
    ws1.range('E7').value = fiscal_period + " " + fiscal_year
    ws1.range('E8').value = (revenues / 1000000)
    ws1.range('E9').value = (cost_of_revenues / 1000000)
    ws1.range('E10').value = (gross_margin / 1000000)
    ws1.range('E11').value = (net_income_loss / 1000000)
    ws1.range('E12').value = (opperating_expenses / 1000000)
    ws1.range('E13').value = gross_margin_percent
    #ws1.range('E14').value = diluted_earnings_per_share

    # yearly Q2_1
    ws1.range('G5').value = 'Year over Year (Q2)'
    joe = earningsDataYearByQuarter_1_1
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    print('Diluted EPS =', diluted_earnings_per_share)
    ws1.range('H6').value = end_date
    ws1.range('H7').value = fiscal_period + " " + fiscal_year
    ws1.range('H8').value = (revenues / 1000000)
    ws1.range('H9').value = (cost_of_revenues / 1000000)
    ws1.range('H10').value = (gross_margin / 1000000)
    ws1.range('H11').value = (net_income_loss / 1000000)
    ws1.range('H12').value = (opperating_expenses / 1000000)
    ws1.range('H13').value = gross_margin_percent
    ws1.range('H14').value = diluted_earnings_per_share

    # yearly Q2_2
    joe = earningsDataYearByQuarter_1_2
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    print('Diluted EPS =', diluted_earnings_per_share)
    ws1.range('I6').value = end_date
    ws1.range('I7').value = fiscal_period + " " + fiscal_year
    ws1.range('I8').value = (revenues / 1000000)
    ws1.range('I9').value = (cost_of_revenues / 1000000)
    ws1.range('I10').value = (gross_margin / 1000000)
    ws1.range('I11').value = (net_income_loss / 1000000)
    ws1.range('I12').value = (opperating_expenses / 1000000)
    ws1.range('I13').value = gross_margin_percent
    ws1.range('I14').value = diluted_earnings_per_share

    # yearly Q2_3
    joe = earningsDataYearByQuarter_1_3
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
#    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    #print('Diluted EPS =', diluted_earnings_per_share)

    ws1.range('J6').value = end_date
    ws1.range('J7').value = fiscal_period + " " + fiscal_year
    ws1.range('J8').value = (revenues / 1000000)
    ws1.range('J9').value = (cost_of_revenues / 1000000)
    ws1.range('J10').value = (gross_margin / 1000000)
    ws1.range('J11').value = (net_income_loss / 1000000)
    ws1.range('J12').value = (opperating_expenses / 1000000)
    ws1.range('J13').value = gross_margin_percent
    #ws1.range('J14').value = diluted_earnings_per_share

    # yearly Q2_1
    ws1.range('L5').value = 'Year over Year (Q1)'
    joe = earningsDataYearByQuarter_2_1
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    #diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    #print('Diluted EPS =', diluted_earnings_per_share)
    ws1.range('M6').value = end_date
    ws1.range('M7').value = fiscal_period + " " + fiscal_year
    ws1.range('M8').value = (revenues / 1000000)
    ws1.range('M9').value = (cost_of_revenues / 1000000)
    ws1.range('M10').value = (gross_margin / 1000000)
    ws1.range('M11').value = (net_income_loss / 1000000)
    ws1.range('M12').value = (opperating_expenses / 1000000)
    ws1.range('M13').value = gross_margin_percent
   # ws1.range('M14').value = diluted_earnings_per_share

    # yearly Q2_2
    joe = earningsDataYearByQuarter_2_2
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    #diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    #print('Diluted EPS =', diluted_earnings_per_share)
    ws1.range('N6').value = end_date
    ws1.range('N7').value = fiscal_period + " " + fiscal_year
    ws1.range('N8').value = (revenues / 1000000)
    ws1.range('N9').value = (cost_of_revenues / 1000000)
    ws1.range('N10').value = (gross_margin / 1000000)
    ws1.range('N11').value = (net_income_loss / 1000000)
    ws1.range('N12').value = (opperating_expenses / 1000000)
    ws1.range('N13').value = gross_margin_percent
    #ws1.range('N14').value = diluted_earnings_per_share

    # yearly Q2_3
    joe = earningsDataYearByQuarter_2_3
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
   # diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    #print('Diluted EPS =', diluted_earnings_per_share)

    ws1.range('O6').value = end_date
    ws1.range('O7').value = fiscal_period + " " + fiscal_year
    ws1.range('O8').value = (revenues / 1000000)
    ws1.range('O9').value = (cost_of_revenues / 1000000)
    ws1.range('O10').value = (gross_margin / 1000000)
    ws1.range('O11').value = (net_income_loss / 1000000)
    ws1.range('O12').value = (opperating_expenses / 1000000)
    ws1.range('O13').value = gross_margin_percent
    #ws1.range('O14').value = diluted_earnings_per_share


    #CHARTS
    sht = ws1
    chart = sht.charts.add(50, 220)
    chart.set_source_data(sht.range("A7:E12"))
    chart.chart_type = 'bar_clustered'
    chart.api[1].SetElement(2)
    chart.api[1].ChartTitle.Text = tiker + ' Quarter by Quarter'
    chart.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart1 = sht.charts.add(50, 660)
    chart1.set_source_data(sht.range("A37:K37,A38:K38"))
    chart1.chart_type = 'column_clustered'
    chart1.api[1].SetElement(2)
    chart1.api[1].ChartTitle.Text = tiker + ' YOY Revenues'
    chart1.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart1.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart1 = sht.charts.add(450, 660)
    chart1.set_source_data(sht.range("A37:K37,A41:K41"))
    chart1.chart_type = 'column_clustered'
    chart1.api[1].SetElement(2)
    chart1.api[1].ChartTitle.Text = tiker + ' YOY Net Income'
    chart1.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart1.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart1 = sht.charts.add(850, 660)
    chart1.set_source_data(sht.range("A37:K37,A40:K40"))
    chart1.chart_type = 'column_clustered'
    chart1.api[1].SetElement(2)
    chart1.api[1].ChartTitle.Text = tiker + ' YOY Gross Margin'
    chart1.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart1.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart1 = sht.charts.add(1250, 660)
    chart1.set_source_data(sht.range("A37:K37,A39:K39"))
    chart1.chart_type = 'column_clustered'
    chart1.api[1].SetElement(2)
    chart1.api[1].ChartTitle.Text = tiker + ' YOY Cost of Revenues'
    chart1.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart1.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart1 = sht.charts.add(1650, 660)
    chart1.set_source_data(sht.range("A37:K37,A42:K42"))
    chart1.chart_type = 'column_clustered'
    chart1.api[1].SetElement(2)
    chart1.api[1].ChartTitle.Text = tiker + ' YOY Operating Expenses'
    chart1.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart1.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart2 = sht.charts.add(450, 220)
    chart2.set_source_data(sht.range("G7:J12"))
    chart2.chart_type = 'bar_clustered'
    chart2.api[1].SetElement(2)
    chart2.api[1].ChartTitle.Text = tiker + ' Year over Year (Q2)'
    chart2.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart2.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart2 = sht.charts.add(850, 220)
    chart2.set_source_data(sht.range("L7:O12"))
    chart2.chart_type = 'bar_clustered'
    chart2.api[1].SetElement(2)
    chart2.api[1].ChartTitle.Text = tiker + ' Year over Year (Q1)'
    chart2.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart2.api[1].Axes(2).AxisTitle.Text = "$ Millions"

elif quarter == 3:

    ws1.range('G8').value = 'Revenues               '
    ws1.range('G9').value = 'Cost of Revenues               '
    ws1.range('G10').value = 'Gross Margin              '
    ws1.range('G11').value = 'Net Income                '
    ws1.range('G12').value = 'Operating Expenses                '
    ws1.range('G13').value = 'Gross Margin Percent'
    ws1.range('G14').value = 'Diluted EPS'

    ws1.range('L8').value = 'Revenues               '
    ws1.range('L9').value = 'Cost of Revenues               '
    ws1.range('L10').value = 'Gross Margin              '
    ws1.range('L11').value = 'Net Income                '
    ws1.range('L12').value = 'Operating Expenses                '
    ws1.range('L13').value = 'Gross Margin Percent'
    ws1.range('L14').value = 'Diluted EPS'

    ws1.range('Q8').value = 'Revenues               '
    ws1.range('Q9').value = 'Cost of Revenues               '
    ws1.range('Q10').value = 'Gross Margin              '
    ws1.range('Q11').value = 'Net Income                '
    ws1.range('Q12').value = 'Operating Expenses                '
    ws1.range('Q13').value = 'Gross Margin Percent'
    ws1.range('Q14').value = 'Diluted EPS'

    joe = earningsData_3
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    print('Diluted EPS =', diluted_earnings_per_share)


    ws1.range('D6').value = end_date
    ws1.range('D7').value = fiscal_period + " " + fiscal_year
    ws1.range('D8').value =  (revenues / 1000000)
    ws1.range('D9').value =  (cost_of_revenues / 1000000)
    ws1.range('D10').value =  (gross_margin / 1000000)
    ws1.range('D11').value =  (net_income_loss / 1000000)
    ws1.range('D12').value =  (opperating_expenses / 1000000)
    ws1.range('D13').value = gross_margin_percent
    ws1.range('D14').value = diluted_earnings_per_share

    print('calc Q4 here')
    print('')
    print('total 1')
    joe = earningsDataTotal_1_Q3
    start_date_Total_1 = joe['start_date']
    end_date_Total_1 = joe['end_date']
    fiscal_period_Total_1 = joe['fiscal_period']
    fiscal_year_Total_1 = joe['fiscal_year']
    revenues_Total_1 = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues_Total_1 = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin_Total_1 = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_Total_1_percent = (gross_margin_Total_1 / revenues_Total_1) * 100
    opperating_expenses_Total_1 = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss_Total_1 = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share_Total_1 = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print(tiker, 'Earnings Data:')
    print('')
    print('Current quarter')
    print('Fiscal Period and Year =', fiscal_period_Total_1, fiscal_year_Total_1)
    print('Fiscal Period =', fiscal_period_Total_1)

    print('')
    print('total 2')
    joe = earningsDataTotal_2_Q3
    start_date_Total_2 = joe['start_date']
    end_date_Total_2 = joe['end_date']
    fiscal_period_Total_2 = joe['fiscal_period']
    fiscal_year_Total_2 = joe['fiscal_year']
    revenues_Total_2 = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues_Total_2 = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin_Total_2 = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_Total_1_percent = (gross_margin_Total_2 / revenues_Total_2) * 100
    opperating_expenses_Total_2 = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss_Total_2 = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share_Total_2 = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print(tiker, 'Earnings Data:')
    print('')
    print('Current quarter')
    print('Fiscal Period and Year =', fiscal_period_Total_2, fiscal_year_Total_2)
    print('Fiscal Period =', fiscal_period_Total_2)

    print('')
    print('total 3')
    joe = earningsDataTotal_3_Q3
    start_date_Total_3 = joe['start_date']
    end_date_Total_3 = joe['end_date']
    fiscal_period_Total_3 = joe['fiscal_period']
    fiscal_year_Total_3 = joe['fiscal_year']
    revenues_Total_3 = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues_Total_3 = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin_Total_3 = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_Total_3_percent = (gross_margin_Total_3 / revenues_Total_3) * 100
    opperating_expenses_Total_3 = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss_Total_3 = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share_Total_3 = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print(tiker, 'Earnings Data:')
    print('')
    print('Current quarter')
    print('Fiscal Period and Year =', fiscal_period_Total_3, fiscal_year_Total_3)
    print('Fiscal Period =', fiscal_period_Total_3)

    Q4revenue = revenues_Annual_1 - (revenues_Total_1 + revenues_Total_2 + revenues_Total_3)
    print('Q4 revenue =', Q4revenue)

    Q4cost_of_revenues = cost_of_revenues_Annual_1 - (
                cost_of_revenues_Total_1 + cost_of_revenues_Total_2 + cost_of_revenues_Total_3)
    print('Q4 cost_of_revenues =', Q4cost_of_revenues)

    Q4gross_margin = gross_margin_Annual_1 - (gross_margin_Total_1 + gross_margin_Total_2 + gross_margin_Total_3)
    print('Q4 gross_margin =', Q4gross_margin)

    Q1_Q3net_income_loss = (net_income_loss_Total_1 + net_income_loss_Total_2 + net_income_loss_Total_3)

    Q4net_income_loss = net_income_loss_Annual_1 - (
                net_income_loss_Total_1 + net_income_loss_Total_2 + net_income_loss_Total_3)

    print('Q3 net_income_loss =', net_income_loss_Total_1 / 1000000)
    print('Q2 net_income_loss =', net_income_loss_Total_2 / 1000000)
    print('Q1 net_income_loss =', net_income_loss_Total_3 / 1000000)
    print('Q1-Q3 net_income_loss =', Q1_Q3net_income_loss / 1000000)
    print('2020 net_income_loss =', net_income_loss_Annual_1 / 1000000)
    print('Q4 CALCULATED net_income_loss =', (net_income_loss_Annual_1 - Q1_Q3net_income_loss) / 1000000)
    print('Q4 net_income_loss =', Q4net_income_loss / 1000000)

    Q4opperating_expenses = opperating_expenses_Annual_1 - (
                opperating_expenses_Total_1 + opperating_expenses_Total_2 + opperating_expenses_Total_3)
    print('Q4 opperating_expenses =', Q4opperating_expenses)

    ws1.range('E6').value = end_date_Annual_1
    ws1.range('E7').value = "Q4" + " " + fiscal_year_Total_3
    ws1.range('E8').value = (Q4revenue / 1000000)
    ws1.range('E9').value = (Q4cost_of_revenues / 1000000)
    ws1.range('E10').value = (Q4gross_margin / 1000000)
    ws1.range('E11').value = (Q4net_income_loss / 1000000)
    ws1.range('E12').value = (Q4opperating_expenses / 1000000)
    # ws1.range('D13').value = gross_margin_percent
    # ws1.range('D14').value = diluted_earnings_per_share


    # yearly Q3_1
    ws1.range('G5').value = 'Year over Year (Q3)'
    joe = earningsDataYearByQuarter_1_1
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    print('Diluted EPS =', diluted_earnings_per_share)
    ws1.range('H6').value = end_date
    ws1.range('H7').value = fiscal_period + " " + fiscal_year
    ws1.range('H8').value = (revenues / 1000000)
    ws1.range('H9').value = (cost_of_revenues / 1000000)
    ws1.range('H10').value = (gross_margin / 1000000)
    ws1.range('H11').value = (net_income_loss / 1000000)
    ws1.range('H12').value = (opperating_expenses / 1000000)
    ws1.range('H13').value = gross_margin_percent
    ws1.range('H14').value = diluted_earnings_per_share

    # yearly Q3_2
    joe = earningsDataYearByQuarter_1_2
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    print('Diluted EPS =', diluted_earnings_per_share)
    ws1.range('I6').value = end_date
    ws1.range('I7').value = fiscal_period + " " + fiscal_year
    ws1.range('I8').value = (revenues / 1000000)
    ws1.range('I9').value = (cost_of_revenues / 1000000)
    ws1.range('I10').value = (gross_margin / 1000000)
    ws1.range('I11').value = (net_income_loss / 1000000)
    ws1.range('I12').value = (opperating_expenses / 1000000)
    ws1.range('I13').value = gross_margin_percent
    ws1.range('I14').value = diluted_earnings_per_share

    # yearly Q3_3
    joe = earningsDataYearByQuarter_1_3
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    print('Diluted EPS =', diluted_earnings_per_share)

    ws1.range('J6').value = end_date
    ws1.range('J7').value = fiscal_period + " " + fiscal_year
    ws1.range('J8').value = (revenues / 1000000)
    ws1.range('J9').value = (cost_of_revenues / 1000000)
    ws1.range('J10').value = (gross_margin / 1000000)
    ws1.range('J11').value = (net_income_loss / 1000000)
    ws1.range('J12').value = (opperating_expenses / 1000000)
    ws1.range('J13').value = gross_margin_percent
    ws1.range('J14').value = diluted_earnings_per_share

    # yearly Q2_1
    ws1.range('L5').value = 'Year over Year (Q2)'
    joe = earningsDataYearByQuarter_2_1
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    print('Diluted EPS =', diluted_earnings_per_share)
    ws1.range('M6').value = end_date
    ws1.range('M7').value = fiscal_period + " " + fiscal_year
    ws1.range('M8').value = (revenues / 1000000)
    ws1.range('M9').value = (cost_of_revenues / 1000000)
    ws1.range('M10').value = (gross_margin / 1000000)
    ws1.range('M11').value = (net_income_loss / 1000000)
    ws1.range('M12').value = (opperating_expenses / 1000000)
    ws1.range('M13').value = gross_margin_percent
    ws1.range('M14').value = diluted_earnings_per_share

    # yearly Q2_2
    joe = earningsDataYearByQuarter_2_2
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    print('Diluted EPS =', diluted_earnings_per_share)
    ws1.range('N6').value = end_date
    ws1.range('N7').value = fiscal_period + " " + fiscal_year
    ws1.range('N8').value = (revenues / 1000000)
    ws1.range('N9').value = (cost_of_revenues / 1000000)
    ws1.range('N10').value = (gross_margin / 1000000)
    ws1.range('N11').value = (net_income_loss / 1000000)
    ws1.range('N12').value = (opperating_expenses / 1000000)
    ws1.range('N13').value = gross_margin_percent
    ws1.range('N14').value = diluted_earnings_per_share

    # yearly Q2_3
    joe = earningsDataYearByQuarter_2_3
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    print('Diluted EPS =', diluted_earnings_per_share)

    ws1.range('O6').value = end_date
    ws1.range('O7').value = fiscal_period + " " + fiscal_year
    ws1.range('O8').value = (revenues / 1000000)
    ws1.range('O9').value = (cost_of_revenues / 1000000)
    ws1.range('O10').value = (gross_margin / 1000000)
    ws1.range('O11').value = (net_income_loss / 1000000)
    ws1.range('O12').value = (opperating_expenses / 1000000)
    ws1.range('O13').value = gross_margin_percent
    ws1.range('O14').value = diluted_earnings_per_share




    # yearly Q1_1
    ws1.range('Q5').value = 'Year over Year (Q1)'
    joe = earningsDataYearByQuarter_3_1
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    print('Diluted EPS =', diluted_earnings_per_share)
    ws1.range('R6').value = end_date
    ws1.range('R7').value = fiscal_period + " " + fiscal_year
    ws1.range('R8').value = (revenues / 1000000)
    ws1.range('R9').value = (cost_of_revenues / 1000000)
    ws1.range('R10').value = (gross_margin / 1000000)
    ws1.range('R11').value = (net_income_loss / 1000000)
    ws1.range('R12').value = (opperating_expenses / 1000000)
    ws1.range('R13').value = gross_margin_percent
    ws1.range('R14').value = diluted_earnings_per_share

    # yearly Q1_2
    joe = earningsDataYearByQuarter_3_2
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    print('Diluted EPS =', diluted_earnings_per_share)
    ws1.range('S6').value = end_date
    ws1.range('S7').value = fiscal_period + " " + fiscal_year
    ws1.range('S8').value = (revenues / 1000000)
    ws1.range('S9').value = (cost_of_revenues / 1000000)
    ws1.range('S10').value = (gross_margin / 1000000)
    ws1.range('S11').value = (net_income_loss / 1000000)
    ws1.range('S12').value = (opperating_expenses / 1000000)
    ws1.range('S13').value = gross_margin_percent
    ws1.range('S14').value = diluted_earnings_per_share

    # yearly Q1_3
    joe = earningsDataYearByQuarter_3_3
    start_date = joe['start_date']
    end_date = joe['end_date']
    fiscal_period = joe['fiscal_period']
    fiscal_year = joe['fiscal_year']
    revenues = joe['financials']['income_statement']['revenues']['value']
    cost_of_revenues = joe['financials']['income_statement']['cost_of_revenue']['value']
    gross_margin = joe['financials']['income_statement']['gross_profit']['value']
    gross_margin_percent = (gross_margin / revenues) * 100
    opperating_expenses = joe['financials']['income_statement']['operating_expenses']['value']
    net_income_loss = joe['financials']['income_statement']['net_income_loss']['value']
    diluted_earnings_per_share = joe['financials']['income_statement']['diluted_earnings_per_share']['value']
    print('')
    print('Previous Quarter')
    print('Fiscal Period and Year =', fiscal_period, fiscal_year)
    print('Time Period:', start_date, '-', end_date)
    print('Revenues =', revenues)
    print('Cost of Revenues =', cost_of_revenues)
    print('Gross Margin =', gross_margin)
    print('Gross Margin Mercent =', gross_margin_percent, '%')
    print('Opperating Expenses =', opperating_expenses)
    print('Net Income =', net_income_loss)
    print('Diluted EPS =', diluted_earnings_per_share)

    ws1.range('T6').value = end_date
    ws1.range('T7').value = fiscal_period + " " + fiscal_year
    ws1.range('T8').value = (revenues / 1000000)
    ws1.range('T9').value = (cost_of_revenues / 1000000)
    ws1.range('T10').value = (gross_margin / 1000000)
    ws1.range('T11').value = (net_income_loss / 1000000)
    ws1.range('T12').value = (opperating_expenses / 1000000)
    ws1.range('T13').value = gross_margin_percent
    ws1.range('T14').value = diluted_earnings_per_share

    # CHARTS
    sht = ws1
    chart = sht.charts.add(50, 220)
    chart.set_source_data(sht.range("A7:E12"))
    chart.chart_type = 'bar_clustered'
    chart.api[1].SetElement(2)
    chart.api[1].ChartTitle.Text = tiker + ' Quarter by Quarter'
    chart.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart1 = sht.charts.add(50, 660)
    chart1.set_source_data(sht.range("A37:K37,A38:K38"))
    chart1.chart_type = 'column_clustered'
    chart1.api[1].SetElement(2)
    chart1.api[1].ChartTitle.Text = tiker + ' YOY Revenues'
    chart1.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart1.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart1 = sht.charts.add(450, 660)
    chart1.set_source_data(sht.range("A37:K37,A41:K41"))
    chart1.chart_type = 'column_clustered'
    chart1.api[1].SetElement(2)
    chart1.api[1].ChartTitle.Text = tiker + ' YOY Net Income'
    chart1.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart1.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart1 = sht.charts.add(850, 660)
    chart1.set_source_data(sht.range("A37:K37,A40:K40"))
    chart1.chart_type = 'column_clustered'
    chart1.api[1].SetElement(2)
    chart1.api[1].ChartTitle.Text = tiker + ' YOY Gross Margin'
    chart1.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart1.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart1 = sht.charts.add(1250, 660)
    chart1.set_source_data(sht.range("A37:K37,A39:K39"))
    chart1.chart_type = 'column_clustered'
    chart1.api[1].SetElement(2)
    chart1.api[1].ChartTitle.Text = tiker + ' YOY Cost of Revenues'
    chart1.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart1.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart1 = sht.charts.add(1650, 660)
    chart1.set_source_data(sht.range("A37:K37,A42:K42"))
    chart1.chart_type = 'column_clustered'
    chart1.api[1].SetElement(2)
    chart1.api[1].ChartTitle.Text = tiker + ' YOY Operating Expenses'
    chart1.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart1.api[1].Axes(2).AxisTitle.Text = "$ Millions"





    chart2 = sht.charts.add(450, 220)
    chart2.set_source_data(sht.range("G7:J12"))
    chart2.chart_type = 'bar_clustered'
    chart2.api[1].SetElement(2)
    chart2.api[1].ChartTitle.Text = tiker + ' Year over Year (Q3)'
    chart2.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart2.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart2 = sht.charts.add(850, 220)
    chart2.set_source_data(sht.range("L7:O12"))
    chart2.chart_type = 'bar_clustered'
    chart2.api[1].SetElement(2)
    chart2.api[1].ChartTitle.Text = tiker + ' Year over Year (Q2)'
    chart2.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart2.api[1].Axes(2).AxisTitle.Text = "$ Millions"

    chart2 = sht.charts.add(1250, 220)
    chart2.set_source_data(sht.range("Q7:T12"))
    chart2.chart_type = 'bar_clustered'
    chart2.api[1].SetElement(2)
    chart2.api[1].ChartTitle.Text = tiker + ' Year over Year (Q1)'
    chart2.api[1].Axes(2).HasTitle = True  # This line creates the Y axis label.
    chart2.api[1].Axes(2).AxisTitle.Text = "$ Millions"

Range('A1').api.Font.Bold = True
Range('B1').api.Font.Bold = True
Range('B35').api.Font.Bold = True

if quarter == 2:
    Range('B5').api.Font.Bold = True
    Range('G5').api.Font.Bold = True
    Range('L5').api.Font.Bold = True
if quarter == 3:
    Range('B5').api.Font.Bold = True
    Range('G5').api.Font.Bold = True
    Range('L5').api.Font.Bold = True
    Range('Q5').api.Font.Bold = True
Range('A1:T44').tautofit()