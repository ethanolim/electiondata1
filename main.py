import webbrowser

import bs4
import ezsheets as ezsheets
import openpyxl as openpyxl
from pip._vendor import requests

openpyxl
bs4
webbrowser
requests
ezsheets

# list of the things that get copied into url to get data from NYT

stateKeys = ["alabama","alaska","arizona","arkansas","california","colorado",
"connecticut","delaware","district-of-columbia","florida","georgia","hawaii","idaho","illinois",
"indiana","iowa","kansas","kentucky","louisiana","maine","maryland",
"massachusetts","michigan","minnesota","mississippi","missouri","montana",
"nebraska","nevada","new-hampshire","new-jersey","new-mexico","new-york",
"north-carolina","north-dakota","ohio","oklahoma","oregon","pennsylvania",
"rhode-island","south-carolina","south-dakota","tennessee","texas","utah",
"vermont","virginia","washington","west-virginia","wisconsin","wyoming"]

percentStatesDone = 0.0

wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = 'Data by State'
wb.save('Python 2016_2020 Data Comparison Project1.xlsx')

sheet['A1'] = 'State'
sheet['B1'] = '2016 Trump Percentage'
sheet['C1'] = '2016 Clinton Percentage'
sheet['D1'] = '2016 Split D'

sheet['F1'] = '2020 Trump Percentage'
sheet['G1'] = '2020 Biden Percentage'
sheet['H1'] = '2020 Split D'

sheet['J1'] = '2016-2020 Shift D'

row = 1

# loops through all state keys to pull 2016 data

for i in stateKeys:

    row = row+1
    res = requests.get('https://www.nytimes.com/elections/2016/results/' + i)
    soup = bs4.BeautifulSoup(res.text, 'html.parser')
    winnerRaw = str(soup.select('a.eln-item:nth-child(1) > span:nth-child(2) > span:nth-child(2)'))

    if 'district' in i: # DC needs a different slice because it's the only 'state' with Trump in single digits

        winner = 'Clinton'
        trumpPercent = str(soup.select('tr.eln-trump-8639:nth-child(2) > td:nth-child(5)'))[56:59]
        clintonPercent = str(soup.select('tr.eln-clinton-1746:nth-child(1) > td:nth-child(5)'))[56:60]
        split = round(float(clintonPercent) - float(trumpPercent), 1)

    elif ('Clinton' in winnerRaw) and ('district' not in i):

        winner = 'Clinton'
        trumpPercent = str(soup.select('tr.eln-trump-8639:nth-child(2) > td:nth-child(5)'))[56:60]
        clintonPercent = str(soup.select('tr.eln-clinton-1746:nth-child(1) > td:nth-child(5)'))[56:60]
        split = round(float(clintonPercent) - float(trumpPercent), 1)

    else:

        winner = 'Trump'
        trumpPercent = str(soup.select('div.eln-results-container:nth-child(1) > table:nth-child(1) > tbody:nth-child(2) > tr:nth-child(1) > td:nth-child(5)'))[56:60]
        clintonPercent = str(soup.select('tr.eln-clinton-1746:nth-child(2) > td:nth-child(5)'))[56:60]
        split = round(float(clintonPercent) - float(trumpPercent), 1)

    # writes results to spreadsheet

    sheet['A' + str(row)] = i
    sheet['B' + str(row)] = trumpPercent
    sheet['C' + str(row)] = clintonPercent
    sheet['D' + str(row)] = split
    percentStatesDone = percentStatesDone + 0.980392
    print('Pulling state data: ' + str(round(percentStatesDone, 2)) + '%')

row = 1

# loops through all state keys to pull 2020 data

for i in stateKeys:

    row = row+1
    res = requests.get('https://www.nytimes.com/interactive/2020/11/03/us/elections/results-' + i + '.html')
    soup = bs4.BeautifulSoup(res.text, 'html.parser')
    winnerRaw = str(soup.select('.e-card-inner > h2:nth-child(4)'))

    if 'district' in i: # same shit with DC

        winner = 'Biden'
        trumpPercent = str(soup.select('.e-donald-j-trump > td:nth-child(5) > span:nth-child(1) > span:nth-child(1)'))[29:32]
        bidenPercent = str(soup.select('.e-joseph-r-biden > td:nth-child(5) > span:nth-child(1) > span:nth-child(1)'))[29:32]
        split = round(float(bidenPercent) - float(trumpPercent), 1)

    elif ('Biden' in winnerRaw) and ('district' not in i):

        winner = 'Biden'
        trumpPercent = str(soup.select('.e-donald-j-trump > td:nth-child(5) > span:nth-child(1) > span:nth-child(1)'))[29:33]
        bidenPercent = str(soup.select('.e-joseph-r-biden > td:nth-child(5) > span:nth-child(1) > span:nth-child(1)'))[29:33]
        split = round(float(bidenPercent) - float(trumpPercent), 1)

    else:

        winner = 'Trump'
        trumpPercent = str(soup.select('tr.e-trump-8639:nth-child(1) > td:nth-child(5) > span:nth-child(1) > span:nth-child(1)'))[29:33]
        bidenPercent = str(soup.select('tr.e-biden-1036:nth-child(2) > td:nth-child(5) > span:nth-child(1) > span:nth-child(1)'))[29:33]
        split = round(float(bidenPercent) - float(trumpPercent), 1)

    # writes to spreadsheet

    sheet['F' + str(row)] = trumpPercent
    sheet['G' + str(row)] = bidenPercent
    sheet['H' + str(row)] = split
    percentStatesDone = percentStatesDone + 0.980392
    print('Pulling state data: ' + str(round(percentStatesDone, 2)) + '%')

row = 1

# loops through the 2016 and 2020 splits to calculate the swing, writes it to a new column

for i in (stateKeys):

    row = row+1
    swing = round((sheet['H' + str(row)].value - sheet['D' + str(row)].value), 1)
    sheet['J' + str(row)] = swing

print('Done with state data.')

wb.create_sheet(title='Data by Congressional District')

sheet = wb['Data by Congressional District']

sheet['A1'] = 'Congressional District'
sheet['B1'] = '2016 Trump Percentage'
sheet['C1'] = '2016 Clinton Percentage'
sheet['D1'] = '2016 Split D'

sheet['F1'] = '2020 Trump Percentage'
sheet['G1'] = '2020 Biden Percentage'
sheet['H1'] = '2020 Split D'

sheet['J1'] = '2016-2020 Shift D'

print('Pulling congressional district data...')

ss = ezsheets.Spreadsheet('https://docs.google.com/spreadsheets/d/1XbUXnI9OyfAuhP5P3vWtMuGc5UJlrhXbzZo3AwMuHtk/edit#gid=0')
cdsheet = ss[0]

for x in range (1, 436):

    dName = cdsheet['A' + str(x+2)]
    sheet['A' + str(x+1)] = dName

    dClinton16 = float(cdsheet['F' + str(x+2)])
    sheet['C' + str(x + 1)] = dClinton16

    dTrump16 = float(cdsheet['G' + str(x+2)])
    sheet['B' + str(x + 1)] = dTrump16

    sheet['D' + str(x+1)] = round(dClinton16 - dTrump16, 1)

    if cdsheet['D' + str(x + 2)] == '':

        sheet['G' + str(x + 1)] = ''
        sheet['F' + str(x + 1)] = ''
        sheet['H' + str(x + 1)] = ''
        sheet['J' + str(x + 1)] = ''

    else:

        dBiden20 = float(cdsheet['D' + str(x+2)])
        sheet['G' + str(x + 1)] = dBiden20

        dTrump20 = float(cdsheet['E' + str(x+2)])
        sheet['F' + str(x + 1)] = dTrump20

        sheet['H' + str(x + 1)] = round(dBiden20 - dTrump20, 1)

        sheet['J' + str(x+1)] = (sheet['H' + str(x + 1)].value - sheet['D' + str(x+1)].value)

wb.save('Python 2016_2020 Data Comparison Project1.xlsx')

print('Done!')