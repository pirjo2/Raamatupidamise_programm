import PySimpleGUI as sg
import xlwt
import xlsxwriter
import openpyxl
from xlwt import Workbook
sg.theme('LightBlue6')  # akna värvilahenduse muutmine

def arvuta(bruto, sots, pens, tooandja, tootaja, tulumaks, toolisemaksud, neto, maksud, tooandjamaks, tulum, tmv, aasta, pens2, tootaja2, sots2, tooandja2):
    bruto = float(bruto)
    aasta = bruto * 12
    if sots == True:    # sotsiaalmaksu arvutamine
        if bruto * 0.33 > 178.2:
            tooandjamaks += bruto * 0.33
            maksud += bruto * 0.33
            sots2 = bruto * 0.33
        else:    # sotsiaalmaksu minimaalne kohustus on 178.2 eurot kuus
            tooandjamaks += 178.2
            maksud += 178.2
            sots2 = 178.2
    if pens == True:    # pensioni arvutamine
        neto += bruto * 0.02
        maksud += bruto * 0.02
        toolisemaksud += bruto * 0.02
        pens2 = bruto * 0.02
    if tooandja == True:    # töötuskindlustusmakse arvutamine, tööandja
        tooandjamaks += bruto * 0.008
        maksud += bruto * 0.008
        tooandja2 = bruto * 0.008
    if tootaja == True:    # töötuskindlustusmakse arvutamine, töötaja
        neto += bruto * 0.016
        maksud += bruto * 0.016
        toolisemaksud += bruto * 0.016
        tootaja2 = bruto * 0.016
    if tulumaks == True:    # tulumaksuvaba summa arvutamine
        if aasta <= 6000:
            tmv = 500
        if aasta > 6000 and aasta <= 14400:
            tmv = 500
        if aasta > 14400 and aasta <= 25200:
            tmv = 6000 - 6000 / 10800 * (aasta - 14400)
        if aasta > 25200:
            tmv = 0
    if aasta >= 6000:    # tulumaksu arvutamine
        tulum = 0
    if aasta > 6000 and aasta <= 14400:
        tulum = (aasta - tmv - toolisemaksud * 12 - 6000) * 0.2
    if aasta > 14400 and aasta <= 25200:
        tulum = (aasta - tmv - toolisemaksud * 12) * 0.2
    if aasta > 25200:
        tulum = (aasta - toolisemaksud) * 0.2
    tulum = round(tulum / 12, 2)    # ümardamine ja lisaarvutused
    neto = round(bruto - neto - tulum, 2)
    maksud += tulum
    tooandjamaks += bruto
    maksud = round(maksud, 2)
    tooandjamaks = round(tooandjamaks, 2)
    tulum = round(tulum, 2)
    tmv = round(tmv, 2)
    toolisemaksud = round(toolisemaksud, 2)
    return [neto, maksud, tooandjamaks, tulum, tmv, toolisemaksud, pens2, tootaja2, sots2, tooandja2]

toolisemaksud = 0.0    # maksud, mis lähevad töölise bruto palgast maha
neto = 0.0    # bruto - maksud = neto, selle saab puhtalt kätte
maksud = 0.0    # kõik maksud kokku
tooandjamaks = 0.0    # tööandja kohustus maksta kõik kokku
tulum = 0.0    # tulumaks
tmv = 0.0    # tulumaksuvaba
aasta = 0.0    # aastatulu
pens2 = 0.0    # pension
tootaja2 = 0.0    # töötaja töötuskindlustusmakse
sots2 = 0.0    #sotsiaalmakse
tooandja2 = 0.0    # tööandja töötuskindlustusmakse


neto3 = 0.0    #Excelisse kirjutamiseks
pension3 = 0.0    #Excelisse kirjutamiseks
tootuskindlustustootaja3 = 0.0    #Excelisse kirjutamiseks
tulumaks3 = 0.0    #Excelisse kirjutamiseks
tootuskindlustustooandja3 = 0.0    #Excelisse kirjutamiseks
sots3 = 0.0    #Excelisse kirjutamiseks
kokku3 = 0.0    #Excelisse kirjutamiseks
lehenimi =""
aa = ""
# file = open
layout = [
    [sg.Text('PALGAKALKULAATOR'), sg.Text(size=(16,1), key = 'tekstisilt1')],
    [sg.Text('Sisestage töötaja nimi: '), sg.Text(size=(12,1), key = 'nimesilt1'),
    sg.InputText('Eesnimi', size = (12,1), do_not_clear = True, key = 'nimelahter1'),
    sg.InputText('Perenimi', size = (12,1), do_not_clear = True, key = 'nimelahter2')],
    [sg.Text('Sisestage tööandja nimi: '), sg.Text(size=(11,1), key = 'nimesilt2'),
    sg.InputText('Eesnimi', size = (12,1), do_not_clear = True, key = 'nimelahter3'),
    sg.InputText('Perenimi', size = (12,1), do_not_clear = True, key = 'nimelahter4')],
    [sg.Text('Sisestage kuu ja aasta: '), sg.Text(size=(12,1), key = 'kuupaevasilt1'),
    sg.InputText('Kuu', size = (12,1), do_not_clear = True, key = 'kuupaevalahter1'),
    sg.InputText('Aasta', size = (12,1), do_not_clear = True, key = 'kuupaevalahter2')],
    [sg.Text('Vali maksud: '), sg.Text(size=(12,1), key = 'tekstisilt2')],
    [sg.Checkbox('Maksuvaba tulu', default = True, key = 'tulumaks')],
    [sg.Checkbox('Sotsiaalmaks', default = True, key = 'sotsmaks'), sg.Checkbox('Kogumispension', default = True, key = 'pens')],
    [sg.Text('Töötuskindlustusmaksed:'), sg.Text(size=(12,1), key = 'tekstisilt3')],
    [sg.Checkbox('Tööandja', default = True, key = 'tooandja'), sg.Checkbox('Töötaja', default = True, key = 'tootaja')],
    [sg.Text('Sisestage bruto palk: '), sg.Text(size=(16,1), key = 'tekstisilt4'),
    sg.InputText('EUR', size = (9,1), do_not_clear = True, key = 'bruto')],
    [sg.Button('Kalkuleeri', key = 'button')],
    [sg.Text('TÖÖTAJA:'), sg.Text(size=(12,1))],
    [sg.Text('Palk netos: '), sg.Text(size=(12,1), key = 'neto2')],
    [sg.Text('Kogumispension: '), sg.Text(size=(12,1), key = 'pension')],
    [sg.Text('Töötuskindlustusmaks (töötaja): '), sg.Text(size=(12,1), key = 'tootuskindlustus2')],
    [sg.Text('Tulumaks: '), sg.Text(size=(12,1), key = 'tulumaks2')],
    [sg.Text('TÖÖANDJA:'), sg.Text(size=(12,1))],
    [sg.Text('Tööandja kulud kokku: '), sg.Text(size=(12,1), key = 'kokku2')],
    [sg.Text('Sotsiaalmaks: '), sg.Text(size=(12,1), key = 'sots2')],
    [sg.Text('Töötuskindlustusmaks (tööandja): '), sg.Text(size=(12,1), key = 'tooandja3')],
    [sg.Button('Lisa andmed faili', key = 'button1'),
    sg.Exit('Välju')]
    ]

window = sg.Window('Raamatupidamine 2000', layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Välju':
        break
    if event == 'button':    # kui klikatakse "Kalkuleeri nuppu
        list = arvuta(values['bruto'], values['sotsmaks'], values['pens'], values['tooandja'], values['tootaja'], values['tulumaks'], toolisemaksud, neto, maksud, tooandjamaks, tulum, tmv, aasta, pens2, tootaja2, sots2, tooandja2)
        #print(list)
        window['neto2'].update(str(list[0]))    # teksti uuendamine peale kalkuleerimist
        neto3 = list[0]
        window['pension'].update(str(list[6]))
        pension3 = list[6]
        window['tootuskindlustus2'].update(str(list[7]))
        tootuskindlustustootaja3 = list[7]
        window['tulumaks2'].update(str(list[3]))
        tulumaks3 = list[3]
        window['kokku2'].update(str(list[2]))
        kokku3 = list[2]
        window['sots2'].update(str(list[8]))
        sots3 = list[8]
        window['tooandja3'].update(str(list[9]))
        tootuskindlustustooandja3 = list[9]
        list.clear()
    #print(event, values)
    if event == 'button1':
        lehenimi = str(values['kuupaevalahter1']) + "." + str(values['kuupaevalahter2'])
        aa = "Andmed_Excelis.xlsx"
        try:
        
        #if open(aa) == True:
            
            #on vaja avada exceli fail kui on olemas ja kui ei ole, siis luua ja avada.
        
            #hea oleks saada yhele reale niimodi, et sheeti pikka teksti ei tuleks 2 korda.
            #wb = Workbook()
        
            workbook = openpyxl.load_workbook('Andmed_Excelis.xlsx')
            #lehenimi = str(values['kuupaevalahter1']) + "." + str(values['kuupaevalahter2'])
        except:
            workbook = xlsxwriter.Workbook("Andmed_Excelis.xlsx")
            #worksheet = workbook.add_worksheet(lehenimi)
            #file = open("Andmed_Excelis.xlsx", "x")
            
        
            workbook.close()
            workbook = openpyxl.load_workbook('Andmed_Excelis.xlsx')
            #lehenimi = str(values['kuupaevalahter1']) + "." + str(values['kuupaevalahter2'])
        if not lehenimi in workbook.sheetnames:
            #book.create_sheet(lehenimi)

            #sheet1 = wb.add_sheet('Andmed Excelis')
            workbook = xlsxwriter.Workbook("Andmed_Excelis.xlsx")
            worksheet = workbook.add_worksheet(lehenimi)
            
            worksheet.write(1, 0, 'Töötaja')
            worksheet.write(2, 0, 'Tööandja') 
            worksheet.write(0, 0, 'Töötaja/Tööandja') 
            worksheet.write(0, 1, 'Eesnimi') 
            worksheet.write(0, 2, 'Perekonnanimi') 
            worksheet.write(0, 3, 'Bruto Palk') 
            worksheet.write(0, 4, 'Neto Palk')
            worksheet.write(0, 5, 'Kogumispension') 
            worksheet.write(0, 6, 'Töötuskindlustusmaks') 
            worksheet.write(0, 7, 'Tulumaks') 
            worksheet.write(0, 8, 'Sotsiaalmaks')
            worksheet.write(0, 9, 'Tööandja kulud kokku')
            worksheet.write(1, 1, values['nimelahter1'])
            worksheet.write(1, 2, values['nimelahter2']) 
            worksheet.write(1, 3, values['bruto']) 
            worksheet.write(1, 4, neto3) 
            worksheet.write(1, 5, pension3) 
            worksheet.write(1, 6, tootuskindlustustootaja3) 
            worksheet.write(1, 7, tulumaks3)
            worksheet.write(1, 8, '-') 
            worksheet.write(1, 9, '-')  
            worksheet.write(2, 1, values['nimelahter3'])
            worksheet.write(2, 2, values['nimelahter4'])
            worksheet.write(2, 3, '-')
            worksheet.write(2, 4, '-') 
            worksheet.write(2, 5, '-') 
            worksheet.write(2, 6, tootuskindlustustooandja3) 
            worksheet.write(2, 7, '-') 
            worksheet.write(2, 8, sots3) 
            worksheet.write(2, 9, kokku3)
    
            workbook.close()
        #book.save('Andmed_Excelis.xlsx')
window.close()