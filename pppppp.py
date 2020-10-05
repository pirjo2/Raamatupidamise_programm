import PySimpleGUI as sg
sg.theme('LightBlue6')  # akna värvilahenduse muutmine

def arvuta(bruto, sots, pens, tooandja, tootaja, tulumaks, brutoo, neto, maksud, tooandjamaks, tulum, tmv, aasta):
    #brutoo = int(bruto)
    bruto = float(bruto)
    if sots == True:
        tooandjamaks += bruto * 0.33
        maksud += bruto * 0.33
    if pens == True:
        neto = neto - (bruto * 0.02)
        maksud += bruto * 0.02
    if tooandja == True:
        tooandjamaks += bruto * 0.008
        maksud += bruto * 0.008
    if tootaja == True:
        neto = neto - (bruto * 0.016)
        maksud += bruto * 0.016
    if tulumaks == True:
        aasta = bruto * 12
        if aasta <= 6000:
            tmv = 500
        if aasta > 6000 and aasta <= 14400:
            tmv = 500
        if aasta > 14400 and aasta <= 25200:
            tmv = 6000 - 6000 / 10800 * (aasta - 14400)
        if aasta > 25200:
            tmv = 0
        if aasta >= 6000:
            tulum = 0
        if aasta > 6000 and aasta <= 14400:
            tulum = (aasta - tmv - 6000) * 0.2
        if aasta > 14400 and aasta <= 25200:
            tulum = (aasta - 6000) * 0.2 + (aasta - tmv - 14400) * 0.311111
        if aasta > 25200:
            tulum = (aasta - 6000) * 0.2 + (25200-14400) * 0.311111 + (aasta - 25200) * 0.2
        tulum = round(tulum / 12, 2)
        neto = bruto - neto - tulum
        maksud += tulum
    return [neto, maksud, tooandjamaks, tulum, tmv]

brutoo = 0.0
neto = 0.0
maksud = 0.0
tooandjamaks = 0.0
tulum = 0.0
tmv = 0.0
aasta = 0.0

layout = [
    [sg.Text('Palgakalkulaator'), sg.Text(size=(16,1), key='tekstisilt')],
    [sg.Text('Vali maksud: '), sg.Text(size=(12,1), key='tekstisilt')],
    [sg.Checkbox('Sotsiaalmaks', default=True, key = 'sotsmaks'), sg.Checkbox('Kogumispension', default=True, key ='pens')],
    [sg.Text('Töötuskindlustusmaksed:'), sg.Text(size=(12,1), key='tekstisilt')],
    [sg.Checkbox('Tööandja', default=True, key = 'tooandja'), sg.Checkbox('Töötaja', default=True, key = 'tootaja')],
    [sg.Checkbox('Astmeline tulumaks', default=True, key = 'tulumaks')],
    [sg.Text('Sisesta bruto palk: '), sg.Text(size=(16,1), key='tekstisilt'),
    sg.InputText('EUR', size = (9,1), do_not_clear = True, key = 'bruto')],
    [sg.Button('Kalkuleeri', key = 'button'), sg.Exit('Välju')]
    ]

window = sg.Window('Palgakalkulaator', layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Välju':
        break
    if event == 'button':
        list = arvuta(values['bruto'], values['sotsmaks'], values['pens'], values['tooandja'], values['tootaja'], values['tulumaks'], brutoo, neto, maksud, tooandjamaks, tulum, tmv, aasta)
        print(list)
    print(event, values)

window.close()