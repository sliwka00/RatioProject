import PySimpleGUI as sg
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import subprocess
import sys
lista=[]

df = pd.read_excel(r'abc.xlsx')


df = df.replace('-',np.nan)   #zamienia  "-" na Nan w komórkach gdzie nie ma ceny
df = df.astype({'DKR':float})  #zamienia kolumne DKR na floaty (dane były jako string)
df['kontrakt short'] = df['Kontrakt'].str.split("_").str[-1]       #Skraca nazwe kontraktu do uniwersalnego (dla base i peak) żeby je sparować
df['Data']=pd.to_datetime(df['Data'], format='%d-%m-%Y')
df['wolumen'] = [float(str(val).replace(u'\xa0','').replace(',','.')) for val in df['wolumen'].values]   #wyrzucenie dziwnych znaków z wolumenu i zamiana na float
df3 = df[['Data','DKR','typ','wolumen','kontrakt short']]  #stworzenie skróconego df bez zbędnych kolumn
df_base = df3[df3['typ'] == 'BASE']     #stworzenie df dla base
df_peak = df3[df3['typ'] == 'PEAK']
df_wsp = pd.merge(df_base,df_peak, on=['Data','kontrakt short'])  #połączenie df_base i df_peak dzieki temu można dodać kolumne ratio
df_wsp['ratio']=df_wsp['DKR_y']/df_wsp['DKR_x']  #kolumna z ratio

# Pętla do uzupełniania listy produktów, które znajdują sie w pliku zródłowym
for produkt in df['kontrakt short']:
    if produkt not in lista and "W-" not in produkt:
        lista.append(produkt)
    else:
        continue
lista.sort()
print(lista)

def draw_ratio2(produkt):    # wyświetla ratio + 2 słupki wolumenowe base i peak
    df_temp=df_wsp[df_wsp['kontrakt short']==produkt]
    data=df_temp['Data']
    ratio=df_temp['ratio']
    wol_peak=df_temp['wolumen_y']
    wol_base=df_temp['wolumen_x']
    # Tworzenie figury i osi
    fig, ax1 = plt.subplots()
    ax1.bar(data, wol_peak, color='red', alpha=0.5)
    ax1.set_ylabel('Wolumen peak->czerwony \n wolumen base ->zielony ')

    ax3=ax1.twinx()
    ax3.bar(data, wol_base, color='green', alpha=0.5)
    ax3.axes.get_yaxis().set_visible(False)
    ax3.set_ylabel('Wolumen base')

    ax2 = ax1.twinx()
    ax2.plot(data, ratio, marker='o', linestyle='-', color='blue')
    ax2.set_title(produkt)
    ax2.set_xlabel('Data')
    ax2.set_ylabel('Ratio Peak/Base')
    fig.autofmt_xdate(rotation=35, ha='right')    #rotuje daty wyświetlane pod wykresem
    plt.show()

#Layout Okna GUI
sg.theme("Black") #gotowe motywy z kolorystyka do podejrzenia w internecie

label=sg.Text("Wykres Ratio")

all_label=sg.Text("Lista Produktów")
all_combo=sg.Combo(lista, font=('Arial Bold', 14),  expand_x=True, enable_events=True, key='all_droplist')
download_button=sg.Button(button_text="zaciągnij brakujące dane z TGE",key="download")

window=sg.Window("Wykresy",
                 layout=[[label],
                        [all_label,all_combo],
                        [download_button]])


while True:
    event,values=window.read()
    print(f'events: {event}')
    print(f'values: {values}')
    match event:
        case sg.WIN_CLOSED:  # co się stanie po zamknięciu okna gui
            break
        case 'all_droplist':
            draw_ratio2(values['all_droplist'])
        case'download':
            subprocess.Popen([sys.executable,"main.py"])   #uruchamia skrypt do ściągniecią danych ze strony po kliknięciu w przycisk
            subprocess.Popen([sys.executable,"ratio.py"])     #uruchamiam GUI jeszcze raz,żeby zaciągniete nowe dane do excela były już dostępne
window.close()

