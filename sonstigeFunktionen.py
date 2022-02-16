# sonstigeFunktionen.py
# Funktionen für das Kreisdiagramm und den Dokumentepfad

# Importe
import os
import os.path
import webbrowser as wb
import plotly.express as px
import plotly.graph_objects as go
import dash_bootstrap_components as dbc
import pandas as pd
import DBconnection
import pyodbc
from logger import write_log

# Variablen
data = [['Monatswechsel', 10], ['AbfrageMonat', 10], ['AnzeigeMonat', 10], ['Schaetzung', 10], [
    'Abrechnung', 10], ['MLJ', 10], ['PN', 10], ['LF', 10], ['Zahlungsausgabe', 20]]
df_pie = pd.DataFrame(data, columns=['Phase', 'Prozessanteil'])
template_color = dict(layout=go.Layout(paper_bgcolor='rgba(240,252,244, 0)'))
file_name = "sonstigeFunktionen.py"


###################################################################
# Gibt die Anzahl der Dokumente aus dem Pfad zurück
def dokumenteVorhanden(customerNumber):
    anzahlDoku = 0
    nummerKunde = customerNumber
    if customerNumber == 99917:
        nummerKunde = "C:\Ordner"
    else:
        nummerKunde = "C:\Ordner"

    pfad = nummerKunde
    if os.path.exists(pfad):
        for path in os.listdir(pfad):
            anzahlDoku += 1
        farbeDokumentUtil = str(anzahlDoku)
    else:
        farbeDokumentUtil = "0"
        write_log(file_name, 'Kein Dokument vorhanden')
    return farbeDokumentUtil

###################################################################
# Öffnet den Dokumenteordner, falls er vorhanden ist
def dokumentPfadOpen(customerNumber):
    pfad = 'C:\Ordner'
    if os.path.exists(pfad):
        nummerKunde = customerNumber
        wb.open('C:\Ordner')
        hidden = "hidden"
    else:
        hidden = "hidden"
        write_log(file_name, 'Ordner nicht vorhanden')
    return hidden

###################################################################
# Kreisdiagramm für den Lohnbuchhaltungsprozess
def create_pie_chart(customerNumber):
    cusNumber = customerNumber

    # Definition der Farben für Kreisdiagramm
    if cusNumber is 0:
        farbe = 'lightgrey'
        title_prozent = '-'
    else:
        farbe, title_prozent = aktueller_status(cusNumber)

    # Erstellung Kreisdiagramm
    pie_chart = px.pie(
        df_pie,
        values='Prozessanteil',
        names='Phase',
        hole=0.5,
        color='Phase',
        color_discrete_map={'Monatswechsel': f'{farbe}', 'AbfrageMonat': f'{farbe}', 'AnzeigeMonat':f'{farbe}','Schaetzung':f'{farbe}','Abrechnung':f'{farbe}','MLJ':f'{farbe}','PN':f'{farbe}','LF':f'{farbe}','Zahlungsausgabe':f'{farbe}'},
        hover_name='Phase',
        labels={"label": "the Labels"},
        template=template_color,
        title=f"<b>{title_prozent} %</b>"
    )
    pie_chart.update_layout(title_x=0.43, title_y= 0.47)

    return pie_chart

def aktueller_status(cusNumber):
    number_mandant = cusNumber
    i = 0
    anfrage_läuft = True

    while anfrage_läuft:
        try:
            df_phasen = DBconnection.df_abrechnungsstand_sql(number_mandant)
            anfrage_läuft = False
        except pyodbc.Error as err:
            write_log(file_name, 'Fehler in Methode "aktueller_status": {err}')

    # Farben werden nach Status definiert
    for phase in df_phasen:
        element = df_phasen[phase]

        # Wenn Mandantennummer nicht existiert
        if str(element).__contains__('Kein:e Mandant:in'):
            farbe = 'lightgrey'
            title_prozent = '-'
            return farbe, title_prozent

        # Wenn Mandantennummer existiert
        elif element is not None and 'None' not in str(element):
            i += 1

    if i < 4:
        farbe = 'darkred'
    elif i < 8:
        farbe = 'darkorange'
    else:
        farbe = 'green'

    # Prozentsatz wird berechnet
    title_prozent = int((i*100)/9)

    return farbe, title_prozent

###################################################################
# Änderung der Farbe für ProgressBar bei der Automatisierung
def return_prog_color(prog_int):
    color = ''
    pro_int = int(prog_int)
    if pro_int < 30:
        color = 'danger'
    elif pro_int < 60:
        color = 'warning'
    else:
        color =  'success'
    return color

