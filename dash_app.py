# dash_ap.py
# Beinhaltet das Layout und die Callbacks

# Import Bibliotheken
import dash
from dash import dcc, html, Input, Output, State, dash_table, Dash
import dash_bootstrap_components as dbc
from dash_bootstrap_templates import ThemeSwitchAIO
from dash.long_callback import DiskcacheLongCallbackManager
import pandas as pd
import diskcache

# Import anderer Klassen
import DBconnection
import sonstigeFunktionen
from logger import write_log

# Variablen ariables
# ************************************************************************
#Kundennummer
customerNumber = ""

# Tabelle in mitte
columns = ['Merkmal', 'Ausp', 'Wert', 'Zahl', 'Datum', 'Info']
df_auswertung = pd.DataFrame(columns=columns)

# Tabelle unten
columns = ['Anrede','Vorname','Nachname','Mail','Tel','Mobil','Fax','Serienbrief']
df_kontakte = pd.DataFrame(columns=columns)

# Variabeln für das Layout (inklusive Darkmode)
template_theme1 = "minty"
template_theme2 = "darkly"
url_theme1 = dbc.themes.MINTY
url_theme2 = dbc.themes.DARKLY
template="none"

# Long Callback manager für den Progress Bar
cache = diskcache.Cache("./cache")
long_callback_manager = DiskcacheLongCallbackManager(cache)

# Erstellung App
# ************************************************************************
dbc_css = ("https://use.fontawesome.com/releases/v5.12.1/css/all.css")
app = Dash(__name__, external_stylesheets=[url_theme1, dbc_css,dbc.icons.BOOTSTRAP], long_callback_manager=long_callback_manager)

# file_name für Log
file_name = "dash_app.py"
write_log(file_name, 'Programm ist gestartet')

app.layout = html.Div([
    # Oberer Bereich des Programms
    # Logo Anzeige, Suche nach Mandaten Nummer, Ausgabe Kundendaten
    #########################################################################
    dbc.Row([
        # Progress Bar
        dbc.Row([
            dbc.Col(html.Progress(id="progress_bar"), width=1)], justify="center"),
        dbc.Row([
            # Logo
            dbc.Col(html.Img(src='/assets/logo.png', className='img-fluid', width='400px')),
            # Eingabe für Mandantennummer mit Such-Button
            dbc.Col([
                    dbc.Input(className='form-control mt-4',id='customerNumber', type='number', min=100, max=99999),
                    ], width=2),
            dbc.Col([
                html.Button("Suchen", id="search", className='btn btn-light mt-4')]),
        ]),
        dbc.Row([
            # Darkmode Schalter
            dbc.Col([
                ThemeSwitchAIO(aio_id="theme", themes=[url_theme1, url_theme2],)
            ], width=2),
            # Ein verstecktes Feld, da immer ein Output benötigt wird (gehört zum Darkmode)
            dbc.Col([(html.Div(id='hidden-div-darkmode'))
            ],style={'display': 'none'}),
        ]),
        dbc.Row([
            # White Label
            dbc.Col([html.H3(id='White_Label', className="text-center, text-center, text-danger"),
            ]),
            dbc.Col([
                html.Div([
                    # Mandantenname und Adresse
                    html.H5(id="Mandant", className='text-left, h-50 d-inline-block, mr-auto'),
                    html.H5(id="nameKnd", className='text-left, h-50 d-inline-block'),
                    html.H6(id="strasse", className='text-left'),
                    html.H6(id="PLZ", className='text-left'),
                ], style={'background-color': 'rgba(148,215,193,0.16)', "width": "34rem", "marginTop": 5}),
            ], width=6),
            dbc.Col([
                # Art des Baulohns
                html.H3(id='Art_Bau', className="text-center, float-right, text-primary"),
            ]),
        ]),
    ]),
    dbc.Row([
        # Kundennummer-CRM, Abrechnungsart, Kunde seit, Gekündigt, DMS seit und alktueller Abrechnungsmonat
        dbc.Col([
            html.Div([html.H5(id="kdnrCRM", className='text-left text-primary'),
            ], style={"marginBottom": 15, "marginLeft": 20}),
        ]),
        dbc.Col([html.H5(id="Abrechnungsart", className='text-left'),], width=2),
        dbc.Col([html.H5(id="KundeSeit", className='text-left, text-primary'), ]),
        dbc.Col([html.H5(id="Gekündigt", className='text-left, text-primary'), ]),
        dbc.Col([html.H4("DMS seit", className='text-left'), ]),
        dbc.Col([html.H4(id="DMSseit", className='text-left'), ]),
        dbc.Col([html.H3(id="AnzeigeMonat",className='text-left, text-primary'), ], width=2),
    ]),

    # Mittlerer Bereich des Programms
    # Informationen von KKArtNr bis PWtest, Auswertung, Kreisdiagramm, Lohnabrechner,
    # Schaetzung, Art der Abrechnung
    #########################################################################
    # LINKS ***********************************
    dbc.Row([
        dbc.Col([
            dbc.Row([
                dbc.Row([
                    # Kundengruppe
                    dbc.Col(html.H1(
                        className='fa fa-user mr-2 fa-4x, text-left, text-black, border-bottom border-primary, text-danger', title="KG"), width=2),
                    dbc.Col(html.H2(
                        " ", id="KG", className='text-left, text-black, border-bottom border-primary')),
                ], style={'background-color': 'rgba(148,215,193,0.16)', "marginLeft": 20}, className='col-md-11'),
                dbc.Row([
                    dbc.Col([
                        # Informationen zu Konsoldierung, Doku, Besonderheit, Katalog, altrenative Fibu, Lohnartenrahmen und PW
                        dbc.Row([
                            dbc.Col(html.H5(className='fa fa-chevron-right', title="Konsolidierung"),width=1),
                            dbc.Col(html.H5(" ", id="Konsolidierung"))]),
                        dbc.Row([
                            dbc.Col(html.H5(className='fas fa-gamepad', title="Doku"),width=1),
                            dbc.Col(html.H5(" ", id="Doku"))]),
                        dbc.Row([
                            dbc.Col(html.H5(className='fas fa-check-circle', title="Besonderheit"), width=1),
                            dbc.Col(html.H5(" ", id="Besonderheit"))]),
                        dbc.Row([
                            dbc.Col(html.H5(className='fas fa-book', title="Katalog"),width=1),
                            dbc.Col(html.H5(" ", id="Katalog"))]),
                        dbc.Row([
                            dbc.Col(html.H5(className='fas fa-chart-line', title="alternativ_Fibu"),width=1),
                            dbc.Col(html.H5(" ", id="alternativ_Fibu"))]),
                        dbc.Row([
                            dbc.Col(html.H5(className='fas fa-crop-alt', title="Lohnartenrahmen"),width=1),
                            dbc.Col(html.H5(" ", id="Lohnartenrahmen"))]),
                        dbc.Row([
                            dbc.Col(html.H5(className='fas fa-key', title="PW"),width=1),
                            dbc.Col(html.H5(" ", id="PW"))])
                    ], style={"marginLeft": 20}, className=' border border-primary rounded'),
                ]),
                dbc.Col([
                    dbc.Row([
                        # Button für die Notizen
                        dbc.Button("Notizen", id="notiz"),
                        dbc.Modal([
                            dbc.ModalHeader("Notizen"),
                            dbc.ModalBody(
                                dbc.Row([
                                    dbc.Label(
                                        "Hier werden vergangene Notizen angezeigt"),
                                    dbc.Label("Neue Notizen hinzufügen:", className="mr-2"),
                                    dbc.Textarea(className="mb-3", placeholder="Fügen Sie ein Kommentar hinzu."),
                                    dbc.Button("Speichern", color="primary"),
                                ],)
                            ),
                            dbc.ModalFooter(
                                dbc.Button("Schließen", id="close", className="ml-auto")
                            ),
                        ],
                            id="modal", is_open=False, size="xl", backdrop=True, scrollable=True, centered=True, fade=True
                        ),
                    ]),
                ], style={"width": "5", "marginTop": 10, "marginLeft": 20}),
            ]),
        ]),

        # MITTE ***********************************
        dbc.Col([
            dbc.Row([
                # Kreisdiagramm für den Lohnabrechunungsprozess
                dcc.Graph(id='piechart', figure={})]),
            dbc.Row([
                dbc.Col([
                    # "Prozess abgeschlossen"-Checkbox
                    dcc.Checklist(
                        options=[{'label': ' Prozess abgeschlossen', 'value': 'PA'}], 
                        value=[],
                        labelStyle={"vertical-align": "middle"},
                        style={"font-size": "30px", "font-weight": "bold"},
                    )
                ], width={'size': 8}),
                dbc.Col([
                    # Button für die Tabelle mit den Auswertung 
                    dbc.Button(
                        "Auswertung", id="btn-open-auswertung", size='md'),
                    # Tabelle für die Auswertung
                    dbc.Modal([
                        dbc.ModalHeader("Weitere Informationen"),
                        dbc.ModalBody(
                            dash_table.DataTable(id="tabelle",columns=[{"name": i, "id": i} for i in df_auswertung.columns],),
                        ),
                        dbc.ModalFooter(
                            dbc.Button("Fenster schließen", id="btn-close-auswertung", color="primary")
                        )
                    ],
                        id="modal_auswertung", is_open=False, size="xl", backdrop=True, scrollable=True, centered=True, fade=True
                    ),
                ], width={'size': 4})
            ]),
        ], width={'size': 4}),

        # RECHTS ***********************************
        dbc.Col([
            # Name des Lohnabrechners
            dbc.Row([
                dbc.Col(html.H1(className='fa fa-user mr-2 fa-2x, text-left', title="Lohnabrechner"), width=2),
                dbc.Col(html.H2(" ", id="lohnabrechner", className='text-left, text-black, border-bottom border-primary')),
            ], style={'background-color': 'rgba(148,215,193,0.16)'}),

            # Name des letzten Lohnabrechern und das Datum, bis zu welchem die Abrechnung zu rechnen ist
            dbc.Row([
                dbc.Col(html.H5(className='fa fa-user-circle fa-xs, text-left', title="Letzter Lohnabrechner"), width=1),
                dbc.Col(html.H5(id="letzerLohnabrechner", className="text-primary"))
            ]),
            dbc.Row([
                dbc.Col(html.H5(id="abrechnungsdatum", className="text-primary")),
            ]),
            # Informationen zu dem Schätztermin und der Art der Abrechnung.
            dbc.Row([
                dbc.Col([
                    dbc.Row(html.H4("Schätztermin", className='text-primary')),
                    dbc.Row(html.H4(id="schätzterminInfos1")),
                    dbc.Row(html.H5(id="schätzterminInfos2")),
                    dbc.Row(html.H4("Art der Abrechnung", className='text-primary')),
                    dbc.Row(html.H4(id="artDerAbrechnung1")),
                    dbc.Row(html.H5(id="artDerAbrechnung2")),
                ], style={"height": "300px"}, className="scroll overflow-auto")
            ], className=' border border-primary rounded mt-4'),

            # der Button öffnet einen Dokumenten-Ordner, die Icons sind Platzhalter.
            dbc.Row([
                # Anzahl Dokumente
                dbc.Col(html.H2(id="anzDoku", className='text-warning, mt-4')),
                # Öffnet Dokumentepfad
                dbc.Col(dbc.Button("Dokumente", id="btn-open-pfad", size='md', className='mt-4')),
                #Ein verstecktes Feld, da immer ein Output benötigt wird (gehört zur Ausgabe der Anzahl der Dokumente)
                dbc.Col([(html.Div(id='hidden-div'))],style={'display': 'none'}),
                # Platzhalter Icons
                dbc.Col(html.H3(className='fa fa-download mt-4 fa-2x', title="Dokumente zur Erfassung")),
                dbc.Col(html.H3(className='fa fa-circle mt-4 fa-2x', title="icon1")),
                dbc.Col(html.H3(className='fa fa-circle mt-4 fa-2x', title="icon2")),
                dbc.Col(html.H3(className='fa fa-circle mt-4 fa-2x', title="icon3"))
            ]),
        ]),
    ]),
    # Unterer Bereich des Programms
    # Ansprechpartner:innen, Versand & Anlieferung, Rechnungsempfänger:in
    #########################################################################
    # LINKS ***********************************
    dbc.Row([
        dbc.Col([
            dbc.Card([
                # Asprechpartner:in, Adresse und Kontaktdaten
                dbc.CardHeader("Ansprechpartner:innen", style={'color': "primary", 'font-weight': 'bold', 'font-size': '1.5em'}),
                dbc.CardBody([
                    dcc.Markdown(" ", id="db_ansprechpartnerKontakt", className="card-text"),
                    dbc.Button("Alle Ansprechpartner", id="collapse-button", className="mb-3", color="primary", n_clicks=0,),
                ]),
            ], style={'background-color': 'rgba(148,215,193,0.16)', "height": "200px", 'whiteSpace': 'pre-wrap'}, class_name="scroll overflow-auto", color="primary", outline=True),
        ]),

        # MITTE ***********************************
        dbc.Col([
            dbc.Card([
                html.H4("",id = "versand",
                        className="bi bi-box-seam",title="Versand"),
            ], className='card mb-4 border-0'),
            dbc.Card([
                html.H4(" ", id = "anlieferung_beweg",
                        className="bi bi-truck",title="Anlieferung Bewegungsdaten"),
                dbc.Progress(id= "bewegung_prog", label="0%", value=1, className="mb-3"),
            ], className='card mb-4 border-0'),
            dbc.Card([
                html.H4(" ", id = "anlieferung_stamm", 
                        className="bi bi-truck",title="Anlieferung Stammdaten"),
                dbc.Progress(id= "stamm_prog",label="75%", value=1,className="mb-3"),
            ], className='card mb-4 border-0'),
        ],  width={'size': 3}),

        # RECHTS ***********************************
        dbc.Col([
            dbc.Card([
                # Rechnungsempfänger:in, Adresse und Kontaktdaten
                dbc.CardHeader("Rechnungsempfänger:in", style={'font-weight': 'bold', 'font-size': '1.5em'}),
                dbc.CardBody([
                    html.P(id="db_kontakt_adresse", className="card-title"),
                    html.P(" ", id="db_kontaktName", className="card-text"),
                    html.P(" ", id="db_kontakt", className="card-text"),
                ]),
            ],  style={'background-color': 'rgba(148,215,193,0.16)', "height": "200px", 'whiteSpace': 'pre-wrap'}, class_name="scroll overflow-auto", color="primary", outline=True),
        ])
    ], style={"marginTop": 15}),
    dbc.Row([
        dbc.Collapse(
            dbc.Card(dbc.CardBody(dash_table.DataTable(id="ansprechTabelle", columns=[
                {"name": i, "id": i} for i in df_kontakte.columns],),)),
            id="collapse",
            is_open=False,),
    ]),
], className='mx-5')

# ************************************************************************
# 1 Callback Darkmode
@app.callback(
    Output("hidden-div-darkmode", "children"), Input(ThemeSwitchAIO.ids.switch("theme"), "value"),
)

# 1 Darkmode Methode, um den Schalter zu betätigen
def update_graph_theme(toggle):
    template = template_theme1 if toggle else template_theme2
    return ("hidden")
 
# ************************************************************************
# 2 Callback Mandantendaten
@app.long_callback(
        output = [ 
                # Output OBEN
                Output("White_Label", "children"),
                Output("Mandant", "children"),
                Output("nameKnd", "children"),
                Output("strasse", "children"),
                Output("PLZ", "children"),
                Output("Art_Bau", "children"),
                Output("kdnrCRM","children"),
                Output("Abrechnungsart","children"),
                Output("KundeSeit","children"),
                Output("Gekündigt","children"),
                Output("DMSseit", "children"),
                Output("AnzeigeMonat","children"),
                
                # Output  MITTE - LINKS
                Output("KG", "children"),
                Output("Konsolidierung", "children"),
                Output("Doku", "children"),
                Output("Besonderheit", "children"),
                Output("Katalog", "children"),
                Output("alternativ_Fibu", "children"),
                Output("Lohnartenrahmen", "children"),
                Output("PW", "children"),
                #Outout("old_notes", "children")
                
                # Output  MITTE - RECHTS
                Output("lohnabrechner", "children"),
                Output("letzerLohnabrechner", "children"),
                Output("abrechnungsdatum", "children"),
                Output("schätzterminInfos1", "children"),
                Output("schätzterminInfos2", "children"),
                Output("artDerAbrechnung1", "children"),
                Output("artDerAbrechnung2", "children"),
                Output("anzDoku", "children"),

                # Output  MITTE - MITTE
                Output("tabelle", "data"),
                Output(component_id='piechart', component_property='figure'),

                # Output  UNTEN - LINKS
                Output("db_ansprechpartnerKontakt", "children"),
                Output("ansprechTabelle", "data"),
                
                # Output  UNTEN - RECHTS
                Output("db_kontakt_adresse", "children"),
                Output("db_kontaktName", "children"),
                Output("db_kontakt", "children"),

                # Output UNTEN - MITTE
                Output("versand", "children"),
                Output("anlieferung_beweg", "children"),
                Output("bewegung_prog", "value"), 
                Output("bewegung_prog", "label"),
                Output("bewegung_prog", "color"), 
                Output("anlieferung_stamm", "children"),
                Output("stamm_prog", "value"), 
                Output("stamm_prog", "label"),
                Output("stamm_prog", "color"), 
        ],
        inputs = [
                Input("search", "n_clicks"),
                State("customerNumber", "value")
        ],
        running = [
                (Output("search", "disabled"), True, False),
                (Output("progress_bar", "style"),
                {"visibility": "visible"},
                {"visibility": "hidden"},),
        ],
        progress = [
                Output("progress_bar", "value"), 
                Output("progress_bar", "max")
        ],
    )

# 2 Mandantendaten Methode, die alle SQL-Befehle aufruft
def update_output(set_progress, n_clicks, customerNumber):
        write_log(file_name, 'Callback wurde aufgerufen. Methode "update_output" ')

        # Progress-Bar zeigt an, wenn Callback noch lädt
        total = 10
        for i in range(total):
                set_progress((str(i + 1), str(total)))

        # DB-Abfragen der SQL-Befehle. Rückgabe der Daten in DataFrames *************************
        df_kunde = DBconnection.df_kunde_sql(customerNumber)
        idKunde = df_kunde['IDKunde']
        df_monat = DBconnection.df_monat_sql(customerNumber)
        df_artAbrechnung = DBconnection.df_artAbrechnung_sql(idKunde)
        df_lohnabrechner = DBconnection.df_lohnabrechner_sql(idKunde)
        df_schaetzung = DBconnection.df_schaetzung_sql(idKunde)
        df_auswertung = DBconnection.df_auswertung_sql(idKunde)
        df_ansprechpartner = DBconnection.ansprechpartner_sql(idKunde)
        df_kontakt =  DBconnection.rechnungKontakt_sql(idKunde)
        df_kontakt_adresse = DBconnection.rechnungAdresse_sql(idKunde)
        df_bewegDaten = DBconnection.bewegungsdaten_sql(idKunde)
        df_stammDaten = DBconnection.stammdaten_sql(idKunde)
        df_versand = DBconnection.versand_sql(idKunde)

        # Zuordnung der Daten aus DataFrames auf Felder *************************
        # OBEN
        db_whitelabel =df_kunde['White_Label']
        db_Mandant = df_kunde['Mandant']
        db_nameKnd = df_kunde['nameKnd']
        db_strasse = df_kunde['strasse']
        db_PLZ = df_kunde['PLZ'] + '''  ''' + df_kunde['Ort']
        db_Art_Bau = df_kunde['Art_Bau']
        db_kdnrCRM = df_kunde['kdnrCRM']
        db_Abrechnungsart = df_kunde['Abrechnungsart']
        db_KundeSeit = df_kunde['KundeSeit']
        db_Gekündigt = df_kunde['Gekündigt']
        db_DMSseit = df_kunde['DMSseit']
        db_AnzeigeMonat = df_monat['AnzeigeMonat']

        # MITTE - LINKS
        db_KG = df_kunde['KG']
        db_Konsolidierung = df_kunde['Konsolidierung']
        db_doku = df_kunde['Doku']
        db_besonderheit = df_kunde['Besonderheit']
        db_katalog = df_kunde['Katalog']
        db_alternativFibu = df_kunde['alternativ_Fibu']
        db_Lohnartenrahmen = df_kunde['Lohnartenrahmen']
        db_PW = df_kunde['PW']

        # MITTE - MITTE
        tabelle = df_auswertung.to_dict('records')
        if customerNumber is None:
                pie_chart = sonstigeFunktionen.create_pie_chart(0)
        elif customerNumber is not None:
                pie_chart = sonstigeFunktionen.create_pie_chart(customerNumber)
        else:
                raise dash.exceptions.PreventUpdate 
                write_log(file_name, 'Pie Chart konnte nicht aktualisiert werden!')

        # MITTE - RECHTS
        db_lohnabrechner = df_lohnabrechner['Lohnabrechner']
        db_letzerLohnabrechner = df_kunde['BearbeiterletzterAuftrag']
        db_abrechnungsdatum = df_artAbrechnung['Datum']
        db_schätzterminInfos1 = df_schaetzung['Schaetzung']
        db_schätzterminInfos2 = df_schaetzung['InfoSchaetzung']
        db_artDerAbrechnung1 = df_artAbrechnung['ArtDerAbrechnung']
        db_artDerAbrechnung2 = df_artAbrechnung['InfoArtDerAbrechnung']
        sf_anzahlDokument = sonstigeFunktionen.dokumenteVorhanden(customerNumber)

        # UNTEN - LINKS
        db_ansprechpartnerName = '''
**''' + df_ansprechpartner['Anrede'] + ' '+ df_ansprechpartner['Vorname'] + ' '+  df_ansprechpartner['Nachname'] + '**: '
        db_ansprechpartnerMail = '''
E-Mail: ''' + df_ansprechpartner['Mail']
        db_ansprechpartnerTel = '''
Tel: ''' +df_ansprechpartner['Tel']
        db_ansprechpartnerMob = '''
Mobil: ''' +df_ansprechpartner['Mobil'] 
        db_ansprechpartnerFax = '''
Fax: ''' +df_ansprechpartner['Fax'] 
        db_ansprechpartnerSerienBrief = '''
Serienbrief: ''' +df_ansprechpartner['Serienbrief']
        db_ansprechpartnerKontakt = db_ansprechpartnerName + db_ansprechpartnerMail + db_ansprechpartnerTel + db_ansprechpartnerMob + db_ansprechpartnerFax + db_ansprechpartnerSerienBrief
        ansprechTabelle = df_ansprechpartner.to_dict('records')
        
        # UNTEN MITTE
        db_versand = """    """ + df_versand["Ausp"]
        db_beweg_value = df_bewegDaten['Zahl'].sum()
        anlieferung_beweg = """    """ + df_bewegDaten["Ausp"]
        prog_color_beweg = sonstigeFunktionen.return_prog_color(db_beweg_value)
        db_stamm_string = """    """ + df_stammDaten["Ausp"]
        db_stamm_value = df_stammDaten['Zahl'].sum()
        prog_color_stamm = sonstigeFunktionen.return_prog_color(db_stamm_value)

        # UNTEN - RECHTS
        db_adresse = df_kontakt_adresse['Name1'] + '; ' + df_kontakt_adresse['Name2'] + """
Strasse: """ + df_kontakt_adresse['Strasse'] + """
Ort: """ + df_kontakt_adresse['Ort'] + '''
PLZ: ''' + df_kontakt_adresse['PLZ'] + '''
Kdnr: ''' + df_kontakt_adresse['Kdnr']
        db_kontaktName = df_kontakt['Anrede'] + ' '+ df_kontakt['Vorname'] + ' '+  df_kontakt['Nachname']
        db_kontaktMail = 'E-Mail: ' + df_kontakt['Mail'] 
        db_kontaktTel = '''
Tel:''' + df_kontakt['Tel'] 
        db_kontaktMob = '''
Mobil: ''' +df_kontakt['Mobil'] 
        db_kontaktFax = '''
Fax: ''' +df_kontakt['Fax'] 
        db_kontaktSerienBrief = '''
Serienbrief: ''' +df_kontakt['Serienbrief'] 
        db_kontakt = db_kontaktMail + db_kontaktTel + db_kontaktMob + db_kontaktFax + db_kontaktSerienBrief
        
        # Rückgabe der Wert an die ID ************************* 
        return(db_whitelabel, db_Mandant, db_nameKnd, db_strasse, db_PLZ, db_Art_Bau, db_kdnrCRM, db_Abrechnungsart, db_KundeSeit, db_Gekündigt, db_DMSseit, db_AnzeigeMonat,
        db_KG,db_Konsolidierung, db_doku, db_besonderheit, db_katalog, db_alternativFibu, db_Lohnartenrahmen, db_PW, db_lohnabrechner,
        db_letzerLohnabrechner,db_abrechnungsdatum, db_schätzterminInfos1, db_schätzterminInfos2, db_artDerAbrechnung1, db_artDerAbrechnung2, sf_anzahlDokument, tabelle, pie_chart, 
        db_ansprechpartnerKontakt,ansprechTabelle,
        db_adresse, db_kontaktName,db_kontakt,
        db_versand, anlieferung_beweg, db_beweg_value, f"{db_beweg_value} %", prog_color_beweg, 
        db_stamm_string, db_stamm_value, f"{db_stamm_value} %", prog_color_stamm
        )

# ************************************************************************
# 3 Callback Notizen
@app.callback(
        Output("modal", "is_open"),
        [Input("notiz", "n_clicks"), 
        Input("close", "n_clicks")],
        [State("modal", "is_open")],
)

# Notzien Methode, öffnet und schließt das Notizfeld
def toggle_modal_n(n1, n2, is_open):
        write_log(file_name, 'Callback wurde aufgerufen. Methode "toggle_modal_n" ')

        if n1 or n2:
                return not is_open
        return is_open

# ************************************************************************
# 4 Callback Auswertung
@app.callback(
        Output("modal_auswertung", "is_open"),
        [Input("btn-open-auswertung", "n_clicks"), 
        Input("btn-close-auswertung", "n_clicks")],
        [State("modal_auswertung", "is_open")],
)

# 4 Auswertung Methode, öffnet und schließt die Auswertung
def toggle_modal_a(n1, n2, is_open):
        write_log(file_name, 'Callback wurde aufgerufen. Methode "toggle_modal_a" ')

        if n1 or n2:
                return not is_open
        return is_open

# ************************************************************************
# 5 Callback Dokumentepfad
@app.callback(
        Output("hidden-div", "children"),
        Input("btn-open-pfad", "n_clicks"),
        State("customerNumber", "value")
)

# 5 Dokumentepfad Methode, öffnet die Methode "dokumentPfadOpen", damit sich der Dokumentenpfad öffnet
def openPath(n_clicks, number):
        write_log(file_name, 'Callback wurde aufgerufen. Methode "openPath" ')

        if number is None:
                hidden = "hidden" 
        else:
                hidden = sonstigeFunktionen.dokumentPfadOpen(customerNumber) 
        return hidden 

# ************************************************************************
# 6 Callback Ansprechpartner
@app.callback(
    Output("collapse", "is_open"),
    [Input("collapse-button", "n_clicks")],
    [State("collapse", "is_open")],
)

# 6 Ansprechpartner Methode, öffnet und schließt die Liste der Ansprechpartner
def toggle_collapse(n, is_open):
        write_log(file_name, 'Callback wurde aufgerufen. Methode "toggle_collapse" ')
        if n:
                return not is_open
        return is_open     

# Run app - hier kann der Port geändert werden
# ************************************************************************
if __name__ == "__main__":
    app.run_server(port=6050, debug=True)