# DBconnection.py
# Diese Datei beinhaltet alle SQL-Befehle

# Importe
from numpy import number
import numpy as np
import pyodbc
import pandas as pd
from logger import write_log

# Verbindung mir der Datenbank
driver = ""  # Treiber
host = "" # sql server wo die DB drauf ist
database = ""  # Name der DB
user = ""  # user
password = ""  # pw
connection = pyodbc.connect('DRIVER=' + driver + ';SERVER=' + host +
                            ';DATABASE=' + database + ';UID=' + user + ';PWD=' + password)

# Variablen
num_mandant = 0
num_idKunde = 0
id_kunde = 0
i = 0
file_name = "DBconnection.py"


# 1 SQL Kunde gibt die dazugehörigen Daten des Mandaten zurück
##########################################################################################################################
def df_kunde_sql(number_mandant):
    if connection:
        write_log(file_name, 'connection ok: df_kunde_sql')

        # Variablen
        cursor = connection.cursor()
        fields_kunde_list = ['IDKunde', 'nameKnd', 'strasse', 'Ort', 'PLZ', 'Mandant', 'kdnrCRM', 'White_Label', 'KG', 'KundeSeit', 'Gekündigt', 'versandRE', 'Katalog', 'Lohnartenrahmen',
                             'PW', 'alternativ_Fibu', 'Baulohn', 'Art_Bau', 'DMSseit', 'Abrechnungsart', 'Sofort', 'Doku', 'Konsolidierung', 'Besonderheit', 'STBMAN', 'BearbeiterletzterAuftrag']
        sql_return_list = []
        k = 0

        # SQL-Abfrage und Formatierung der Rückgabe in DataFrame
        try:
            cursor.execute(
                f"""SELECT Top 1 c.I_CUSTOMER_M as IDKunde,(c.s_name1 + ' ' + isnull(c.s_name2,'')) AS nameKnd,
                    c.S_STREET as strasse,
                    c.s_town as Ort,
                    c.S_ZIPCODENO as PLZ,
                    c.s_optional3 AS Mandant,
                    c.s_custno AS kdnrCRM,
                    (CASE WHEN  sao.GetTypeLookupValues('customer_m','n_optional7', 2,c.n_optional7, '', 0) LIKE '%white label%' THEN 'White Label' ELSE '' END) AS White_Label,
                    (SELECT k.s_description FROM sao.custgroup_p k WHERE k.i_custgroup_p = c.i_custgroup_p) AS KG,
                    isnull((convert(varchar,day(c.d_optional11)) + '.' + convert(varchar,month(c.d_optional11)) + '.' + convert(varchar,year(c.d_optional11))),'-') as KundeSeit,
                    isnull((convert(varchar,day(c.d_optional12)) + '.' + convert(varchar,month(c.d_optional12)) + '.' + convert(varchar,year(c.d_optional12))),'-') AS Gekündigt,
                    (CASE WHEN (b_optional18 = 1 AND b_optional19 = 0) THEN 'WebCenter' 
                        WHEN (b_optional18 = 0 AND b_optional19 = 1) THEN 'per Mail'  
                        WHEN (b_optional18 = 0 AND b_optional19 = 0) THEN 'per Post'  
                        ELSE  'FEHLER!' END) AS versandRE,
                    isnull(sao.GetTypeLookupValues('customer_m','n_optional30', 2,c.n_optional30, '', 0),'-') AS Katalog, 
                    isnull(c.s_pager,'-') AS Lohnartenrahmen,
                    isnull(c.s_optional5,'-') AS PW, 
                    isnull(c.s_optional2,'-') AS alternativ_Fibu,
                    (CASE WHEN  sao.GetTypeLookupValues('customer_m','n_optional7', 2,c.n_optional7, '', 0) LIKE '%Baulohn%' THEN 'Baulohn' ELSE '' END) AS Baulohn,
                    (SELECT s_description FROM sao.addrorigin ao WHERE ao.i_addrorigin = c.i_addrorigin) AS Art_Bau,
                    isnull((convert(varchar,day(c.d_optional13)) + '.' + convert(varchar,month(c.d_optional13)) + '.' + convert(varchar,year(c.d_optional13))),'-') AS DMSseit,
                    sao.GetTypeLookupValues('customer_m','n_optional28', 2,c.n_optional28, '', 0) AS Abrechnungsart,
                    (CASE  WHEN c.b_optional16 = 0 THEN 'nein'  ELSE 'ja' END) AS Sofort,
                    isnull(c.s_gln,'-') AS Doku,
                    isnull(c.s_creditcardno,'-') AS Konsolidierung,
                    isnull(sao.GetTypeLookupValues('customer_m','n_optional7', 2,c.n_optional7, '', 0),'-') AS Besonderheit,
                    isnull(sao.GetTypeLookupValues('customer_m','n_optional6', 2,c.n_optional6, '', 0),'-')  AS STBMAN,
                    (SELECT TOP 1 (isnull(e.s_name2, '') + ' ' + e.s_name1 ) 
                        FROM    sao.orders_p o, 
                                sao.employee_m e 
                        WHERE o.dt_deleted IS NULL 
                        AND o.i_customer_m = c.I_CUSTOMER_M 
                        AND e.i_employee_m = o.i_employee_m 
                        ORDER BY o.d_orderdate DESC) AS BearbeiterletzterAuftrag 
                    FROM sao.customer_m c 
                    WHERE c.s_optional3 = '{number_mandant}' 
                    and c.dt_deleted is null""")

            sql_return = cursor.fetchall()

            if len(sql_return) > 0:
                while k < 26:
                    for row in sql_return:
                        sql_return_list.append(str(row[k]))
                        k += 1
            # Wenn das Programm startet oder kein Mandant gefunden wurde, wird das angezeigt
            else:
                if (number_mandant is None):
                    sql_return_list = [-1, "Bitte wählen Sie eine Mandantennummer aus.", "", "", "", "",
                                       "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
                else:
                    sql_return_list = [-1, "Es konnte kein Mandant mit dieser Nummer gefunden werden!", "", "",
                                       "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
            dictionary_kunde = dict(zip(fields_kunde_list, sql_return_list))
            df = pd.DataFrame(dictionary_kunde, index=[0])

        except pyodbc.DatabaseError as err:
            write_log(
                file_name, f'Datenbankfehler in Methode "df_kunde_sql": {err}')

        try:
            return(df)
        except UnboundLocalError as errUnbound:
            write_log(
                file_name, f'Keine Nummer eingegeben! In Methode "df_kunde_sql": {errUnbound}')
    else:
        write_log(file_name, 'Connection-Fehler in Methode "df_kunde_sql" ')

# 2 SQL Monat gibt den aktuellen Abrechnungsmonat zurück
##########################################################################################################################
def df_monat_sql(number_mandant):
    if connection:
        write_log(file_name, 'connection ok: df_monat_sql')

        # Variablen
        cursor = connection.cursor()
        fields_monat_list = ['AnzeigeMonat']
        sql_return_list = []
        a = 0

        # SQL-Abfrage und Formatierung der Rückgabe in DataFrame
        try:
            cursor.execute(
                f"""SELECT top 1 (convert(varchar,month(dt_667_Monatswechsel)) + '/' + convert(varchar,year(dt_667_Monatswechsel))) as AnzeigeMonat 
                    from dbo.LD_AbrechnungsMonat 
                    where s_Mandant = '{number_mandant}' 
                    order by dt_667_Monatswechsel desc""")

            sql_return = cursor.fetchall()
            if len(sql_return) > 0:
                while a < 1:
                    for row in sql_return:
                        sql_return_list.append(str(row[a]))
                        a += 1
            else:
                sql_return_list = [""]
            dictionary_monat = dict(zip(fields_monat_list, sql_return_list))
            df = pd.DataFrame(dictionary_monat, index=[0])

        except pyodbc.DatabaseError as err:
            write_log(
                file_name, f'Datenbankfehler in Methode "df_monat_sql": {err}')

        try:
            return(df)
        except UnboundLocalError as errUnbound:
            write_log(
                file_name, f'Keine Nummer eingegeben! In Methode "df_monat_sql": {errUnbound}')
    else:
        write_log(file_name, 'Connection-Fehler in Methode "df_monat_sql" ')

# 3 Abrechnungsstand selektiert die Daten für das Kreisdiagramm
##########################################################################################################################
def df_abrechnungsstand_sql(number_mandant):
    if connection:
        write_log(file_name, 'connection ok: df_abrechnungsstand_sql')

        # Variablen
        cursor = connection.cursor()
        fields_abrechnungsstand_list = ['Monatswechsel', 'AbfrageMonat', 'AnzeigeMonat',
                                        'Schaetzung', 'Abrechnung', 'MLJ', 'PN', 'LF', 'Zahlungsausgabe']
        sql_return_list = []
        a = 0

        # SQL-Abfrage und Formatierung der Rückgabe in DataFrame
        try:
            cursor.execute(
                f"""SELECT top 1 dt_667_Monatswechsel as Monatswechsel, 
                    (year(dt_667_Monatswechsel) * 100 + month(dt_667_Monatswechsel)) as AbfrageMonat, 
                    (convert(varchar,month(dt_667_Monatswechsel)) + '/' + convert(varchar,year(dt_667_Monatswechsel))) as AnzeigeMonat, 
                    dt_666_Schaetzung as Schaetzung, dt_330_Abrechnung as Abrechnung, 
                    dt_341_MLJ as MLJ, 
                    dt_366367_PN as PN, 
                    dt_670_LF as LF, 
                    dt_884_Zahlungsausgabe as Zahlungsausgabe 
                    from dbo.LD_AbrechnungsMonat 
                    where s_Mandant = '{number_mandant}' 
                    order by dt_667_Monatswechsel desc""")

            sql_return = cursor.fetchall()
            if len(sql_return) > 0:
                while a < 9:
                    for row in sql_return:
                        sql_return_list.append(str(row[a]))
                        a += 1
            else:
                sql_return_list = ['Kein:e Mandant:in']
            dictionary_abrechnungsstand = dict(
                zip(fields_abrechnungsstand_list, sql_return_list))
            df = pd.DataFrame(dictionary_abrechnungsstand, index=[0])

        except pyodbc.DatabaseError as err:
            write_log(
                file_name, f'Datenbankfehler in Methode "df_abrechnungsstand_sql": {err}')

        try:
            return(df)
        except UnboundLocalError as errUnbound:
            write_log(
                file_name, f'Keine Nummer eingegeben! In Methode "df_abrechnungsstand_sql": {errUnbound}')
    else:
        write_log(
            file_name, 'Connection-Fehler in Methode "df_abrechnungsstand_sql" ')

# 4 SQL für Art der Abrechnung gibt das Datum, bis wann abgerechnet werden muss,
# die Art der Abrechnung und die dazugehörigen Informationen zurück
##########################################################################################################################
def df_artAbrechnung_sql(id_kunde):
    if connection:
        write_log(file_name, 'connection ok: df_artAbrechnung_sql')

        # Variablen
        cursor = connection.cursor()
        sql_return_list = []
        sql_return_list2 = []
        sql_return_list3 = []

        # SQL-Abfrage und Formatierung der Rückgabe in DataFrame
        try:
            cursor.execute(
                f"""SELECT isnull(((convert(varchar,day(df.d_featuredate))  + '. ') + (case when convert(varchar,month(df.d_featuredate)) = 3 then 'des Vormonats' 
                        when convert(varchar,month(df.d_featuredate)) = 5 then 'des lfd. Monats' 
                        when convert(varchar,month(df.d_featuredate)) = 7 then  'des Folgemonats' 
                        else 'Fehler' end )),'') AS datum, 
                    isnull(c.s_description,'-') AS Ausp, 
                    isnull(df.ls_featuremtext, '-') AS Info 
                    
                    FROM sao.devfeature df, 
                        sao.criterion c,	
                        sao.feature f 
                    
                    WHERE df.i_customer_m = {int(id_kunde)} 
                    AND df.dt_deleted IS NULL 
                    AND df.i_criterion = c.i_criterion 
                    AND f.i_feature = df.i_feature 
                    AND f.i_feature in(9) 
                    ORDER BY f.s_optional2""")

            sql_return = cursor.fetchall()
            if len(sql_return) > 0:
                for row in sql_return:
                    sql_return_list.append(str(row[0]))
                    sql_return_list2.append(str(row[1])+"\n")
                    sql_return_list3.append(str(row[2])+",\n")
                data = {'Datum': sql_return_list, 'ArtDerAbrechnung': sql_return_list2,
                        'InfoArtDerAbrechnung': sql_return_list3}
                df = pd.DataFrame(data)
            else:
                sql_return_list = ["", "", ""]
                sql_return_list2 = ["", "", ""]
                sql_return_list3 = ["", "", ""]
                data = {'Datum': sql_return_list, 'ArtDerAbrechnung': sql_return_list2,
                        'InfoArtDerAbrechnung': sql_return_list3}
                df = pd.DataFrame(data)

        except pyodbc.DatabaseError as err:
            write_log(
                file_name, f'Datenbankfehler in Methode "df_artAbrechnung_sql": {err}')
        try:
            return(df)
        except UnboundLocalError as errUnbound:
            write_log(
                file_name, f'Keine Nummer eingegeben! In Methode "df_artAbrechnung_sql": {errUnbound}')
    else:
        write_log(file_name, 'Connection-Fehler in Methode "df_artAbrechnung_sql" ')

# 5 SQL Lohnabrechner gibt den Lohnabrechner zurück
##########################################################################################################################
def df_lohnabrechner_sql(id_kunde):
    if connection:
        write_log(file_name, 'connection ok: df_lohnabrechner_sql')

        # Variablen
        cursor = connection.cursor()
        fields_lohnabrechner_list = ['Lohnabrechner']
        sql_return_list = []
        l = 0

        # SQL-Abfrage und Formatierung der Rückgabe in DataFrame
        try:
            cursor.execute(
                f"""SELECT top 1 c.s_description AS Ausp 
                    
                    FROM sao.devfeature df, 
                    sao.criterion c, 
                    sao.feature f 
        
                    WHERE df.i_customer_m = {int(id_kunde)} 
                    AND df.dt_deleted IS NULL 
                    AND df.i_criterion = c.i_criterion 
                    AND f.i_feature = df.i_feature 
                    AND  f.I_FEATURE in ('10','17') 
                    ORDER BY f.s_optional2""")

            sql_return = cursor.fetchall()
            if len(sql_return) > 0:
                while l < 1:
                    for row in sql_return:
                        sql_return_list.append(str(row[l]))
                        l += 1
            else:
                sql_return_list = [""]
            dictionary_lohnabrechner = dict(
                zip(fields_lohnabrechner_list, sql_return_list))
            df = pd.DataFrame(dictionary_lohnabrechner, index=[0])

        except pyodbc.DatabaseError as err:
            write_log(
                file_name, f'Datenbankfehler in Methode "df_lohnabrechner_sql": {err}')
        try:
            return(df)
        except UnboundLocalError as errUnbound:
            write_log(
                file_name, f'Keine Nummer eingegeben! In Methode "df_lohnabrechner_sql": {errUnbound}')
    else:
        write_log(file_name, 'Connection-Fehler in Methode "df_lohnabrechner_sql" ')

# 6 SQL Auswertung gibt die Werte für die Tabelle der Auswertung zurück
##########################################################################################################################
def df_auswertung_sql(id_kunde):
    if connection:
        write_log(file_name, 'connection ok: df_auswertung_sql')

        # Variablen
        cursor = connection.cursor()
        sql_return_list = []
        sql_return_list2 = []
        sql_return_list3 = []
        sql_return_list4 = []
        sql_return_list5 = []
        sql_return_list6 = []

        # SQL-Abfrage und Formatierung der Rückgabe in DataFrame
        try:
            cursor.execute(
                f"""SELECT f.s_description AS Merkmal,
                    c.s_description AS Ausp,
                    df.s_featurevalue AS Wert,
                    df.n_featurenumber AS Zahl,
                    (convert(varchar,day(df.d_featuredate))  + '.') AS datum,
                    df.ls_featuremtext AS Info 
                    
                    FROM sao.devfeature df, 
                    sao.criterion c, 
                    sao.feature f 
                    
                    WHERE df.i_customer_m = {int(id_kunde)} 
                    AND df.dt_deleted IS NULL 
                    AND df.i_criterion = c.i_criterion 
                    AND f.i_feature = df.i_feature 
                    AND f.I_FEATURE in (1,2,3,4,5,8,14,12) 
                    ORDER BY f.i_feature""")

            sql_return = cursor.fetchall()

            if len(sql_return) > 0:
                for row in sql_return:
                    sql_return_list.append(str(row[0]))
                    sql_return_list2.append(str(row[1]))
                    sql_return_list3.append(str(row[2]))
                    sql_return_list4.append(str(row[3]))
                    sql_return_list5.append(str(row[4]))
                    sql_return_list6.append(str(row[5]))
                data = {'Merkmal': sql_return_list, 'Ausp': sql_return_list2, 'Wert': sql_return_list3,
                        'Zahl': sql_return_list4, 'Datum': sql_return_list5, 'Info': sql_return_list6}
                df = pd.DataFrame(data)
            else:
                sql_return_list = ["/"]
                sql_return_list2 = ["/"]
                sql_return_list3 = ["/"]
                sql_return_list4 = ["/"]
                sql_return_list5 = ["/"]
                sql_return_list6 = ["/"]
                data = {'Merkmal': sql_return_list, 'Ausp': sql_return_list2, 'Wert': sql_return_list3,
                        'Zahl': sql_return_list4, 'Datum': sql_return_list5, 'Info': sql_return_list6}
                df = pd.DataFrame(data)

        except pyodbc.DatabaseError as err:
            write_log(
                file_name, f'Datenbankfehler in Methode "df_auswertung_sql": {err}')
        try:
            return(df)
        except UnboundLocalError as errUnbound:
            write_log(
                file_name, f'Keine Nummer eingegeben! In Methode "df_auswertung_sql": {errUnbound}')
    else:
        write_log(file_name, 'Connection-Fehler in Methode "df_auswertung_sql" ')

# 7 SQL Schaetzung gibt die Informationen für die Schaetzung zurück
##########################################################################################################################
def df_schaetzung_sql(id_kunde):
    if connection:
        write_log(file_name, 'connection ok: df_schaetzung_sql')

        # Variablen
        cursor = connection.cursor()
        sql_return_list = []
        sql_return_list2 = []

        try:
            cursor.execute(
                f"""SELECT c.s_description AS Ausp, 
                    isnull(df.ls_featuremtext,'') AS Info 
                    
                    FROM sao.devfeature df, 
                    sao.criterion c, 
                    sao.feature f 
                    
                    WHERE df.i_customer_m = {int(id_kunde)} 
                    AND df.dt_deleted IS NULL 
                    AND df.i_criterion = c.i_criterion 
                    AND f.i_feature = df.i_feature 
                    AND f.i_feature in(7) 
                    ORDER BY f.s_optional2""")

            sql_return = cursor.fetchall()

            if len(sql_return) > 0:
                for row in sql_return:
                    sql_return_list.append(str(row[0])+",\n")
                    sql_return_list2.append(str(row[1]))
                data = {'Schaetzung': sql_return_list,
                        'InfoSchaetzung': sql_return_list2}
                df = pd.DataFrame(data)
            else:
                sql_return_list = ["", ""]
                sql_return_list2 = ["", ""]
                data = {'Schaetzung': sql_return_list,
                        'InfoSchaetzung': sql_return_list2}
                df = pd.DataFrame(data)

        except pyodbc.DatabaseError as err:
            write_log(
                file_name, f'Datenbankfehler in Methode "df_schaetzung_sql": {err}')
        try:
            return(df)
        except UnboundLocalError as errUnbound:
            write_log(
                file_name, f'Keine Nummer eingegeben! In Methode "df_schaetzung_sql": {errUnbound}')
    else:
        write_log(file_name, 'Connection-Fehler in Methode "df_schaetzung_sql" ')

# 8 SQL Ansprechpartner
##########################################################################################################################
def ansprechpartner_sql(id_kunde):
    if connection:
        write_log(file_name, 'connection ok: ansprechpartner_sql_test')

        # Variablen
        fields_ansprechpartner_list = [
            'Anrede', 'Vorname', 'Nachname', 'Mail', 'Tel', 'Mobil', 'Fax', 'Serienbrief']
        cursor = connection.cursor()
        sql_return_list = []
        k = 0 
        # SQL-Abfrage und Formatierung der Rückgabe in DataFrame
        try:
            cursor.execute(f"""SELECT isnull(a.s_abbreviation,'-') AS Anrede,
                                isnull(c.s_name2,'-') AS Vorname,
                                isnull(c.s_name1,'-') AS Nachname,
                                isnull(c.s_email,'-') AS Mail,
                                isnull(c.s_phone,'-') AS Tel,
                                isnull(c.s_cellular,'-') AS Mobil,
                                isnull(c.s_fax,'-') AS Fax,
                                (CASE 	WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Buchhaltung%'  
                                                            AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Geschäftsführung%'  
                                                            AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Lohnabrechnung%')
                                            THEN  'Fibu, GF, Lohn'
                                            WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Buchhaltung%'  
                                                            AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Geschäftsführung%'  
                                                            AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Lohnabrechnung%')
                                            THEN  'Fibu, GF'				
                                            WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Buchhaltung%'  
                                                            AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Geschäftsführung%'  
                                                            AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Lohnabrechnung%')
                                            THEN  'Fibu'
                                            WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Buchhaltung%'  
                                                            AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Geschäftsführung%'  
                                                            AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)   LIKE '%AP Lohnabrechnung%')
                                            THEN  'Fibu, Lohn'							
                                            WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Buchhaltung%'  
                                                            AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Geschäftsführung%'  
                                                            AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Lohnabrechnung%')
                                            THEN  'GF'				
                                            WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Buchhaltung%'  
                                                            AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Geschäftsführung%'  
                                                            AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Lohnabrechnung%')
                                            THEN  'Lohn'
                                            WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Buchhaltung%'  
                                                            AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Geschäftsführung%'  
                                                            AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Lohnabrechnung%')
                                            THEN  'GF, Lohn'							
                                            ELSE	''	
                                END) AS Serienbrief
                            FROM 
                                sao.contact_m c,
                                sao.formofaddr_p a

                            WHERE 
                                c.i_customer_m = {int(id_kunde)}
                                AND c.dt_deleted IS NULL
                                AND c.b_notactive = 0
                                AND a.i_formofaddr_p = c.i_formofaddr_p
                                AND (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Buchhaltung%' 
                                        OR sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Geschäftsführung%' 
                                        OR sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Lohnabrechnung%')""")

            sql_return = cursor.fetchall()
            k = 0
            content = []
            if len(sql_return) > 0:
                for row in sql_return:
                    while k < len(fields_ansprechpartner_list):
                        sql_return_list.append(str(row[k]))
                        k += 1
                    content.append(row)
                array = np.array(content)
                df = pd.DataFrame(array, columns=fields_ansprechpartner_list)

                df['Mail'] = "[" + \
                df['Mail'].astype(str)+"](mailto:" + \
                df['Mail'].astype(str) + ")"

                df['Tel'] = "[" + \
                df['Tel'].astype(str)+"](tel:" + \
                df['Tel'].astype(str) + ")"
            else:
                if (num_idKunde is None):
                    sql_return_list = ["Kein:e Ansprechpartner:in", '', '', '', '', '', '', '']
                else:
                    sql_return_list = ['Kein:e Ansprechpartner:in', '', '', '', '', '', '', '']
                array = dict(zip(fields_ansprechpartner_list, sql_return_list))
                df = pd.DataFrame(array, index=[0])

        except pyodbc.DatabaseError as err:
            write_log(
                file_name, f'Datenbankfehler in Methode "ansprechpartner_sql_test": {err}')
        try:
            return(df)
        except UnboundLocalError as errUnbound:
            write_log(
                file_name, f'Keine Nummer eingegeben! In Methode "ansprechpartner_sql_test": {errUnbound}')
    else:
        write_log(
            file_name, 'Connection-Fehler in Methode "ansprechpartner_sql_test" ')

# 9.1 SQL Rechnungsempfaenger Kontakt
##########################################################################################################################
def rechnungKontakt_sql(id_kunde):
    if connection:
        write_log(file_name, 'connection ok: rechnungKontakt_sql')

        # Variablen
        fields_REkontakt_list = [
            'Anrede', 'Vorname', 'Nachname', 'Mail', 'Tel', 'Mobil', 'Fax', 'Serienbrief']
        cursor = connection.cursor()
        sql_return_list = []
        k = 0

        # SQL-Abfrage und Formatierung der Rückgabe in DataFrame
        try:
            cursor.execute(f"""SELECT top 1
                            a.s_abbreviation AS Anrede,
                            c.s_name2 AS Vorname,
                            c.s_name1 AS Nachname,
                            c.s_email AS Mail,
                            c.s_phone AS Tel,
                            c.s_cellular AS Mobil,
                            c.s_fax AS Fax,
                            (CASE 	WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Buchhaltung%'  
                                                        AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Geschäftsführung%'  
                                                        AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Lohnabrechnung%')
                                        THEN  'Fibu, GF, Lohn'
                                        WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Buchhaltung%'  
                                                        AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Geschäftsführung%'  
                                                        AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Lohnabrechnung%')
                                        THEN  'Fibu, GF'				
                                        WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Buchhaltung%'  
                                                        AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Geschäftsführung%'  
                                                        AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Lohnabrechnung%')
                                        THEN  'Fibu'
                                        WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Buchhaltung%'  
                                                        AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Geschäftsführung%'  
                                                        AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)   LIKE '%AP Lohnabrechnung%')
                                        THEN  'Fibu, Lohn'							
                                        WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Buchhaltung%'  
                                                        AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Geschäftsführung%'  
                                                        AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Lohnabrechnung%')
                                        THEN  'GF'				
                                        WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Buchhaltung%'  
                                                        AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Geschäftsführung%'  
                                                        AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Lohnabrechnung%')
                                        THEN  'Lohn'
                                        WHEN (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  NOT LIKE '%AP Buchhaltung%'  
                                                        AND sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Geschäftsführung%'  
                                                        AND  sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Lohnabrechnung%')
                                        THEN  'GF, Lohn'							
                                        ELSE	''	
                            END) AS Serienbrief
                        FROM 
                            sao.contact_m c,
                            sao.formofaddr_p a

                        WHERE 
                            c.i_customer_m =  (SELECT (SELECT c1.i_customer_m FROM sao.customer_m c1 WHERE a1.i_customer_m_01 = c1.i_customer_m) AS Kdnr
                                                    FROM
                                                            sao.address_m a1
                                                    WHERE 
                                                            a1.i_customer_m = {int(id_kunde)}
                                                            AND a1.dt_deleted IS NULL
                                                            AND a1.b_defaults = 1
                                                            AND sao.GetTypeLookupValues('address_m','n_type', 2,a1.n_type, '', 0) LIKE '%rechnungsadresse%') 
                            AND c.dt_deleted IS NULL
                            AND c.b_notactive = 0
                            AND a.i_formofaddr_p = c.i_formofaddr_p
                            AND (sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Buchhaltung%' 
                                    OR sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Geschäftsführung%' 
                                    OR sao.GetTypeLookupValues('contact_m','n_optional7', 2,c.n_optional7, '', 0)  LIKE '%AP Lohnabrechnung%')""")

            sql_return = cursor.fetchall()
            k = 0
            if len(sql_return) > 0:
                for row in sql_return:
                    while k < len(fields_REkontakt_list):
                        sql_return_list.append(str(row[k]))
                        k += 1
            else:
                if (num_idKunde is None):
                    sql_return_list = [
                        "Keine Rechnungsempfänger:innen", '', '', '', '', '', '', '']
                else:
                    sql_return_list = [
                        'Keine Rechnungsempfänger:innen', '', '', '', '', '', '', '']

            dictionary_kontakt = dict(
                zip(fields_REkontakt_list, sql_return_list))
            df = pd.DataFrame(dictionary_kontakt, index=[0])

        except pyodbc.DatabaseError as err:
            write_log(
                file_name, f'Datenbankfehler in Methode "rechnungKontakt_sql": {err}')
        try:
            return(df)
        except UnboundLocalError as errUnbound:
            write_log(
                file_name, f'Keine Nummer eingegeben! In Methode "rechnungKontakt_sql": {errUnbound}')
    else:
        write_log(file_name, 'Connection-Fehler in Methode "rechnungKontakt_sql" ')

# 9.2 SQL Rechnungsempfaenger Adresse
##########################################################################################################################
def rechnungAdresse_sql(id_kunde):
    if connection:
        write_log(file_name, 'connection ok: rechnungAdresse_sql')

        # Variablen
        fields_REkontakt_list = ['Name1', 'Name2',
                                 'Strasse', 'Ort', 'PLZ', 'Kdnr']
        cursor = connection.cursor()
        sql_return_list = []
        k = 0

        # SQL-Abfrage und Formatierung der Rückgabe in DataFrame
        try:
            cursor.execute(f"""SELECT 
                                a.s_name1 AS name1,
                                a.s_name2 AS name2,
                                a.s_street as Strasse,
                                a.s_town as Ort,
                                a.s_zipcodeno as PLZ,
                                (SELECT c.s_custno FROM sao.customer_m c WHERE a.i_customer_m_01 = c.i_customer_m) AS Kdnr
                            FROM
                                sao.address_m a
                            WHERE 
                                a.i_customer_m = {int(id_kunde)}
                                AND a.dt_deleted IS NULL
                                AND a.b_defaults = 1
                                AND sao.GetTypeLookupValues('address_m','n_type', 2,a.n_type, '', 0) LIKE '%rechnungsadresse%'""")

            sql_return = cursor.fetchall()
            k = 0

            if len(sql_return) > 0:
                for row in sql_return:
                    while k < len(fields_REkontakt_list):
                        sql_return_list.append(str(row[k]))
                        k += 1
            else:
                if (num_idKunde is None) and num_idKunde == 0:
                    sql_return_list = [
                        "Keine Adresse", '', '', '', '', '', '', '']
                else:
                    sql_return_list = [
                        'Keine Adresse', '', '', '', '', '', '', '']

            dictionary_kontakt = dict(
                zip(fields_REkontakt_list, sql_return_list))
            df = pd.DataFrame(dictionary_kontakt, index=[0])

        except pyodbc.DatabaseError as err:
            write_log(
                file_name, f'Datenbankfehler in Methode "aktueller_status": {err}')

        try:
            return(df)
        except UnboundLocalError as errUnbound:
            write_log(
                file_name, f'Keine Nummer eingegeben! In Methode "aktueller_status": {errUnbound}')
    else:
        write_log(file_name, 'Connection-Fehler in Methode "aktueller_status" ')

# 10.1 SQL Anlieferung Bewegungsdaten
##########################################################################################################################
def bewegungsdaten_sql(id_kunde):
    if connection:
        write_log(file_name, 'connection ok: bewegungsdaten_sql')

        # Variablen
        fields_Auto_list = ['Zahl', 'Ausp', 'Merkmal', 'Info', 'I_CRITERION']
        cursor = connection.cursor()
        sql_return_list = []
        k = 0 
        # SQL-Abfrage und Formatierung der Rückgabe in DataFrame
        try:
            cursor.execute(f"""SELECT 
                                isnull(df.n_featurenumber,0) AS Zahl,
                                c.s_description AS Ausp,
                                f.s_description AS Merkmal,
                                df.ls_featuremtext AS Info,
                                df.I_CRITERION
                            FROM 
                                sao.devfeature df,
                                sao.criterion c,
                                sao.feature f 
                            WHERE 
                                df.i_customer_m = {int(id_kunde)} 
                                AND df.dt_deleted IS NULL
                                AND df.i_criterion = c.i_criterion
                                AND f.i_feature = df.i_feature
                                AND f.i_feature in(55)
                                and df.I_CRITERION in(232,227)
                            ORDER BY f.s_optional2""")

            sql_return = cursor.fetchall()
            k = 0
            content = []
            if len(sql_return) > 0:
                for row in sql_return:
                    while k < len(fields_Auto_list):
                        sql_return_list.append(str(row[k]))
                        k += 1
                    content.append(row)
                array = np.array(content)
                df = pd.DataFrame(array, columns=fields_Auto_list)
            else:
                if (num_idKunde is None):
                    sql_return_list = ['0', '', '', '', '']
                else:
                    sql_return_list = ['0', '', '', '', '']
                array = dict(zip(fields_Auto_list, sql_return_list))
                df = pd.DataFrame(array, index=[0])

        except pyodbc.DatabaseError as err:
            write_log(
                file_name, f'Datenbankfehler in Methode "bewegungsdaten_sql": {err}')
        try:
            return(df)
        except UnboundLocalError as errUnbound:
            write_log(
                file_name, f'Keine Nummer eingegeben! In Methode "bewegungsdaten_sql": {errUnbound}')
    else:
        write_log(
            file_name, 'Connection-Fehler in Methode "bewegungsdaten_sql" ')

# 10.2 SQL Anlieferung Stammdaten
##########################################################################################################################
def stammdaten_sql(id_kunde):
    if connection:
        write_log(file_name, 'connection ok: stammdaten_sql')

        # Variablen
        fields_Auto_list = ['Zahl', 'Ausp', 'Merkmal', 'Info', 'I_CRITERION']
        cursor = connection.cursor()
        sql_return_list = []
        k = 0 
        # SQL-Abfrage und Formatierung der Rückgabe in DataFrame
        try:
            cursor.execute(f"""SELECT 
                                isnull(df.n_featurenumber,0) AS Zahl,
                                c.s_description AS Ausp,
                                f.s_description AS Merkmal,
                                df.ls_featuremtext AS Info,
                                df.I_CRITERION
                            FROM 
                                sao.devfeature df,
                                sao.criterion c,
                                sao.feature f 
                            WHERE 
                                df.i_customer_m = {int(id_kunde)}
                                AND df.dt_deleted IS NULL
                                AND df.i_criterion = c.i_criterion
                                AND f.i_feature = df.i_feature
                                AND f.i_feature in(55)
                                and df.I_CRITERION in(231)
                            ORDER BY f.s_optional2""")

            sql_return = cursor.fetchall()
            k = 0
            content = []
            if len(sql_return) > 0:
                for row in sql_return:
                    while k < len(fields_Auto_list):
                        sql_return_list.append(str(row[k]))
                        k += 1
                    content.append(row)
                array = np.array(content)
                df = pd.DataFrame(array, columns=fields_Auto_list)

                
            else:
                if (num_idKunde is None):
                    sql_return_list = ['0', '', '', '', '']
                else:
                    sql_return_list = ['0', '', '', '', '']
                array = dict(zip(fields_Auto_list, sql_return_list))
                df = pd.DataFrame(array, index=[0])

        except pyodbc.DatabaseError as err:
            write_log(
                file_name, f'Datenbankfehler in Methode "stammdaten_sql": {err}')
        try:
            return(df)
        except UnboundLocalError as errUnbound:
            write_log(
                file_name, f'Keine Nummer eingegeben! In Methode "stammdaten_sql": {errUnbound}')
    else:
        write_log(
            file_name, 'Connection-Fehler in Methode "stammdaten_sql" ')

# 11 SQL Versand
##########################################################################################################################
def versand_sql(id_kunde):
    if connection:
        write_log(file_name, 'connection ok: versand_sql')

        # Variablen
        fields_Auto_list = ['Ausp', 'Merkmal', 'Info']
        cursor = connection.cursor()
        sql_return_list = []
        k = 0 
        # SQL-Abfrage und Formatierung der Rückgabe in DataFrame
        try:
            cursor.execute(f"""SELECT 
                                c.s_description AS Ausp,
                                f.s_description AS Merkmal,
                                df.ls_featuremtext AS Info
                            FROM 
                                sao.devfeature df,
                                sao.criterion c,
                                sao.feature f 
                            WHERE 
                                df.i_customer_m = {int(id_kunde)} 
                                AND df.dt_deleted IS NULL
                                AND df.i_criterion = c.i_criterion
                                AND f.i_feature = df.i_feature
                                AND f.i_feature in(54)
                            ORDER BY f.s_optional2""")

            sql_return = cursor.fetchall()
            k = 0
            content = []
            if len(sql_return) > 0:
                for row in sql_return:
                    while k < len(fields_Auto_list):
                        sql_return_list.append(str(row[k]))
                        k += 1
                    content.append(row)
                array = np.array(content)
                df = pd.DataFrame(array, columns=fields_Auto_list)
            else:
                if (num_idKunde is None):
                    sql_return_list = [ '', '', '']
                else:
                    sql_return_list = [ '', '', '']
                array = dict(zip(fields_Auto_list, sql_return_list))
                df = pd.DataFrame(array, index=[0])

        except pyodbc.DatabaseError as err:
            write_log(
                file_name, f'Datenbankfehler in Methode "versand_sql": {err}')
        try:
            return(df)
        except UnboundLocalError as errUnbound:
            write_log(
                file_name, f'Keine Nummer eingegeben! In Methode "versand_sql": {errUnbound}')
    else:
        write_log(
            file_name, 'Connection-Fehler in Methode "versand_sql" ')
