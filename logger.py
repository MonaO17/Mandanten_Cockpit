# logger.py
# Gibt Log-Daten in Datei in Big Picture aus

# Importe
from datetime import date, datetime
file_name = "logger.py"

def write_log(file, message):
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")

    try:
        # TODO: Pfad anpassen, wo Log-Datei gespeichert werden soll!
        with open('C:\mandanten_cockpit_log.txt', 'a') as text_file:
            print(f'{dt_string} \t {file} \t {message}', file=text_file)
    except Exception as e:
        print(f"Logging funktioniert nicht! Fehler: {e}")     

