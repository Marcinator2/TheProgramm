
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import datetime
from datetime import datetime
import os

# Aktuelles Verzeichnis ermitteln
current_path = os.getcwd()

def machen(csv_file_path):

    # Lade die CSV-Datei und finde die Zeile mit dem Header
    try:
        df = pd.read_csv(csv_file_path, header=None, sep=';')
        df.columns = ['0', 'Backprogramm', 'Zeitzone', "3", "Start", "Ende", "Ofen", "Filiale", "9", "Dauer", "11", "Gruppe"]
        # Zeichenfolge #;#-1 in Spalte 2 (Index 1 nach Löschen der ersten Spalte) ersetzen
        df["Backprogramm"] = df["Backprogramm"].str.replace('#;#-1', '')
        df["Backprogramm"] = df["Backprogramm"].str.replace('#;#1', '')
        df["Backprogramm"] = df["Backprogramm"].str.replace('#;#2', '')
        df["Backprogramm"] = df["Backprogramm"].str.replace('#;#3', '')
        df["Backprogramm"] = df["Backprogramm"].str.replace('#;#4', '')
        df["Backprogramm"] = df["Backprogramm"].str.replace('#;#5', '')
        df["Backprogramm"] = df["Backprogramm"].str.replace('#;#6', '')
        df["Backprogramm"] = df["Backprogramm"].str.replace('#;#10', '')


        # Duplikate entfernen
        df = df.drop_duplicates()

        # Erste Spalte löschen
        df = df.drop(columns=["0", "Zeitzone", "3", "Ende", "9", "11", "Gruppe"])
        df = df[["Filiale", "Ofen", "Backprogramm","Start","Dauer"]]
        # DataFrame nach der Spalte "Dauer" absteigend sortieren
        df = df.sort_values(by="Dauer", ascending=False)
        df = df.drop_duplicates(subset=df.columns[:-1])
        # Zeilen entfernen, bei denen der Wert in der Spalte "Dauer" kleiner oder gleich 0 ist
        df = df[df["Dauer"] > 0]
        df = df[df["Dauer"] < 600]
        df = df.sort_values(by="Start", ascending=True)

        # Das aktualisierte DataFrame in eine neue CSV-Datei schreiben
        #df.to_csv(os.path.join(current_path, 'bereinigte_datei.csv'), index=False, sep=';', header=True)
        #df.to_excel(os.path.join(current_path, 'bereinigte_datei.xlsx'), index=False, engine='openpyxl')
        df.to_excel(csv_file_path[:-3] +"xlsx", index=False, engine='openpyxl')


    except FileNotFoundError:
        print(f"Datei '{csv_file_path}' nicht gefunden.")
    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}")