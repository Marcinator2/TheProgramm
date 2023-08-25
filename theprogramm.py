
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import datetime
from datetime import datetime
from datetime import datetime, timedelta

# Annahme: Du hast zwei Zeitpunkte im Format "%Y-%m-%d %H:%M:%S"
time_format = "%H:%M:%S"



now = datetime.now()
Start_time = now.strftime("%H:%M:%S")
print("Skript wird gestartet:", Start_time)

# Annahme: Du hast eine CSV-Datei, die du einlesen möchtest
csv_file_path = "./BackprogrammeProGruppe_german.csv"


# Initialisiere die Header-Variable
header_row = None

# Lade die CSV-Datei und finde die Zeile mit dem Header
try:
    with open(csv_file_path, "r") as file:
        # Überspringe Zeilen, bis die Zeile mit dem Header gefunden wird
        for line in file:
            if line.strip() == "Filialname1,Geraetename,BackprogrammName,Start,Differenz":
                header_row = line.strip()
                break
        
        # Lese den Rest der Datei in einen DataFrame ein
        df = pd.read_csv(file, header=None, skiprows=1)
        
    # Setze den Header des DataFrames auf die gespeicherte Header-Zeile
    df.columns = header_row.split(',')
    
    # print("DataFrame nach Überspringen bis zum Header:")
    # print(df)
    
    # Hier kannst du den DataFrame weiterverarbeiten oder abspeichern
    df_filtered = df[df['Differenz'] > 0]

# Sortiere den gefilterten DataFrame nach der Spalte 'Differenz' absteigend
    df_sorted = df_filtered.sort_values(by='Differenz', ascending=False)

# Entferne Zeilen, die Duplikate sind (ohne Berücksichtigung von 'Differenz')
    df_no_duplicates = df_sorted.drop_duplicates(subset=df_sorted.columns.difference(['Differenz']))

# Zähle die Anzahl der Backprogramme
    backprogram_counts = df_no_duplicates['BackprogrammName'].value_counts()

# Sortiere die Zählungen in absteigender Reihenfolge
    sorted_counts = backprogram_counts.sort_values(ascending=False)

# Plotten des Donut-Diagramms
    fig, ax = plt.subplots()

# Erstelle eine Liste von Farben für die Sektoren
    colors = plt.cm.viridis_r(sorted_counts / float(sum(sorted_counts)))

#  # Iteriere durch die Index- und Wertepaare von sorted_counts
#     for idx, (backprogram, count) in enumerate(sorted_counts.items()):
#           ax.pie([count], labels=[backprogram], colors=[colors[idx]], autopct='%1.1f%%', startangle=140, pctdistance=0.85, wedgeprops=dict(width=0.4))

# # Füge einen Kreis in der Mitte hinzu, um ein Donut-Diagramm zu erstellen
#     centre_circle = plt.Circle((0,0),0.70,fc='white')
#     fig.gca().add_artist(centre_circle)



    excel_file_path = './output_data.xlsx'


    # Erstelle einen Excel-Writer
    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
        # Schreibe den DataFrame in ein Excel-Arbeitsblatt mit Header
        df_no_duplicates.to_excel(writer, sheet_name='DataFrame', index=False, startrow=2, header=True)
        # Hole das Arbeitsblatt-Objekt
        worksheet = writer.sheets['DataFrame']
        # Erstelle einen Stil-Objekt für den Header
        header_format = writer.book.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'border': 1})
         # Schreibe die Spaltennamen in den Header
        for col_num, value in enumerate(df_no_duplicates.columns.values):
             worksheet.write(0, col_num, value, header_format)

    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    
    # Wandle die Zeitpunkte in datetime-Objekte um
    start_time = datetime.strptime(Start_time, time_format)
    end_time = datetime.strptime(current_time, time_format)

    # Berechne die Differenz zwischen den Zeitpunkten
    time_difference = end_time - start_time

    # Gib die Differenz aus
    print(f"Skript wird in excel gespeichert. Bischerige Dauer: {time_difference}")


##############################################

# Plotten des einfachen Kuchendiagramms
    plt.pie(backprogram_counts, labels=backprogram_counts.index, autopct='%1.1f%%', startangle=140)

# Zeige das Diagramm an
    ax.axis('off')  # Stellt sicher, dass das Diagramm rund ist
    plt.title('Verteilung der Backprogramme')
    #plt.show()


    # Speichere den Plot als Bild
    plot_image_path = './plot_image.png'
    plt.savefig(plot_image_path, format='png')
   # plt.close()

    # Füge das Plot-Bild in eine Excel-Datei ein
    excel_with_plot_path = './output_data.xlsx'#'./output_with_plot.xlsx'
    wb = Workbook()
    ws = wb.active

    # Füge den DataFrame in das Excel-Arbeitsblatt ein
    for r_idx, row in enumerate(df_no_duplicates.values):
        for c_idx, value in enumerate(row):
            ws.cell(row=r_idx + 1, column=c_idx + 1, value=value)

    # Füge das Plot-Bild in das Excel-Arbeitsblatt ein
    img = Image(plot_image_path)
    ws.add_image(img, 'G1')
    # Speichere die Excel-Datei mit DataFrame und Plot
    


    #print("Skript wird in excel gespeichert. Bischerige Dauer:", type(Dauer))
    
    wb.save(excel_with_plot_path)
    
    print(f"Excel-Datei mit DataFrame und Plot wurde in '{excel_with_plot_path}' gespeichert.")



except FileNotFoundError:
    print(f"Datei '{csv_file_path}' nicht gefunden.")
except Exception as e:
    print(f"Ein Fehler ist aufgetreten: {e}")