import tkinter as tk
from PIL import Image, ImageTk
from tkinter import filedialog
#import verbindungstest


def verbindungstest_machen():
    # Hier können Sie den Code für den Verbindungstest einfügen
    import verbindungstest
    file_path = filedialog.askopenfilename(title="Datei auswählen", filetypes=[("CSV-Dateien", "*.csv")])
    if file_path:
        verbindungstest.machen(file_path)
    tk.messagebox.showinfo(title=None, message="Fertig.")

def bakingreports_machen():
    # Hier können Sie den Code für die Erstellung von Baking-Reports einfügen
    import bakingreports
    file_path = filedialog.askopenfilename(title="Datei auswählen", filetypes=[("CSV-Dateien", "*.csv")])
    if file_path:
        bakingreports.machen(file_path)
    tk.messagebox.showinfo(title=None, message="Fertig.")

def open_file_dialog():
    file_path = filedialog.askopenfilename(title="Datei auswählen", filetypes=[("CSV-Dateien", "*.csv")])
    if file_path:

        verbindungstest_machen(file_path)

# GUI erstellen
root = tk.Tk()
root.title("The Programm")

# Bild laden und skalieren in Prozent
original_image = Image.open("W-Net.png")  # Pfad zum Bild einfügen
# Hier setzen Sie die gewünschte prozentuale Skalierung
scale_percent = 15  # 50% Skalierung
width = int(original_image.width * scale_percent / 100)
height = int(original_image.height * scale_percent / 100)
scaled_image = original_image.resize((width, height), Image.ANTIALIAS)
photo = ImageTk.PhotoImage(scaled_image)
image_label = tk.Label(root, image=photo)
image_label.pack()

# Buttons erstellen
verbindungstest_button = tk.Button(root, text="Verbindungstest machen", command=verbindungstest_machen)
verbindungstest_button.pack()

bakingreports_button = tk.Button(root, text="Bakingreports machen", command=bakingreports_machen)
bakingreports_button.pack()

root.mainloop()