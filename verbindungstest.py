import pandas as pd

# # Hier kannst du Funktionen oder Klassen aus datei2 aufrufen
### datei2.meine_funktion()
# Define the CSV file path
def machen(csv_file_path):
   # csv_file_path = "./ConnectionTest_german.csv"  # Replace with the actual path to your CSV file

    # Read the CSV and set the header
    header_row = 2  # 0-based index of the row containing the header
    df = pd.read_csv(csv_file_path, header=header_row)

    # Save the DataFrame to an Excel file
    excel_file_path  = csv_file_path[:-3] +"xlsx"          #'./output_with_plot.xlsx'


    df.to_excel(excel_file_path, index=False)
    #print(df)