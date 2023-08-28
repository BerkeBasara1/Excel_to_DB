import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
import os
import pyodbc
sheet_list = ['Geçici Araç Durum Özeti', 'Geçici Araç Tahsis Takip 2023', 'Lena BODRUM Geçici Araç', 'Geçici Araç Tahsis Takip']
table_list = ['GeciciAracDurumOzeti', 'GeciciAracTahsisi2023', 'LenaBodrumGeciciArac', 'GeciciAracTahsisTakip']

for i in range(len(sheet_list)):
    excel = pd.read_excel(r'C:\Users\berke\OneDrive - Kadir Has University\Masaüstü\Yüce Auto Geçici Araç Tahsis Takip Listesi 2023.xlsx', sheet_name=sheet_list[i])

    # Düzeltilmiş sütun adlarını oluştur
    cleaned_columns = [col.replace("\n", "_") for col in excel.columns]
    excel.columns = cleaned_columns

    db_engine = create_engine('mssql+pyodbc://@BERKEBASARA\SQLEXPRESS/berke?trusted_connection=yes&driver=ODBC+Driver+17+for+SQL+Server')

    non_empty_columns = [col for col in excel.columns if not excel[col].isnull().all()]

    excel_filtered = excel[non_empty_columns]

    # Drop the "Kritik" column if it exists
    if "Kritik" in excel_filtered.columns:
        excel_filtered = excel_filtered.drop(columns=["Kritik"])

    # Exclude the first column from the filtered DataFrame
    excel_filtered = excel_filtered.iloc[:, 1:]

    # Drop rows with all null values
    excel_filtered = excel_filtered.dropna(how='all')

    # Add a timestamp column with the current date and time
    excel_filtered['Timestamp'] = datetime.now()

    excel_filtered.to_sql(table_list[i], db_engine, if_exists='replace', index=False)

folder_name = "my_folder"  # Change this to the desired folder name
folder_path = r"C:\Users\berke\OneDrive - Kadir Has University\Masaüstü"

# Check if the folder already exists
if not os.path.exists(folder_path):
    os.mkdir(folder_path)
    print(f"Folder '{folder_name}' created successfully.")
else:
    print(f"Folder '{folder_name}' already exists.")
