import pyodbc
import urllib
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from sqlalchemy import create_engine

# SQL Server'dan verileri alma
conn = pyodbc.connect(
    "Driver={SQL Server};"
    "Server=DESKTOP-G1HC1O3\TEST;"
    "Database=turko;"
    "Trusted_Connection=yes;"
)

# SQLAlchemy bağlantısı oluşturma
conn_str = "mssql+pyodbc:///?odbc_connect={}".format(
    urllib.parse.quote_plus(
        "DRIVER={ODBC Driver 17 for SQL Server};SERVER=DESKTOP-G1HC1O3\TEST;DATABASE=turko;Trusted_Connection=yes;"
    )
)
# Store procedure'yi çağır
cursor = conn.cursor()
cursor.execute("{CALL StokListesiNew('2023-04-30')}")

# Sonucu DataFrame'e dönüştür

cursor = conn.cursor()
params = "2022-01-01"
cursor.execute("{CALL StokListesiNew (?)}", params)
rows = cursor.fetchall()

df = pd.DataFrame.from_records(rows, columns=[desc[0] for desc in cursor.description])

# DataFrame'i Excel dosyasına yazdır
writer = pd.ExcelWriter("D:\python\stokreprosedurveri\output2.xlsx", engine="openpyxl")
df.to_excel(
    writer,
    index=False,
    # dtype={"StokAdi": str, "SubeAdi": str, "ınd": float, "StokMiktari": int},
)
writer._save()
