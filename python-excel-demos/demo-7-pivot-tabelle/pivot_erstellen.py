"""
Excel + Python Automatisierung
Projekt: Excel-Pivot-Tabelle mit Python (Demo 7)
Erstellt: August 2026
"""

import pandas as pd
import xlwings as xw
import os

# 1. Testdaten erstellen
if not os.path.exists("verkaufsdaten.xlsx"):
    daten = {
        "Region": ["Ost", "Ost", "West", "West", "Ost", "West", "Ost", "West", "Ost", "West", "Ost", "West"],
        "Produkt": ["A", "B", "A", "B", "A", "B", "A", "B", "A", "B", "A", "B"],
        "Monat": ["Jan", "Jan", "Jan", "Jan", "Feb", "Feb", "Feb", "Feb", "Mär", "Mär", "Mär", "Mär"],
        "Umsatz": [100, 200, 150, 250, 120, 180, 130, 220, 110, 190, 140, 210],
        "Menge": [10, 20, 15, 25, 12, 18, 13, 22, 11, 19, 14, 21]
    }
    pd.DataFrame(daten).to_excel("verkaufsdaten.xlsx", index=False)
    print("📄 Verkaufsdaten erstellt")

# 2. Pivot-Tabelle mit xlwings erstellen
df = pd.read_excel("verkaufsdaten.xlsx")
app = xw.App(visible=False)
wb = app.books.add()

ws_daten = wb.sheets[0]
ws_daten.name = "Rohdaten"
ws_daten.range('A1').value = df

ws_pivot = wb.sheets.add("PivotAnalyse")
pivot_cache = wb.api.PivotCaches().Create(
    SourceType=1,
    SourceData=ws_daten.range('A1').expand().api
)
pivot_table = pivot_cache.CreatePivotTable(
    TableDestination=ws_pivot.range('A3').api,
    TableName="Verkaufspivot"
)

pivot_table.PivotFields("Region").Orientation = 1  # Zeilen
pivot_table.PivotFields("Produkt").Orientation = 2  # Spalten
pivot_table.AddDataField(pivot_table.PivotFields("Umsatz"), "Summe Umsatz", -4157)

wb.save("pivot_bericht.xlsx")
wb.close()
app.quit()

print("✅ Pivot-Bericht erstellt: 'pivot_bericht.xlsx'")
