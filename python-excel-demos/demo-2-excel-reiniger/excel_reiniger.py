"""
Excel + Python Automatisierung
Projekt: Excel-Reiniger (Demo 2)
Erstellt: August 2026
"""

import pandas as pd
import numpy as np

# 1. Testdaten mit Fehlern erstellen
fehlerhafte_daten = {
    "Name": ["Anna", "Ben", "Anna", "Clara", "David", "Ben"],
    "Datum": ["2026-01-01", "01.02.2026", "2026-01-01", "2026-03-15", "15.04.2026", "01.02.2026"],
    "Verkauf": [1200, 850, 1200, 2100, 450, 850],
    "Mitarbeiter-ID": ["A001", "B002", "A001", "C003", "D004", "B002"]
}
df_roh = pd.DataFrame(fehlerhafte_daten)

# 2. Bereinigung
df = df_roh.drop_duplicates()
df["Name"] = df["Name"].fillna("Unbekannt")
df["Datum"] = pd.to_datetime(df["Datum"]).dt.strftime("%Y-%m-%d")

# 3. Saubere Excel speichern
df.to_excel("bereinigte_daten.xlsx", index=False)
print("✅ Bereinigte Datei gespeichert!")
