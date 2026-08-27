"""
Excel + Python Automatisierung
Projekt: Mitarbeiter-Bonus-Rechner (Demo 1)
Erstellt: August 2026
"""

import pandas as pd

# 1. Daten erstellen
daten = {
    "Name": ["Anna", "Ben", "Clara", "David"],
    "Abteilung": ["Vertrieb", "IT", "Marketing", "Vertrieb"],
    "Gehalt": [4200, 5100, 3800, 4900]
}
df = pd.DataFrame(daten)

# 2. Bonus & Jahresgehalt berechnen
df["Bonus"] = df["Gehalt"] * 0.1
df["Jahresgehalt"] = df["Gehalt"] * 12

# 3. In Excel speichern (2 Sheets)
with pd.ExcelWriter("mitarbeiter_mit_bonus.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Alle Mitarbeiter", index=False)
    df[df["Abteilung"] == "Vertrieb"].to_excel(writer, sheet_name="Nur Vertrieb", index=False)

print("✅ Excel-Datei mit Bonus erstellt: 'mitarbeiter_mit_bonus.xlsx'")
print(df)
