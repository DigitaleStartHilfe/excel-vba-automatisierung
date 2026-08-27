"""
Excel + Python Automatisierung
Projekt: Automatischer E-Mail-Versand (Demo 3)
Erstellt: August 2026
"""

import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import datetime

# 1. Excel-Report erstellen
daten = {
    "Monat": ["Januar", "Februar", "März"],
    "Umsatz": [12500, 14800, 13200]
}
df = pd.DataFrame(daten)
dateiname = f"Monatsreport_{datetime.datetime.now().strftime('%Y-%m-%d')}.xlsx"
df.to_excel(dateiname, index=False)
print(f"📊 Report erstellt: {dateiname}")

# 2. E-Mail versenden (SMTP-Konfiguration notwendig!)
def send_report(empfaenger):
    msg = MIMEMultipart()
    msg["From"] = "ihre-email@gmail.com"
    msg["To"] = empfaenger
    msg["Subject"] = "Monatsreport"
    msg.attach(MIMEText("Ihr Report im Anhang.", "plain"))
    
    with open(dateiname, "rb") as datei:
        teil = MIMEBase("application", "octet-stream")
        teil.set_payload(datei.read())
        encoders.encode_base64(teil)
        teil.add_header("Content-Disposition", f"attachment; filename={dateiname}")
        msg.attach(teil)
    
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login("ihre-email@gmail.com", "app-passwort")
    server.send_message(msg)
    server.quit()
    print("✅ E-Mail erfolgreich versendet!")

send_report("empfaenger@beispiel.de")
