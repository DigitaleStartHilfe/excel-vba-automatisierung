"""
Excel + Python Automatisierung
Projekt: Web-Scraping mit Excel-Export (Demo 5)
Erstellt: August 2026
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd

# 1. Webseite aufrufen
url = "https://books.toscrape.com/"
response = requests.get(url)

# 2. HTML parsen
soup = BeautifulSoup(response.text, 'html.parser')
produkte = []

for buch in soup.select('article.product_pod'):
    titel = buch.find('h3').find('a').get('title', 'Kein Titel')
    preis = buch.find('p', class_='price_color').text if buch.find('p', class_='price_color') else 'Kein Preis'
    bewertung = buch.find('p', class_='star-rating').get('class')[1] if buch.find('p', class_='star-rating') else 'Keine Bewertung'
    produkte.append({'Titel': titel, 'Preis': preis, 'Bewertung': bewertung})

# 3. In Excel speichern
df = pd.DataFrame(produkte)
df.to_excel("produkte.xlsx", index=False)
print(f"✅ {len(produkte)} Produkte exportiert!")
