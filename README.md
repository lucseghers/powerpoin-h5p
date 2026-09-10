# PowerPoint → H5P Course Presentation (Streamlit)

## Starten

Open een terminal in deze map en voer uit:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Daarna opent de app normaal automatisch in de browser.

## Werking

1. Upload een `.pptx`.
2. Standaard wordt `templates/default_template.h5p` gebruikt.
3. Wil je voor één omzetting een ander template gebruiken, vink dan
   **Een ander H5P-template gebruiken** aan en upload een `.h5p`.
4. Klik op **PowerPoint omzetten naar H5P**.
5. Download het aangemaakte `.h5p`-bestand.

## Standaardtemplate permanent wijzigen

Vervang:

`templates/default_template.h5p`

door een ander H5P Course Presentation-bestand en behoud dezelfde bestandsnaam.

## Momenteel ondersteund

- PowerPoint-tekstvakken en titels
- afbeeldingen
- positie en afmetingen
- basisopmaak van tekst

## Nog niet ondersteund

- SmartArt
- tabellen
- grafieken
- gewone PowerPoint-vormen
- animaties en overgangen
