# Devis in sjabloon gieten

Webapp om ongeformatteerde Word-tekst automatisch in de opmaak van een
referentiedocument te gieten (lettertype, vet/cursief/onderstreping, kaders,
uitlijning van bedragen).

## Bestanden

- `app.py` — Streamlit-interface (upload, verwerk, download)
- `devis_engine.py` — kernlogica die de tekst herkent en in de opmaak giet
- `requirements.txt` — benodigde Python-pakketten

## Lokaal draaien

```bash
pip install -r requirements.txt
streamlit run app.py
```

Dit opent de app in je browser op `http://localhost:8501`.

## Online zetten via Streamlit Community Cloud (gratis)

1. Zet dit hele project (alle 3 bestanden) in een nieuwe GitHub-repository.
2. Ga naar [share.streamlit.io](https://share.streamlit.io) en meld je aan
   met je GitHub-account.
3. Klik op **New app**, kies je repository en selecteer `app.py` als
   hoofdbestand.
4. Klik **Deploy**. Na een minuut krijg je een permanente link
   (bijvoorbeeld `jouwapp.streamlit.app`) die je met de dossierbeheerster
   kan delen — zij hoeft enkel die link te openen in haar browser, geen
   installatie nodig.

## Gebruik

1. Open de app-link.
2. Upload bij **Sjabloon** het opgemaakte referentiedocument (bv. het
   Casella-devis met de juiste lay-out).
3. Upload bij **Brondocument** het nieuwe, ongeformatteerde tekstbestand.
4. Klik **Verwerk document**.
5. Klik **Download resultaat** om het nieuwe, correct opgemaakte
   Word-bestand te downloaden.

## Aanpassen aan een ander documenttype

De herkenningslogica in `devis_engine.py` is afgestemd op devis-documenten
met de structuur: titel > subtitel A/B/C > verdiepingsniveau > ruimtenaam >
Contenant/Contenu + bedrag > bullets. Voor een ander documenttype pas je
`FLOOR_PATTERNS`, `MAIN_TITLE_PATTERNS` en `classify_line()` aan.

## Troubleshooting

**"Kon deze structuur-onderdelen niet herkennen in het sjabloon"**
Het sjabloon-document moet minstens één voorbeeld bevatten van elk van deze
onderdelen: hoofdtitel, onderlijnde titel, subtitel (A./B./...),
verdiepingsniveau (REZ-DE-CHAUSSEE e.d.), een ruimtenaam die direct gevolgd
wordt door "Contenant" of "Contenu", een bedrag met €-teken, en minstens één
bullet-regel met een ':' erin. Ontbreekt een van die onderdelen in het
sjabloon, pas dan de patronen in `devis_engine.py` aan of vul het sjabloon
aan met een voorbeeld van dat onderdeel.

**De app start niet lokaal**
Controleer of je Python 3.9 of hoger gebruikt en of de installatie van
`requirements.txt` zonder fouten is verlopen (`pip install -r requirements.txt`).
