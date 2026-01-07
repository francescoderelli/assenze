# Presenze Report

Script Python che legge `Presenze_Base.xlsx` (export presenze) e genera un report Excel con:
- Ore annue per Risorsa e Tipologia (pivot)
- Foglio "percentuali presenza" basato su **ore lavorabili** (totale - CIG - festività)
- Didascalia laterale con spiegazione colonne

## Installazione
```bash
pip install -r requirements.txt
