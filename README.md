# Fielmann Mitarbeiterbefragung - Delta-Analyse 2023/2024

Dieses Projekt enthält ein Python-Skript zur Analyse und Berechnung der Veränderungen (Deltas) zwischen den Mitarbeiterbefragungen von Fielmann aus den Jahren 2023 und 2024.

## Funktionen

- Extraktion der gemeinsamen Frage-Items aus beiden Befragungen
- Berechnung der absoluten Differenz (Delta) für jedes Item pro Organisationseinheit
- Pivot-Style Darstellung mit einer Zeile pro Organisationseinheit
- Gruppierung der Daten nach Items in Spalten (2023, 2024, Delta)
- Export der Ergebnisse in eine formatierte Excel-Datei

## Voraussetzungen

Das Skript benötigt die folgenden Python-Bibliotheken:
- pandas
- openpyxl
- numpy

Installation:
```bash
pip install pandas openpyxl numpy
```

## Nutzung

1. Legen Sie die Excel-Dateien der Mitarbeiterbefragungen im Ordner `survey_2024_2023` ab
2. Führen Sie das Skript aus:
```bash
python mab_delta_analysis.py
```
3. Die Ergebnisse werden in der Datei `Fielmann_MAB_Delta_2023_2024_pro_OrgEinheit_Pivot.xlsx` gespeichert

## Ausgabeformat

Die Ausgabedatei hat eine Pivot-Struktur:
- **Zeilen**: Jede Zeile repräsentiert eine Organisationseinheit
- **Spalten**: Für jedes Item gibt es drei Spalten:
  - **[Item] 2023**: Der Wert für dieses Item in 2023
  - **[Item] 2024**: Der Wert für dieses Item in 2024
  - **[Item] Delta**: Die Differenz zwischen 2024 und 2023

Die Items werden in alphabetischer Reihenfolge sortiert, und die Organisationseinheiten werden ebenfalls alphabetisch sortiert.

## Hinweise

- Die Excel-Datei ist mit bedingter Formatierung versehen, sodass positive Veränderungen grün und negative Veränderungen rot dargestellt werden
- Das Skript berücksichtigt nur Items (Fragen), die in beiden Befragungen identisch vorkommen
- Es werden nur Organisationseinheiten berücksichtigt, die in beiden Befragungen vorhanden sind
- Lange Item-Namen werden auf 30 Zeichen gekürzt und mit "..." versehen, um die Lesbarkeit zu verbessern 