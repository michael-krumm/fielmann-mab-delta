# Fielmann Mitarbeiterbefragung - Delta-Analyse 2023/2024

Dieses Projekt enthält ein Python-Skript zur Analyse und Berechnung der Veränderungen (Deltas) zwischen den Mitarbeiterbefragungen von Fielmann aus den Jahren 2023 und 2024.

## Funktionen

- Extraktion der gemeinsamen Frage-Items aus beiden Befragungen
- Berechnung der absoluten Differenz (Delta) für jedes Item
- Berechnung der prozentualen Veränderung
- Sortierung der Items nach Größe der Veränderung
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
3. Die Ergebnisse werden in der Datei `Fielmann_MAB_Delta_2023_2024.xlsx` gespeichert

## Ausgabeformat

Die Ausgabedatei enthält folgende Spalten:
- **Item**: Die Frage aus der Mitarbeiterbefragung
- **Wert 2023**: Der Durchschnittswert aus 2023
- **Wert 2024**: Der Durchschnittswert aus 2024
- **Absolutes Delta**: Die absolute Differenz zwischen 2024 und 2023
- **Prozentuale Veränderung**: Die prozentuale Veränderung

Die Zeilen sind nach dem absoluten Delta absteigend sortiert, sodass die größten Veränderungen oben stehen.

## Hinweise

- Die Excel-Datei ist mit bedingter Formatierung versehen, sodass positive Veränderungen grün und negative Veränderungen rot dargestellt werden
- Das Skript berücksichtigt nur Items (Fragen), die in beiden Befragungen identisch vorkommen 