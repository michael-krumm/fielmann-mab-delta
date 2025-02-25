import pandas as pd
import os
import numpy as np
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

# Pfade zu den Excel-Dateien
data_dir = 'survey_2024_2023'
file_2023 = os.path.join(data_dir, 'Mitarbeiterbefragung 2023-T1911-20230702-results-total.xlsx')
file_2024 = os.path.join(data_dir, 'Mitarbeiterbefragung 2024-F2505-Ergebnisexport-Gesamtergebnisse-2024-06-23T235959.xlsx')
output_file = 'Fielmann_MAB_Delta_2023_2024.xlsx'

def analyze_employee_surveys():
    print("Lese Excel-Dateien...")
    try:
        # Excel-Dateien laden
        excel_2023 = pd.ExcelFile(file_2023)
        excel_2024 = pd.ExcelFile(file_2024)
        
        # Verfügbare Blätter anzeigen
        print(f"Verfügbare Blätter in 2023: {excel_2023.sheet_names}")
        print(f"Verfügbare Blätter in 2024: {excel_2024.sheet_names}")
        
        # Für beide Dateien ist das Blatt 'M' zu verwenden
        sheet_2023 = 'M'
        sheet_2024 = 'M'
        
        # Laden der Daten
        df_2023 = pd.read_excel(file_2023, sheet_name=sheet_2023)
        df_2024 = pd.read_excel(file_2024, sheet_name=sheet_2024)
        
        # Identifiziere alle Items in beiden Datensätzen
        # Die Items sind alle Spalten, die weder die initialen Metadaten-Spalten noch die letzte Spalte sind
        meta_columns_2023 = ['Sort', 'ID', 'Name', 'Level 1', 'Level 2', 'Level 3', 'Level 4', 'Level 5', 
                            'Level 6', 'Level 7', 'Level 8', 'Level 9', 'Level 10', 'Level 11', 'Typ', 'n', 'N']
        meta_columns_2024 = ['Sort', 'ID', 'Name', 'Level 1', 'Level 2', 'Level 3', 'Level 4', 'Level 5', 
                            'Level 6', 'Level 7', 'Level 8', 'Level 9', 'Level 10', 'Level 11', 'Level 12', 'Typ', 'n', 'N']
        
        # Extrahiere Items aus beiden Datensätzen
        items_2023 = [col for col in df_2023.columns if col not in meta_columns_2023]
        items_2024 = [col for col in df_2024.columns if col not in meta_columns_2024]
        
        # Entferne die letzte Spalte von 2023, da es die Umfrage-Frage ist
        if "The staff survey (2022) led to positive changes in my working environment." in items_2023:
            items_2023.remove("The staff survey (2022) led to positive changes in my working environment.")
            
        # Entferne spezielle Spalten aus 2024, die keine direkten Items sind
        special_columns_2024 = ["Information", "Decision making", "Learning from mistakes", "Collaboration", "Leadership", 
                                "The employee survey (2023) led to positive changes in my working environment."]
        items_2024 = [item for item in items_2024 if item not in special_columns_2024]
        
        # Finde gemeinsame Items
        common_items = set(items_2023).intersection(set(items_2024))
        print(f"Anzahl gemeinsamer Items: {len(common_items)}")
        
        # Erstelle einen neuen DataFrame für die Analyse mit den gemeinsamen Items
        delta_data = []
        
        # Für jedes gemeinsame Item den Durchschnittswert und das Delta berechnen
        for item in sorted(common_items):
            # Werte für 2023 und 2024 extrahieren (erste Zeile enthält die Gesamtwerte)
            value_2023 = df_2023.iloc[0][item]
            value_2024 = df_2024.iloc[0][item]
            
            # Delta berechnen
            delta = value_2024 - value_2023
            
            # Prozentuale Veränderung berechnen
            percent_change = (delta / value_2023) * 100 if value_2023 != 0 else np.nan
            
            # Daten zur Liste hinzufügen
            delta_data.append({
                'Item': item,
                'Wert 2023': value_2023,
                'Wert 2024': value_2024,
                'Absolutes Delta': delta,
                'Prozentuale Veränderung': percent_change
            })
        
        # Erstelle den DataFrame
        delta_df = pd.DataFrame(delta_data)
        
        # Sortiere nach absolutem Delta (absteigend)
        delta_df = delta_df.sort_values(by='Absolutes Delta', ascending=False)
        
        # Formatierung für Excel-Export
        writer = pd.ExcelWriter(output_file, engine='openpyxl')
        delta_df.to_excel(writer, sheet_name='Delta Analyse', index=False)
        
        # Zugriff auf das Arbeitsblatt für Formatierung
        workbook = writer.book
        worksheet = writer.sheets['Delta Analyse']
        
        # Spaltenbreiten anpassen
        worksheet.column_dimensions['A'].width = 60  # Item-Spalte
        for col in ['B', 'C', 'D', 'E']:
            worksheet.column_dimensions[col].width = 20
        
        # Bedingte Formatierung für Delta-Spalte
        worksheet.conditional_formatting.add(
            f'D2:D{len(delta_data) + 1}',
            ColorScaleRule(
                start_type='min',
                start_color='FF9999',  # Rot für negative Werte
                mid_type='num',
                mid_value=0,
                mid_color='FFFFFF',    # Weiß für 0
                end_type='max',
                end_color='99FF99'     # Grün für positive Werte
            )
        )
        
        # Formatiere Zahlen
        for row in range(2, len(delta_data) + 2):
            for col in range(2, 6):
                cell = worksheet.cell(row=row, column=col)
                if col in [2, 3, 4]:  # Werte und Delta
                    cell.number_format = '0.00'
                elif col == 5:  # Prozentuale Veränderung
                    cell.number_format = '0.00%'
        
        # Überschriften formatieren
        header_font = Font(bold=True, size=12)
        header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        for col in range(1, 6):
            cell = worksheet.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
        
        # Rahmen für alle Zellen
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        for row in range(1, len(delta_data) + 2):
            for col in range(1, 6):
                worksheet.cell(row=row, column=col).border = thin_border
        
        # Speichern der Datei
        writer._save()
        print(f"Analyse abgeschlossen. Ergebnisse wurden in {output_file} gespeichert.")
        
    except Exception as e:
        print(f"Fehler beim Verarbeiten der Dateien: {e}")
        import traceback
        traceback.print_exc()
        
if __name__ == "__main__":
    analyze_employee_surveys() 