#!/usr/bin/env python3
"""
Script to create the Customer Monitoring Platform Excel Template
Generates an .xlsx file with multiple sheets and data validation rules
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

def create_customer_monitoring_template(output_path=None):
    """
    Create the complete Excel template with all sheets and validation
    
    Args:
        output_path (str, optional): Path where the .xlsx file should be saved.
                                    If None, saves to current directory.
    
    Returns:
        str: Path to the created file
    """
    
    # Determine output path
    if output_path is None:
        output_path = os.path.join(os.getcwd(), "Kundenüberwachung_Vorlage.xlsx")
    
    # Create workbook
    wb = Workbook()
    
    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    # Create sheets
    ws_main = wb.create_sheet("Kundenübersicht", 0)
    ws_categories = wb.create_sheet("Kategorien", 1)
    ws_field_config = wb.create_sheet("Feldkonfiguration", 2)
    ws_status = wb.create_sheet("Status_Optionen", 3)
    
    # Setup styles
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    config_header_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    config_header_font = Font(bold=True, color="FFFFFF", size=11)
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_alignment = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # ========== Main Sheet (Kundenübersicht) ==========
    print("Creating main sheet...")
    
    # Headers for main sheet
    headers = [
        "Kunden-Nr.",
        "Kundenname",
        "Status",
        "Leistungskategorie",
        "Feld 1 (Overlap/Anzahl/Detailgrad)",
        "Feld 2 (Kamera/Dauer/Format)",
        "Feld 3 (Auflösung/Musik/Texturierung)",
        "Feld 4",
        "Feld 5",
        "Feld 6"
    ]
    
    # Write headers
    for col, header in enumerate(headers, start=1):
        cell = ws_main.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_alignment
        cell.border = thin_border
    
    # Set column widths
    column_widths = [15, 25, 20, 25, 30, 30, 30, 20, 20, 20]
    for col, width in enumerate(column_widths, start=1):
        ws_main.column_dimensions[get_column_letter(col)].width = width
    
    # Freeze first row
    ws_main.freeze_panes = "A2"
    
    # Add example data
    example_data = [
        ["K001", "Musterfirma GmbH", "In Bearbeitung", "Bild-Aufnahme", "80/80", "RGB", "45MP", "", "", ""],
        ["K002", "Beispiel AG", "Genehmigung ausstehend", "Drohnenshow", "50", "15", "Ja", "", "", ""],
        ["K003", "Test Solutions", "In Bearbeitung", "3D-Modell erstellen", "Hoch", "OBJ", "Ja", "", "", ""],
    ]
    
    for row_idx, row_data in enumerate(example_data, start=2):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws_main.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = left_alignment
            cell.border = thin_border
    
    # ========== Categories Sheet ==========
    print("Creating categories sheet...")
    
    ws_categories.cell(row=1, column=1, value="Verfügbare Leistungskategorien")
    ws_categories.cell(row=1, column=1).fill = config_header_fill
    ws_categories.cell(row=1, column=1).font = config_header_font
    ws_categories.cell(row=1, column=1).alignment = center_alignment
    
    categories = [
        "Bild-Aufnahme",
        "Drohnenshow",
        "3D-Modell erstellen",
        "Video-Produktion",
        "Inspektion",
        "Vermessung"
    ]
    
    for idx, category in enumerate(categories, start=2):
        cell = ws_categories.cell(row=idx, column=1, value=category)
        cell.border = thin_border
        cell.alignment = left_alignment
    
    ws_categories.column_dimensions['A'].width = 30
    
    # ========== Field Configuration Sheet ==========
    print("Creating field configuration sheet...")
    
    # Headers for field configuration
    field_headers = [
        "Kategorie",
        "Feld 1 Name",
        "Feld 1 Optionen",
        "Feld 2 Name",
        "Feld 2 Optionen",
        "Feld 3 Name",
        "Feld 3 Optionen",
        "Feld 4 Name",
        "Feld 4 Optionen"
    ]
    
    for col, header in enumerate(field_headers, start=1):
        cell = ws_field_config.cell(row=1, column=col, value=header)
        cell.fill = config_header_fill
        cell.font = config_header_font
        cell.alignment = center_alignment
        cell.border = thin_border
    
    # Field configuration data
    field_config = [
        [
            "Bild-Aufnahme",
            "Overlap",
            "60/60;80/80;70/70",
            "Kamera",
            "RGB;Multispektral;Thermal",
            "Auflösung",
            "20MP;45MP;100MP",
            "Flughöhe",
            "50m;100m;120m;150m"
        ],
        [
            "Drohnenshow",
            "Anzahl Drohnen",
            "10;25;50;100;200",
            "Dauer (Min)",
            "5;10;15;20;30",
            "Musik",
            "Ja;Nein",
            "Indoor/Outdoor",
            "Indoor;Outdoor;Beides"
        ],
        [
            "3D-Modell erstellen",
            "Detailgrad",
            "Niedrig;Mittel;Hoch;Sehr Hoch",
            "Format",
            "OBJ;FBX;STL;GLTF",
            "Texturierung",
            "Ja;Nein",
            "Polygonanzahl",
            "Low Poly;Medium Poly;High Poly"
        ],
        [
            "Video-Produktion",
            "Auflösung",
            "1080p;4K;6K;8K",
            "FPS",
            "24;30;60;120",
            "Länge (Min)",
            "1;3;5;10;15",
            "Nachbearbeitung",
            "Basic;Standard;Premium"
        ],
        [
            "Inspektion",
            "Objekttyp",
            "Gebäude;Brücke;Windrad;Solaranlage;Sonstiges",
            "Sensor",
            "RGB;Thermal;Multispektral",
            "Berichtstyp",
            "Standard;Detailliert;Umfassend",
            "Urgenz",
            "Normal;Hoch;Sehr Hoch"
        ],
        [
            "Vermessung",
            "Fläche",
            "< 1 ha;1-5 ha;5-10 ha;10-50 ha;> 50 ha",
            "Genauigkeit",
            "±5cm;±3cm;±2cm;±1cm",
            "Ausgabeformat",
            "DXF;SHP;KML;GeoTIFF",
            "Höhenmodell",
            "DTM;DSM;Beides"
        ]
    ]
    
    for row_idx, row_data in enumerate(field_config, start=2):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws_field_config.cell(row=row_idx, column=col_idx, value=value)
            cell.border = thin_border
            cell.alignment = left_alignment
    
    # Set column widths for field configuration
    for col in range(1, 10):
        ws_field_config.column_dimensions[get_column_letter(col)].width = 22
    
    # ========== Status Options Sheet ==========
    print("Creating status options sheet...")
    
    ws_status.cell(row=1, column=1, value="Verfügbare Status-Optionen")
    ws_status.cell(row=1, column=1).fill = config_header_fill
    ws_status.cell(row=1, column=1).font = config_header_font
    ws_status.cell(row=1, column=1).alignment = center_alignment
    
    status_options = [
        "In Bearbeitung",
        "Abgeschlossen",
        "Wartend",
        "Genehmigung ausstehend",
        "Storniert"
    ]
    
    for idx, status in enumerate(status_options, start=2):
        cell = ws_status.cell(row=idx, column=1, value=status)
        cell.border = thin_border
        cell.alignment = left_alignment
    
    ws_status.column_dimensions['A'].width = 30
    
    # ========== Add Data Validation ==========
    print("Adding data validation rules...")
    
    # Status dropdown (Column C in main sheet)
    status_dv = DataValidation(
        type="list",
        formula1="=Status_Optionen!$A$2:$A$6",
        allow_blank=True
    )
    status_dv.error = "Bitte wählen Sie einen Status aus der Liste"
    status_dv.errorTitle = "Ungültige Eingabe"
    ws_main.add_data_validation(status_dv)
    status_dv.add("C2:C1000")
    
    # Category dropdown (Column D in main sheet)
    category_dv = DataValidation(
        type="list",
        formula1="=Kategorien!$A$2:$A$7",
        allow_blank=True
    )
    category_dv.error = "Bitte wählen Sie eine Kategorie aus der Liste"
    category_dv.errorTitle = "Ungültige Eingabe"
    ws_main.add_data_validation(category_dv)
    category_dv.add("D2:D1000")
    
    # Field dropdowns (Columns E-J)
    # Note: We create example data validations for the first few rows using the field_config data
    # Excel requires comma-separated values in formula1, so we convert from semicolon-separated config
    
    # Map example rows to their categories
    example_row_categories = {
        2: "Bild-Aufnahme",
        3: "Drohnenshow",
        4: "3D-Modell erstellen"
    }
    
    for row_num, category in example_row_categories.items():
        if category in [cfg[0] for cfg in field_config]:
            # Find the configuration for this category
            category_cfg = next(cfg for cfg in field_config if cfg[0] == category)
            
            # Process each field (skip category name at index 0, then pairs of name/options)
            for field_idx in range(1, len(category_cfg), 2):
                if field_idx + 1 < len(category_cfg):
                    field_name = category_cfg[field_idx]
                    field_options = category_cfg[field_idx + 1]
                    
                    # Convert semicolon-separated to comma-separated for Excel validation
                    excel_options = field_options.replace(';', ',')
                    
                    # Calculate column (E=5, F=6, G=7, H=8, etc.)
                    col_offset = (field_idx - 1) // 2
                    col_letter = get_column_letter(5 + col_offset)
                    
                    # Create data validation
                    field_dv = DataValidation(
                        type="list",
                        formula1=f'"{excel_options}"',
                        allow_blank=True
                    )
                    field_dv.error = f"Bitte wählen Sie einen Wert für {field_name} aus der Liste"
                    field_dv.errorTitle = "Ungültige Eingabe"
                    ws_main.add_data_validation(field_dv)
                    field_dv.add(f"{col_letter}{row_num}")
    
    # ========== Add Instructions Sheet ==========
    print("Creating instructions sheet...")
    
    ws_instructions = wb.create_sheet("Anleitung", 4)
    
    instructions_text = [
        ["KUNDENÜBERWACHUNGSPLATTFORM - SCHNELLANLEITUNG", "", ""],
        ["", "", ""],
        ["Schritt 1: Neuen Kunden hinzufügen", "", ""],
        ["", "• Fügen Sie eine neue Zeile im Blatt 'Kundenübersicht' hinzu", ""],
        ["", "• Geben Sie Kunden-Nr. (Spalte A) und Kundenname (Spalte B) ein", ""],
        ["", "", ""],
        ["Schritt 2: Status wählen", "", ""],
        ["", "• Klicken Sie auf Spalte C und wählen Sie den Status aus dem Dropdown", ""],
        ["", "• Verfügbare Optionen: In Bearbeitung, Abgeschlossen, Wartend, etc.", ""],
        ["", "", ""],
        ["Schritt 3: Leistungskategorie wählen", "", ""],
        ["", "• Klicken Sie auf Spalte D und wählen Sie die Kategorie", ""],
        ["", "• Verfügbare Kategorien: Bild-Aufnahme, Drohnenshow, 3D-Modell, etc.", ""],
        ["", "", ""],
        ["Schritt 4: Kategoriefelder ausfüllen", "", ""],
        ["", "• Je nach Kategorie haben die Spalten E-J unterschiedliche Bedeutungen", ""],
        ["", "• Wählen Sie die passenden Werte aus den Dropdown-Menüs", ""],
        ["", "• Nicht relevante Felder können leer bleiben", ""],
        ["", "", ""],
        ["WICHTIGE HINWEISE:", "", ""],
        ["", "• Spalten A-C sind für alle Kunden gleich", ""],
        ["", "• Ab Spalte D sind die Felder kategoriespezifisch", ""],
        ["", "• Schauen Sie in die Spaltenüberschrift um zu sehen, welches Feld es ist", ""],
        ["", "", ""],
        ["KONFIGURATION ANPASSEN:", "", ""],
        ["", "• Neue Kategorien: Blatt 'Kategorien' bearbeiten", ""],
        ["", "• Neue Feldoptionen: Blatt 'Feldkonfiguration' bearbeiten", ""],
        ["", "• Neue Statuswerte: Blatt 'Status_Optionen' bearbeiten", ""],
        ["", "", ""],
        ["Für ausführliche Dokumentation siehe DOKUMENTATION.md", "", ""],
    ]
    
    for row_idx, row_data in enumerate(instructions_text, start=1):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws_instructions.cell(row=row_idx, column=col_idx, value=value)
            if row_idx == 1:
                cell.font = Font(bold=True, size=14, color="4472C4")
                cell.alignment = center_alignment
            elif "Schritt" in str(value) or "WICHTIGE" in str(value) or "KONFIGURATION" in str(value):
                cell.font = Font(bold=True, size=11)
            else:
                cell.alignment = left_alignment
    
    ws_instructions.column_dimensions['A'].width = 5
    ws_instructions.column_dimensions['B'].width = 70
    ws_instructions.column_dimensions['C'].width = 5
    
    # Merge title cell
    ws_instructions.merge_cells('A1:C1')
    
    # ========== Save Workbook ==========
    print(f"Saving workbook to {output_path}...")
    wb.save(output_path)
    print(f"✓ Template created successfully: {output_path}")
    
    return output_path

if __name__ == "__main__":
    print("=" * 60)
    print("Customer Monitoring Platform - Template Generator")
    print("=" * 60)
    print()
    
    try:
        output_file = create_customer_monitoring_template()
        print()
        print("=" * 60)
        print("SUCCESS!")
        print("=" * 60)
        print(f"Template file: {output_file}")
        print()
        print("Next steps:")
        print("1. Open the .xlsx file in Excel or upload to Google Sheets")
        print("2. Read the 'Anleitung' sheet for quick start guide")
        print("3. Review DOKUMENTATION.md for detailed documentation")
        print("=" * 60)
    except Exception as e:
        print(f"ERROR: Failed to create template")
        print(f"Details: {e}")
        import traceback
        traceback.print_exc()
        exit(1)
