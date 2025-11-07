# Kundenüberwachungsplattform / Customer Monitoring Platform

Dieses Repository enthält eine strukturierte Lösung zur Überwachung von Kundenprojekten für ISO 9001 Compliance, implementiert als Excel-Vorlage mit intelligenten Dropdown-Menüs und dynamischen Feldern.

This repository contains a structured solution for monitoring customer projects for ISO 9001 compliance, implemented as an Excel template with intelligent dropdown menus and dynamic fields.

## 📁 Dateien / Files

- **`Kundenüberwachung_Vorlage.xlsx`** - Die fertige Excel-Vorlage / The ready-to-use Excel template
- **`DOKUMENTATION.md`** - Ausführliche deutsche Dokumentation / Comprehensive German documentation
- **`create_template.py`** - Python-Script zur Generierung der Vorlage / Python script to generate the template

## 🚀 Schnellstart / Quick Start

### Deutsch

1. Öffnen Sie `Kundenüberwachung_Vorlage.xlsx` in Excel oder laden Sie es in Google Sheets hoch
2. Lesen Sie das Blatt "Anleitung" für eine Kurzanleitung
3. Konsultieren Sie `DOKUMENTATION.md` für detaillierte Informationen

### English

1. Open `Kundenüberwachung_Vorlage.xlsx` in Excel or upload to Google Sheets
2. Read the "Anleitung" (Instructions) sheet for a quick guide
3. Consult `DOKUMENTATION.md` for detailed information (in German)

## 📋 Funktionen / Features

### Deutsch

- **Fixierte Spalten**: Kunden-Nr., Kundenname, Status
- **Dynamische Leistungskategorien**: Verschiedene Servicekategorien wählbar
- **Intelligente Dropdowns**: Automatische Anpassung der Feldoptionen basierend auf Kategorie
- **Konfigurierbar**: Einfaches Hinzufügen neuer Kategorien und Optionen
- **ISO 9001 konform**: Strukturierte Dokumentation und Nachverfolgbarkeit

### English

- **Fixed columns**: Customer number, customer name, status
- **Dynamic service categories**: Various service categories selectable
- **Intelligent dropdowns**: Automatic adjustment of field options based on category
- **Configurable**: Easy addition of new categories and options
- **ISO 9001 compliant**: Structured documentation and traceability

## 🏗️ Struktur / Structure

Die Vorlage besteht aus 5 Blättern / The template consists of 5 sheets:

1. **Kundenübersicht** - Hauptblatt für Kundendaten / Main sheet for customer data
2. **Kategorien** - Verfügbare Leistungskategorien / Available service categories
3. **Feldkonfiguration** - Feldnamen und Optionen pro Kategorie / Field names and options per category
4. **Status_Optionen** - Verfügbare Statuswerte / Available status values
5. **Anleitung** - Schnellanleitung / Quick guide

## 🔄 Vorlage neu generieren / Regenerate Template

Falls Sie die Vorlage anpassen und neu generieren möchten / If you want to customize and regenerate the template:

```bash
# Install dependencies / Abhängigkeiten installieren
pip install openpyxl

# Run the generator / Generator ausführen
python3 create_template.py
```

## 📖 Kategorien / Categories

Die Vorlage enthält folgende vordefinierte Kategorien / The template includes the following predefined categories:

- **Bild-Aufnahme** / Image Capture
- **Drohnenshow** / Drone Show
- **3D-Modell erstellen** / 3D Model Creation
- **Video-Produktion** / Video Production
- **Inspektion** / Inspection
- **Vermessung** / Surveying

Jede Kategorie hat eigene Felder mit spezifischen Dropdown-Optionen / Each category has its own fields with specific dropdown options.

## 🛠️ Anpassung / Customization

### Neue Kategorie hinzufügen / Add New Category

1. Öffnen Sie das Blatt "Kategorien" / Open the "Kategorien" sheet
2. Fügen Sie die neue Kategorie zur Liste hinzu / Add the new category to the list
3. Öffnen Sie das Blatt "Feldkonfiguration" / Open the "Feldkonfiguration" sheet
4. Fügen Sie eine Zeile mit Felddefinitionen hinzu / Add a row with field definitions

### Feldoptionen ändern / Change Field Options

1. Öffnen Sie das Blatt "Feldkonfiguration" / Open the "Feldkonfiguration" sheet
2. Bearbeiten Sie die Optionen (getrennt durch Semikolon) / Edit the options (separated by semicolon)
3. Optionen werden sofort in den Dropdowns verfügbar / Options become immediately available in dropdowns

## 📊 Verwendung mit Google Sheets / Use with Google Sheets

Die Vorlage ist kompatibel mit Google Sheets / The template is compatible with Google Sheets:

1. Laden Sie die .xlsx-Datei in Google Drive hoch / Upload the .xlsx file to Google Drive
2. Google Sheets konvertiert sie automatisch / Google Sheets converts it automatically
3. Alle Dropdowns und Validierungen funktionieren / All dropdowns and validations work

Für erweiterte Funktionen können Sie Google Apps Script verwenden (siehe DOKUMENTATION.md) / For advanced features, you can use Google Apps Script (see DOKUMENTATION.md)

## 📝 Lizenz / License

Dieses Projekt ist open source und frei verwendbar / This project is open source and free to use.

## 🤝 Beitragen / Contributing

Verbesserungsvorschläge und Pull Requests sind willkommen! / Suggestions for improvement and pull requests are welcome!

---

**Version**: 1.0  
**Erstellt für**: ISO 9001 Compliance Monitoring  
**Created for**: ISO 9001 Compliance Monitoring