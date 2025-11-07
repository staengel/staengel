# Google Sheets Integration - Apps Script Guide

## Übersicht

Diese Anleitung zeigt, wie Sie die Excel-Vorlage in Google Sheets verwenden und mit Google Apps Script erweitern können, um vollautomatische Spaltenanpassungen zu implementieren.

## Excel zu Google Sheets konvertieren

### Schritt 1: Datei hochladen

1. Öffnen Sie [Google Drive](https://drive.google.com)
2. Klicken Sie auf "Neu" → "Datei-Upload"
3. Wählen Sie `Kundenüberwachung_Vorlage.xlsx`
4. Rechtsklick auf die hochgeladene Datei → "Öffnen mit" → "Google Tabellen"
5. Die Datei wird automatisch konvertiert

### Schritt 2: Formel-Anpassungen (falls nötig)

Die meisten Formeln funktionieren direkt. Falls Anpassungen nötig sind:

**Excel → Google Sheets Übersetzungen:**
- `VERGLEICH()` → `MATCH()`
- `INDIREKT()` → `INDIRECT()`
- Semikolon `;` in Listen → Komma `,`

## Apps Script für automatische Spaltenanpassung

### Grundlegendes Script

Dieses Script passt die Spaltenüberschriften automatisch an die gewählte Kategorie an.

#### Script installieren

1. Öffnen Sie Ihre Google Sheets Datei
2. Klicken Sie auf "Erweiterungen" → "Apps Script"
3. Löschen Sie den Standardcode
4. Fügen Sie folgenden Code ein:

```javascript
/**
 * Kundenüberwachungsplattform - Automatische Spaltenanpassung
 * Dieses Script passt Spaltenüberschriften und Validierungen basierend auf der gewählten Kategorie an
 */

// Konfiguration
const CONFIG = {
  MAIN_SHEET_NAME: "Kundenübersicht",
  CONFIG_SHEET_NAME: "Feldkonfiguration",
  CATEGORY_COLUMN: 4, // Spalte D
  FIRST_DYNAMIC_COLUMN: 5, // Spalte E
  HEADER_ROW: 1
};

// Feldkonfiguration für jede Kategorie
const FIELD_CONFIG = {
  "Bild-Aufnahme": [
    { name: "Overlap", options: ["60/60", "80/80", "70/70"] },
    { name: "Kamera", options: ["RGB", "Multispektral", "Thermal"] },
    { name: "Auflösung", options: ["20MP", "45MP", "100MP"] },
    { name: "Flughöhe", options: ["50m", "100m", "120m", "150m"] }
  ],
  "Drohnenshow": [
    { name: "Anzahl Drohnen", options: ["10", "25", "50", "100", "200"] },
    { name: "Dauer (Min)", options: ["5", "10", "15", "20", "30"] },
    { name: "Musik", options: ["Ja", "Nein"] },
    { name: "Indoor/Outdoor", options: ["Indoor", "Outdoor", "Beides"] }
  ],
  "3D-Modell erstellen": [
    { name: "Detailgrad", options: ["Niedrig", "Mittel", "Hoch", "Sehr Hoch"] },
    { name: "Format", options: ["OBJ", "FBX", "STL", "GLTF"] },
    { name: "Texturierung", options: ["Ja", "Nein"] },
    { name: "Polygonanzahl", options: ["Low Poly", "Medium Poly", "High Poly"] }
  ],
  "Video-Produktion": [
    { name: "Auflösung", options: ["1080p", "4K", "6K", "8K"] },
    { name: "FPS", options: ["24", "30", "60", "120"] },
    { name: "Länge (Min)", options: ["1", "3", "5", "10", "15"] },
    { name: "Nachbearbeitung", options: ["Basic", "Standard", "Premium"] }
  ],
  "Inspektion": [
    { name: "Objekttyp", options: ["Gebäude", "Brücke", "Windrad", "Solaranlage", "Sonstiges"] },
    { name: "Sensor", options: ["RGB", "Thermal", "Multispektral"] },
    { name: "Berichtstyp", options: ["Standard", "Detailliert", "Umfassend"] },
    { name: "Urgenz", options: ["Normal", "Hoch", "Sehr Hoch"] }
  ],
  "Vermessung": [
    { name: "Fläche", options: ["< 1 ha", "1-5 ha", "5-10 ha", "10-50 ha", "> 50 ha"] },
    { name: "Genauigkeit", options: ["±5cm", "±3cm", "±2cm", "±1cm"] },
    { name: "Ausgabeformat", options: ["DXF", "SHP", "KML", "GeoTIFF"] },
    { name: "Höhenmodell", options: ["DTM", "DSM", "Beides"] }
  ]
};

/**
 * Wird ausgelöst, wenn eine Zelle bearbeitet wird
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Nur im Hauptblatt und nur bei Kategorie-Spalte reagieren
  if (sheet.getName() !== CONFIG.MAIN_SHEET_NAME) return;
  if (range.getColumn() !== CONFIG.CATEGORY_COLUMN) return;
  if (range.getRow() === CONFIG.HEADER_ROW) return; // Nicht bei Header
  
  const row = range.getRow();
  const category = range.getValue();
  
  if (category && FIELD_CONFIG[category]) {
    updateRowForCategory(sheet, row, category);
  }
}

/**
 * Aktualisiert eine Zeile basierend auf der gewählten Kategorie
 */
function updateRowForCategory(sheet, row, category) {
  const fields = FIELD_CONFIG[category];
  
  if (!fields) {
    Logger.log(`Keine Konfiguration für Kategorie: ${category}`);
    return;
  }
  
  // Vorhandene Werte in dynamischen Spalten löschen
  const numDynamicCols = 6; // Spalten E-J
  const clearRange = sheet.getRange(
    row,
    CONFIG.FIRST_DYNAMIC_COLUMN,
    1,
    numDynamicCols
  );
  clearRange.clearContent();
  clearRange.clearDataValidations();
  
  // Für jedes Feld der Kategorie
  fields.forEach((field, index) => {
    const col = CONFIG.FIRST_DYNAMIC_COLUMN + index;
    const cell = sheet.getRange(row, col);
    
    // Datenvalidierung setzen
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(field.options, true)
      .setAllowInvalid(false)
      .setHelpText(`Wählen Sie ${field.name}`)
      .build();
    
    cell.setDataValidation(rule);
    
    // Optional: Tooltip hinzufügen
    cell.setNote(`${field.name}: ${field.options.join(", ")}`);
  });
  
  // Feedback
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Felder für "${category}" wurden angepasst`,
    "Kategorie aktualisiert",
    3
  );
}

/**
 * Menü-Option: Alle Zeilen aktualisieren
 */
function updateAllRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  
  const lastRow = sheet.getLastRow();
  let updatedCount = 0;
  
  // Durch alle Zeilen iterieren (ab Zeile 2, Header überspringen)
  for (let row = 2; row <= lastRow; row++) {
    const category = sheet.getRange(row, CONFIG.CATEGORY_COLUMN).getValue();
    
    if (category && FIELD_CONFIG[category]) {
      updateRowForCategory(sheet, row, category);
      updatedCount++;
    }
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${updatedCount} Zeile(n) wurden aktualisiert`,
    "Massenaktualisierung",
    3
  );
}

/**
 * Fügt ein benutzerdefiniertes Menü hinzu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 Kundenüberwachung')
    .addItem('Alle Zeilen aktualisieren', 'updateAllRows')
    .addItem('Hilfefunktionen anzeigen', 'showHelp')
    .addSeparator()
    .addItem('Neue Kategorie hinzufügen', 'showAddCategoryDialog')
    .addToUi();
}

/**
 * Zeigt Hilfe-Dialog
 */
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  const helpText = `
KUNDENÜBERWACHUNGSPLATTFORM - HILFE

Automatische Funktionen:
• Wenn Sie eine Kategorie in Spalte D wählen, werden die Felder automatisch angepasst
• Dropdown-Menüs werden für die jeweilige Kategorie erstellt
• Alte Werte werden gelöscht

Manuelle Funktionen:
• "Alle Zeilen aktualisieren": Aktualisiert alle vorhandenen Zeilen basierend auf ihrer Kategorie
• "Neue Kategorie hinzufügen": Öffnet Dialog zum Hinzufügen neuer Kategorien

Verfügbare Kategorien:
${Object.keys(FIELD_CONFIG).join('\n')}

Bei Problemen:
1. Überprüfen Sie, ob die Kategorie korrekt geschrieben ist
2. Stellen Sie sicher, dass Spalte D die Kategorie enthält
3. Verwenden Sie "Alle Zeilen aktualisieren" um Inkonsistenzen zu beheben
  `;
  
  ui.alert('Hilfe', helpText, ui.ButtonSet.OK);
}

/**
 * Dialog zum Hinzufügen neuer Kategorien
 */
function showAddCategoryDialog() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    'Neue Kategorie hinzufügen',
    'Um eine neue Kategorie hinzuzufügen:\n\n' +
    '1. Öffnen Sie das Blatt "Kategorien" und fügen Sie die Kategorie hinzu\n' +
    '2. Öffnen Sie das Blatt "Feldkonfiguration" und definieren Sie die Felder\n' +
    '3. Bearbeiten Sie das Apps Script (Erweiterungen → Apps Script)\n' +
    '4. Fügen Sie die Kategorie in FIELD_CONFIG hinzu\n\n' +
    'Beispiel:\n' +
    '"Neue Kategorie": [\n' +
    '  { name: "Feld 1", options: ["Option A", "Option B"] },\n' +
    '  { name: "Feld 2", options: ["Option X", "Option Y"] }\n' +
    ']',
    ui.ButtonSet.OK
  );
}

/**
 * Test-Funktion
 */
function testConfiguration() {
  Logger.log("Testing configuration...");
  Logger.log("Available categories:");
  
  Object.keys(FIELD_CONFIG).forEach(category => {
    Logger.log(`\n${category}:`);
    FIELD_CONFIG[category].forEach(field => {
      Logger.log(`  - ${field.name}: ${field.options.join(", ")}`);
    });
  });
  
  Logger.log("\nConfiguration test complete!");
}
```

#### Script speichern und verwenden

1. Klicken Sie auf "Speichern" (Disketten-Symbol)
2. Benennen Sie das Projekt: "Kundenüberwachung Script"
3. Schließen Sie den Apps Script Editor
4. Laden Sie die Tabelle neu (F5)
5. Ein neues Menü "🔧 Kundenüberwachung" erscheint
6. Beim ersten Mal müssen Sie die Berechtigung erteilen

## Erweiterte Funktionen

### Automatische Spaltenüberschriften

Um auch die Spaltenüberschriften dynamisch anzupassen, erweitern Sie das Script:

```javascript
/**
 * Aktualisiert Spaltenüberschriften basierend auf der Kategorie in Zeile
 */
function updateHeadersForRow(sheet, row, category) {
  const fields = FIELD_CONFIG[category];
  
  if (!fields) return;
  
  fields.forEach((field, index) => {
    const col = CONFIG.FIRST_DYNAMIC_COLUMN + index;
    const headerCell = sheet.getRange(CONFIG.HEADER_ROW, col);
    headerCell.setValue(field.name);
    
    // Formatierung
    headerCell.setFontWeight("bold");
    headerCell.setBackground("#4472C4");
    headerCell.setFontColor("#FFFFFF");
  });
}
```

### Bedingte Formatierung

Fügen Sie visuelle Hinweise für verschiedene Status hinzu:

```javascript
function applyConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.MAIN_SHEET_NAME);
  
  const statusColumn = 3; // Spalte C
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(2, statusColumn, lastRow - 1, 1);
  
  // Regel für "Abgeschlossen" - Grün
  const completedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Abgeschlossen")
    .setBackground("#D4EDDA")
    .setFontColor("#155724")
    .setRanges([range])
    .build();
  
  // Regel für "Wartend" - Gelb
  const waitingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Wartend")
    .setBackground("#FFF3CD")
    .setFontColor("#856404")
    .setRanges([range])
    .build();
  
  // Regel für "Storniert" - Rot
  const canceledRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Storniert")
    .setBackground("#F8D7DA")
    .setFontColor("#721C24")
    .setRanges([range])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(completedRule, waitingRule, canceledRule);
  sheet.setConditionalFormatRules(rules);
}
```

### Zeitstempel für Änderungen

Erfassen Sie automatisch, wann Änderungen vorgenommen wurden:

```javascript
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  if (sheet.getName() !== CONFIG.MAIN_SHEET_NAME) return;
  
  const row = range.getRow();
  const col = range.getColumn();
  
  // Zeitstempel in Spalte K hinzufügen
  if (row > 1 && col <= 10) { // Nur Datenzeilen und relevante Spalten
    const timestampCol = 11; // Spalte K
    const timestamp = new Date();
    sheet.getRange(row, timestampCol).setValue(timestamp);
  }
  
  // Kategorie-spezifische Logik
  if (col === CONFIG.CATEGORY_COLUMN && row > 1) {
    const category = range.getValue();
    if (category && FIELD_CONFIG[category]) {
      updateRowForCategory(sheet, row, category);
    }
  }
}
```

## Trigger einrichten

### Automatische Backups

1. Im Apps Script Editor: "Trigger" (Uhr-Symbol)
2. "Trigger hinzufügen"
3. Funktion wählen: `createBackup`
4. Ereignisquelle: "Zeitgesteuert"
5. Art des zeitbasierten Triggers: "Wöchentlich"
6. Speichern

```javascript
function createBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupName = `Backup_${ss.getName()}_${new Date().toISOString().split('T')[0]}`;
  
  ss.copy(backupName);
  
  Logger.log(`Backup erstellt: ${backupName}`);
}
```

## Fehlerbehebung

### Problem: Script läuft nicht

**Lösung:**
1. Überprüfen Sie die Berechtigungen (Erweiterungen → Apps Script → Berechtigung überprüfen)
2. Stellen Sie sicher, dass alle Blattnamen korrekt sind
3. Überprüfen Sie das Protokoll (Apps Script → Ansicht → Logs)

### Problem: Dropdown-Optionen falsch

**Lösung:**
1. Überprüfen Sie FIELD_CONFIG im Script
2. Stellen Sie sicher, dass die Kategorie exakt übereinstimmt (Groß-/Kleinschreibung)
3. Verwenden Sie "Alle Zeilen aktualisieren" aus dem Menü

### Problem: Langsame Performance

**Lösung:**
1. Begrenzen Sie die Anzahl der Zeilen
2. Verwenden Sie Batch-Operationen
3. Optimieren Sie die Script-Ausführung

## Best Practices

1. **Regelmäßige Backups**: Erstellen Sie wöchentliche Backups
2. **Dokumentation**: Kommentieren Sie Ihre Script-Änderungen
3. **Testing**: Testen Sie neue Kategorien in einer Kopie
4. **Versionierung**: Verwenden Sie die Versionsverwaltung in Apps Script
5. **Berechtigungen**: Beschränken Sie Bearbeitungsrechte auf autorisierte Benutzer

## Weitere Ressourcen

- [Google Apps Script Dokumentation](https://developers.google.com/apps-script)
- [Spreadsheet Service Reference](https://developers.google.com/apps-script/reference/spreadsheet)
- [Best Practices für Apps Script](https://developers.google.com/apps-script/guides/support/best-practices)

---

**Hinweis**: Dieses Script ist ein Ausgangspunkt. Passen Sie es an Ihre spezifischen Anforderungen an.
