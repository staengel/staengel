# Kundenüberwachungsplattform - Dokumentation

## Übersicht

Diese Lösung bietet eine strukturierte Plattform zur Überwachung von Kundenprojekten für ISO 9001 Compliance, implementiert als Excel-Arbeitsmappe mit intelligenten Dropdown-Menüs und dynamischen Feldern.

## Konzept und Struktur

### 1. Hauptblatt (Kundenübersicht)

Das Hauptblatt enthält eine Zeile pro Kunde mit folgenden Spalten:

#### Fixierte Spalten (1-3)
- **Spalte A (Kunden-Nr.)**: Eindeutige Kundennummer
- **Spalte B (Kundenname)**: Name des Kunden
- **Spalte C (Status)**: Aktueller Projektstatus (z.B. "In Bearbeitung", "Abgeschlossen", "Wartend")

#### Dynamische Spalten (ab Spalte 4)
- **Spalte D (Leistungskategorie)**: Dropdown zur Auswahl der Servicekategorie
- **Spalten E-J+ (Kategoriefelder)**: Dynamische Felder, die sich je nach gewählter Kategorie ändern

### 2. Konfigurationsblätter

Die Lösung verwendet zusätzliche Blätter zur Konfiguration:

#### Blatt "Kategorien"
Definiert die verfügbaren Leistungskategorien:
- Bild-Aufnahme
- Drohnenshow
- 3D-Modell erstellen
- Video-Produktion
- Inspektion
- Vermessung

#### Blatt "Feldkonfiguration"
Definiert für jede Kategorie die zugehörigen Felder und deren Dropdown-Optionen:

| Kategorie | Feld 1 | Optionen 1 | Feld 2 | Optionen 2 | Feld 3 | Optionen 3 |
|-----------|--------|------------|--------|------------|--------|------------|
| Bild-Aufnahme | Overlap | 60/60;80/80;70/70 | Kamera | RGB;Multispektral;Thermal | Auflösung | 20MP;45MP;100MP |
| Drohnenshow | Anzahl Drohnen | 10;25;50;100;200 | Dauer (Min) | 5;10;15;20;30 | Musik | Ja;Nein |
| 3D-Modell erstellen | Detailgrad | Niedrig;Mittel;Hoch;Sehr Hoch | Format | OBJ;FBX;STL;GLTF | Texturierung | Ja;Nein |

#### Blatt "Status-Optionen"
Definiert die verfügbaren Statuswerte für Spalte C:
- In Bearbeitung
- Abgeschlossen
- Wartend
- Genehmigung ausstehend
- Storniert

### 3. Datenvalidierung

Die Lösung verwendet Excel-Datenvalidierung für:
1. **Kategorien-Dropdown (Spalte D)**: Verweist auf die Liste im Blatt "Kategorien"
2. **Status-Dropdown (Spalte C)**: Verweist auf die Liste im Blatt "Status-Optionen"
3. **Dynamische Feld-Dropdowns (Spalten E+)**: Verwenden bedingte Formeln basierend auf der gewählten Kategorie

## Implementierungsansatz

### Methode 1: Einfache Lösung mit manueller Anpassung

**Vorteil**: Einfach zu verstehen und zu pflegen
**Nachteil**: Erfordert manuelle Anpassung der Spaltenüberschriften

In dieser Methode:
1. Der Benutzer wählt die Kategorie in Spalte D
2. Die Spaltenüberschriften (Zeile 1) zeigen die Feldnamen für jede Kategorie
3. Die Dropdown-Optionen in den Feldern (E, F, G, etc.) sind kategoriespezifisch konfiguriert
4. Der Benutzer muss die relevanten Felder für die gewählte Kategorie ausfüllen

### Methode 2: Erweiterte Lösung mit Formeln

**Vorteil**: Vollautomatisch, Spaltenüberschriften passen sich an
**Nachteil**: Komplexere Formel-Logik

In dieser Methode (für zukünftige Erweiterung mit Google Sheets Apps Script oder Excel VBA):
1. Makros/Scripts erkennen die gewählte Kategorie
2. Spaltenüberschriften werden automatisch angepasst
3. Datenvalidierungsregeln werden dynamisch erstellt
4. Nur relevante Felder werden angezeigt

## Verwendung der Vorlage

### Schritt 1: Neue Zeile für Kunden hinzufügen
1. Fügen Sie eine neue Zeile unter den vorhandenen Kunden ein
2. Geben Sie die Kunden-Nr. in Spalte A ein
3. Geben Sie den Kundennamen in Spalte B ein
4. Wählen Sie den Status in Spalte C aus dem Dropdown

### Schritt 2: Kategorie wählen
1. Klicken Sie auf Spalte D der Kundenzeile
2. Wählen Sie die gewünschte Leistungskategorie aus dem Dropdown

### Schritt 3: Kategorie-spezifische Felder ausfüllen
1. Schauen Sie sich die Spaltenüberschriften an (E, F, G, etc.)
2. Für jedes relevante Feld Ihrer gewählten Kategorie:
   - Klicken Sie auf die Zelle
   - Wählen Sie den gewünschten Wert aus dem Dropdown
3. Nicht relevante Felder (für andere Kategorien) können leer bleiben

### Schritt 4: Konfiguration erweitern
Um neue Kategorien oder Feldoptionen hinzuzufügen:

1. **Neue Kategorie hinzufügen**:
   - Gehen Sie zum Blatt "Kategorien"
   - Fügen Sie die neue Kategorie in der Liste hinzu
   - Gehen Sie zum Blatt "Feldkonfiguration"
   - Fügen Sie eine neue Zeile mit der Kategorie und ihren Feldern hinzu

2. **Feldoptionen ändern**:
   - Gehen Sie zum Blatt "Feldkonfiguration"
   - Ändern Sie die Optionen in der entsprechenden Zelle (getrennt durch Semikolon)

3. **Neue Statuswerte hinzufügen**:
   - Gehen Sie zum Blatt "Status-Optionen"
   - Fügen Sie den neuen Status zur Liste hinzu

## Technische Details

### Datenvalidierung in Excel

Die Vorlage verwendet Excel-Datenvalidierung mit folgenden Formeln:

1. **Kategorien-Dropdown (Spalte D)**:
   ```
   =Kategorien!$A$2:$A$7
   ```

2. **Status-Dropdown (Spalte C)**:
   ```
   =Status_Optionen!$A$2:$A$6
   ```

3. **Dynamische Feld-Dropdowns** (Beispiel für Spalte E):
   ```
   =INDIREKT("Feldkonfiguration!$C$"&VERGLEICH($D2;Feldkonfiguration!$A:$A;0))
   ```
   Diese Formel sucht die Zeile für die gewählte Kategorie und gibt die entsprechenden Optionen zurück.

### Spaltenstruktur im Hauptblatt

| Spalte | Zweck | Typ | Datenquelle |
|--------|-------|-----|-------------|
| A | Kunden-Nr. | Text/Zahl | Manuell |
| B | Kundenname | Text | Manuell |
| C | Status | Dropdown | Status_Optionen Blatt |
| D | Kategorie | Dropdown | Kategorien Blatt |
| E | Feld 1 | Dropdown | Feldkonfiguration (basierend auf Kategorie) |
| F | Feld 2 | Dropdown | Feldkonfiguration (basierend auf Kategorie) |
| G | Feld 3 | Dropdown | Feldkonfiguration (basierend auf Kategorie) |
| H | Feld 4 | Dropdown | Feldkonfiguration (basierend auf Kategorie) |
| I | Feld 5 | Dropdown | Feldkonfiguration (basierend auf Kategorie) |
| J | Feld 6 | Dropdown | Feldkonfiguration (basierend auf Kategorie) |

## Migration zu Google Sheets

Um diese Lösung in Google Sheets zu verwenden:

1. **Datei hochladen**:
   - Öffnen Sie Google Drive
   - Laden Sie die .xlsx Datei hoch
   - Google Sheets konvertiert sie automatisch

2. **Formeln anpassen** (falls nötig):
   - Excel `VERGLEICH` → Google Sheets `MATCH`
   - Excel `INDIREKT` → Google Sheets `INDIRECT`
   - Excel `;` in Listen → Google Sheets `,`

3. **Erweiterte Funktionen mit Apps Script**:
   Für vollautomatische Spaltenanpassung können Sie Google Apps Script verwenden:
   
   ```javascript
   function onEdit(e) {
     var sheet = e.source.getActiveSheet();
     var range = e.range;
     
     // Prüfen ob Kategorie-Spalte (D) bearbeitet wurde
     if (range.getColumn() == 4 && sheet.getName() == "Kundenübersicht") {
       var row = range.getRow();
       var category = range.getValue();
       
       // Spaltenüberschriften basierend auf Kategorie aktualisieren
       updateFieldHeaders(sheet, row, category);
       
       // Datenvalidierungen aktualisieren
       updateValidations(sheet, row, category);
     }
   }
   ```

## Vorteile dieser Struktur

1. **Modular**: Neue Kategorien und Felder können einfach hinzugefügt werden
2. **Datenintegrität**: Dropdown-Menüs verhindern Tippfehler
3. **ISO 9001 konform**: Strukturierte Dokumentation und Nachverfolgbarkeit
4. **Skalierbar**: Kann auf viele Kunden und Kategorien erweitert werden
5. **Wartbar**: Zentrale Konfiguration in separaten Blättern
6. **Flexibel**: Jeder Kunde kann unterschiedliche Leistungskategorien haben

## Einschränkungen und Lösungen

### Einschränkung 1: Fixe Spaltenanzahl
**Problem**: Excel hat eine begrenzte Anzahl von Spalten
**Lösung**: Bei Bedarf mehrere Zeilen pro Kunde verwenden oder zu einer Datenbank-Lösung migrieren

### Einschränkung 2: Komplexe bedingte Validierung
**Problem**: Excel-Datenvalidierung ist für sehr komplexe Szenarien limitiert
**Lösung**: Verwendung von VBA-Makros (Excel) oder Apps Script (Google Sheets) für erweiterte Funktionalität

### Einschränkung 3: Mehrere Kategorien pro Kunde
**Problem**: Aktuell ist nur eine Kategorie pro Zeile möglich
**Lösung**: Mehrere Zeilen für denselben Kunden verwenden oder erweiterte Lösung mit Sub-Tabellen implementieren

## Zukünftige Erweiterungen

1. **Automatische Berichtserstellung**: Pivot-Tabellen für Übersichten nach Kategorie oder Status
2. **Bedingte Formatierung**: Visuelle Hervorhebung von kritischen Status
3. **Zeitstempel**: Automatische Erfassung von Erstellungs- und Änderungsdaten
4. **Verantwortlichkeiten**: Zusätzliche Spalte für zuständige Mitarbeiter
5. **Kommentare/Notizen**: Freitextfeld für zusätzliche Informationen
6. **Dashboard**: Separates Blatt mit Statistiken und Übersichten
7. **Versionierung**: Historie von Statusänderungen

## Support und Wartung

### Häufige Probleme

1. **Dropdown zeigt keine Optionen**:
   - Prüfen Sie, ob die Konfigurationsblätter vorhanden sind
   - Prüfen Sie, ob die Zellbereiche korrekt benannt sind

2. **Falsche Optionen im Dropdown**:
   - Prüfen Sie die Feldkonfiguration für die entsprechende Kategorie
   - Stellen Sie sicher, dass Optionen mit Semikolon getrennt sind

3. **Formeln funktionieren nicht**:
   - Prüfen Sie, ob Blattnamen korrekt sind (keine Leerzeichen oder Sonderzeichen)
   - Prüfen Sie die Zellenreferenzen in den Formeln

### Wartungsaufgaben

- **Monatlich**: Überprüfen Sie abgeschlossene Projekte und archivieren Sie bei Bedarf
- **Quartalsweise**: Überprüfen Sie Kategorien und Felder auf Aktualität
- **Jährlich**: Erstellen Sie ein Backup und überprüfen Sie die Struktur auf Optimierungspotenzial

## Datenschutz und Sicherheit

Bei der Verwendung dieser Plattform beachten Sie:

1. **Zugriffskontrolle**: Beschränken Sie den Zugriff auf autorisierte Mitarbeiter
2. **Backups**: Erstellen Sie regelmäßige Backups der Datei
3. **Versionskontrolle**: Bei Google Sheets automatisch, bei Excel manuell
4. **Personenbezogene Daten**: Beachten Sie DSGVO-Anforderungen
5. **Audit Trail**: Dokumentieren Sie wichtige Änderungen

## Kontakt und Feedback

Für Fragen, Verbesserungsvorschläge oder Support wenden Sie sich an den Systemadministrator.

---

**Version**: 1.0  
**Datum**: November 2025  
**Erstellt für**: ISO 9001 Compliance Monitoring
