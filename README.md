# Davengo Ergebnisse Exporter

Ein JavaScript-Tool zum Exportieren von Lauf-Ergebnissen aus [Davengo](https://www.davengo.com) in Excel-Dateien mit mehreren Arbeitsblättern.

## Features

- ✅ Automatisches Abrufen aller Ergebnisse (mit Pagination)
- ✅ Export in Excel-Format (.xlsx)
- ✅ Mehrere Arbeitsblätter für verschiedene Ansichten:
  - **Gesamt**: Alle Ergebnisse
  - **Männer**: Alle Ergebnisse der Männer
  - **Frauen**: Alle Ergebnisse der Frauen
  - **Altersklassen**: Je ein Arbeitsblatt pro Altersklasse (z.B. "AK M20", "AK W30")
- ✅ Automatische Formatierung (Spaltenbreiten, Header, Textausrichtung)
- ✅ Gefrorene Header-Zeilen für bessere Navigation

## Voraussetzungen

- Node.js (Version 14 oder höher)
- npm (wird mit Node.js mitgeliefert)

## Installation

1. Repository klonen oder herunterladen
2. Dependencies installieren:

```bash
npm install
```

## Verwendung

### Grundlegende Verwendung

```bash
node index.js <laufname> <kategorie>
```

### Parameter

- **laufname** (erforderlich): Der Name des Laufs in der Davengo-URL
  - Beispiel: `nikolauslauf-2024`
- **kategorie** (erforderlich): Die Kategorie des Laufs
  - Beispiel: `Nikolauslauf`
- **output-datei** (optional): Name der Ausgabedatei
  - Standard: `<laufname>_<kategorie>_ergebnisse.xlsx`

### Beispiele

```bash
# Standard-Verwendung
node index.js nikolauslauf-2024 Nikolauslauf

# Mit eigenem Dateinamen
node index.js nikolauslauf-2024 Nikolauslauf ergebnisse.xlsx

# Anderer Lauf
node index.js marathon-2024 Marathon
```

## Ausgabe

Das Programm erstellt eine Excel-Datei mit folgenden Arbeitsblättern:

1. **Gesamt**: Alle Teilnehmer-Ergebnisse
2. **Männer**: Gefiltert nach männlichen Teilnehmern
3. **Frauen**: Gefiltert nach weiblichen Teilnehmern
4. **AK [Altersklasse]**: Ein Arbeitsblatt pro Altersklasse (z.B. "AK M", "AK M20", "AK W30")

### Spalten in den Arbeitsblättern

- Platz (Gesamtplatzierung)
- M (Platz Männer)
- W (Platz Frauen)
- Nr. (Startnummer)
- Vorname
- Nachname
- AK-Platz (Platz in Altersklasse)
- AK (Altersklasse)
- Verein
- Brutto (Bruttozeit)

Die genauen Spalten hängen von den Felddefinitionen der API-Response ab.

## Technische Details

### API-Endpoint

Das Tool verwendet folgenden Davengo API-Endpoint:

```
POST https://www.davengo.com/event/result/:laufname/search/list
```

### Request Body

```json
{
    "type": "simple",
    "term": null,
    "category": "Kategorie",
    "offset": 0
}
```

### Pagination

Das Tool lädt automatisch alle Ergebnisse, auch wenn mehrere Seiten vorhanden sind. Die Pagination erfolgt über den `offset`-Parameter in der Response.

## Fehlerbehandlung

Bei Fehlern zeigt das Programm:
- Detaillierte Fehlermeldungen
- HTTP-Status-Codes (falls vorhanden)
- Informationen zum aktuellen Offset bei Pagination-Fehlern

## Abhängigkeiten

- **exceljs** (^4.4.0): Erstellung von Excel-Dateien
- **axios** (^1.6.0): HTTP-Requests zur Davengo API

## Lizenz

ISC

## Support

Bei Problemen oder Fragen:
1. Überprüfen Sie, ob der Laufname und die Kategorie korrekt sind
2. Stellen Sie sicher, dass Sie eine Internetverbindung haben
3. Prüfen Sie, ob die Davengo-API erreichbar ist

