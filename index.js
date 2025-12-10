const ExcelJS = require('exceljs');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

/**
 * Ruft alle Ergebnisse von der Davengo API ab (mit Pagination)
 * Gibt sowohl die Ergebnisse als auch die Felddefinitionen zurück
 */
async function fetchAllResults(laufname, category) {
    const allResults = [];
    let offset = 0;
    let hasMore = true;
    let fields = null;

    console.log(`Lade Ergebnisse für Lauf: ${laufname}, Kategorie: ${category}...`);

    while (hasMore) {
        try {
            const response = await axios.post(
                `https://www.davengo.com/event/result/${laufname}/search/list`,
                {
                    type: "simple",
                    term: null,
                    category: category,
                    offset: offset
                },
                {
                    headers: {
                        'Content-Type': 'application/json'
                    }
                }
            );

            const data = response.data;
            
            // Speichere Felddefinitionen aus dem ersten Request
            if (!fields && data.fields) {
                fields = data.fields;
            }
            
            if (data.results && data.results.length > 0) {
                allResults.push(...data.results);
                console.log(`  Geladen: ${allResults.length} Ergebnisse (Offset: ${offset})`);
            }

            // Prüfe ob es weitere Ergebnisse gibt
            if (data.navigation && data.navigation.nextOffset) {
                offset = data.navigation.nextOffset;
            } else {
                hasMore = false;
            }

            // Wenn keine Ergebnisse mehr kommen, stoppe
            if (!data.results || data.results.length === 0) {
                hasMore = false;
            }

        } catch (error) {
            console.error(`Fehler beim Abrufen der Daten (Offset: ${offset}):`, error.message);
            if (error.response) {
                console.error(`  Status: ${error.response.status}`);
                console.error(`  Daten:`, error.response.data);
            }
            hasMore = false;
        }
    }

    console.log(`Insgesamt ${allResults.length} Ergebnisse geladen.\n`);
    return { results: allResults, fields: fields || [] };
}

/**
 * Erstellt ein Excel-Workbook mit allen Ergebnissen
 */
async function createExcelFile(results, fields, outputPath) {
    const workbook = new ExcelJS.Workbook();
    
    // Helper-Funktion zum Erstellen eines Sheets
    function createSheet(workbook, sheetName, data, fields) {
        const sheet = workbook.addWorksheet(sheetName);
        
        // Header-Zeile erstellen
        const headerRow = sheet.addRow(fields.map(f => f.title));
        headerRow.font = { bold: true };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
        };
        
        // Daten-Zeilen hinzufügen
        data.forEach(result => {
            const row = fields.map(field => {
                const value = result[field.id];
                // Spezielle Behandlung für Zeit-Werte
                if (field.id === 'bruttoTime' && value) {
                    return value;
                }
                return value || '';
            });
            sheet.addRow(row);
        });
        
        // Spaltenbreiten anpassen
        fields.forEach((field, index) => {
            const column = sheet.getColumn(index + 1);
            column.width = Math.max(15, field.title.length + 2);
            
            // Textausrichtung setzen
            if (field.textAlignment === 'RIGHT') {
                column.alignment = { horizontal: 'right' };
            } else {
                column.alignment = { horizontal: 'left' };
            }
        });
        
        // Header-Zeile einfrieren
        sheet.views = [{
            state: 'frozen',
            ySplit: 1
        }];
    }
    
    // 1. Sheet: Gesamt - Alle Ergebnisse
    console.log('Erstelle Sheet: Gesamt...');
    createSheet(workbook, 'Gesamt', results, fields);
    
    // 2. Sheet: Männer - Alle Ergebnisse der Männer
    console.log('Erstelle Sheet: Männer...');
    const menResults = results.filter(r => r.rankMale !== null && r.rankMale !== undefined);
    if (menResults.length > 0) {
        createSheet(workbook, 'Männer', menResults, fields);
    }
    
    // 3. Sheet: Frauen - Alle Ergebnisse der Frauen
    console.log('Erstelle Sheet: Frauen...');
    const womenResults = results.filter(r => r.rankFemale !== null && r.rankFemale !== undefined);
    if (womenResults.length > 0) {
        createSheet(workbook, 'Frauen', womenResults, fields);
    }
    
    // 4. Sheets für Altersklassen
    // Sammle alle eindeutigen Altersklassen
    const ageGroups = new Set();
    results.forEach(r => {
        if (r.ageGroupShort) {
            ageGroups.add(r.ageGroupShort);
        }
    });
    
    // Sortiere Altersklassen (M/W ohne Zahl zuerst, dann jüngste zuerst)
    const sortedAgeGroups = Array.from(ageGroups).sort((a, b) => {
        // Zuerst nach Geschlecht (M vor W)
        const aGender = a.startsWith('M') ? 'M' : (a.startsWith('W') ? 'W' : '');
        const bGender = b.startsWith('M') ? 'M' : (b.startsWith('W') ? 'W' : '');
        
        if (aGender !== bGender) {
            if (aGender === 'M') return -1;
            if (aGender === 'W') return 1;
            return aGender.localeCompare(bGender);
        }
        
        // Prüfe ob Altersklasse eine Zahl enthält
        const hasAge = (str) => /\d+/.test(str);
        const aHasAge = hasAge(a);
        const bHasAge = hasAge(b);
        
        // Altersklassen ohne Zahl (nur "M" oder "W") kommen zuerst
        if (aHasAge !== bHasAge) {
            return aHasAge ? 1 : -1; // Ohne Zahl zuerst
        }
        
        // Wenn beide eine Zahl haben, sortiere nach Alter (niedrigste zuerst)
        if (aHasAge && bHasAge) {
            const extractAge = (str) => {
                const match = str.match(/\d+/);
                return match ? parseInt(match[0], 10) : 0;
            };
            
            const aAge = extractAge(a);
            const bAge = extractAge(b);
            
            if (aAge !== bAge) {
                return aAge - bAge;
            }
        }
        
        // Falls gleiches Alter oder beide ohne Zahl, alphabetisch
        return a.localeCompare(b);
    });
    
    console.log(`Erstelle ${sortedAgeGroups.length} Sheets für Altersklassen...`);
    sortedAgeGroups.forEach(ageGroup => {
        const ageGroupResults = results.filter(r => r.ageGroupShort === ageGroup);
        if (ageGroupResults.length > 0) {
            // Sheet-Name: Altersklasse (z.B. "AK M20")
            const sheetName = `AK ${ageGroup}`;
            console.log(`  Erstelle Sheet: ${sheetName}...`);
            createSheet(workbook, sheetName, ageGroupResults, fields);
        }
    });
    
    // Excel-Datei speichern
    console.log(`\nSpeichere Excel-Datei: ${outputPath}...`);
    await workbook.xlsx.writeFile(outputPath);
    console.log('✓ Excel-Datei erfolgreich erstellt!');
}

/**
 * Hauptfunktion
 */
async function main() {
    // Parameter aus Kommandozeile lesen
    const args = process.argv.slice(2);
    
    if (args.length < 2) {
        console.log('Verwendung: node index.js <laufname> <kategorie> [output-datei]');
        console.log('Beispiel: node index.js nikolauslauf-2024 Nikolauslauf');
        console.log('Beispiel: node index.js nikolauslauf-2024 Nikolauslauf ergebnisse.xlsx');
        process.exit(1);
    }
    
    const laufname = args[0];
    const category = args[1];
    const outputFile = args[2] || `${laufname}_${category}_ergebnisse.xlsx`;
    
    try {
        // Alle Ergebnisse abrufen
        const { results, fields } = await fetchAllResults(laufname, category);
        
        if (results.length === 0) {
            console.log('Keine Ergebnisse gefunden.');
            process.exit(1);
        }
        
        if (!fields || fields.length === 0) {
            console.log('Warnung: Keine Felddefinitionen gefunden. Verwende Standard-Felder.');
            // Fallback auf Standard-Felder
            const defaultFields = [
                { id: 'rankTotal', title: 'Platz', textAlignment: 'LEFT', priority: true },
                { id: 'rankMale', title: 'M', textAlignment: 'LEFT', priority: false },
                { id: 'rankFemale', title: 'W', textAlignment: 'LEFT', priority: false },
                { id: 'startNo', title: 'Nr.', textAlignment: 'LEFT', priority: true },
                { id: 'firstName', title: 'Vorname', textAlignment: 'LEFT', priority: true },
                { id: 'lastName', title: 'Nachname', textAlignment: 'LEFT', priority: true },
                { id: 'rankAgeGroup', title: 'AK-Platz', textAlignment: 'LEFT', priority: false },
                { id: 'ageGroupShort', title: 'AK', textAlignment: 'LEFT', priority: false },
                { id: 'teamName', title: 'Verein', textAlignment: 'LEFT', priority: false },
                { id: 'bruttoTime', title: 'Brutto', textAlignment: 'RIGHT', priority: true }
            ];
            await createExcelFile(results, defaultFields, outputFile);
        } else {
            // Excel-Datei erstellen
            await createExcelFile(results, fields, outputFile);
        }
        
        console.log(`\n✓ Fertig! Datei gespeichert: ${path.resolve(outputFile)}`);
        
    } catch (error) {
        console.error('Fehler:', error.message);
        process.exit(1);
    }
}

// Programm starten
if (require.main === module) {
    main();
}

module.exports = { fetchAllResults, createExcelFile };

