const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname)));

// ================================================================
// FUNZIONE PER OTTENERE LE DATE (supporta entrambi i formati)
// ================================================================
function getDates() {
    const today = new Date();
    const dates = [];
    
    for (let i = 0; i < 3; i++) {
        const d = new Date(today);
        d.setDate(d.getDate() + i);
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        
        // Aggiungi ENTRAMBI i formati!
        dates.push(`${year}-${month}-${day}`);  // Formato AAAA-MM-GG
        dates.push(`${day}/${month}/${year}`);  // Formato GG/MM/AAAA (quello del tuo Excel!)
    }
    
    return dates;
}

// ================================================================
// FUNZIONE PER LEGGERE IL FILE EXCEL
// ================================================================
function readExcelFile(filePath, sheetName) {
    try {
        if (!fs.existsSync(filePath)) {
            console.log(`⚠️ File non trovato: ${filePath}`);
            return null;
        }
        const workbook = XLSX.readFile(filePath);
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) {
            console.log(`⚠️ Foglio "${sheetName}" non trovato`);
            return null;
        }
        return XLSX.utils.sheet_to_json(sheet);
    } catch (error) {
        console.error(`❌ Errore lettura ${sheetName}:`, error.message);
        return null;
    }
}

// ================================================================
// ENDPOINT: PARTITE
// ================================================================
app.get('/api/matches', (req, res) => {
    try {
        const filePath = path.join(__dirname, 'prono.xlsx');
        
        if (!fs.existsSync(filePath)) {
            console.log('❌ File prono.xlsx non trovato in:', filePath);
            return res.json({ 
                success: true, 
                matches: [], 
                warning: 'File prono.xlsx non trovato' 
            });
        }

        const data = readExcelFile(filePath, 'Partite');
        if (!data || data.length === 0) {
            console.log('⚠️ Nessun dato nel foglio "Partite"');
            return res.json({ 
                success: true, 
                matches: [], 
                warning: 'Nessun dato nel foglio Partite' 
            });
        }

        const targetDates = getDates();
        console.log('📅 Date target:', targetDates.join(', '));
        console.log(`📊 Trovate ${data.length} righe nel file Excel`);

        // Filtra le partite per le date target (supporta entrambi i formati)
        const matches = data
            .filter(row => {
                const dateStr = String(row.Data || row['Data'] || '').trim();
                if (!dateStr) {
                    return false;
                }
                // Controlla se la data corrisponde a una delle date target
                const found = targetDates.some(target => dateStr.includes(target));
                if (found) {
                    console.log(`✅ Trovata partita: ${dateStr} - ${row.Campionato || ''}`);
                }
                return found;
            })
            .map(row => {
                // Gestisci i nomi delle colonne (alcuni potrebbero avere spazi)
                const campionato = row.Campionato || row['Campionato'] || '';
                const giornata = row.Giornata || row['Giornata'] || row['Turno'] || '';
                const data = row.Data || row['Data'] || '';
                const ora = row.Ora || row['Ora'] || '';
                const casa = row['Squadra Casa'] || row['Squadra'] || row['Home'] || '';
                const ospite = row['Squadra Ospite'] || row['Ospite'] || row['Away'] || '';
                
                return {
                    campionato: String(campionato).trim(),
                    giornata: String(giornata).trim(),
                    data: String(data).trim(),
                    ora: String(ora).trim(),
                    casa: String(casa).trim(),
                    ospite: String(ospite).trim(),
                    gol_casa: parseInt(row['Gol Casa'] || row['GC'] || 0),
                    gol_ospite: parseInt(row['Gol Ospite'] || row['GO'] || 0)
                };
            })
            .filter(m => m.campionato && m.casa && m.ospite);

        // Ordina per data e ora
        matches.sort((a, b) => (a.data + a.ora).localeCompare(b.data + b.ora));

        console.log(`✅ Trovate ${matches.length} partite per le date target`);
        if (matches.length > 0) {
            console.log('📋 Partite trovate:');
            matches.forEach(m => {
                console.log(`   - ${m.data} ${m.ora}: ${m.casa} vs ${m.ospite} (${m.campionato})`);
            });
        }
        
        res.json({
            success: true,
            matches: matches,
            total: matches.length,
            dates: targetDates
        });

    } catch (error) {
        console.error('❌ Errore /api/matches:', error);
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// ================================================================
// ENDPOINT: CLASSIFICA
// ================================================================
app.get('/api/standings', (req, res) => {
    try {
        const filePath = path.join(__dirname, 'prono.xlsx');
        
        if (!fs.existsSync(filePath)) {
            return res.json({ success: true, standings: [] });
        }

        const data = readExcelFile(filePath, 'Classifica');
        if (!data) {
            return res.json({ success: true, standings: [] });
        }
        
        // Mappa i dati della classifica
        const standings = data.map(row => ({
            Campionato: String(row.Campionato || row['Campionato'] || '').trim(),
            Rk: row.Rk || row['Rk'] || row.Pos || '-',
            Squad: String(row.Squad || row['Squad'] || row.Squadra || '').trim(),
            MP: row.MP || row['MP'] || row.PG || '-',
            W: row.W || row['W'] || row.V || '-',
            D: row.D || row['D'] || row.P || '-',
            L: row.L || row['L'] || row.S || '-',
            GF: row.GF || row['GF'] || 0,
            GA: row.GA || row['GA'] || 0,
            GD: row.GD || row['GD'] || row.Diff || '-',
            Pts: row.Pts || row['Pts'] || row.Punti || '-',
            'Last 5': row['Last 5'] || row.Last5 || row.Forma || ''
        }));
        
        res.json({
            success: true,
            standings: standings,
            total: standings.length
        });
    } catch (error) {
        console.error('❌ Errore /api/standings:', error);
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// ================================================================
// ENDPOINT: TEST
// ================================================================
app.get('/api/test', (req, res) => {
    const filePath = path.join(__dirname, 'prono.xlsx');
    const exists = fs.existsSync(filePath);
    
    let sheets = [];
    let sampleData = [];
    if (exists) {
        try {
            const workbook = XLSX.readFile(filePath);
            sheets = workbook.SheetNames;
            
            // Leggi i primi 5 dati per test
            const sheet = workbook.Sheets['Partite'];
            if (sheet) {
                const data = XLSX.utils.sheet_to_json(sheet);
                sampleData = data.slice(0, 5).map(row => ({
                    Data: row.Data || row['Data'] || '',
                    Campionato: row.Campionato || row['Campionato'] || '',
                    'Squadra Casa': row['Squadra Casa'] || row['Squadra'] || '',
                    'Squadra Ospite': row['Squadra Ospite'] || row['Ospite'] || ''
                }));
            }
        } catch (e) {
            console.error('Errore test:', e);
        }
    }
    
    res.json({
        success: true,
        fileExists: exists,
        filePath: filePath,
        cwd: __dirname,
        sheets: sheets,
        sampleData: sampleData,
        dates: getDates()
    });
});

// ================================================================
// AVVIO SERVER
// ================================================================
app.listen(PORT, '0.0.0.0', () => {
    console.log('\n' + '='.repeat(60));
    console.log('✅ SERVER AVVIATO CON SUCCESSO');
    console.log('='.repeat(60));
    console.log(`📍 Porta: ${PORT}`);
    console.log(`📁 Cartella: ${__dirname}`);
    
    const filePath = path.join(__dirname, 'prono.xlsx');
    if (fs.existsSync(filePath)) {
        console.log(`📊 prono.xlsx: ✅ TROVATO`);
        try {
            const workbook = XLSX.readFile(filePath);
            console.log(`📋 Fogli trovati: ${workbook.SheetNames.join(', ')}`);
            
            // Mostra un esempio dei dati
            const sheet = workbook.Sheets['Partite'];
            if (sheet) {
                const data = XLSX.utils.sheet_to_json(sheet);
                console.log(`📊 Righe nel foglio "Partite": ${data.length}`);
                if (data.length > 0) {
                    const sample = data[0];
                    console.log('📋 Esempio dati:', JSON.stringify(sample, null, 2).slice(0, 200) + '...');
                }
            }
        } catch (e) {
            console.log(`⚠️ Errore lettura file: ${e.message}`);
        }
    } else {
        console.log(`📊 prono.xlsx: ❌ NON TROVATO`);
        console.log(`   Metti il file in: ${filePath}`);
    }
    
    console.log('\n🌐 URL di accesso:');
    console.log(`   → http://localhost:${PORT}`);
    console.log('='.repeat(60) + '\n');
});