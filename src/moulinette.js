const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Répertoires
const inputDir = path.join(__dirname, '../', 'convertir');
const outputDir = path.join(__dirname, '../', 'completer');
const errorLogDir = path.join(__dirname, '../', 'logs'); // Répertoire pour les fichiers de log

// Vérifie si les répertoires de sortie et de log existent, sinon les crée
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
}
if (!fs.existsSync(errorLogDir)) {
    fs.mkdirSync(errorLogDir);
}

const phoneNumberRegex = /^[0-9]{9,15}$/; // Numéro de téléphone valide (9 à 15 chiffres)
const invalidCharactersRegex = /[^0-9\s.]/; // Caractères invalides dans les numéros de téléphone (hors chiffres, espaces et points)

let colStyle = {};
let cellStyle = {};

// Fonction pour écrire les erreurs dans un fichier texte uniquement s'il y a des erreurs
function writeErrorLog(fileName, sheetLogs) {
    const logFilePath = path.join(errorLogDir, `${path.basename(fileName, path.extname(fileName))}_errors.txt`);
    const logStream = fs.createWriteStream(logFilePath, { flags: 'w' });

    for (const [sheetName, errors] of Object.entries(sheetLogs)) {
        if (errors.length === 0) continue; // Ne crée pas de section pour les feuilles sans erreurs

        logStream.write(`Feuille : ${sheetName}\n`);
        errors.forEach(error => {
            logStream.write(`  Cellule : ${error.cellRef}\n`);
            logStream.write(`  Valeur : ${error.value}\n`);
            logStream.write(`  Raison : ${error.reason}\n`);
            logStream.write(`\n`);
        });
        logStream.write(`\n`);
    }

    logStream.end();
}

// Fonction pour traiter une feuille
async function processSheet(workbook, sheetName, filePath, sheetLogs) {
    // Initialisation des compteurs par feuille
    let invalidReasons = {
        invalidType: 0,
        invalidCharacters: 0,
        invalidFormat: 0
    };
    let correctedCount = 0;
    let errors = [];

    const worksheet = workbook.Sheets[sheetName];

    // Sauvegarder les styles des colonnes et des cellules
    colStyle[sheetName] = worksheet['!cols'] ? worksheet['!cols'].slice(0, 7) : [];
    cellStyle[sheetName] = {};

    for (let cell in worksheet) {
        if (cell[0] === '!') continue; // Ignore les propriétés internes
        const cellRef = cell;
        const cellData = worksheet[cell];
        if (cellData && cellData.s) {
            cellStyle[sheetName][cellRef] = cellData.s;
        }
    }

    // Convertir la feuille en JSON
    const data = xlsx.utils.sheet_to_json(worksheet, { defval: null, header: 1 });

    // Identifier les indices des colonnes "Téléphone" et "Mobile"
    const headers = data[0];
    let phoneColIndex = headers.findIndex(header => /Téléphone/i.test(header));
    let mobileColIndex = headers.findIndex(header => /Mobile/i.test(header));

    if (phoneColIndex === -1 && mobileColIndex === -1) {
        console.log(`Feuille "${sheetName}" : Colonnes "Téléphone" et "Mobile" non trouvées dans ${filePath}`);
        return;
    }

    // Retranscrire l'index en lettre alphabétique
    const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    let phoneColString = '';
    let mobileColString = '';

    if (phoneColIndex >= 0) {
        let quotient = Math.floor(phoneColIndex / 26);
        let remainder = phoneColIndex % 26;
        phoneColString = alphabet.charAt(remainder);

        while (quotient > 0) {
            remainder = quotient % 26;
            quotient = Math.floor(quotient / 26);
            phoneColString = alphabet.charAt(remainder - 1) + phoneColString;
        }

        console.log(`Feuille "${sheetName}" : Colonnes Téléphone : ${phoneColString}`);
    } else {
        console.log(`Feuille "${sheetName}" : Colonne Téléphone : non trouvée`);
    }

    if (mobileColIndex >= 0) {
        let quotient = Math.floor(mobileColIndex / 26);
        let remainder = mobileColIndex % 26;
        mobileColString = alphabet.charAt(remainder);

        while (quotient > 0) {
            remainder = quotient % 26;
            quotient = Math.floor(quotient / 26);
            mobileColString = alphabet.charAt(remainder - 1) + mobileColString;
        }

        console.log(`Feuille "${sheetName}" : Colonnes Mobile : ${mobileColString}`);
    } else {
        console.log(`Feuille "${sheetName}" : Colonne Mobile : non trouvée`);
    }

    const uniqueNumbers = new Set();
    const rowsToRemove = new Set();

    // Fonction pour traiter un numéro de téléphone
    function processPhoneNumber(phoneNumber, cellRef) {
        if (!phoneNumber || phoneNumber === null) {
            return null; // Ne rien faire si la cellule est vide
        }

        if (typeof phoneNumber === 'number') {
            phoneNumber = phoneNumber.toString();
        }

        if (typeof phoneNumber !== 'string') {
            invalidReasons.invalidType++;
            errors.push({ cellRef, value: phoneNumber + `(${typeof(phoneNumber)})`, reason: 'Type de données invalide' });
            return null;
        }

        if (invalidCharactersRegex.test(phoneNumber)) {
            invalidReasons.invalidCharacters++;
            errors.push({ cellRef, value: phoneNumber, reason: 'Caractères invalides détectés' });
            return null;
        }

        const originalPhoneNumber = phoneNumber;
        let cleanedPhoneNumber = phoneNumber.replace(/[\s.]+/g, ''); // Retirer les espaces et les points

        if (cleanedPhoneNumber !== originalPhoneNumber) {
            correctedCount++; // Compte comme corrigé si des espaces ont été supprimés
        }

        if (phoneNumberRegex.test(cleanedPhoneNumber)) {
            if (!cleanedPhoneNumber.startsWith('0')) {
                cleanedPhoneNumber = '0' + cleanedPhoneNumber;
                correctedCount++;
            }
            return cleanedPhoneNumber;
        } else {
            invalidReasons.invalidFormat++;
            errors.push({ cellRef, value: phoneNumber, reason: 'Format invalide' });
            return null; // Indiquer que la valeur est invalide
        }
    }

    const colIndexes = [phoneColIndex, mobileColIndex].filter(index => index >= 0); // Filtrer les colonnes valides

    // Parcourir les numéros de téléphone dans les colonnes dynamiques
    data.forEach((row, rowIndex) => {
        colIndexes.forEach((colIndex) => {
            if (colIndex >= 0 && row[colIndex] !== undefined && rowIndex > 0) { // Ignorer la ligne d'en-tête
                const cellRef = `${alphabet.charAt(colIndex % 26)}${rowIndex + 1}`;
                let processedNumber = processPhoneNumber(row[colIndex], cellRef);

                if (processedNumber !== null) {
                    if (uniqueNumbers.has(processedNumber)) {
                        rowsToRemove.add(rowIndex); // Marquer la ligne pour suppression
                    } else {
                        uniqueNumbers.add(processedNumber);
                        row[colIndex] = processedNumber;
                    }
                }
            }
        });
    });

    // Supprimer les lignes en doublon
    const cleanedData = data.filter((row, index) => !rowsToRemove.has(index));

    // Convertir les données JSON de retour en feuille de calcul
    const newWorksheet = xlsx.utils.aoa_to_sheet(cleanedData);

    // Restaurer les 7 premières largeurs des colonnes
    if (colStyle[sheetName] && colStyle[sheetName].length > 0) {
        newWorksheet['!cols'] = colStyle[sheetName];
    }

    // Restaurer les styles des cellules
    for (let cellRef in cellStyle[sheetName]) {
        if (newWorksheet[cellRef]) {
            newWorksheet[cellRef].s = cellStyle[sheetName][cellRef];
        }
    }

    workbook.Sheets[sheetName] = newWorksheet;

    // Ajouter les erreurs au journal des feuilles
    if (errors.length > 0) {
        sheetLogs[sheetName] = errors;
    }

    console.log(`Feuille "${sheetName}" traitée dans ${filePath}`);
    console.log(`  Nombre de cellules corrigées : ${correctedCount}`);
    console.log(`  Nombre de lignes supprimées en raison de doublons : ${rowsToRemove.size}`);
    console.log(`  Détails sur les raisons d'invalidité :`);
    console.log(`    - Type de données invalide : ${invalidReasons.invalidType}`);
    console.log(`    - Caractères invalides détectés : ${invalidReasons.invalidCharacters}`);
    console.log(`    - Format invalide : ${invalidReasons.invalidFormat}`);
}

// Fonction pour traiter un fichier
async function processFile(filePath) {
    // Initialiser le journal des feuilles pour ce fichier
    let sheetLogs = {};

    // Lire le fichier Excel
    const workbook = xlsx.readFile(filePath, { cellStyles: true, cellText: true });

    for (const sheetName of workbook.SheetNames) {
        console.log(`\nTraitement de la feuille "${sheetName}" dans ${filePath}`);
        await processSheet(workbook, sheetName, filePath, sheetLogs);
    }

    // Déterminer le chemin du fichier de sortie
    const outputFilePath = path.join(outputDir, path.basename(filePath));

    // Écrire le nouveau fichier Excel
    xlsx.writeFile(workbook, outputFilePath);

    // Écrire les erreurs dans un fichier texte uniquement si des erreurs existent
    writeErrorLog(filePath, sheetLogs);

    console.log(`Fichier ${filePath} traité avec succès!\n\n`);
}

// Lire tous les fichiers dans le répertoire 'convertir'
fs.readdir(inputDir, async (err, files) => {
    if (err) {
        console.error('Erreur lors de la lecture du répertoire', err);
        return;
    }

    // Filtrer les fichiers Excel
    const excelFiles = files.filter(file => file.endsWith('.xlsx') || file.endsWith('.xlsm') || file.endsWith('.xls'));

    if (excelFiles.length === 0) {
        console.log('Aucun fichier Excel trouvé dans le répertoire de conversion.');
        return;
    }

    let totalFiles = excelFiles.length;

    for (const file of excelFiles) {
        if (file.startsWith('~$')) {
            totalFiles--;
        }
    }

    console.log(`\nNombre total de fichiers à traiter : ${totalFiles}\n`);

    // Traiter chaque fichier séquentiellement
    for (const file of excelFiles) {
        const filePath = path.join(inputDir, file);

        if (!file.startsWith('~$')) {
            console.log(`Traitement du fichier : ${filePath}`);
            await processFile(filePath);
        }
    }

    console.log('Tous les fichiers ont été traités avec succès!');
});
