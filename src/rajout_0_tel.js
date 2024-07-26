const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Répertoires
const inputDir = path.join(__dirname, '../telephone', 'convertir');
const outputDir = path.join(__dirname, '../telephone', 'completer');

// Vérifie si le répertoire de sortie existe, sinon le crée
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
}

// Regex pour les numéros de téléphone valides
const phoneNumberRegex = /^[0-9]{9,15}$/;
const invalidCharactersRegex = /[^0-9\s.]/;

// Fonction pour traiter un fichier
async function processFile(filePath) {
    // Lire le fichier Excel
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convertir la feuille en JSON
    const data = xlsx.utils.sheet_to_json(worksheet);

    let correctedCount = 0;
    let invalidCount = 0;

    // Fonction pour traiter un numéro de téléphone
    function processPhoneNumber(phoneNumber) {
        if (typeof phoneNumber !== 'string') {
            console.log(`Le numéro n'est pas une chaîne de caractères : ${phoneNumber} (${typeof phoneNumber}). Conversion en chaîne de caractères.`);
            phoneNumber = phoneNumber.toString();
        }

        // Vérifier s'il y a des caractères invalides
        if (invalidCharactersRegex.test(phoneNumber)) {
            console.log(`Valeur invalide trouvée pour le numéro : ${phoneNumber}`);
            invalidCount++;
            return null; // Indiquer que la valeur est invalide
        }

        let cleanedPhoneNumber = phoneNumber.replace(/[\s.]+/g, ''); // Retirer les espaces et les points

        if (cleanedPhoneNumber !== phoneNumber) {
            console.log(`Points ou espaces supprimés du numéro : ${phoneNumber}`);
        }

        if (phoneNumberRegex.test(cleanedPhoneNumber)) {
            if (!cleanedPhoneNumber.startsWith('0')) {
                cleanedPhoneNumber = '0' + cleanedPhoneNumber;
                console.log(`Numéro de téléphone corrigé : ${cleanedPhoneNumber}`);
            }
            correctedCount++;
            return cleanedPhoneNumber;
        } else {
            console.log(`Valeur invalide trouvée pour le numéro : ${phoneNumber}`);
            invalidCount++;
            return null; // Indiquer que la valeur est invalide
        }
    }

    // Parcourir les numéros de téléphone dans les colonnes "Téléphone" et "Mobile"
    data.forEach((row) => {
        let phoneNumber = row['Téléphone'] || row['Mobile'];

        if (phoneNumber) {
            let processedNumber = processPhoneNumber(phoneNumber);

			if (processedNumber !== null) {
                if (row['Téléphone'] !== undefined) {
                    row['Téléphone'] = processedNumber;
                }

				if (row['Mobile'] !== undefined) {
                    row['Mobile'] = processedNumber;
                }
            }
        }
    });

    // Convertir de nouveau en feuille de calcul
    const newWorksheet = xlsx.utils.json_to_sheet(data);
    workbook.Sheets[sheetName] = newWorksheet;

    // Déterminer le chemin du fichier de sortie
    const outputFilePath = path.join(outputDir, path.basename(filePath));

    // Écrire le nouveau fichier Excel
    xlsx.writeFile(workbook, outputFilePath);

    console.log(`Fichier traité : ${outputFilePath}`);
    console.log(`Nombre de cellules corrigées : ${correctedCount}`);
    console.log(`Nombre de cellules invalides : ${invalidCount}`);
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

    // Traiter chaque fichier séquentiellement
    for (const file of excelFiles) {
        const filePath = path.join(inputDir, file);
        console.log(`Traitement du fichier : ${filePath}`);
        await processFile(filePath);
    }

    console.log('Tout les fichiers ont été traités avec succès!');
});
