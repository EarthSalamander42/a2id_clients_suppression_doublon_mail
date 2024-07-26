const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Répertoires
const inputDir = path.join(__dirname, 'tel_convertir');
const outputDir = path.join(__dirname, 'tel_completer');

// Vérifie si le répertoire de sortie existe, sinon le crée
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir);
}

// Fonction pour vérifier si une colonne contient des numéros de téléphone
function isPhoneNumberColumn(data, column) {
  const phoneNumberRegex = /^[0-9]{9,15}$/; // Ajuster selon les formats de numéros de téléphone
  let phoneCount = 0;
  let nonEmptyCount = 0;

  data.forEach((row) => {
    const value = row[column];
    if (value !== undefined && value !== null) {
      nonEmptyCount++;
      if (typeof value === 'string' && phoneNumberRegex.test(value.replace(/\s+/g, ''))) {
        phoneCount++;
      }
    }
  });

  // On considère que c'est une colonne de numéros de téléphone si plus de 80% des valeurs non vides correspondent
  return (phoneCount / nonEmptyCount) > 0.8;
}

// Fonction pour traiter un fichier
async function processFile(filePath) {
  // Lire le fichier Excel
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convertir la feuille en JSON
  const data = xlsx.utils.sheet_to_json(worksheet);

  // Déterminer les colonnes de numéros de téléphone
  const columns = Object.keys(data[0]);
  const phoneNumberColumns = columns.filter(column => isPhoneNumberColumn(data, column));

  // Parcourir les numéros de téléphone et ajouter le 0 si nécessaire
  data.forEach((row) => {
    phoneNumberColumns.forEach((column) => {
      if (row[column] && typeof row[column] === 'string') {
        if (!row[column].startsWith('0')) {
          row[column] = '0' + row[column];
        }
      }
    });
  });

  // Convertir de nouveau en feuille de calcul
  const newWorksheet = xlsx.utils.json_to_sheet(data);
  workbook.Sheets[sheetName] = newWorksheet;

  // Déterminer le chemin du fichier de sortie
  const outputFilePath = path.join(outputDir, path.basename(filePath));

  // Écrire le nouveau fichier Excel
  xlsx.writeFile(workbook, outputFilePath);

  console.log(`Fichier traité : ${outputFilePath}`);
}

// Lire tous les fichiers dans le répertoire 'convertir'
fs.readdir(inputDir, async (err, files) => {
  if (err) {
    console.error('Erreur lors de la lecture du répertoire', err);
    return;
  }

  // Filtrer les fichiers Excel
  const excelFiles = files.filter(file => file.endsWith('.xlsx'));

  // Traiter chaque fichier séquentiellement
  for (const file of excelFiles) {
    const filePath = path.join(inputDir, file);
    console.log(`Traitement du fichier : ${filePath}`);
    await processFile(filePath);
  }

  console.log('Tous les fichiers ont été traités avec succès!');
});
