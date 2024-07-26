const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Fonction pour lire un fichier Excel et extraire les adresses e-mails
function extractEmailsFromSheet(sheet) {
	const emailSet = new Set();

	// Obtenir les données de la feuille sous forme de tableau JSON
	const sheetData = xlsx.utils.sheet_to_json(sheet, { header: 1 });

	// Parcourir chaque ligne de la feuille
	sheetData.forEach(row => {
		row.forEach(cell => {
			// Utiliser une expression régulière pour trouver les e-mails
			const emailPattern = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
			const match = cell && cell.toString().match(emailPattern);
			if (match) {
				emailSet.add(match[0]);
			}
		});
	});

	return Array.from(emailSet);
}

// Fonction pour écrire les e-mails dans un fichier texte
function writeEmailsToFile(emailMap, outputFile) {
	let output = '';

	emailMap.forEach((emails, sheetName) => {
		output += `\n\n=== Feuille : ${sheetName} ===\n`;
		output += emails.join('\n') + '\n';
	});

	fs.writeFileSync(outputFile, output);
    console.log(`Emails extraits avec succès dans ${outputFile}`);
}

// Répertoire contenant les fichiers Excel
const directoryPath = path.join(__dirname, 'mail_export');

// Fichier de sortie
const outputFilePath = path.join(__dirname + '/mail_export', 'emails.txt');

// Extraire les e-mails et les écrire dans le fichier de sortie
(async function() {
	const emailMap = new Map();

	// Lire tous les fichiers dans le répertoire
	const files = fs.readdirSync(directoryPath);

	for (const file of files) {
		const filePath = path.join(directoryPath, file);

		// Vérifier si le fichier est un fichier Excel
		if (filePath.endsWith('.xlsx') || filePath.endsWith('.xls')) {
			console.log(`Traitement du fichier : ${file}`);
			const workbook = xlsx.readFile(filePath);

			// Parcourir chaque feuille du fichier Excel
			workbook.SheetNames.forEach(sheetName => {
				console.log(`  Lecture de la feuille : ${sheetName}`);
				const worksheet = workbook.Sheets[sheetName];
				
				// Extraire les e-mails de la feuille
				const emails = extractEmailsFromSheet(worksheet);

				// Stocker les e-mails de la feuille dans la map
				emailMap.set(sheetName, emails);
			});
		} else {
			console.log(`  Le fichier ${file} n'est pas un fichier Excel, il est ignoré.`);
		}
	}

	// Écrire les e-mails dans le fichier de sortie
	writeEmailsToFile(emailMap, outputFilePath);
})();
