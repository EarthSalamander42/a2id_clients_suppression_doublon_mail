# Utilisation du script d'extraction de mails depuis un fichier Excel

Ce guide explique comment utiliser le script `export_mail_doublon.js` pour extraire des adresses e-mail à partir d'un fichier Excel et supprimer les doublons.

## Prérequis
Avant d'utiliser le script, assurez-vous d'avoir les éléments suivants :
- NodeJS installé sur votre machine
- L'autorisation d'éxecuter des scritps sur votre machine. Utiliser la commande suivante pour régler le soucis : ```Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser```

## Étapes

1. Installer Github Desktop.
2. Cloner ce répertoire.
3. Assurez-vous que votre fichier Excel est dans le dossier ```convertir```.
4. Double-cliquez sur le raccourci nommé ```moulinette.bat```
5. Le script va supprimer les lignes en doublons en se basant sur les numéros de téléphones du fichier Excel et les enregistrer dans un nouveau fichier excel dans le répertoire `completer`.
6. Le script va également rajouter un 0 si manquant au début d'un numéro de téléphone, supprimer les espaces en trop, et supprimer les points.
