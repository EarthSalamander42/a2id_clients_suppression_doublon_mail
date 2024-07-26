# Utilisation du script d'extraction de mails depuis un fichier Excel

Ce guide explique comment utiliser le script `export_mail_doublon.js` pour extraire des adresses e-mail à partir d'un fichier Excel et supprimer les doublons.

## Prérequis
Avant d'utiliser le script, assurez-vous d'avoir les éléments suivants :
- NodeJS installé sur votre machine
- L'autorisation d'éxecuter des scritps sur votre machine. Utiliser la commande suivante pour régler le soucis : ```Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser```

## Étapes

1. Installer Github Desktop.
2. Cloner ce répertoire.
3. Assurez-vous que votre fichier Excel contenant les adresses e-mail est dans le dossier ```mail_export```.
4. Double-cliquez sur le raccourci nommé ```exporter.bat```
5. Le script va extraire les adresses e-mail du fichier Excel et les enregistrer dans un nouveau fichier `mails.txt`.
6. Les doublons seront automatiquement supprimés par feuille Excel dans le fichier `mails.txt`.

Et voilà ! Vous avez maintenant extrait les adresses e-mail depuis votre fichier Excel et supprimé les doublons.
