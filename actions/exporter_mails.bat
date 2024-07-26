@echo off
cd /d "%~dp0"
node ../src/export_mail_doublon.js
notepad ../mail_export/mails.txt
pause
