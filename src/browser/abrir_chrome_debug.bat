@echo off
echo Abrindo o Chrome em Modo de Depuracao (Porta 9222)...
echo Pode fechar esta janela preta apos o Chrome abrir.

:: Cria a pasta de perfil se ela nao existir para evitar erros
if not exist "C:\selenium\ChromeProfile" mkdir "C:\selenium\ChromeProfile"

:: Comando para abrir o Chrome
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\selenium\ChromeProfile"

exit