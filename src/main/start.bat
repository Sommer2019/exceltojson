@echo off
setlocal enabledelayedexpansion

:: Pfad zum Java-Programm und zur Excel-Datei
set "javaPath=java"
set "excelToJsonClass=src/main/java/ExcelToJsonConverter"
set "folder=out/tests"

:: Schritt 1: Führe das Java-Programm aus, um die Excel-Datei in JSON zu konvertieren
echo Starte ExcelToJsonConverter...
%javaPath% -cp . %excelToJsonClass%

:: Überprüfe, ob das Java-Programm erfolgreich ausgeführt wurde
if %ERRORLEVEL% neq 0 (
    echo Fehler beim Ausführen von ExcelToJsonConverter.java
    exit /b %ERRORLEVEL%
)

:: Schritt 2: Alle JSON-Dateien im Ordner durchgehen und Tests ausführen
for %%f in (%folder%\*.json) do (
    echo Starte Test mit Datei: %%f
    :: Test ausführen und die JSON-Datei als Umgebungsvariable übergeben
    npx cypress run --env fixture=%%f

    :: Warten, bis der Test abgeschlossen ist, bevor der nächste startet
    echo Warte auf den nächsten Test...
    timeout /t 5 >nul
)

echo Alle Tests abgeschlossen!
pause
