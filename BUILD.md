# SuperPUWEtty2 - Build Anleitung

## Voraussetzungen

1. **Windows 10/11**
2. **Visual Studio Build Tools** oder **Visual Studio 2019/2022**
3. **PuTTY** installiert (https://www.putty.org/)
4. **Git** (optional, für Updates)

## Installation der Build Tools

Falls MSBuild noch nicht installiert ist:

### Option 1: Visual Studio Build Tools (empfohlen, ~6 GB)
```powershell
# Download und Installation
Invoke-WebRequest -Uri "https://aka.ms/vs/17/release/vs_buildtools.exe" -OutFile "vs_buildtools.exe"
.\vs_buildtools.exe --add Microsoft.VisualStudio.Workload.MSBuildTools --quiet --wait
```

### Option 2: Visual Studio Community (vollständige IDE, ~10 GB)
Download: https://visualstudio.microsoft.com/de/downloads/
- Workload auswählen: ".NET Desktop-Entwicklung"

## Projekt bauen

### Methode 1: Batch-Script (einfach)
```cmd
build.bat
```

### Methode 2: PowerShell-Script (empfohlen)
```powershell
# Falls Execution Policy Fehler auftritt:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Build starten:
.\build.ps1
```

### Methode 3: Manuell mit MSBuild
```cmd
# NuGet Pakete wiederherstellen
nuget restore SuperPUWEtty2.sln

# Projekt bauen
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" SuperPUWEtty2.sln /t:Rebuild /p:Configuration=Release /p:Platform="Any CPU"
```

## Ausgabe

Nach erfolgreichem Build findest du die EXE hier:
```
SuperPUWEtty2\bin\Release\SuperPUWEtty2.exe
```

## Git Update (für spätere Änderungen)

```bash
git pull origin main
.\build.ps1
```

## Fehlerbehandlung

### Problem: "MSBuild not found"
**Lösung:** Installiere Visual Studio Build Tools (siehe oben)

### Problem: "NuGet not found"
**Lösung:**
```powershell
Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile "nuget.exe"
```

### Problem: Build-Fehler wegen fehlender Dependencies
**Lösung:**
```cmd
# Lösche alte Build-Artefakte
rmdir /s /q SuperPUWEtty2\bin SuperPUWEtty2\obj

# NuGet Pakete neu installieren
nuget restore SuperPUWEtty2.sln -Force

# Neu bauen
.\build.ps1
```

### Problem: "Access Denied" bei PowerShell Script
**Lösung:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Problem: Installer-Projekt (SuperPUWEtty2Installer) schlägt fehl
**Lösung:** Das Installer-Projekt ist optional. Wenn du es nicht brauchst:
1. Öffne `SuperPUWEtty2.sln` in einem Texteditor
2. Kommentiere die Zeilen mit "SuperPUWEtty2Installer" aus
3. Oder installiere WiX Toolset: https://wixtoolset.org/

## Schnellstart (komplette Anleitung)

```powershell
# 1. Build Tools installieren (einmalig)
Invoke-WebRequest -Uri "https://aka.ms/vs/17/release/vs_buildtools.exe" -OutFile "vs_buildtools.exe"
.\vs_buildtools.exe --add Microsoft.VisualStudio.Workload.MSBuildTools --quiet --wait

# 2. Repository klonen/ziehen
git pull

# 3. Bauen
.\build.ps1
```

## Entwicklung

Für die Entwicklung empfehle ich Visual Studio Community:
1. Öffne `SuperPUWEtty2.sln` in Visual Studio
2. Drücke `F5` zum Debuggen
3. Drücke `Ctrl+Shift+B` zum Bauen

## Support

Bei Problemen:
1. Überprüfe, ob alle Voraussetzungen installiert sind
2. Lösche `bin/` und `obj/` Ordner
3. Führe `nuget restore` erneut aus
4. Baue die Solution neu
