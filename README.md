# PortableWinFormsRecorder v1 (C# / WinForms)

Includes:
- GUI (Record / Stop / Run)
- UI Inspector (live hover + overlay highlight)
- Assertions: AssertExists / AssertTextEquals / AssertTextContains
- Image-based fallback for weak/custom controls (captures element screenshots to `script_assets/`)
- CLI mode (record/run/print)

## Build
```bat
cd PortableWinFormsRecorder
dotnet restore
dotnet build -c Release
```

## Run GUI
```bat
dotnet run -c Release
```

## CLI record
```bat
dotnet run -c Release -- record --process MyWinFormsApp --out script.json --capture-images
```

## CLI run
```bat
dotnet run -c Release -- run --script script.json --data data.csv
```

## Assertions in script.json
```json
{ "action": "AssertExists", "target": { "automationId": "lblStatus" } }
{ "action": "AssertTextEquals", "target": { "automationId": "lblStatus" }, "value": "Ready" }
{ "action": "AssertTextContains", "target": { "automationId": "lblStatus" }, "value": "Error" }
```

## Portable publish (single EXE)
```bat
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:IncludeNativeLibrariesForSelfExtract=true
```

Ship:
- Published EXE
- `script.json`
- `data.csv`
- `script_assets/` folder next to the script (for image fallback)