$WebFile = "https://somewhere.on.the.internet/scanner.ps1"

$LocalFile = "$env:APPDATA\scanner.ps1"

Invoke-WebRequest -Uri $WebFile -OutFile $LocalFile   

PowerShell -ExecutionPolicy bypass -File  $LocalFile