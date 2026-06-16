# FetchUde.ps1 -- best-effort download of the UDE.CSharp encoding-detection
# library (Ude.dll) used by EdSharp's content-based charset autodetection.
#
# UDE.CSharp 1.1.0 is the Mozilla Universal Charset Detector port. It targets
# .NET 4.0 as a single dependency-free assembly (Ude.dll), so it references
# cleanly under the bare csc build with no NuGet restore. The .nupkg is just a
# zip; this script downloads it and extracts Ude.dll beside the build.
#
# This step NEVER fails the build. If Ude.dll cannot be fetched, the compile
# proceeds without the HAVEUDE symbol and EdSharp still opens files using
# byte-order-mark detection with a UTF-8-with-BOM default; only content-based
# detection of BOM-less, non-UTF-8 files is unavailable until Ude.dll is added.
#
# To use a newer/alternative library instead (for example UTF.Unknown), drop
# its assembly in as Ude.dll-compatible, or adjust the /reference in
# BuildEdSharp.cmd and the Ude.CharsetDetector call in EdSharp.cs accordingly.

$ErrorActionPreference = "Stop"
try {
    if (Test-Path "Ude.dll") { Write-Host "Ude.dll already present; skipping fetch."; exit 0 }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $url = "https://www.nuget.org/api/v2/package/UDE.CSharp/1.1.0"
    $nupkg = Join-Path $env:TEMP "UDE.CSharp.1.1.0.nupkg.zip"
    Write-Host "Downloading UDE.CSharp 1.1.0 ..."
    Invoke-WebRequest -Uri $url -OutFile $nupkg -UseBasicParsing
    $dir = Join-Path $env:TEMP "ude_extract"
    if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [IO.Compression.ZipFile]::ExtractToDirectory($nupkg, $dir)
    $dll = Get-ChildItem -Path $dir -Recurse -Filter "Ude.dll" | Select-Object -First 1
    if ($dll) {
        Copy-Item $dll.FullName "Ude.dll" -Force
        Write-Host "Ude.dll fetched and placed beside the build."
    } else {
        Write-Host "Ude.dll not found inside the package; content detection will be disabled."
    }
}
catch {
    Write-Host ("UDE fetch skipped (" + $_.Exception.Message + "); content detection will be disabled.")
}
exit 0
