param([string]$Root = (Split-Path -Parent $MyInvocation.MyCommand.Path))

# ModernizePandocConfig.ps1
# Brings the Pandoc command lines in EdSharp.ini (and EdSharp.inix, if present)
# up to date for Pandoc 2.0+/3.x. It is idempotent and only rewrites lines that
# still contain an old token, and it backs the file up once before its first
# change. Run automatically by FetchConvertTools.ps1 whenever Pandoc is
# installed or upgraded; safe to run by hand too.
#
# Changes applied only to lines that invoke pandoc.exe:
#   " -S"            (the removed --smart flag)   -> deleted
#   markdown_github  (old GitHub-flavored name)   -> gfm
#   --reference-docx (renamed option)             -> --reference-doc
#
# These are the breaking renames between Pandoc 1.x and 2.0+. Smart typography
# is now an input extension (on by default for pandoc's markdown), so dropping
# -S is correct; gfm and --reference-doc are the current spellings.

function Update-File($path) {
  if (-not (Test-Path -LiteralPath $path)) { return }
  $text = [IO.File]::ReadAllText($path)               # preserve ANSI/default encoding
  if ($text -notmatch '(?m)pandoc\.exe') { return }
  $changed = $false
  $out = New-Object System.Text.StringBuilder
  foreach ($line in ($text -split "(?<=`n)")) {        # keep line endings
    $new = $line
    if ($line -match 'pandoc\.exe') {
      $new = $new -replace ' -S(?= )', ''              # " -S" before a space
      $new = $new -replace ' -S(?=\r?$)', ''           # " -S" at line end
      $new = $new -replace 'markdown_github', 'gfm'
      $new = $new -replace '--reference-docx', '--reference-doc'
    }
    if ($new -ne $line) { $changed = $true }
    [void]$out.Append($new)
  }
  if ($changed) {
    $bak = $path + ".pandoc-bak"
    if (-not (Test-Path -LiteralPath $bak)) { Copy-Item -LiteralPath $path -Destination $bak -Force }
    [IO.File]::WriteAllText($path, $out.ToString())
    Write-Host ("[pandoc] modernized flags in {0} (backup: {1})." -f (Split-Path -Leaf $path), (Split-Path -Leaf $bak))
  } else {
    Write-Host ("[pandoc] {0} already up to date." -f (Split-Path -Leaf $path))
  }
}

Update-File (Join-Path $Root "EdSharp.ini")
Update-File (Join-Path $Root "EdSharp.inix")
Write-Host "[pandoc] Note: if EdSharp keeps an active EdSharp.ini in a separate user data folder, apply the same change there (or copy this folder's EdSharp.ini over it)."
