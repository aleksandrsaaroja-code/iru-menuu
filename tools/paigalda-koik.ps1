# Seadistab kogu automaatika uhe korraga (ei vaja admini oigusi):
#  1) toob uusimad failid
#  2) lisab Avalda-menuu.ps1-sse kirjavigade automaatse parandamise (varukoopia: Avalda-menuu.ps1.bak)
#  3) paneb arvude valvuri kaivituma igal sisselogimisel (Windowsi kaivituskaust) ja kaivitab selle kohe
# Kaivita:  cd $HOME\iru-menuu ; .\tools\paigalda-koik.ps1
$ErrorActionPreference = 'Continue'   # git kirjutab tavateated (nt "From ...") veavoogu, need ei ole vead
$repo = Split-Path $PSScriptRoot
Write-Host "1/3 Toon uusimad failid ..."
git -C $repo pull --rebase --autostash
if (-not (Test-Path (Join-Path $repo 'tools\kontrolli_menuu.py'))) { throw 'tools\kontrolli_menuu.py puudub, kontrolli git pull' }

Write-Host "2/3 Menuu avaldamise skript ..."
$av = Get-ChildItem "$HOME\Desktop" -Recurse -Filter 'Avalda-menuu.ps1' -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $av) { Write-Host "  Avalda-menuu.ps1 ei leitud Tolaualt, jatan vahele" }
else {
  $avDir = $av.DirectoryName
  git -C $avDir pull --rebase --autostash 2>&1 | Out-Null     # tools\ kaust jouab ka sinna
  $txt = [IO.File]::ReadAllText($av.FullName)
  if ($txt -match 'kontrolli_menuu\.py') { Write-Host "  kirjavigade parandus on juba lisatud" }
  else {
    Copy-Item $av.FullName ($av.FullName + '.bak') -Force
    $line = "  try { py (Join-Path `$PSScriptRoot 'tools\kontrolli_menuu.py') --paranda | Out-Null; Add-Content `$log ((Get-Date).ToString('s') + ' kirjavigade kontroll: ' + (Get-Content (Join-Path `$PSScriptRoot 'kirjavead.txt') -TotalCount 1)) } catch {}"
    $pat = "(?m)^(\s*&\s*\(Join-Path \`$PSScriptRoot 'Uuenda-menuu\.ps1'\)\s*)$"
    if ($txt -notmatch $pat) { throw 'Ei leidnud Avalda-menuu.ps1-st rida Uuenda-menuu.ps1 kohta' }
    $ev = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $m.Groups[1].Value.TrimEnd() + "`r`n" + $line }
    $new = [regex]::Replace($txt, $pat, $ev)
    [IO.File]::WriteAllText($av.FullName, $new, (New-Object Text.UTF8Encoding($false)))
    Write-Host "  lisatud kirjavigade automaatne parandus (varukoopia .bak)"
  }
}

Write-Host "3/3 Arvude valvur ..."
$cmd = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup\Iru kook valvur.cmd'
$vp = Join-Path $repo 'tools\valvur.ps1'
Set-Content -Path $cmd -Encoding ASCII -Value ('powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $vp + '"')
$run = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like '*valvur.ps1*' }
if (-not $run) { Start-Process 'cmd.exe' -ArgumentList ('/c "' + $cmd + '"') -WindowStyle Hidden; Write-Host "  valvur kaivitatud" } else { Write-Host "  valvur juba tootab" }
Write-Host "VALMIS. Menuu ja arvud uuenevad nuud ise."
