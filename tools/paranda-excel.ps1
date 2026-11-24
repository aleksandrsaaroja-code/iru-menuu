# Parandab kirjavead menuu Exceli failides (X:\Toitlustusosakond\Menuud\Nadal 1-5\*.xlsm).
# Reeglid: tools\parandused.txt. ESIMENE KORD ILMA VOTMETA = ainult naitab, mida muudaks (midagi ei salvestata).
# Salvestamiseks:  .\paranda-excel.ps1 -Salvesta     (teeb enne varukoopia kausta "varukoopia")
# Sulge enne Excelis menuu failid. Vajab Excelit (COM). Makrod EI kaivitu.
param([switch]$Salvesta)
$ErrorActionPreference = 'Stop'

$rules = @{}
Get-Content (Join-Path $PSScriptRoot 'parandused.txt') -Encoding UTF8 | ForEach-Object {
  if ($_ -match '^\s*([^#=]+)=(.+)$') { $rules[$matches[1].Trim().ToUpper()] = $matches[2].Trim().ToUpper() }
}
function Fix([string]$t) {
  $t = [regex]::Replace($t, '\s+', ' ').Trim()
  $t = [regex]::Replace($t, '\s+,', ',')
  $t = [regex]::Replace($t, ',(?=\S)', ', ')
  foreach ($k in $rules.Keys) {
    $t = [regex]::Replace($t, '(?<![\p{L}-])' + [regex]::Escape($k) + '(?![\p{L}-])', $rules[$k])
  }
  return $t
}

$root = Get-ChildItem 'X:\Toitlustusosakond' -Directory | Where-Object { $_.Name -like 'Men*' } | Select-Object -First 1
$dir = (Get-ChildItem $root.FullName -Directory | Where-Object { $_.Name -like '*1-5' } | Select-Object -First 1).FullName
$files = Get-ChildItem $dir -Filter '*.xlsm' | Where-Object { $_.Name -notlike '~*' } | Sort-Object Name
if (-not $files) { throw 'Menuu faile ei leitud' }

if ($Salvesta) {
  $bk = Join-Path $dir 'varukoopia'
  New-Item -ItemType Directory -Force $bk | Out-Null
  $stamp = (Get-Date).ToString('yyyyMMdd-HHmm')
  foreach ($f in $files) { Copy-Item $f.FullName (Join-Path $bk ($f.BaseName + '_' + $stamp + '.xlsm')) }
  Write-Host "Varukoopiad: $bk"
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 3   # makrod keelatud
$total = 0
try {
  foreach ($f in $files) {
    $wb = $xl.Workbooks.Open($f.FullName, 0, (-not $Salvesta))
    if ($Salvesta -and $wb.ReadOnly) { Write-Host ("HOIATUS: " + $f.Name + " on avatud kellelgi, jatan vahele"); $wb.Close($false); continue }
    $changed = 0
    foreach ($ws in $wb.Worksheets) {
      $first = $ws.UsedRange.Row
      $last = $first + $ws.UsedRange.Rows.Count - 1
      for ($r = $first; $r -le $last; $r++) {
        $c = $ws.Cells.Item($r, 2)
        if ($c.HasFormula) { continue }
        $v = $c.Value2
        if ($v -is [string]) {
          $n = Fix $v
          if ($n -cne $v) {
            Write-Host ($f.BaseName + ' / ' + $ws.Name + ' / B' + $r + ' : "' + $v + '" -> "' + $n + '"')
            if ($Salvesta) { $c.Value2 = $n }
            $changed++
          }
        }
      }
    }
    if ($Salvesta -and $changed) { $wb.Save() }
    $wb.Close($false)
    $total += $changed
  }
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
if ($Salvesta) { Write-Host "Parandatud ridu kokku: $total (salvestatud)" } else { Write-Host "Muudaks ridu kokku: $total (EI salvestanud; kaivita -Salvesta, et salvestada)" }
