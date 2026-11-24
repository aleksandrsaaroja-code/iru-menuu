# Vaatab toidulehte; kui see salvestatakse, uuendab kokkade äpi arvud (umbes 25 sekundi pärast).
$file = "X:\Statistika\Toidulehed_2026 UUS.xlsx"
$last = $null
try { $last = (Get-Item $file).LastWriteTime } catch {}
while ($true) {
  Start-Sleep -Seconds 15
  try { $t = (Get-Item $file).LastWriteTime } catch { continue }
  if ($t -ne $last) {
    Start-Sleep -Seconds 10            # oota, kuni salvestamine lõpeb
    try { $last = (Get-Item $file).LastWriteTime } catch { continue }
    & "$PSScriptRoot\uuenda_arvud.bat"
  }
}
