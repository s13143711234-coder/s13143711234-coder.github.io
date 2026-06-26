# update-smm.ps1
# Run after SMM DATA PRO refresh + Excel save. Reads 4 datasets and pushes to GitHub.

$Excel = "D:\新桌面\ks-锂电资料\20. AI应用\【看板3】锂电材料价格-20260626.xlsx"
$Html  = "C:\Users\97484\Claude\股票池跟踪\publish\lidian-preview.html"
Set-Location (Split-Path $Html)

function fv($v) {
    if ($null -eq $v) { return "null" }
    try {
        $d = [double]$v
        if ([double]::IsNaN($d) -or [double]::IsInfinity($d)) { return "null" }
        if ([Math]::Abs($d) -ge 100) { return ([Math]::Round($d,0)).ToString() }
        return ([Math]::Round($d,4)).ToString()
    } catch { return "null" }
}
function Jarr($items) { "[" + ($items -join ",") + "]" }

function ReadCols($ws, [hashtable]$cols) {
    $list = [System.Collections.Generic.List[hashtable]]::new()
    $r = 5
    while ($r -le 2000) {
        $dt = $ws.Cells.Item($r,1).Text
        if ($dt -eq "") { break }
        $row = @{ _d = $dt }
        foreach ($kv in $cols.GetEnumerator()) { $row[$kv.Key] = $ws.Cells.Item($r,$kv.Value).Value2 }
        $list.Add($row); $r++
    }
    $a = $list.ToArray(); [array]::Reverse($a); return $a
}

function ConvDate($s, $sep) {
    if ($s -match "^(\d{4})-(\d{2})-\d{2}$") { return $Matches[1] + $sep + [int]$Matches[2] }
    return $s
}

Write-Host "Opening Excel (read-only)..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($Excel, 0, $true)

# CELLPRICE
Write-Host "CELLPRICE..."
$wsW = $wb.Sheets.Item("SMM周度")
$cp = ReadCols $wsW @{s280=2;s314=3;p174=5;t158=6}
$cpD = Jarr ($cp | ForEach-Object { '"' + $_._d + '"' })
$newCP = "const CELLPRICE={dates:" + $cpD + ",s280:" + (Jarr ($cp|%{fv $_.s280})) + ",s314:" + (Jarr ($cp|%{fv $_.s314})) + ",p174:" + (Jarr ($cp|%{fv $_.p174})) + ",t158:" + (Jarr ($cp|%{fv $_.t158})) + "};"

# LFPFEE
Write-Host "LFPFEE..."
$wsLF = $wb.Sheets.Item("SMM周度-铁锂加工费")
$lf = ReadCols $wsLF @{d250=2;d260=3}
$lfD = Jarr ($lf | ForEach-Object { '"' + $_._d + '"' })
$newLF = "const LFPFEE={dates:" + $lfD + ",d250:" + (Jarr ($lf|%{fv $_.d250})) + ",d260:" + (Jarr ($lf|%{fv $_.d260})) + "};"

# INVENTORY + LITCO
Write-Host "INVENTORY + LITCO..."
$wsM = $wb.Sheets.Item("SMM月度")
$mo = ReadCols $wsM @{lfp=2;ncm=3;es=4;ev=5;prod=10;demand=11;import=12;export=13;balance=14}
$invD = Jarr ($mo | ForEach-Object { '"' + (ConvDate $_._d "/") + '"' })
$newINV = "const INVENTORY={dates:" + $invD + ",lfp:" + (Jarr ($mo|%{fv $_.lfp})) + ",ncm:" + (Jarr ($mo|%{fv $_.ncm})) + ",es:" + (Jarr ($mo|%{fv $_.es})) + ",ev:" + (Jarr ($mo|%{fv $_.ev})) + "};"
$litD = Jarr ($mo | ForEach-Object { '"' + (ConvDate $_._d "-") + '"' })
$newLIT = "const LITCO={dates:" + $litD + ",demand:" + (Jarr ($mo|%{fv $_.demand})) + ",import:" + (Jarr ($mo|%{fv $_.import})) + ",export:" + (Jarr ($mo|%{fv $_.export})) + ",balance:" + (Jarr ($mo|%{fv $_.balance})) + ",prod:" + (Jarr ($mo|%{fv $_.prod})) + "};"

$wb.Close($false); $xl.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "Excel closed."

# Update HTML
Write-Host "Updating HTML..."
$c = [IO.File]::ReadAllText($Html, [Text.Encoding]::UTF8)
function RC($html,$name,$nv) { [regex]::Replace($html, ([regex]::Escape("const $name={") + "[^\r\n]+\};"), $nv) }
$c = RC $c "CELLPRICE" $newCP
$c = RC $c "LFPFEE"    $newLF
$c = RC $c "INVENTORY" $newINV
$c = RC $c "LITCO"     $newLIT
[IO.File]::WriteAllText($Html, $c, [Text.Encoding]::UTF8)
Write-Host "HTML updated."

# Git
$today = Get-Date -Format "yyyy-MM-dd"
git add lidian-preview.html
git commit -m "SMM auto-update $today"
$maxTries=20
for ($i=1;$i -le $maxTries;$i++) {
    $r2 = git push 2>&1
    if ($LASTEXITCODE -eq 0) { Write-Host "Push OK (attempt $i)"; break }
    Write-Host "Retry $i/$maxTries in 18s..."; if ($i -lt $maxTries) { Start-Sleep -Seconds 18 }
}
Write-Host "Done: https://s13143711234-coder.github.io/lidian-preview.html"
