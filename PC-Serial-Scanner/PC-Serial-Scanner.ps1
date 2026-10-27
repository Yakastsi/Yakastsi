#Requires -Version 5.1
<#
.SYNOPSIS
  Сбор серийных номеров комплектующих ПК (Windows 10/11).
.NOTES
  Yakastsi — rgb.by
  Портативный скрипт: можно запускать с флешки без установки.
#>

$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$BrandName = 'Yakastsi'
$BrandUrl  = 'rgb.by'
$AppTitle  = 'PC Serial Scanner'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Stamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$HostName = $env:COMPUTERNAME
$ReportsDir = Join-Path $ScriptRoot 'Reports'
if (-not (Test-Path $ReportsDir)) {
    New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null
}

$BaseName = "Serials_${HostName}_${Stamp}"
$TxtPath  = Join-Path $ReportsDir "$BaseName.txt"
$CsvPath  = Join-Path $ReportsDir "$BaseName.csv"
$HtmlPath = Join-Path $ReportsDir "$BaseName.html"

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-Safe {
    param([scriptblock]$Script, $Default = $null)
    try { & $Script } catch { $Default }
}

function Norm-Serial {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $v = $Value.Trim()
    $junk = @(
        'To Be Filled By O.E.M.',
        'To be filled by O.E.M.',
        'Default string',
        'Default String',
        'System Serial Number',
        'System Product Name',
        'None',
        'N/A',
        'NA',
        '0',
        '00000000',
        '0000000000000000',
        'FFFFFFFF',
        'OEM',
        'O.E.M.',
        'Chassis Serial Number',
        'Base Board Serial Number',
        'Empty'
    )
    if ($junk -contains $v) { return $null }
    if ($v -match '^\s*$') { return $null }
    return $v
}

function Add-Row {
    param(
        [System.Collections.Generic.List[object]]$List,
        [string]$Category,
        [string]$Device,
        [string]$Manufacturer,
        [string]$Model,
        [string]$Serial,
        [string]$Extra = ''
    )
    $serialNorm = Norm-Serial $Serial
    $List.Add([pscustomobject]@{
        Category     = $Category
        Device       = $Device
        Manufacturer = if ($Manufacturer) { $Manufacturer.Trim() } else { '' }
        Model        = if ($Model) { $Model.Trim() } else { '' }
        Serial       = if ($serialNorm) { $serialNorm } else { '— не указан / скрыт производителем —' }
        SerialRaw    = $Serial
        HasRealSerial= [bool]$serialNorm
        Extra        = $Extra
    })
}

$Rows = New-Object System.Collections.Generic.List[object]
$IsAdmin = Test-IsAdmin
$CollectedAt = Get-Date
$Os = Get-Safe { Get-CimInstance Win32_OperatingSystem }
$Cs = Get-Safe { Get-CimInstance Win32_ComputerSystem }

Write-Host "  Хост: $HostName"
Write-Host "  Админ-права: $(if ($IsAdmin) { 'да' } else { 'нет (часть данных может быть недоступна)' })"
Write-Host ""

# --- Система / корпус ---
Write-Host "  [1/12] Система и корпус..."
$Enclosure = Get-Safe { Get-CimInstance Win32_SystemEnclosure }
$Product = Get-Safe { Get-CimInstance Win32_ComputerSystemProduct }
Add-Row $Rows 'Система' 'ПК / корпус' `
    $(if ($Product) { $Product.Vendor } else { $Cs.Manufacturer }) `
    $(if ($Product) { $Product.Name } else { $Cs.Model }) `
    $(if ($Product) { $Product.IdentifyingNumber } elseif ($Enclosure) { $Enclosure.SerialNumber } else { $null }) `
    $(if ($Product -and $Product.UUID) { "UUID: $($Product.UUID)" } else { '' })

if ($Enclosure -and $Enclosure.SerialNumber -and $Product -and
    (Norm-Serial $Enclosure.SerialNumber) -and
    ((Norm-Serial $Enclosure.SerialNumber) -ne (Norm-Serial $Product.IdentifyingNumber))) {
    Add-Row $Rows 'Система' 'Корпус (Enclosure)' $Cs.Manufacturer $Cs.Model $Enclosure.SerialNumber
}

# --- BIOS ---
Write-Host "  [2/12] BIOS..."
$Bios = Get-Safe { Get-CimInstance Win32_BIOS }
if ($Bios) {
    Add-Row $Rows 'BIOS' 'BIOS / прошивка' $Bios.Manufacturer $Bios.SMBIOSBIOSVersion $Bios.SerialNumber `
        $(if ($Bios.ReleaseDate) { "Дата: $($Bios.ReleaseDate.ToString('yyyy-MM-dd'))" } else { '' })
}

# --- Материнская плата ---
Write-Host "  [3/12] Материнская плата..."
$Board = Get-Safe { Get-CimInstance Win32_BaseBoard }
if ($Board) {
    Add-Row $Rows 'Материнская плата' 'Motherboard' $Board.Manufacturer $Board.Product $Board.SerialNumber `
        $(if ($Board.Version) { "Версия: $($Board.Version)" } else { '' })
}

# --- CPU ---
Write-Host "  [4/12] Процессор..."
$Cpus = @(Get-Safe { Get-CimInstance Win32_Processor } @())
$i = 0
foreach ($cpu in $Cpus) {
    $i++
    $extra = @()
    if ($cpu.NumberOfCores) { $extra += "Ядер: $($cpu.NumberOfCores)" }
    if ($cpu.NumberOfLogicalProcessors) { $extra += "Потоков: $($cpu.NumberOfLogicalProcessors)" }
    if ($cpu.MaxClockSpeed) { $extra += "Макс. МГц: $($cpu.MaxClockSpeed)" }
    if ($cpu.ProcessorId) { $extra += "ProcessorId: $($cpu.ProcessorId)" }
    Add-Row $Rows 'Процессор' "CPU #$i" $cpu.Manufacturer $cpu.Name $cpu.SerialNumber ($extra -join '; ')
}

# --- ОЗУ ---
Write-Host "  [5/12] Оперативная память..."
$Mem = @(Get-Safe { Get-CimInstance Win32_PhysicalMemory } @())
$i = 0
foreach ($m in $Mem) {
    $i++
    $gb = if ($m.Capacity) { [math]::Round($m.Capacity / 1GB, 2) } else { $null }
    $extra = @()
    if ($gb) { $extra += "$gb ГБ" }
    if ($m.Speed) { $extra += "$($m.Speed) МГц" }
    if ($m.ConfiguredClockSpeed) { $extra += "настроено: $($m.ConfiguredClockSpeed) МГц" }
    if ($m.DeviceLocator) { $extra += "слот: $($m.DeviceLocator)" }
    if ($m.BankLabel) { $extra += "банк: $($m.BankLabel)" }
    if ($m.PartNumber) { $extra += "Part#: $($m.PartNumber.Trim())" }
    $model = if ($m.PartNumber) { $m.PartNumber.Trim() } else { "$gb ГБ" }
    Add-Row $Rows 'ОЗУ' "Модуль RAM #$i" $m.Manufacturer $model $m.SerialNumber ($extra -join '; ')
}

# --- Диски ---
Write-Host "  [6/12] Накопители..."
$Disks = @(Get-Safe { Get-CimInstance Win32_DiskDrive } @())
$i = 0
foreach ($d in $Disks) {
    $i++
    $gb = if ($d.Size) { [math]::Round($d.Size / 1GB, 2) } else { $null }
    $extra = @()
    if ($gb) { $extra += "$gb ГБ" }
    if ($d.InterfaceType) { $extra += $d.InterfaceType }
    if ($d.MediaType) { $extra += $d.MediaType }
    if ($d.FirmwareRevision) { $extra += "FW: $($d.FirmwareRevision)" }
    if ($d.PNPDeviceID) { $extra += "PNP: $($d.PNPDeviceID)" }
    Add-Row $Rows 'Накопитель' "Диск #$i" $d.Manufacturer $d.Model $d.SerialNumber ($extra -join '; ')
}

# Storage cmdlets (often better serials for NVMe/SSD)
$PhysDisks = @(Get-Safe { Get-PhysicalDisk } @())
foreach ($pd in $PhysDisks) {
    $sn = Norm-Serial $pd.SerialNumber
    if (-not $sn) { continue }
    $exists = $Rows | Where-Object { $_.Category -eq 'Накопитель' -and (Norm-Serial $_.SerialRaw) -eq $sn }
    if ($exists) { continue }
    $gb = if ($pd.Size) { [math]::Round($pd.Size / 1GB, 2) } else { $null }
    Add-Row $Rows 'Накопитель' "PhysicalDisk $($pd.DeviceId)" $pd.Manufacturer $pd.FriendlyName $pd.SerialNumber `
        ("$gb ГБ; $($pd.MediaType); $($pd.BusType)")
}

# --- Видео ---
Write-Host "  [7/12] Видеокарта..."
$Gpus = @(Get-Safe { Get-CimInstance Win32_VideoController } @())
$i = 0
foreach ($g in $Gpus) {
    if ($g.Name -match 'Microsoft Basic|Remote Desktop|Virtual') { continue }
    $i++
    $extra = @()
    if ($g.AdapterRAM -and $g.AdapterRAM -gt 0) {
        # AdapterRAM can wrap for >4GB; still show if sensible
        $vram = [math]::Round([uint64]([math]::Abs([int64]$g.AdapterRAM)) / 1MB)
        if ($vram -gt 0 -and $vram -lt 100000) { $extra += "~$vram МБ VRAM (WMI)" }
    }
    if ($g.DriverVersion) { $extra += "драйвер: $($g.DriverVersion)" }
    if ($g.PNPDeviceID) { $extra += "PNP: $($g.PNPDeviceID)" }
    # Serial rarely in WMI for GPU
    Add-Row $Rows 'Видеокарта' "GPU #$i" $g.AdapterCompatibility $g.Name $null ($extra -join '; ')
}

# NVIDIA serial via nvidia-smi if present
$Nvidia = Get-Command nvidia-smi -ErrorAction SilentlyContinue
if ($Nvidia) {
    $nOut = Get-Safe { & nvidia-smi --query-gpu=index,name,serial,uuid,gpu_part_number --format=csv,noheader 2>$null }
    if ($nOut) {
        foreach ($line in @($nOut)) {
            $parts = $line -split ',\s*'
            if ($parts.Count -ge 3) {
                Add-Row $Rows 'Видеокарта' "NVIDIA GPU $($parts[0])" 'NVIDIA' $parts[1] $parts[2] `
                    $(if ($parts.Count -ge 5) { "UUID: $($parts[3]); Part: $($parts[4])" } elseif ($parts.Count -ge 4) { "UUID: $($parts[3])" } else { '' })
            }
        }
    }
}

# --- Мониторы ---
Write-Host "  [8/12] Мониторы..."
$Mons = @(Get-Safe { Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -ErrorAction SilentlyContinue } @())
$i = 0
foreach ($mon in $Mons) {
    $i++
    $decode = {
        param($bytes)
        if (-not $bytes) { return '' }
        -join ($bytes | Where-Object { $_ -ne 0 -and $_ -ge 32 -and $_ -le 126 } | ForEach-Object { [char]$_ })
    }
    $mfr = & $decode $mon.ManufacturerName
    $name = & $decode $mon.UserFriendlyName
    $sn = & $decode $mon.SerialNumberID
    $week = $mon.WeekOfManufacture
    $year = $mon.YearOfManufacture
    $extra = @()
    if ($year) { $extra += "год: $year" }
    if ($week) { $extra += "неделя: $week" }
    Add-Row $Rows 'Монитор' "Монитор #$i" $mfr $name $sn ($extra -join '; ')
}

# Fallback Win32_DesktopMonitor
if ($Mons.Count -eq 0) {
    $Desk = @(Get-Safe { Get-CimInstance Win32_DesktopMonitor } @())
    $i = 0
    foreach ($m in $Desk) {
        if (-not $m.Name) { continue }
        $i++
        Add-Row $Rows 'Монитор' "Монитор #$i" $m.MonitorManufacturer $m.Name $m.SerialNumber ''
    }
}

# --- Сеть ---
Write-Host "  [9/12] Сетевые адаптеры..."
$Nics = @(Get-Safe {
    Get-CimInstance Win32_NetworkAdapter |
        Where-Object { $_.PhysicalAdapter -eq $true -and $_.MACAddress -and $_.PNPDeviceID -notmatch 'ROOT\\' }
} @())
$i = 0
foreach ($n in $Nics) {
    $i++
    $extra = @()
    if ($n.MACAddress) { $extra += "MAC: $($n.MACAddress)" }
    if ($n.Speed) { $extra += "скорость: $([math]::Round($n.Speed/1MB,0)) Мбит/с" }
    Add-Row $Rows 'Сеть' "NIC #$i" $n.Manufacturer $n.Name $n.MACAddress ($extra -join '; ')
}

# --- Батарея (ноутбуки) ---
Write-Host "  [10/12] Батарея..."
$Bats = @(Get-Safe { Get-CimInstance Win32_Battery } @())
$i = 0
foreach ($b in $Bats) {
    $i++
    Add-Row $Rows 'Батарея' "Battery #$i" $b.Manufacturer $b.Name $b.DeviceID ''
}
$BatStatic = @(Get-Safe { Get-CimInstance -Namespace root\wmi -ClassName BatteryStaticData -ErrorAction SilentlyContinue } @())
foreach ($b in $BatStatic) {
    $sn = if ($b.SerialNumber) { [string]$b.SerialNumber } else { $null }
    $mfr = if ($b.ManufactureName) { -join ($b.ManufactureName | ForEach-Object { [char]$_ }) } else { '' }
    Add-Row $Rows 'Батарея' 'BatteryStaticData' $mfr '' $sn ''
}

# --- TPM ---
Write-Host "  [11/12] TPM..."
$Tpm = Get-Safe { Get-CimInstance -Namespace root\cimv2\Security\MicrosoftTpm -ClassName Win32_Tpm -ErrorAction SilentlyContinue }
if ($Tpm) {
    Add-Row $Rows 'TPM' 'Trusted Platform Module' '' "Spec $($Tpm.SpecVersion)" $Tpm.ManufacturerIdTxt `
        "ManufacturerVersion: $($Tpm.ManufacturerVersion); Enabled: $($Tpm.IsEnabled_InitialValue)"
}

# --- Прочее: оптика, звук (без серийников часто) ---
Write-Host "  [12/12] Дополнительно..."
$Cd = @(Get-Safe { Get-CimInstance Win32_CDROMDrive } @())
$i = 0
foreach ($c in $Cd) {
    $i++
    Add-Row $Rows 'Оптический привод' "ODD #$i" $c.Manufacturer $c.Name $c.SerialNumber $c.MediaType
}

# SMBIOS memory array / chassis asset tags via msinfo if needed — skip

# ============================================
# Reports
# ============================================
Write-Host ""
Write-Host "  Формирование отчётов..."

$realCount = @($Rows | Where-Object { $_.HasRealSerial }).Count
$totalCount = $Rows.Count

# --- TXT ---
$txt = New-Object System.Text.StringBuilder
[void]$txt.AppendLine('=' * 72)
[void]$txt.AppendLine("  $AppTitle")
[void]$txt.AppendLine("  $BrandName — $BrandUrl")
[void]$txt.AppendLine('=' * 72)
[void]$txt.AppendLine("ПК:            $HostName")
[void]$txt.AppendLine("Пользователь:  $env:USERNAME")
[void]$txt.AppendLine("ОС:            $($Os.Caption) $($Os.Version) ($($Os.OSArchitecture))")
[void]$txt.AppendLine("Производитель: $($Cs.Manufacturer)")
[void]$txt.AppendLine("Модель:        $($Cs.Model)")
[void]$txt.AppendLine("Дата сбора:    $($CollectedAt.ToString('yyyy-MM-dd HH:mm:ss'))")
[void]$txt.AppendLine("Админ-права:   $(if ($IsAdmin) { 'да' } else { 'нет' })")
[void]$txt.AppendLine("Найдено позиций: $totalCount  |  с реальным серийником: $realCount")
[void]$txt.AppendLine('=' * 72)
[void]$txt.AppendLine('')

$grouped = $Rows | Group-Object Category
foreach ($g in $grouped) {
    [void]$txt.AppendLine(('-' * 72))
    [void]$txt.AppendLine("  $($g.Name)")
    [void]$txt.AppendLine(('-' * 72))
    foreach ($r in $g.Group) {
        [void]$txt.AppendLine("Устройство:    $($r.Device)")
        if ($r.Manufacturer) { [void]$txt.AppendLine("Производитель: $($r.Manufacturer)") }
        if ($r.Model)        { [void]$txt.AppendLine("Модель:        $($r.Model)") }
        [void]$txt.AppendLine("Серийный №:    $($r.Serial)")
        if ($r.Extra)        { [void]$txt.AppendLine("Дополнительно: $($r.Extra)") }
        [void]$txt.AppendLine('')
    }
}

[void]$txt.AppendLine('=' * 72)
[void]$txt.AppendLine('Примечания:')
[void]$txt.AppendLine('• У CPU, GPU и части устройств серийник часто скрыт и недоступен через Windows.')
[void]$txt.AppendLine('• «— не указан / скрыт производителем —» значит BIOS/драйвер не отдал номер.')
[void]$txt.AppendLine('• Для сверки с гарантией важнее всего: материнская плата, ОЗУ, диски, корпус/система, мониторы.')
[void]$txt.AppendLine('• Запуск от администратора даёт более полные данные.')
[void]$txt.AppendLine('')
[void]$txt.AppendLine("Сделано: $BrandName — $BrandUrl")
[void]$txt.AppendLine('=' * 72)

$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($TxtPath, $txt.ToString(), $utf8Bom)

# --- CSV ---
$Rows | Select-Object Category, Device, Manufacturer, Model, Serial, Extra |
    Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

# --- HTML ---
function Html-Enc([string]$s) {
    if ($null -eq $s) { return '' }
    [System.Net.WebUtility]::HtmlEncode($s)
}

$htmlRows = foreach ($r in $Rows) {
    $snClass = if ($r.HasRealSerial) { 'ok' } else { 'miss' }
    @"
<tr>
  <td>$(Html-Enc $r.Category)</td>
  <td>$(Html-Enc $r.Device)</td>
  <td>$(Html-Enc $r.Manufacturer)</td>
  <td>$(Html-Enc $r.Model)</td>
  <td class="$snClass">$(Html-Enc $r.Serial)</td>
  <td class="extra">$(Html-Enc $r.Extra)</td>
</tr>
"@
}

$html = @"
<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$AppTitle — $HostName — $BrandName</title>
<style>
  :root {
    --bg: #0f1419;
    --panel: #1a222c;
    --line: #2a3542;
    --text: #e8eef4;
    --muted: #8b9aab;
    --accent: #3dd6c6;
    --accent2: #5b8def;
    --ok: #3dd68c;
    --miss: #c4a35a;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: "Segoe UI", "Candara", "Trebuchet MS", sans-serif;
    background:
      radial-gradient(1200px 600px at 10% -10%, #1a3a45 0%, transparent 55%),
      radial-gradient(900px 500px at 100% 0%, #1a2a4a 0%, transparent 50%),
      var(--bg);
    color: var(--text);
    line-height: 1.45;
    min-height: 100vh;
  }
  header {
    padding: 2rem 1.5rem 1.25rem;
    border-bottom: 1px solid var(--line);
    background: linear-gradient(180deg, rgba(26,34,44,.9), transparent);
  }
  .brand {
    font-size: 1.75rem;
    font-weight: 700;
    letter-spacing: 0.02em;
    margin: 0 0 0.25rem;
  }
  .brand span { color: var(--accent); }
  .tagline {
    color: var(--muted);
    margin: 0 0 1rem;
    font-size: 0.95rem;
  }
  .meta {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 0.75rem 1.25rem;
    max-width: 1100px;
  }
  .meta div {
    background: var(--panel);
    border: 1px solid var(--line);
    border-radius: 10px;
    padding: 0.65rem 0.85rem;
  }
  .meta .k { display:block; color: var(--muted); font-size: 0.75rem; text-transform: uppercase; letter-spacing: .06em; }
  .meta .v { font-weight: 600; word-break: break-word; }
  main { padding: 1.25rem 1.5rem 3rem; max-width: 1200px; margin: 0 auto; }
  h2 { font-size: 1.1rem; margin: 1.5rem 0 0.75rem; font-weight: 600; }
  .note {
    background: rgba(91,141,239,.08);
    border: 1px solid rgba(91,141,239,.25);
    border-radius: 10px;
    padding: 0.85rem 1rem;
    color: var(--muted);
    font-size: 0.9rem;
  }
  .wrap { overflow-x: auto; border: 1px solid var(--line); border-radius: 12px; background: var(--panel); }
  table { width: 100%; border-collapse: collapse; font-size: 0.9rem; }
  th, td { padding: 0.65rem 0.75rem; text-align: left; border-bottom: 1px solid var(--line); vertical-align: top; }
  th { color: var(--muted); font-weight: 600; font-size: 0.75rem; text-transform: uppercase; letter-spacing: .04em; background: rgba(0,0,0,.2); position: sticky; top: 0; }
  tr:last-child td { border-bottom: 0; }
  tr:hover td { background: rgba(61,214,198,.04); }
  td.ok { color: var(--ok); font-family: Consolas, "Courier New", monospace; font-weight: 600; }
  td.miss { color: var(--miss); font-style: italic; }
  td.extra { color: var(--muted); font-size: 0.82rem; max-width: 280px; }
  footer {
    margin-top: 2rem;
    padding-top: 1rem;
    border-top: 1px solid var(--line);
    color: var(--muted);
    font-size: 0.85rem;
  }
  footer a { color: var(--accent); text-decoration: none; }
  footer a:hover { text-decoration: underline; }
  .badge {
    display: inline-block;
    background: rgba(61,214,198,.12);
    color: var(--accent);
    border: 1px solid rgba(61,214,198,.35);
    border-radius: 6px;
    padding: 0.15rem 0.5rem;
    font-size: 0.8rem;
    margin-left: 0.35rem;
  }
</style>
</head>
<body>
<header>
  <p class="brand">$BrandName <span>— $BrandUrl</span></p>
  <p class="tagline">$AppTitle — серийные номера комплектующих для сверки с гарантией</p>
  <div class="meta">
    <div><span class="k">Компьютер</span><span class="v">$(Html-Enc $HostName)</span></div>
    <div><span class="k">ОС</span><span class="v">$(Html-Enc ($Os.Caption + ' ' + $Os.OSArchitecture))</span></div>
    <div><span class="k">Модель ПК</span><span class="v">$(Html-Enc ($Cs.Manufacturer + ' / ' + $Cs.Model))</span></div>
    <div><span class="k">Дата сбора</span><span class="v">$(Html-Enc $CollectedAt.ToString('yyyy-MM-dd HH:mm:ss'))</span></div>
    <div><span class="k">С серийником</span><span class="v">$realCount из $totalCount<span class="badge">$(if ($IsAdmin) { 'admin' } else { 'user' })</span></span></div>
  </div>
</header>
<main>
  <div class="note">
    Сверьте номера с гарантией продавца. Реальные серийники подсвечены зелёным.
    У процессора и многих видеокарт Windows часто не отдаёт серийный номер — это нормально.
    Для полного отчёта запускайте от имени администратора.
  </div>
  <h2>Комплектующие</h2>
  <div class="wrap">
    <table>
      <thead>
        <tr>
          <th>Категория</th>
          <th>Устройство</th>
          <th>Производитель</th>
          <th>Модель</th>
          <th>Серийный номер</th>
          <th>Дополнительно</th>
        </tr>
      </thead>
      <tbody>
$($htmlRows -join "`n")
      </tbody>
    </table>
  </div>
  <footer>
    Сделано: <strong>$BrandName</strong> — <a href="https://$BrandUrl">$BrandUrl</a><br>
    Файлы рядом: $(Html-Enc (Split-Path $TxtPath -Leaf)), $(Html-Enc (Split-Path $CsvPath -Leaf))
  </footer>
</main>
</body>
</html>
"@

[System.IO.File]::WriteAllText($HtmlPath, $html, $utf8Bom)

# Open report
Write-Host ""
Write-Host "  TXT:  $TxtPath"
Write-Host "  CSV:  $CsvPath"
Write-Host "  HTML: $HtmlPath"
Write-Host ""
Write-Host "  Сделано: $BrandName — $BrandUrl"
Write-Host ""

try { Start-Process $HtmlPath } catch { }
try { Start-Process notepad.exe $TxtPath } catch { }

exit 0
