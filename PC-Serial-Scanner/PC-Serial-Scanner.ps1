#Requires -Version 5.1
<#
.SYNOPSIS
  Максимальный сбор серийных номеров комплектующих ПК (Windows 10/11).
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
    # trim nulls / non-printable leftovers from SMBIOS
    $v = ($v -replace '[\x00-\x1F]', '').Trim()
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
        '0000000000',
        '000000000000',
        '0000000000000000',
        'FFFFFFFF',
        'FFFFFFFFFFFFFFFF',
        'OEM',
        'O.E.M.',
        'Chassis Serial Number',
        'Base Board Serial Number',
        'Empty',
        'Unknown',
        'Not Specified',
        'Not Available',
        'No Enclosure',
        'Invalid',
        'undefined'
    )
    if ($junk -contains $v) { return $null }
    if ($v -match '^(0+|F+)$') { return $null }
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
        [string]$Extra = '',
        [string]$Note = ''
    )
    $serialNorm = Norm-Serial $Serial
    $display = if ($serialNorm) {
        $serialNorm
    } elseif ($Note) {
        $Note
    } else {
        '— не указан / скрыт производителем —'
    }
    $List.Add([pscustomobject]@{
        Category     = $Category
        Device       = $Device
        Manufacturer = if ($Manufacturer) { $Manufacturer.Trim() } else { '' }
        Model        = if ($Model) { $Model.Trim() } else { '' }
        Serial       = $display
        SerialRaw    = $Serial
        HasRealSerial= [bool]$serialNorm
        Extra        = $Extra
    })
}

function Get-UsbSerialFromPnp {
    param([string]$PnpId)
    if (-not $PnpId) { return $null }
    # USB\VID_xxxx&PID_xxxx\SERIAL  or  SCSI\...\SERIAL
    if ($PnpId -match '\\([^\\]+)$') {
        $tail = $Matches[1]
        if ($tail -match '^(?:[0-9A-F]{4}&)?(?:[0-9A-F]{4}&)?[0-9A-F]+$' -and $tail.Length -lt 8) { return $null }
        if ($tail -match '^[&0-9]+$') { return $null }
        if ($tail -match '^(5&|6&|7&|[0-9]+&)') { return $null } # instance path, not serial
        return $tail
    }
    return $null
}

function Get-MemoryTypeName {
    param($Type, $SmbiosType)
    # SMBIOS 3.2+ MemoryDevice.SMBIOSMemoryType is more accurate
    $map = @{
        20 = 'DDR'; 21 = 'DDR2'; 22 = 'DDR2 FB-DIMM'; 24 = 'DDR3'
        26 = 'DDR4'; 27 = 'LPDDR'; 28 = 'LPDDR2'; 29 = 'LPDDR3'
        30 = 'LPDDR4'; 34 = 'DDR5'; 35 = 'LPDDR5'
    }
    if ($SmbiosType -and $map.ContainsKey([int]$SmbiosType)) { return $map[[int]$SmbiosType] }
    $legacy = @{
        0='Unknown'; 1='Other'; 2='DRAM'; 3='Synchronous DRAM'; 4='Cache DRAM'
        5='EDO'; 6='EDRAM'; 7='VRAM'; 8='SRAM'; 9='RAM'
        10='ROM'; 11='Flash'; 12='EEPROM'; 13='FEPROM'; 14='EPROM'
        15='CDRAM'; 16='3DRAM'; 17='SDRAM'; 18='SGRAM'; 19='RDRAM'
        20='DDR'; 21='DDR2'; 22='DDR2 FB-DIMM'; 24='DDR3'; 25='FBD2'; 26='DDR4'
    }
    if ($Type -and $legacy.ContainsKey([int]$Type)) { return $legacy[[int]$Type] }
    return ''
}

function Get-FormFactorName {
    param($ff)
    $map = @{ 0='Unknown'; 8='DIMM'; 12='SO-DIMM'; 13='Micro-DIMM'; 22='72-pin SODIMM' }
    if ($ff -and $map.ContainsKey([int]$ff)) { return $map[[int]$ff] }
    return "FormFactor=$ff"
}

$Rows = New-Object System.Collections.Generic.List[object]
$SeenSerialKeys = New-Object 'System.Collections.Generic.HashSet[string]'
$IsAdmin = Test-IsAdmin
$CollectedAt = Get-Date
$Os = Get-Safe { Get-CimInstance Win32_OperatingSystem }
$Cs = Get-Safe { Get-CimInstance Win32_ComputerSystem }

function Add-UniqueRow {
    param(
        [string]$Category, [string]$Device, [string]$Manufacturer,
        [string]$Model, [string]$Serial, [string]$Extra = '', [string]$Note = ''
    )
    $sn = Norm-Serial $Serial
    if ($sn) {
        $key = "$Category|$sn|$Model"
        if (-not $SeenSerialKeys.Add($key)) { return }
    }
    Add-Row $Rows $Category $Device $Manufacturer $Model $Serial $Extra $Note
}

Write-Host "  Хост: $HostName"
Write-Host "  Админ-права: $(if ($IsAdmin) { 'да' } else { 'нет (часть данных может быть недоступна)' })"
Write-Host ""

# ========== 1. Система / корпус ==========
Write-Host "  [1/16] Система и корпус..."
$Enclosure = Get-Safe { Get-CimInstance Win32_SystemEnclosure }
$Product = Get-Safe { Get-CimInstance Win32_ComputerSystemProduct }
$System = Get-Safe { Get-CimInstance Win32_SystemEnclosure | Select-Object -First 1 }

Add-UniqueRow 'Система' 'ПК (System Product)' `
    $(if ($Product) { $Product.Vendor } else { $Cs.Manufacturer }) `
    $(if ($Product) { $Product.Name } else { $Cs.Model }) `
    $(if ($Product) { $Product.IdentifyingNumber } else { $null }) `
    $(@(
        if ($Product -and $Product.UUID) { "UUID: $($Product.UUID)" }
        if ($Product -and $Product.Version) { "Версия: $($Product.Version)" }
        if ($Cs -and $Cs.SystemSKUNumber) { "SKU: $($Cs.SystemSKUNumber)" }
        if ($Cs -and $Cs.OEMStringArray) { "OEM: $($Cs.OEMStringArray -join ', ')" }
    ) -join '; ')

if ($Enclosure) {
    $asset = $null
    if ($Enclosure.SMBIOSAssetTag) { $asset = Norm-Serial ([string]$Enclosure.SMBIOSAssetTag) }
    Add-UniqueRow 'Система' 'Корпус (Chassis / Enclosure)' $Cs.Manufacturer $Cs.Model $Enclosure.SerialNumber `
        $(@(
            if ($asset) { "Asset Tag: $asset" }
            if ($Enclosure.ChassisTypes) { "Тип корпуса: $($Enclosure.ChassisTypes -join ',')" }
        ) -join '; ')
    if ($asset) {
        Add-UniqueRow 'Система' 'Asset Tag корпуса' $Cs.Manufacturer $Cs.Model $asset
    }
}

# ========== 2. BIOS ==========
Write-Host "  [2/16] BIOS..."
$Bios = Get-Safe { Get-CimInstance Win32_BIOS }
if ($Bios) {
    Add-UniqueRow 'BIOS' 'BIOS / UEFI' $Bios.Manufacturer $Bios.SMBIOSBIOSVersion $Bios.SerialNumber `
        $(@(
            if ($Bios.ReleaseDate) { "Дата: $($Bios.ReleaseDate.ToString('yyyy-MM-dd'))" }
            if ($Bios.SMBIOSMajorVersion) { "SMBIOS: $($Bios.SMBIOSMajorVersion).$($Bios.SMBIOSMinorVersion)" }
            if ($Bios.Version) { "Version: $($Bios.Version)" }
        ) -join '; ')
}

# ========== 3. Материнская плата ==========
Write-Host "  [3/16] Материнская плата..."
$Board = Get-Safe { Get-CimInstance Win32_BaseBoard }
if ($Board) {
    Add-UniqueRow 'Материнская плата' 'Motherboard' $Board.Manufacturer $Board.Product $Board.SerialNumber `
        $(@(
            if ($Board.Version) { "Версия: $($Board.Version)" }
            if ($Board.Tag) { "Tag: $($Board.Tag)" }
            if ($Board.Model) { "Model: $($Board.Model)" }
        ) -join '; ')
}
$Mb = Get-Safe { Get-CimInstance Win32_MotherboardDevice }
if ($Mb) {
    Add-UniqueRow 'Материнская плата' 'MotherboardDevice' $Mb.Manufacturer $Mb.Name $Mb.DeviceID `
        $(if ($Mb.PrimaryBusType) { "Шина: $($Mb.PrimaryBusType)/$($Mb.SecondaryBusType)" } else { '' }) `
        'DeviceID (не всегда серийник платы)'
}

# ========== 4. Процессор ==========
Write-Host "  [4/16] Процессор..."
$Cpus = @(Get-Safe { Get-CimInstance Win32_Processor } @())
$cpuIdx = 0
foreach ($cpu in $Cpus) {
    $cpuIdx++
    $extra = New-Object System.Collections.Generic.List[string]
    if ($cpu.SocketDesignation) { [void]$extra.Add("Сокет: $($cpu.SocketDesignation)") }
    if ($cpu.NumberOfCores) { [void]$extra.Add("Ядер: $($cpu.NumberOfCores)") }
    if ($cpu.NumberOfLogicalProcessors) { [void]$extra.Add("Потоков: $($cpu.NumberOfLogicalProcessors)") }
    if ($cpu.MaxClockSpeed) { [void]$extra.Add("Макс. МГц: $($cpu.MaxClockSpeed)") }
    if ($cpu.CurrentClockSpeed) { [void]$extra.Add("Текущ. МГц: $($cpu.CurrentClockSpeed)") }
    if ($cpu.L2CacheSize) { [void]$extra.Add("L2: $($cpu.L2CacheSize) КБ") }
    if ($cpu.L3CacheSize) { [void]$extra.Add("L3: $($cpu.L3CacheSize) КБ") }
    if ($cpu.Family) { [void]$extra.Add("Family: $($cpu.Family)") }
    if ($cpu.Stepping) { [void]$extra.Add("Stepping: $($cpu.Stepping)") }
    if ($cpu.Revision) { [void]$extra.Add("Revision: $($cpu.Revision)") }
    if ($cpu.PartNumber) { [void]$extra.Add("Part#: $($cpu.PartNumber)") }

    # SerialNumber у CPU почти всегда пустой (Intel/AMD отключают Unique ID)
    $cpuSerial = Norm-Serial $cpu.SerialNumber
    $cpuNote = if (-not $cpuSerial) {
        'Серийник CPU обычно скрыт Intel/AMD. Ниже — ProcessorId и PartNumber как идентификаторы.'
    } else { '' }

    Add-UniqueRow 'Процессор' "CPU планка/сокет #$cpuIdx — $($cpu.Name)" `
        $cpu.Manufacturer $cpu.Name $cpu.SerialNumber ($extra -join '; ') $cpuNote

    # ProcessorId — уникальный ID из CPUID (не заводской S/N гарантии, но полезен)
    if ($cpu.ProcessorId) {
        Add-UniqueRow 'Процессор' "CPU #$cpuIdx ProcessorId (CPUID)" `
            $cpu.Manufacturer $cpu.Name $cpu.ProcessorId `
            'Это ProcessorId из CPUID, не всегда совпадает с гарантийным S/N на коробке'
    }
    if ($cpu.PartNumber -and (Norm-Serial $cpu.PartNumber)) {
        Add-UniqueRow 'Процессор' "CPU #$cpuIdx Part Number" `
            $cpu.Manufacturer $cpu.Name $cpu.PartNumber 'Part Number из SMBIOS'
    }
    # UniqueId если есть
    if ($cpu.UniqueId) {
        Add-UniqueRow 'Процессор' "CPU #$cpuIdx UniqueId" $cpu.Manufacturer $cpu.Name $cpu.UniqueId
    }
}

# Registry CPU identifiers
$regCpu = Get-Safe {
    Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0' -ErrorAction Stop
}
if ($regCpu) {
    $idStr = @()
    if ($regCpu.ProcessorNameString) { $idStr += $regCpu.ProcessorNameString }
    if ($regCpu.Identifier) { $idStr += "Identifier: $($regCpu.Identifier)" }
    if ($regCpu.'VendorIdentifier') { $idStr += "Vendor: $($regCpu.'VendorIdentifier')" }
    if ($regCpu.'~MHz') { $idStr += "~$($regCpu.'~MHz') МГц" }
    Add-UniqueRow 'Процессор' 'CPU (реестр Windows)' `
        $(if ($regCpu.'VendorIdentifier') { $regCpu.'VendorIdentifier' } else { '' }) `
        $(if ($regCpu.ProcessorNameString) { $regCpu.ProcessorNameString } else { '' }) `
        $(if ($regCpu.Identifier) { $regCpu.Identifier } else { $null }) `
        ($idStr -join '; ') `
        'Идентификатор из реестра (не заводской S/N)'
}

# ========== 5. ОЗУ — каждая планка отдельно ==========
Write-Host "  [5/16] Оперативная память (каждая планка)..."
$MemArray = @(Get-Safe { Get-CimInstance Win32_PhysicalMemoryArray } @())
foreach ($arr in $MemArray) {
    $maxGb = if ($arr.MaxCapacity) { [math]::Round($arr.MaxCapacity / 1MB, 1) } else { $null } # MaxCapacity in KB
    Add-UniqueRow 'ОЗУ' 'Массив памяти (слоты)' '' '' $arr.SerialNumber `
        $(@(
            "Слотов: $($arr.MemoryDevices)"
            if ($maxGb) { "Макс. ёмкость: ~$maxGb ГБ" }
        ) -join '; ')
}

$Mem = @(Get-Safe { Get-CimInstance Win32_PhysicalMemory } @())
# Also try namespace alternative
if ($Mem.Count -eq 0) {
    $Mem = @(Get-Safe { Get-WmiObject Win32_PhysicalMemory } @())
}

$totalRamGb = 0
$stickCount = $Mem.Count
Write-Host "         Найдено планок ОЗУ: $stickCount"

$stickIdx = 0
foreach ($m in $Mem) {
    $stickIdx++
    $gb = if ($m.Capacity) { [math]::Round($m.Capacity / 1GB, 2) } else { 0 }
    $totalRamGb += $gb
    $typeName = Get-MemoryTypeName $m.MemoryType $m.SMBIOSMemoryType
    $ff = Get-FormFactorName $m.FormFactor
    $mfr = if ($m.Manufacturer) { $m.Manufacturer.Trim() } else { '' }
    $part = if ($m.PartNumber) { $m.PartNumber.Trim() } else { '' }
    $sn = $m.SerialNumber
    # Some vendors pad serial with spaces
    if ($sn) { $sn = $sn.Trim() }

    $extra = New-Object System.Collections.Generic.List[string]
    [void]$extra.Add("$gb ГБ")
    if ($typeName) { [void]$extra.Add($typeName) }
    if ($ff) { [void]$extra.Add($ff) }
    if ($m.Speed) { [void]$extra.Add("JEDEC: $($m.Speed) МГц") }
    if ($m.ConfiguredClockSpeed) { [void]$extra.Add("сейчас: $($m.ConfiguredClockSpeed) МГц") }
    if ($m.ConfiguredVoltage) { [void]$extra.Add("$($m.ConfiguredVoltage) мВ") }
    if ($m.DeviceLocator) { [void]$extra.Add("слот: $($m.DeviceLocator)") }
    if ($m.BankLabel) { [void]$extra.Add("банк: $($m.BankLabel)") }
    if ($m.Tag) { [void]$extra.Add("tag: $($m.Tag)") }
    if ($m.DataWidth) { [void]$extra.Add("DataWidth: $($m.DataWidth)") }
    if ($m.TotalWidth) { [void]$extra.Add("TotalWidth: $($m.TotalWidth)") }
    if ($part) { [void]$extra.Add("Part#: $part") }
    if ($m.MinVoltage) { [void]$extra.Add("Vmin: $($m.MinVoltage)") }
    if ($m.MaxVoltage) { [void]$extra.Add("Vmax: $($m.MaxVoltage)") }

    $slotLabel = if ($m.DeviceLocator) { $m.DeviceLocator } else { "#$stickIdx" }
    $model = if ($part) { $part } else { "$gb ГБ $typeName" }

    Add-UniqueRow 'ОЗУ' "Планка ОЗУ #$stickIdx (слот $slotLabel)" `
        $mfr $model $sn ($extra -join '; ') `
        $(if (-not (Norm-Serial $sn)) { 'Серийник планки не отдан SMBIOS — смотрите наклейку на модуле' } else { '' })
}

# Empty slots hint via MemoryArray MemoryDevices vs stick count
foreach ($arr in $MemArray) {
    if ($arr.MemoryDevices -and $stickCount -lt [int]$arr.MemoryDevices) {
        $empty = [int]$arr.MemoryDevices - $stickCount
        Add-UniqueRow 'ОЗУ' "Пустые слоты ОЗУ: $empty" '' '' $null `
            "Всего слотов: $($arr.MemoryDevices); занято: $stickCount" `
            'Свободные слоты без модулей'
    }
}

# ========== 6. Накопители ==========
Write-Host "  [6/16] Накопители..."
$Disks = @(Get-Safe { Get-CimInstance Win32_DiskDrive } @())
$diskIdx = 0
foreach ($d in $Disks) {
    $diskIdx++
    $gb = if ($d.Size) { [math]::Round($d.Size / 1GB, 2) } else { $null }
    $extra = @(
        if ($gb) { "$gb ГБ" }
        if ($d.InterfaceType) { $d.InterfaceType }
        if ($d.MediaType) { $d.MediaType }
        if ($d.FirmwareRevision) { "FW: $($d.FirmwareRevision)" }
        if ($d.Partitions) { "разделов: $($d.Partitions)" }
        if ($d.Index -ne $null) { "Index: $($d.Index)" }
        if ($d.PNPDeviceID) { "PNP: $($d.PNPDeviceID)" }
        if ($d.Status) { "Status: $($d.Status)" }
    ) -join '; '
    $sn = $d.SerialNumber
    if ($sn) { $sn = $sn.Trim() }
    Add-UniqueRow 'Накопитель' "Диск #$diskIdx — $($d.Model)" $d.Manufacturer $d.Model $sn $extra

    $pnpSn = Get-UsbSerialFromPnp $d.PNPDeviceID
    if ($pnpSn -and (Norm-Serial $pnpSn) -ne (Norm-Serial $sn)) {
        Add-UniqueRow 'Накопитель' "Диск #$diskIdx PNP Serial" $d.Manufacturer $d.Model $pnpSn 'Из PNPDeviceID'
    }
}

$PhysDisks = @(Get-Safe { Get-PhysicalDisk -ErrorAction SilentlyContinue } @())
foreach ($pd in $PhysDisks) {
    $sn = if ($pd.SerialNumber) { $pd.SerialNumber.Trim() } else { $null }
    $gb = if ($pd.Size) { [math]::Round($pd.Size / 1GB, 2) } else { $null }
    Add-UniqueRow 'Накопитель' "PhysicalDisk $($pd.DeviceId) — $($pd.FriendlyName)" `
        $pd.Manufacturer $pd.Model $sn `
        ("$gb ГБ; $($pd.MediaType); $($pd.BusType); Health: $($pd.HealthStatus); FW: $($pd.FirmwareVersion)")
}

$MsftDisks = @(Get-Safe {
    Get-CimInstance -Namespace root\Microsoft\Windows\Storage -ClassName MSFT_Disk -ErrorAction SilentlyContinue
} @())
foreach ($d in $MsftDisks) {
    $sn = if ($d.SerialNumber) { [string]$d.SerialNumber.Trim() } else { $null }
    $gb = if ($d.Size) { [math]::Round($d.Size / 1GB, 2) } else { $null }
    Add-UniqueRow 'Накопитель' "MSFT_Disk $($d.Number) — $($d.FriendlyName)" `
        $d.Manufacturer $d.Model $sn ("$gb ГБ; Bus: $($d.BusType); Path: $($d.Path)")
}

$MsftPhys = @(Get-Safe {
    Get-CimInstance -Namespace root\Microsoft\Windows\Storage -ClassName MSFT_PhysicalDisk -ErrorAction SilentlyContinue
} @())
foreach ($d in $MsftPhys) {
    $sn = if ($d.SerialNumber) { [string]$d.SerialNumber.Trim() } else { $null }
    Add-UniqueRow 'Накопитель' "MSFT_PhysicalDisk — $($d.FriendlyName)" `
        $d.Manufacturer $d.Model $sn ("FW: $($d.FirmwareVersion); Media: $($d.MediaType)")
}

# Volumes / logical (no serial usually, but model link)
$Logical = @(Get-Safe { Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } } @())
foreach ($l in $Logical) {
    $gb = if ($l.Size) { [math]::Round($l.Size / 1GB, 2) } else { $null }
    Add-UniqueRow 'Раздел' "Том $($l.DeviceID)" '' $l.VolumeName $l.VolumeSerialNumber `
        ("$gb ГБ; FS: $($l.FileSystem) — серийник ТОМА, не диска")
}

# ========== 7. Видео ==========
Write-Host "  [7/16] Видеокарта..."
$Gpus = @(Get-Safe { Get-CimInstance Win32_VideoController } @())
$gpuIdx = 0
foreach ($g in $Gpus) {
    if ($g.Name -match 'Microsoft Basic|Remote Desktop|Virtual Display|Basic Render') { continue }
    $gpuIdx++
    $extra = @(
        if ($g.DriverVersion) { "драйвер: $($g.DriverVersion)" }
        if ($g.DriverDate) { "дата драйвера: $($g.DriverDate.ToString('yyyy-MM-dd'))" }
        if ($g.VideoProcessor) { "GPU: $($g.VideoProcessor)" }
        if ($g.PNPDeviceID) { "PNP: $($g.PNPDeviceID)" }
        if ($g.AdapterDACType) { "DAC: $($g.AdapterDACType)" }
    ) -join '; '
    Add-UniqueRow 'Видеокарта' "GPU #$gpuIdx — $($g.Name)" $g.AdapterCompatibility $g.Name $null $extra `
        'WMI почти никогда не отдаёт S/N видеокарты — см. nvidia-smi / наклейку'
    $pnpSn = Get-UsbSerialFromPnp $g.PNPDeviceID
    if ($pnpSn) {
        Add-UniqueRow 'Видеокарта' "GPU #$gpuIdx PNP ID tail" $g.AdapterCompatibility $g.Name $pnpSn
    }
}

$Nvidia = Get-Command nvidia-smi -ErrorAction SilentlyContinue
if ($Nvidia) {
    $nOut = Get-Safe { & nvidia-smi --query-gpu=index,name,serial,uuid,gpu_part_number,pci.bus_id --format=csv,noheader 2>$null }
    foreach ($line in @($nOut)) {
        if (-not $line) { continue }
        $parts = @($line -split ',\s*')
        if ($parts.Count -ge 3) {
            Add-UniqueRow 'Видеокарта' "NVIDIA GPU $($parts[0]) — $($parts[1])" 'NVIDIA' $parts[1] $parts[2] `
                $(@(
                    if ($parts.Count -ge 4) { "UUID: $($parts[3])" }
                    if ($parts.Count -ge 5) { "Part: $($parts[4])" }
                    if ($parts.Count -ge 6) { "PCI: $($parts[5])" }
                ) -join '; ')
        }
    }
}

# AMD: try atiadlxx / registry board serial — limited; try PnP
$AmdReg = Get-Safe {
    Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}' -ErrorAction SilentlyContinue |
        ForEach-Object { Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue }
}
foreach ($a in @($AmdReg)) {
    if (-not $a.DriverDesc) { continue }
    if ($a.DriverDesc -match 'Microsoft Basic') { continue }
    $cand = @($a.HardwareInformation_SerialNumber, $a.SerialNumber, $a.AdapterString) | Where-Object { $_ }
    foreach ($c in $cand) {
        if (Norm-Serial ([string]$c)) {
            Add-UniqueRow 'Видеокарта' "GPU реестр — $($a.DriverDesc)" 'AMD/GPU' $a.DriverDesc ([string]$c) 'Из реестра класса Display'
        }
    }
}

# ========== 8. Мониторы ==========
Write-Host "  [8/16] Мониторы..."
$Mons = @(Get-Safe { Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -ErrorAction SilentlyContinue } @())
$monIdx = 0
foreach ($mon in $Mons) {
    $monIdx++
    $decode = {
        param($bytes)
        if (-not $bytes) { return '' }
        -join ($bytes | Where-Object { $_ -ne 0 -and $_ -ge 32 -and $_ -le 126 } | ForEach-Object { [char]$_ })
    }
    $mfr = (& $decode $mon.ManufacturerName).Trim()
    $name = (& $decode $mon.UserFriendlyName).Trim()
    $sn = (& $decode $mon.SerialNumberID).Trim()
    $prod = ''
    if ($mon.ProductCodeID) { $prod = (& $decode $mon.ProductCodeID).Trim() }
    $extra = @(
        if ($prod) { "ProductCode: $prod" }
        if ($mon.YearOfManufacture) { "год: $($mon.YearOfManufacture)" }
        if ($mon.WeekOfManufacture) { "неделя: $($mon.WeekOfManufacture)" }
        if ($mon.InstanceName) { "Instance: $($mon.InstanceName)" }
    ) -join '; '
    Add-UniqueRow 'Монитор' "Монитор #$monIdx — $name" $mfr $(if ($name) { $name } else { $prod }) $sn $extra
}

$MonParams = @(Get-Safe { Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams -ErrorAction SilentlyContinue } @())
# informational only if needed

if ($Mons.Count -eq 0) {
    $Desk = @(Get-Safe { Get-CimInstance Win32_DesktopMonitor } @())
    $monIdx = 0
    foreach ($m in $Desk) {
        if (-not $m.Name) { continue }
        $monIdx++
        Add-UniqueRow 'Монитор' "Монитор #$monIdx" $m.MonitorManufacturer $m.Name $m.SerialNumber ''
    }
}

# ========== 9. Сеть ==========
Write-Host "  [9/16] Сетевые адаптеры..."
$Nics = @(Get-Safe {
    Get-CimInstance Win32_NetworkAdapter |
        Where-Object { $_.MACAddress -and $_.PNPDeviceID -notmatch 'ROOT\\|MSFT_|VPN|Virtual|Hyper-V|VMware|Loopback' }
} @())
$nicIdx = 0
foreach ($n in $Nics) {
    $nicIdx++
    $extra = @(
        if ($n.MACAddress) { "MAC: $($n.MACAddress)" }
        if ($n.PermanentAddress) { "Permanent: $($n.PermanentAddress)" }
        if ($n.Speed) { "скорость: $([math]::Round($n.Speed/1MB,0)) Мбит/с" }
        if ($n.AdapterType) { $n.AdapterType }
        if ($n.PNPDeviceID) { "PNP: $($n.PNPDeviceID)" }
        if ($n.NetConnectionID) { "Имя: $($n.NetConnectionID)" }
    ) -join '; '
    # MAC as identifier (network serial equivalent)
    Add-UniqueRow 'Сеть' "NIC #$nicIdx — $($n.Name)" $n.Manufacturer $n.Name $n.MACAddress $extra
    if ($n.PermanentAddress -and $n.PermanentAddress -ne ($n.MACAddress -replace ':', '')) {
        Add-UniqueRow 'Сеть' "NIC #$nicIdx Permanent MAC" $n.Manufacturer $n.Name $n.PermanentAddress
    }
    $pnpSn = Get-UsbSerialFromPnp $n.PNPDeviceID
    if ($pnpSn -and $pnpSn -ne $n.MACAddress) {
        Add-UniqueRow 'Сеть' "NIC #$nicIdx PNP Serial" $n.Manufacturer $n.Name $pnpSn
    }
}

# ========== 10. Батарея ==========
Write-Host "  [10/16] Батарея..."
$Bats = @(Get-Safe { Get-CimInstance Win32_Battery } @())
$batIdx = 0
foreach ($b in $Bats) {
    $batIdx++
    Add-UniqueRow 'Батарея' "Battery #$batIdx — $($b.Name)" $b.Manufacturer $b.Name $b.DeviceID `
        $(@(
            if ($b.Chemistry) { "Chemistry: $($b.Chemistry)" }
            if ($b.DesignCapacity) { "Design: $($b.DesignCapacity)" }
            if ($b.EstimatedChargeRemaining -ne $null) { "Заряд: $($b.EstimatedChargeRemaining)%" }
        ) -join '; ')
}
$BatStatic = @(Get-Safe { Get-CimInstance -Namespace root\wmi -ClassName BatteryStaticData -ErrorAction SilentlyContinue } @())
foreach ($b in $BatStatic) {
    $sn = if ($b.SerialNumber) { [string]$b.SerialNumber } else { $null }
    $mfr = ''
    if ($b.ManufactureName) {
        $mfr = -join ($b.ManufactureName | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ })
    }
    $dev = ''
    if ($b.DeviceName) {
        $dev = -join ($b.DeviceName | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ })
    }
    Add-UniqueRow 'Батарея' "BatteryStatic — $dev" $mfr $dev $sn ''
}

# ========== 11. TPM / Secure ==========
Write-Host "  [11/16] TPM..."
$Tpm = Get-Safe { Get-CimInstance -Namespace root\cimv2\Security\MicrosoftTpm -ClassName Win32_Tpm -ErrorAction SilentlyContinue }
if ($Tpm) {
    Add-UniqueRow 'TPM' 'Trusted Platform Module' $Tpm.ManufacturerIdTxt "Spec $($Tpm.SpecVersion)" $Tpm.ManufacturerIdTxt `
        "ManufacturerVersion: $($Tpm.ManufacturerVersion); Enabled: $($Tpm.IsEnabled_InitialValue); Activated: $($Tpm.IsActivated_InitialValue)"
}

# ========== 12. Оптика / звук ==========
Write-Host "  [12/16] Оптика и звук..."
$Cd = @(Get-Safe { Get-CimInstance Win32_CDROMDrive } @())
$i = 0
foreach ($c in $Cd) {
    $i++
    Add-UniqueRow 'Оптический привод' "ODD #$i — $($c.Name)" $c.Manufacturer $c.Name $c.SerialNumber $c.MediaType
    $pnpSn = Get-UsbSerialFromPnp $c.PNPDeviceID
    if ($pnpSn) { Add-UniqueRow 'Оптический привод' "ODD #$i PNP" $c.Manufacturer $c.Name $pnpSn }
}

$Sound = @(Get-Safe { Get-CimInstance Win32_SoundDevice } @())
$i = 0
foreach ($s in $Sound) {
    $i++
    Add-UniqueRow 'Звук' "Audio #$i — $($s.Name)" $s.Manufacturer $s.Name $s.DeviceID `
        $(if ($s.PNPDeviceID) { "PNP: $($s.PNPDeviceID)" } else { '' }) `
        'DeviceID звуковой карты'
}

# ========== 13. USB-устройства с серийниками ==========
Write-Host "  [13/16] USB-устройства..."
$Usb = @(Get-Safe { Get-CimInstance Win32_USBHub } @())
$i = 0
foreach ($u in $Usb) {
    $pnpSn = Get-UsbSerialFromPnp $u.PNPDeviceID
    if (-not $pnpSn -and -not (Norm-Serial $u.DeviceID)) { continue }
    $i++
    Add-UniqueRow 'USB' "USB Hub/Device #$i — $($u.Description)" '' $u.Name $(if ($pnpSn) { $pnpSn } else { $u.DeviceID }) `
        "PNP: $($u.PNPDeviceID)"
}

$UsbControllers = @(Get-Safe { Get-CimInstance Win32_USBController } @())
foreach ($u in $UsbControllers) {
    Add-UniqueRow 'USB' "USB Controller — $($u.Name)" $u.Manufacturer $u.Name $u.DeviceID "PNP: $($u.PNPDeviceID)" 'ID контроллера'
}

# PnP entities with likely serial in instance path (USB, SCSI, STORAGE)
Write-Host "  [14/16] PnP-устройства с серийниками..."
$Pnp = @(Get-Safe {
    Get-CimInstance Win32_PnPEntity |
        Where-Object {
            $_.PNPDeviceID -match '^(USB\\|SCSI\\|STORAGE\\|DISPLAY\\|HDAUDIO\\|PCI\\)(.*)\\(?![\d&]+$)' -and
            $_.PNPDeviceID -notmatch '\\{?[0-9A-Fa-f]{8}-' -and
            $_.Name -notmatch 'Generic|Root Hub|Composite|HID-compliant|USB Input'
        }
} @())
$pnpCount = 0
foreach ($p in $Pnp) {
    $pnpSn = Get-UsbSerialFromPnp $p.PNPDeviceID
    if (-not (Norm-Serial $pnpSn)) { continue }
    # skip if already captured same serial in any category
    $already = $false
    foreach ($r in $Rows) {
        if ((Norm-Serial $r.SerialRaw) -eq (Norm-Serial $pnpSn)) { $already = $true; break }
    }
    if ($already) { continue }
    $pnpCount++
    if ($pnpCount -gt 80) { break } # safety cap
    $bus = ($p.PNPDeviceID -split '\\')[0]
    Add-UniqueRow 'PnP / прочее' "$bus — $($p.Name)" $p.Manufacturer $p.Name $pnpSn "PNP: $($p.PNPDeviceID)"
}

# ========== 15. Принтеры / сканеры / прочее ==========
Write-Host "  [15/16] Принтеры и порты..."
$Printers = @(Get-Safe { Get-CimInstance Win32_Printer | Where-Object { -not $_.Network -or $_.Local } } @())
foreach ($pr in $Printers) {
    if ($pr.Name -match 'OneNote|Fax|Microsoft|XPS|PDF') { continue }
    Add-UniqueRow 'Принтер' $pr.Name $pr.DriverName $pr.Name $pr.DeviceID $pr.PortName 'DeviceID принтера'
}

$SerialPorts = @(Get-Safe { Get-CimInstance Win32_SerialPort } @())
foreach ($sp in $SerialPorts) {
    Add-UniqueRow 'Порты' "COM — $($sp.Name)" $sp.Manufacturer $sp.Name $sp.DeviceID $sp.PNPDeviceID
}

$Parallel = @(Get-Safe { Get-CimInstance Win32_ParallelPort } @())
foreach ($pp in $Parallel) {
    Add-UniqueRow 'Порты' "LPT — $($pp.Name)" '' $pp.Name $pp.DeviceID ''
}

# Keyboard / Pointing (rarely have serial)
$Kb = @(Get-Safe { Get-CimInstance Win32_Keyboard } @())
$i = 0
foreach ($k in $Kb) {
    $i++
    $pnpSn = Get-UsbSerialFromPnp $k.PNPDeviceID
    Add-UniqueRow 'Ввод' "Клавиатура #$i" '' $k.Name $(if ($pnpSn) { $pnpSn } else { $k.DeviceID }) "PNP: $($k.PNPDeviceID)"
}
$Mice = @(Get-Safe { Get-CimInstance Win32_PointingDevice } @())
$i = 0
foreach ($m in $Mice) {
    $i++
    $pnpSn = Get-UsbSerialFromPnp $m.PNPDeviceID
    if (-not $pnpSn) { continue }
    Add-UniqueRow 'Ввод' "Мышь/указатель #$i" $m.Manufacturer $m.Name $pnpSn "PNP: $($m.PNPDeviceID)"
}

# ========== 16. SCSI / RAID / Cache / Cooling extras ==========
Write-Host "  [16/16] SCSI, охлаждение, доп. SMBIOS..."
$Scsi = @(Get-Safe { Get-CimInstance Win32_SCSIController } @())
foreach ($s in $Scsi) {
    Add-UniqueRow 'Контроллер' "SCSI/RAID — $($s.Name)" $s.Manufacturer $s.Name $s.DeviceID "PNP: $($s.PNPDeviceID)"
}

$Fans = @(Get-Safe { Get-CimInstance Win32_Fan -ErrorAction SilentlyContinue } @())
foreach ($f in $Fans) {
    Add-UniqueRow 'Охлаждение' "Вентилятор — $($f.Name)" '' $f.Name $f.DeviceID ''
}
$Probe = @(Get-Safe { Get-CimInstance Win32_TemperatureProbe -ErrorAction SilentlyContinue } @())
foreach ($t in $Probe) {
    Add-UniqueRow 'Охлаждение' "Датчик — $($t.Name)" '' $t.Name $t.DeviceID ''
}

$Cache = @(Get-Safe { Get-CimInstance Win32_CacheMemory -ErrorAction SilentlyContinue } @())
foreach ($c in $Cache) {
    Add-UniqueRow 'Кэш' "Cache — $($c.Purpose)" '' $c.Purpose $c.SerialNumber ("$([math]::Round($c.MaxCapacity/1KB,1)) КБ")
}

$OnBoard = @(Get-Safe { Get-CimInstance Win32_OnBoardDevice -ErrorAction SilentlyContinue } @())
foreach ($o in $OnBoard) {
    Add-UniqueRow 'На плате' $o.Description $o.Manufacturer $o.Description $o.SerialNumber $o.Tag
}

$PortConn = @(Get-Safe { Get-CimInstance Win32_PortConnector -ErrorAction SilentlyContinue } @())
# too many without serials — skip listing all

$SysSlot = @(Get-Safe { Get-CimInstance Win32_SystemSlot -ErrorAction SilentlyContinue } @())
foreach ($s in $SysSlot) {
    if (-not $s.SerialNumber -and -not $s.Tag) { continue }
    Add-UniqueRow 'Слот' "Slot $($s.SlotDesignation)" '' $s.SlotDesignation $s.SerialNumber $s.Tag
}

# IDE / FC controllers
$Ide = @(Get-Safe { Get-CimInstance Win32_IDEController -ErrorAction SilentlyContinue } @())
foreach ($x in $Ide) {
    $pnpSn = Get-UsbSerialFromPnp $x.PNPDeviceID
    if ($pnpSn) { Add-UniqueRow 'Контроллер' "IDE — $($x.Name)" $x.Manufacturer $x.Name $pnpSn }
}
$Fc = @(Get-Safe { Get-CimInstance Win32_DiskDriveToDiskPartition -ErrorAction SilentlyContinue } @())
# skip mapping noise

# PCMCIA / Tape
$Tape = @(Get-Safe { Get-CimInstance Win32_TapeDrive -ErrorAction SilentlyContinue } @())
foreach ($t in $Tape) {
    Add-UniqueRow 'Накопитель' "Tape — $($t.Name)" $t.Manufacturer $t.Name $t.SerialNumber ''
}

# ============================================
# Summary helpers
# ============================================
$ramRows = @($Rows | Where-Object { $_.Device -like 'Планка ОЗУ*' })
$cpuRows = @($Rows | Where-Object { $_.Category -eq 'Процессор' -and $_.HasRealSerial })
$realCount = @($Rows | Where-Object { $_.HasRealSerial }).Count
$totalCount = $Rows.Count

Write-Host ""
Write-Host "  Планок ОЗУ: $stickCount  |  Всего объём: $totalRamGb ГБ"
Write-Host "  Позиций в отчёте: $totalCount  |  с идентификатором: $realCount"
Write-Host "  Формирование отчётов..."

# --- TXT ---
$txt = New-Object System.Text.StringBuilder
[void]$txt.AppendLine('=' * 78)
[void]$txt.AppendLine("  $AppTitle — максимальный сбор серийников")
[void]$txt.AppendLine("  $BrandName — $BrandUrl")
[void]$txt.AppendLine('=' * 78)
[void]$txt.AppendLine("ПК:            $HostName")
[void]$txt.AppendLine("Пользователь:  $env:USERNAME")
[void]$txt.AppendLine("ОС:            $($Os.Caption) $($Os.Version) ($($Os.OSArchitecture))")
[void]$txt.AppendLine("Производитель: $($Cs.Manufacturer)")
[void]$txt.AppendLine("Модель:        $($Cs.Model)")
[void]$txt.AppendLine("Дата сбора:    $($CollectedAt.ToString('yyyy-MM-dd HH:mm:ss'))")
[void]$txt.AppendLine("Админ-права:   $(if ($IsAdmin) { 'да' } else { 'нет' })")
[void]$txt.AppendLine("ОЗУ планок:    $stickCount  ($totalRamGb ГБ)")
[void]$txt.AppendLine("Позиций:       $totalCount  |  с реальным ID/S/N: $realCount")
[void]$txt.AppendLine('=' * 78)
[void]$txt.AppendLine('')
[void]$txt.AppendLine('КРАТКО ДЛЯ ГАРАНТИИ:')
[void]$txt.AppendLine('-' * 78)

# Quick warranty block: board, each RAM, each disk, system, CPU
$warrantyCats = @('Система','BIOS','Материнская плата','Процессор','ОЗУ','Накопитель','Монитор','Видеокарта')
foreach ($cat in $warrantyCats) {
    $items = @($Rows | Where-Object { $_.Category -eq $cat -and ($_.HasRealSerial -or $_.Device -like 'Планка ОЗУ*' -or $_.Device -like 'CPU планка*' -or $_.Device -like 'GPU #*') })
    if ($items.Count -eq 0) { continue }
    [void]$txt.AppendLine("[$cat]")
    foreach ($r in $items) {
        [void]$txt.AppendLine("  • $($r.Device)")
        if ($r.Manufacturer) { [void]$txt.AppendLine("    Производитель: $($r.Manufacturer)") }
        if ($r.Model) { [void]$txt.AppendLine("    Модель: $($r.Model)") }
        [void]$txt.AppendLine("    S/N: $($r.Serial)")
    }
    [void]$txt.AppendLine('')
}

[void]$txt.AppendLine('=' * 78)
[void]$txt.AppendLine('ПОЛНЫЙ СПИСОК')
[void]$txt.AppendLine('=' * 78)

$grouped = $Rows | Group-Object Category
foreach ($g in $grouped) {
    [void]$txt.AppendLine('')
    [void]$txt.AppendLine(('-' * 78))
    [void]$txt.AppendLine("  $($g.Name)")
    [void]$txt.AppendLine(('-' * 78))
    foreach ($r in $g.Group) {
        [void]$txt.AppendLine("Устройство:    $($r.Device)")
        if ($r.Manufacturer) { [void]$txt.AppendLine("Производитель: $($r.Manufacturer)") }
        if ($r.Model)        { [void]$txt.AppendLine("Модель:        $($r.Model)") }
        [void]$txt.AppendLine("Серийный №:    $($r.Serial)")
        if ($r.Extra)        { [void]$txt.AppendLine("Дополнительно: $($r.Extra)") }
        [void]$txt.AppendLine('')
    }
}

[void]$txt.AppendLine('=' * 78)
[void]$txt.AppendLine('Примечания:')
[void]$txt.AppendLine('• Каждая планка ОЗУ выводится отдельно (Планка ОЗУ #1, #2, …) со своим S/N.')
[void]$txt.AppendLine('• У CPU заводской серийник Windows почти никогда не отдаёт (отключён в Intel/AMD).')
[void]$txt.AppendLine('  Для процессора смотрите ProcessorId / PartNumber и наклейку на коробке/кулере.')
[void]$txt.AppendLine('• У многих GPU S/N только через nvidia-smi или с наклейки на плате.')
[void]$txt.AppendLine('• Volume SerialNumber раздела — НЕ серийник SSD/HDD.')
[void]$txt.AppendLine('• Запуск от администратора даёт более полные данные.')
[void]$txt.AppendLine('')
[void]$txt.AppendLine("Сделано: $BrandName — $BrandUrl")
[void]$txt.AppendLine('=' * 78)

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

function Rows-To-Html($items) {
    foreach ($r in $items) {
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
}

$highlight = @($Rows | Where-Object {
    $_.Category -in @('Материнская плата','ОЗУ','Накопитель','Процессор','Система','BIOS','Монитор','Видеокарта') -and
    ($_.HasRealSerial -or $_.Device -like 'Планка ОЗУ*' -or $_.Device -like 'CPU планка*' -or $_.Device -like 'Диск #*' -or $_.Device -like 'ПК *' -or $_.Device -like 'Motherboard*' -or $_.Device -like 'GPU #*' -or $_.Device -like 'NVIDIA*' -or $_.Device -like 'Монитор #*' -or $_.Device -like 'BIOS*')
})

$htmlHighlight = (Rows-To-Html $highlight) -join "`n"
$htmlAll = (Rows-To-Html $Rows) -join "`n"

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
  }
  .brand { font-size: 1.75rem; font-weight: 700; margin: 0 0 0.25rem; }
  .brand span { color: var(--accent); }
  .tagline { color: var(--muted); margin: 0 0 1rem; font-size: 0.95rem; }
  .meta {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
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
  main { padding: 1.25rem 1.5rem 3rem; max-width: 1280px; margin: 0 auto; }
  h2 { font-size: 1.1rem; margin: 1.5rem 0 0.75rem; }
  .note {
    background: rgba(61,214,198,.08);
    border: 1px solid rgba(61,214,198,.25);
    border-radius: 10px;
    padding: 0.85rem 1rem;
    color: var(--muted);
    font-size: 0.9rem;
  }
  .wrap { overflow-x: auto; border: 1px solid var(--line); border-radius: 12px; background: var(--panel); margin-bottom: 0.5rem; }
  table { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
  th, td { padding: 0.6rem 0.7rem; text-align: left; border-bottom: 1px solid var(--line); vertical-align: top; }
  th { color: var(--muted); font-weight: 600; font-size: 0.72rem; text-transform: uppercase; letter-spacing: .04em; background: rgba(0,0,0,.2); }
  tr:hover td { background: rgba(61,214,198,.04); }
  td.ok { color: var(--ok); font-family: Consolas, "Courier New", monospace; font-weight: 600; }
  td.miss { color: var(--miss); font-style: italic; }
  td.extra { color: var(--muted); font-size: 0.8rem; max-width: 300px; }
  footer { margin-top: 2rem; padding-top: 1rem; border-top: 1px solid var(--line); color: var(--muted); font-size: 0.85rem; }
  footer a { color: var(--accent); text-decoration: none; }
</style>
</head>
<body>
<header>
  <p class="brand">$BrandName <span>— $BrandUrl</span></p>
  <p class="tagline">$AppTitle — все доступные серийники комплектующих (ОЗУ попланочно, CPU, диски…)</p>
  <div class="meta">
    <div><span class="k">Компьютер</span><span class="v">$(Html-Enc $HostName)</span></div>
    <div><span class="k">ОС</span><span class="v">$(Html-Enc ($Os.Caption + ' ' + $Os.OSArchitecture))</span></div>
    <div><span class="k">Модель ПК</span><span class="v">$(Html-Enc ($Cs.Manufacturer + ' / ' + $Cs.Model))</span></div>
    <div><span class="k">Планки ОЗУ</span><span class="v">$stickCount шт. / $totalRamGb ГБ</span></div>
    <div><span class="k">С ID/S/N</span><span class="v">$realCount из $totalCount</span></div>
    <div><span class="k">Дата</span><span class="v">$(Html-Enc $CollectedAt.ToString('yyyy-MM-dd HH:mm:ss'))</span></div>
  </div>
</header>
<main>
  <div class="note">
    <strong>ОЗУ:</strong> каждая планка отдельной строкой («Планка ОЗУ #1», «#2»…) — у двух модулей будет два серийника.<br>
    <strong>CPU:</strong> заводской S/N Windows обычно не видит (так задумано Intel/AMD). В отчёте есть ProcessorId / PartNumber; для гарантии сверьте наклейку на коробке.<br>
    Зелёным — реальные номера. Запускайте от администратора.
  </div>

  <h2>Главное для сверки с гарантией</h2>
  <div class="wrap">
    <table>
      <thead>
        <tr>
          <th>Категория</th><th>Устройство</th><th>Производитель</th>
          <th>Модель</th><th>Серийный номер</th><th>Дополнительно</th>
        </tr>
      </thead>
      <tbody>
$htmlHighlight
      </tbody>
    </table>
  </div>

  <h2>Полный список (все найденные запчасти)</h2>
  <div class="wrap">
    <table>
      <thead>
        <tr>
          <th>Категория</th><th>Устройство</th><th>Производитель</th>
          <th>Модель</th><th>Серийный номер / ID</th><th>Дополнительно</th>
        </tr>
      </thead>
      <tbody>
$htmlAll
      </tbody>
    </table>
  </div>

  <footer>
    Сделано: <strong>$BrandName</strong> — <a href="https://$BrandUrl">$BrandUrl</a><br>
    Файлы: $(Html-Enc (Split-Path $TxtPath -Leaf)), $(Html-Enc (Split-Path $CsvPath -Leaf))
  </footer>
</main>
</body>
</html>
"@

[System.IO.File]::WriteAllText($HtmlPath, $html, $utf8Bom)

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
