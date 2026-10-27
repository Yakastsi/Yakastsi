#Requires -Version 5.1
<#
.SYNOPSIS
  Maximal hardware serial collection for Windows 10/11 warranty checks.
.NOTES
  Yakastsi - rgb.by
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
$WarrantyPath = Join-Path $ReportsDir "$BaseName-GARANTIYA.txt"

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
    if ($null -eq $Value) { return $null }
    $v = [string]$Value
    $v = ($v -replace '[\x00-\x1F\x7F]', '').Trim()
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    $junk = @(
        'To Be Filled By O.E.M.', 'To be filled by O.E.M.', 'Default string', 'Default String',
        'System Serial Number', 'System Product Name', 'None', 'N/A', 'NA', 'n/a',
        '0', '00000000', '0000000000', '000000000000', '0000000000000000',
        'FFFFFFFF', 'FFFFFFFFFFFFFFFF', 'OEM', 'O.E.M.',
        'Chassis Serial Number', 'Base Board Serial Number', 'Base Board',
        'Empty', 'Unknown', 'Not Specified', 'Not Available', 'No Enclosure',
        'Invalid', 'undefined', 'No Asset Tag', 'NoAssetTag', 'ASSET TAG',
        'Motherboard', 'NULL', 'null', 'None.', 'S/N', 'Serial Number',
        '123456789', 'ABCDEF', 'XXXXXX', 'xxxxxxx'
    )
    foreach ($j in $junk) { if ($v -eq $j) { return $null } }
    if ($v -match '^(0+|F+|X+|x+)$') { return $null }
    if ($v -match '^(No Asset Tag|Base Board|Motherboard)$') { return $null }
    # DeviceID-like paths are not factory serials for warranty summary
    return $v
}

function Is-DevicePathId {
    param([string]$v)
    if (-not $v) { return $false }
    return [bool]($v -match '^(PCI\\|USB\\|ACPI\\|ROOT\\|SWD\\|BTH\\|HDAUDIO\\|STORAGE\\|SCSI\\|\\\\)')
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
        [string]$Note = '',
        [bool]$Warranty = $false
    )
    $serialNorm = Norm-Serial $Serial
    if ($serialNorm -and (Is-DevicePathId $serialNorm)) {
        # keep as info but not as warranty serial
        if (-not $Note) { $Note = 'Это DeviceID/PNP, не заводской S/N для гарантии' }
        $Warranty = $false
    }
    $display = if ($serialNorm) { $serialNorm }
               elseif ($Note) { $Note }
               else { [char]0x2014 + ' не указан / скрыт производителем ' + [char]0x2014 }

    $List.Add([pscustomobject]@{
        Category      = $Category
        Device        = $Device
        Manufacturer  = if ($Manufacturer) { $Manufacturer.Trim() } else { '' }
        Model         = if ($Model) { $Model.Trim() } else { '' }
        Serial        = $display
        SerialRaw     = $Serial
        HasRealSerial = [bool]$serialNorm
        Extra         = $Extra
        Warranty      = [bool]$Warranty
    })
}

function Get-UsbSerialFromPnp {
    param([string]$PnpId)
    if (-not $PnpId) { return $null }
    if ($PnpId -match '\\([^\\]+)$') {
        $tail = $Matches[1]
        if ($tail -match '^(5&|6&|7&|[0-9]+&)') { return $null }
        if ($tail -match '^[&0-9]+$') { return $null }
        if ($tail.Length -lt 4) { return $null }
        return $tail
    }
    return $null
}

function Get-MemoryTypeName {
    param($Type, $SmbiosType)
    $map = @{
        20='DDR'; 21='DDR2'; 22='DDR2 FB-DIMM'; 24='DDR3'; 26='DDR4'
        27='LPDDR'; 28='LPDDR2'; 29='LPDDR3'; 30='LPDDR4'; 34='DDR5'; 35='LPDDR5'
    }
    if ($SmbiosType -and $map.ContainsKey([int]$SmbiosType)) { return $map[[int]$SmbiosType] }
    if ($Type -and $map.ContainsKey([int]$Type)) { return $map[[int]$Type] }
    return ''
}

function Get-FormFactorName {
    param($ff)
    $map = @{ 8='DIMM'; 12='SO-DIMM'; 13='Micro-DIMM' }
    if ($ff -and $map.ContainsKey([int]$ff)) { return $map[[int]$ff] }
    return ''
}

function Get-RamVendorName {
    param([string]$Raw)
    if (-not $Raw) { return '' }
    $t = $Raw.Trim()
    # Already a name
    if ($t -match '[A-Za-z]{3,}' -and $t -notmatch '^\d+$') { return $t }
    # Common WMI numeric OEM ids seen on Lenovo etc.
    $map = @{
        '1302' = 'Patriot'
        '1315' = 'Crucial'
        '0198' = 'Kingston'
        '198'  = 'Kingston'
        '04CD' = 'G.Skill'
        '04cd' = 'G.Skill'
        '029E' = 'Corsair'
        '2E'   = 'Corsair'
        '80AD' = 'Hynix/SK hynix'
        '80CE' = 'Samsung'
        '859B' = 'Crucial/Micron'
        '00CE' = 'Samsung'
        '002C' = 'Micron'
        'AD00' = 'Hynix'
        'CE00' = 'Samsung'
        '2C00' = 'Micron'
        '9912' = 'A-DATA'
        '0C91' = 'A-DATA'
        '7F7F7F7F7F510000' = 'Transcend'
    }
    if ($map.ContainsKey($t)) { return "$($map[$t]) ($t)" }
    return $t
}

# ---------- Raw SMBIOS parser (Type 1/2/3/4/17) ----------
function Get-SmbiosString {
    param([byte[]]$Data, [int]$StructStart, [int]$StructLen, [int]$StringIndex)
    if ($StringIndex -le 0) { return $null }
    $pos = $StructStart + $StructLen
    $idx = 1
    while ($pos -lt $Data.Length) {
        if ($Data[$pos] -eq 0) {
            # end of string table (double-null) or empty string
            if (($pos + 1) -lt $Data.Length -and $Data[$pos + 1] -eq 0) { return $null }
            $pos++
            continue
        }
        $end = $pos
        while ($end -lt $Data.Length -and $Data[$end] -ne 0) { $end++ }
        if ($idx -eq $StringIndex) {
            $len = $end - $pos
            if ($len -le 0) { return $null }
            return [System.Text.Encoding]::ASCII.GetString($Data, $pos, $len).Trim()
        }
        $idx++
        $pos = $end + 1
        if ($pos -lt $Data.Length -and $Data[$pos] -eq 0) { break }
    }
    return $null
}

function Get-RawSmbiosRecords {
    $raw = Get-Safe {
        Get-CimInstance -Namespace root\wmi -ClassName MSSmBios_RawSMBiosTables -ErrorAction Stop |
            Select-Object -First 1
    }
    if (-not $raw -or -not $raw.SMBiosData) { return @() }

    $data = [byte[]]$raw.SMBiosData
    $list = New-Object System.Collections.Generic.List[object]
    $i = 0
    $guard = 0
    while ($i + 4 -lt $data.Length -and $guard -lt 5000) {
        $guard++
        $type = $data[$i]
        $len  = $data[$i + 1]
        if ($len -lt 4) { break }
        $handle = [BitConverter]::ToUInt16($data, $i + 2)

        # find end of string table (double NUL)
        $p = $i + $len
        while ($p + 1 -lt $data.Length -and -not ($data[$p] -eq 0 -and $data[$p + 1] -eq 0)) { $p++ }
        $next = [Math]::Min($p + 2, $data.Length)

        $getStr = {
            param($strIdx)
            Get-SmbiosString -Data $data -StructStart $i -StructLen $len -StringIndex $strIdx
        }

        switch ($type) {
            1 { # System
                if ($len -ge 0x19) {
                    $list.Add([pscustomobject]@{
                        Kind='System'; Handle=$handle
                        Manufacturer=(& $getStr $data[$i+0x04])
                        Product=(& $getStr $data[$i+0x05])
                        Version=(& $getStr $data[$i+0x06])
                        Serial=(& $getStr $data[$i+0x07])
                        SKU=(& $getStr $(if ($len -gt 0x19) { $data[$i+0x19] } else { 0 }))
                        Family=(& $getStr $(if ($len -gt 0x1A) { $data[$i+0x1A] } else { 0 }))
                    })
                }
            }
            2 { # Baseboard
                if ($len -ge 0x08) {
                    $list.Add([pscustomobject]@{
                        Kind='Baseboard'; Handle=$handle
                        Manufacturer=(& $getStr $data[$i+0x04])
                        Product=(& $getStr $data[$i+0x05])
                        Version=(& $getStr $data[$i+0x06])
                        Serial=(& $getStr $data[$i+0x07])
                        AssetTag=(& $getStr $(if ($len -gt 0x08) { $data[$i+0x08] } else { 0 }))
                    })
                }
            }
            3 { # Chassis
                if ($len -ge 0x09) {
                    $list.Add([pscustomobject]@{
                        Kind='Chassis'; Handle=$handle
                        Manufacturer=(& $getStr $data[$i+0x04])
                        Serial=(& $getStr $data[$i+0x07])
                        AssetTag=(& $getStr $data[$i+0x08])
                    })
                }
            }
            4 { # Processor
                if ($len -ge 0x21) {
                    $list.Add([pscustomobject]@{
                        Kind='Processor'; Handle=$handle
                        Socket=(& $getStr $data[$i+0x04])
                        Manufacturer=(& $getStr $data[$i+0x07])
                        Version=(& $getStr $data[$i+0x10])
                        Serial=(& $getStr $data[$i+0x20])
                        AssetTag=(& $getStr $data[$i+0x21])
                        PartNumber=(& $getStr $(if ($len -gt 0x22) { $data[$i+0x22] } else { 0 }))
                    })
                }
            }
            17 { # Memory Device
                if ($len -ge 0x1B) {
                    $sizeRaw = [BitConverter]::ToUInt16($data, $i + 0x0C)
                    $sizeMb = 0
                    if ($sizeRaw -ne 0xFFFF) {
                        if (($sizeRaw -band 0x8000) -ne 0) { $sizeMb = ($sizeRaw -band 0x7FFF) } # KB units rare
                        else { $sizeMb = $sizeRaw }
                    }
                    # SMBIOS 2.7+ extended size at 0x1C
                    if ($sizeRaw -eq 0x7FFF -and $len -ge 0x20) {
                        $sizeMb = [BitConverter]::ToUInt32($data, $i + 0x1C)
                    }
                    $list.Add([pscustomobject]@{
                        Kind='Memory'; Handle=$handle
                        SizeMB=$sizeMb
                        DeviceLocator=(& $getStr $data[$i+0x10])
                        BankLocator=(& $getStr $data[$i+0x11])
                        Manufacturer=(& $getStr $data[$i+0x17])
                        Serial=(& $getStr $data[$i+0x18])
                        AssetTag=(& $getStr $data[$i+0x19])
                        PartNumber=(& $getStr $data[$i+0x1A])
                        Speed=$(if ($len -ge 0x17) { [BitConverter]::ToUInt16($data, $i + 0x15) } else { 0 })
                        FormFactor=$data[$i+0x0E]
                        MemoryType=$data[$i+0x12]
                    })
                }
            }
        }

        if ($type -eq 127) { break } # End of table
        $i = $next
    }
    return $list.ToArray()
}

function Find-NvidiaSmi {
    $cmds = @(
        'nvidia-smi'
        "${env:ProgramFiles}\NVIDIA Corporation\NVSMI\nvidia-smi.exe"
        "${env:ProgramFiles(x86)}\NVIDIA Corporation\NVSMI\nvidia-smi.exe"
        "${env:SystemRoot}\System32\nvidia-smi.exe"
    )
    foreach ($c in $cmds) {
        $cmd = Get-Command $c -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd.Source }
        if (Test-Path -LiteralPath $c) { return $c }
    }
    return $null
}

# ===================== collect =====================
$Rows = New-Object System.Collections.Generic.List[object]
$SeenSerialKeys = New-Object 'System.Collections.Generic.HashSet[string]'
$IsAdmin = Test-IsAdmin
$CollectedAt = Get-Date
$Os = Get-Safe { Get-CimInstance Win32_OperatingSystem }
$Cs = Get-Safe { Get-CimInstance Win32_ComputerSystem }

function Add-UniqueRow {
    param(
        [string]$Category, [string]$Device, [string]$Manufacturer,
        [string]$Model, [string]$Serial, [string]$Extra = '', [string]$Note = '',
        [bool]$Warranty = $false
    )
    $sn = Norm-Serial $Serial
    if ($sn) {
        $key = "$Category|$sn|$(Norm-Serial $Model)"
        if (-not $SeenSerialKeys.Add($key)) { return }
    } elseif (-not $Warranty -and -not $Note -and -not $Serial) {
        # still add empty warranty-critical rows
    }
    Add-Row $Rows $Category $Device $Manufacturer $Model $Serial $Extra $Note $Warranty
}

Write-Host "  Host: $HostName"
Write-Host "  Admin: $(if ($IsAdmin) { 'yes' } else { 'no' })"
Write-Host ""

Write-Host "  [0/17] Raw SMBIOS tables..."
$Smbios = @(Get-RawSmbiosRecords)
Write-Host "         SMBIOS records parsed: $($Smbios.Count)"

# ========== 1. System ==========
Write-Host "  [1/17] System / chassis..."
$Enclosure = Get-Safe { Get-CimInstance Win32_SystemEnclosure }
$Product = Get-Safe { Get-CimInstance Win32_ComputerSystemProduct }
$smbSys = @($Smbios | Where-Object Kind -eq 'System' | Select-Object -First 1)
$sysSerial = $null
if ($Product) { $sysSerial = $Product.IdentifyingNumber }
if (-not (Norm-Serial $sysSerial) -and $smbSys) { $sysSerial = $smbSys.Serial }

Add-UniqueRow 'Система' 'ПК (System Product)' `
    $(if ($Product) { $Product.Vendor } elseif ($smbSys) { $smbSys.Manufacturer } else { $Cs.Manufacturer }) `
    $(if ($Product) { $Product.Name } elseif ($smbSys) { $smbSys.Product } else { $Cs.Model }) `
    $sysSerial `
    $(@(
        if ($Product -and $Product.UUID) { "UUID: $($Product.UUID)" }
        if ($Product -and $Product.Version) { "Версия: $($Product.Version)" }
        if ($Cs -and $Cs.SystemSKUNumber) { "SKU: $($Cs.SystemSKUNumber)" }
    ) -join '; ') -Warranty $true

$smbCh = @($Smbios | Where-Object Kind -eq 'Chassis' | Select-Object -First 1)
if ($Enclosure -or $smbCh) {
    $chSn = if ($Enclosure) { $Enclosure.SerialNumber } else { $null }
    if (-not (Norm-Serial $chSn) -and $smbCh) { $chSn = $smbCh.Serial }
    if (Norm-Serial $chSn) {
        Add-UniqueRow 'Система' 'Корпус (Chassis)' $Cs.Manufacturer $Cs.Model $chSn '' -Warranty $true
    }
    $asset = $null
    if ($Enclosure -and $Enclosure.SMBIOSAssetTag) { $asset = $Enclosure.SMBIOSAssetTag }
    if (-not (Norm-Serial $asset) -and $smbCh) { $asset = $smbCh.AssetTag }
    if (Norm-Serial $asset) {
        Add-UniqueRow 'Система' 'Asset Tag' $Cs.Manufacturer $Cs.Model $asset '' -Warranty $true
    }
}

# ========== 2. BIOS ==========
Write-Host "  [2/17] BIOS..."
$Bios = Get-Safe { Get-CimInstance Win32_BIOS }
if ($Bios) {
    Add-UniqueRow 'BIOS' 'BIOS / UEFI' $Bios.Manufacturer $Bios.SMBIOSBIOSVersion $Bios.SerialNumber `
        $(@(
            if ($Bios.ReleaseDate) { "Дата: $($Bios.ReleaseDate.ToString('yyyy-MM-dd'))" }
            if ($Bios.SMBIOSMajorVersion) { "SMBIOS: $($Bios.SMBIOSMajorVersion).$($Bios.SMBIOSMinorVersion)" }
        ) -join '; ') -Warranty $true
}

# ========== 3. Board ==========
Write-Host "  [3/17] Motherboard..."
$Board = Get-Safe { Get-CimInstance Win32_BaseBoard }
$smbBoard = @($Smbios | Where-Object Kind -eq 'Baseboard' | Select-Object -First 1)
$boardSn = if ($Board) { $Board.SerialNumber } else { $null }
if (-not (Norm-Serial $boardSn) -and $smbBoard) { $boardSn = $smbBoard.Serial }
Add-UniqueRow 'Материнская плата' 'Motherboard' `
    $(if ($Board) { $Board.Manufacturer } elseif ($smbBoard) { $smbBoard.Manufacturer } else { '' }) `
    $(if ($Board) { $Board.Product } elseif ($smbBoard) { $smbBoard.Product } else { '' }) `
    $boardSn `
    $(if ($Board -and $Board.Version) { "Версия: $($Board.Version)" } else { '' }) -Warranty $true

# ========== 4. CPU ==========
Write-Host "  [4/17] CPU..."
$Cpus = @(Get-Safe { Get-CimInstance Win32_Processor } @())
$smbCpus = @($Smbios | Where-Object Kind -eq 'Processor')
$cpuIdx = 0
foreach ($cpu in $Cpus) {
    $cpuIdx++
    $smbCpu = $null
    if ($smbCpus.Count -ge $cpuIdx) { $smbCpu = $smbCpus[$cpuIdx - 1] }
    elseif ($smbCpus.Count -eq 1) { $smbCpu = $smbCpus[0] }

    $cpuSn = Norm-Serial $cpu.SerialNumber
    if (-not $cpuSn -and $smbCpu) { $cpuSn = Norm-Serial $smbCpu.Serial }

    $extra = @(
        if ($cpu.SocketDesignation) { "Сокет: $($cpu.SocketDesignation)" }
        "Ядер: $($cpu.NumberOfCores)"; "Потоков: $($cpu.NumberOfLogicalProcessors)"
        "Макс. МГц: $($cpu.MaxClockSpeed)"
        if ($cpu.ProcessorId) { "ProcessorId: $($cpu.ProcessorId)" }
    ) -join '; '

    if ($cpuSn) {
        Add-UniqueRow 'Процессор' "CPU #$cpuIdx — $($cpu.Name)" $cpu.Manufacturer $cpu.Name $cpuSn $extra -Warranty $true
    } else {
        Add-UniqueRow 'Процессор' "CPU #$cpuIdx — $($cpu.Name)" $cpu.Manufacturer $cpu.Name $null $extra `
            'Заводской S/N CPU скрыт Intel/AMD в SMBIOS. Для гарантии: ProcessorId + коробка/наклейка.' -Warranty $true
    }

    if ($cpu.ProcessorId) {
        Add-UniqueRow 'Процессор' "CPU #$cpuIdx ProcessorId" $cpu.Manufacturer $cpu.Name $cpu.ProcessorId `
            'CPUID — не гарантийный S/N с коробки, но стабильный ID этого CPU' -Warranty $true
    }
    $part = $null
    if ($smbCpu) { $part = Norm-Serial $smbCpu.PartNumber }
    if (-not $part) { $part = Norm-Serial $cpu.PartNumber }
    if ($part) {
        Add-UniqueRow 'Процессор' "CPU #$cpuIdx PartNumber" $cpu.Manufacturer $cpu.Name $part '' -Warranty $true
    }
}

# ========== 5. RAM per stick ==========
Write-Host "  [5/17] RAM modules..."
$Mem = @(Get-Safe { Get-CimInstance Win32_PhysicalMemory } @())
$smbMemAll = @($Smbios | Where-Object { $_.Kind -eq 'Memory' -and $_.SizeMB -gt 0 })
$stickCount = [Math]::Max($Mem.Count, $smbMemAll.Count)
if ($Mem.Count -eq 0 -and $smbMemAll.Count -gt 0) {
    # synthesize from SMBIOS only
    $Mem = @()
}
$totalRamGb = 0
Write-Host "         WMI sticks: $($Mem.Count); SMBIOS sticks: $($smbMemAll.Count)"

$maxSticks = [Math]::Max($Mem.Count, $smbMemAll.Count)
for ($stickIdx = 1; $stickIdx -le $maxSticks; $stickIdx++) {
    $m = if ($stickIdx -le $Mem.Count) { $Mem[$stickIdx - 1] } else { $null }
    $sm = $null
    if ($m) {
        # match SMBIOS by locator
        $loc = if ($m.DeviceLocator) { $m.DeviceLocator.Trim() } else { '' }
        $sm = @($smbMemAll | Where-Object { $_.DeviceLocator -and ($_.DeviceLocator.Trim() -eq $loc) } | Select-Object -First 1)
        if (-not $sm -and $stickIdx -le $smbMemAll.Count) { $sm = $smbMemAll[$stickIdx - 1] }
    } else {
        $sm = $smbMemAll[$stickIdx - 1]
    }

    $gb = 0
    if ($m -and $m.Capacity) { $gb = [math]::Round($m.Capacity / 1GB, 2) }
    elseif ($sm -and $sm.SizeMB) { $gb = [math]::Round($sm.SizeMB / 1024, 2) }
    $totalRamGb += $gb

    $mfrRaw = if ($m -and $m.Manufacturer) { $m.Manufacturer } elseif ($sm) { $sm.Manufacturer } else { '' }
    $mfr = Get-RamVendorName $mfrRaw
    $part = ''
    if ($m -and $m.PartNumber) { $part = $m.PartNumber.Trim() }
    if ((-not $part) -and $sm -and $sm.PartNumber) { $part = $sm.PartNumber.Trim() }

    $sn = $null
    if ($m) { $sn = Norm-Serial $m.SerialNumber }
    if (-not $sn -and $sm) { $sn = Norm-Serial $sm.Serial }
    # Sometimes serial is numeric with leading spaces only in one source
    if (-not $sn -and $m -and $m.SerialNumber) {
        $hex = -join ([byte[]][char[]]([string]$m.SerialNumber).ToCharArray() | ForEach-Object { '{0:X2}' -f [int]$_ })
        if ($hex -and $hex -notmatch '^(00)+$') {
            # don't use hex as serial; just leave empty
        }
    }

    $typeName = ''
    $ff = ''
    $speed = ''
    $locator = ''
    if ($m) {
        $typeName = Get-MemoryTypeName $m.MemoryType $m.SMBIOSMemoryType
        $ff = Get-FormFactorName $m.FormFactor
        if ($m.ConfiguredClockSpeed) { $speed = "$($m.ConfiguredClockSpeed) МГц" }
        elseif ($m.Speed) { $speed = "$($m.Speed) МГц" }
        $locator = $m.DeviceLocator
    } elseif ($sm) {
        $typeName = Get-MemoryTypeName $sm.MemoryType $null
        $ff = Get-FormFactorName $sm.FormFactor
        if ($sm.Speed) { $speed = "$($sm.Speed) МГц" }
        $locator = $sm.DeviceLocator
    }

    $extra = @(
        "$gb ГБ"
        if ($typeName) { $typeName }
        if ($ff) { $ff }
        if ($speed) { $speed }
        if ($locator) { "слот: $locator" }
        if ($m -and $m.BankLabel) { "банк: $($m.BankLabel)" }
        if ($part) { "Part#: $part" }
        if ($mfrRaw -and $mfr -ne $mfrRaw) { "SMBIOS mfr code: $mfrRaw" }
        if ($sm -and (Norm-Serial $sm.Serial)) { "SMBIOS S/N: $($sm.Serial)" }
        if ($m -and (Norm-Serial $m.SerialNumber) -and $sm -and (Norm-Serial $sm.Serial) -and ((Norm-Serial $m.SerialNumber) -ne (Norm-Serial $sm.Serial))) {
            "WMI S/N: $($m.SerialNumber.Trim())"
        }
    ) -join '; '

    $model = if ($part) { $part } else { "$gb ГБ $typeName" }
    $slotLabel = if ($locator) { $locator } else { "#$stickIdx" }

    if ($sn) {
        Add-UniqueRow 'ОЗУ' "Планка ОЗУ #$stickIdx (слот $slotLabel)" $mfr $model $sn $extra -Warranty $true
    } else {
        # No serial in firmware — put Part# as warranty fallback identity
        $note = "S/N планки НЕ прописан в SMBIOS/SPD (часто у Patriot/OEM). Для гарантии используйте Part#: $part и наклейку на модуле."
        Add-UniqueRow 'ОЗУ' "Планка ОЗУ #$stickIdx (слот $slotLabel)" $mfr $model $(if ($part) { "PART:$part" } else { $null }) $extra $note -Warranty $true
    }
}

# ========== 6. Disks ==========
Write-Host "  [6/17] Storage..."
$Disks = @(Get-Safe { Get-CimInstance Win32_DiskDrive } @())
$diskIdx = 0
foreach ($d in $Disks) {
    $diskIdx++
    $gb = if ($d.Size) { [math]::Round($d.Size / 1GB, 2) } else { $null }
    $sn = if ($d.SerialNumber) { $d.SerialNumber.Trim() } else { $null }
    $isRemovable = ($d.MediaType -match 'Removable') -or ($d.InterfaceType -eq 'USB') -or ($d.Model -match 'USB Flash|USB Device')
    $extra = @("$gb ГБ"; $d.InterfaceType; $d.MediaType; if ($d.FirmwareRevision) { "FW: $($d.FirmwareRevision)" }; if ($isRemovable) { 'СЪЁМНЫЙ/USB' }) -join '; '
    Add-UniqueRow 'Накопитель' "Диск #$diskIdx — $($d.Model)" $d.Manufacturer $d.Model $sn $extra -Warranty (-not $isRemovable)
}

foreach ($pd in @(Get-Safe { Get-PhysicalDisk -ErrorAction SilentlyContinue } @())) {
    $sn = if ($pd.SerialNumber) { $pd.SerialNumber.Trim() } else { $null }
    $gb = if ($pd.Size) { [math]::Round($pd.Size / 1GB, 2) } else { $null }
    $isUsb = ($pd.BusType -eq 'USB')
    Add-UniqueRow 'Накопитель' "PhysicalDisk $($pd.DeviceId) — $($pd.FriendlyName)" `
        $pd.Manufacturer $(if ($pd.Model) { $pd.Model } else { $pd.FriendlyName }) $sn `
        ("$gb ГБ; $($pd.MediaType); $($pd.BusType); FW: $($pd.FirmwareVersion)") -Warranty (-not $isUsb)
}

# Prefer Storage namespace for any missed serials
foreach ($d in @(Get-Safe { Get-CimInstance -Namespace root\Microsoft\Windows\Storage -ClassName MSFT_PhysicalDisk -ErrorAction SilentlyContinue } @())) {
    $sn = if ($d.SerialNumber) { ([string]$d.SerialNumber).Trim() } else { $null }
    if (-not (Norm-Serial $sn)) { continue }
    $isUsb = ($d.BusType -eq 7) # USB
    Add-UniqueRow 'Накопитель' "MSFT_PhysicalDisk — $($d.FriendlyName)" $d.Manufacturer $d.Model $sn `
        ("FW: $($d.FirmwareVersion)") -Warranty (-not $isUsb)
}

foreach ($l in @(Get-Safe { Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } } @())) {
    $gb = if ($l.Size) { [math]::Round($l.Size / 1GB, 2) } else { $null }
    Add-UniqueRow 'Раздел' "Том $($l.DeviceID)" '' $l.VolumeName $l.VolumeSerialNumber `
        ("$gb ГБ; FS: $($l.FileSystem) — это S/N ТОМА, не диска") -Warranty $false
}

# ========== 7. GPU ==========
Write-Host "  [7/17] GPU..."
$Gpus = @(Get-Safe { Get-CimInstance Win32_VideoController } @())
$gpuIdx = 0
foreach ($g in $Gpus) {
    if ($g.Name -match 'Microsoft Basic|Remote Desktop|Virtual Display|Basic Render') { continue }
    $gpuIdx++
    $extra = @(
        if ($g.DriverVersion) { "драйвер: $($g.DriverVersion)" }
        if ($g.PNPDeviceID) { "PNP: $($g.PNPDeviceID)" }
        if ($g.VideoProcessor) { "chip: $($g.VideoProcessor)" }
    ) -join '; '
    Add-UniqueRow 'Видеокарта' "GPU #$gpuIdx — $($g.Name)" $g.AdapterCompatibility $g.Name $null $extra `
        'Windows/WMI не отдаёт заводской S/N GPU. NVIDIA: см. ниже nvidia-smi; иначе наклейка на чипе/коробке.' -Warranty $true
}

$nvsmi = Find-NvidiaSmi
if ($nvsmi) {
    Write-Host "         nvidia-smi: $nvsmi"
    $nOut = Get-Safe { & $nvsmi --query-gpu=index,name,serial,uuid,gpu_part_number,pci.bus_id --format=csv,noheader 2>$null }
    $any = $false
    foreach ($line in @($nOut)) {
        if (-not $line) { continue }
        $parts = @($line -split ',\s*')
        if ($parts.Count -ge 3) {
            $any = $true
            $gpuSerial = $parts[2]
            if ($gpuSerial -match 'N/A|Unknown') { $gpuSerial = $null }
            Add-UniqueRow 'Видеокарта' "NVIDIA GPU $($parts[0]) — $($parts[1])" 'NVIDIA' $parts[1] $gpuSerial `
                $(@(
                    if ($parts.Count -ge 4) { "UUID: $($parts[3])" }
                    if ($parts.Count -ge 5) { "Part: $($parts[4])" }
                    if ($parts.Count -ge 6) { "PCI: $($parts[5])" }
                ) -join '; ') `
                $(if (-not (Norm-Serial $gpuSerial)) { 'nvidia-smi вернул N/A — у мобильных GPU S/N часто недоступен' } else { '' }) `
                -Warranty $true
        }
    }
    if (-not $any) {
        Add-UniqueRow 'Видеокарта' 'nvidia-smi' 'NVIDIA' '' $null 'Найден, но не вернул данные (старый драйвер?)' -Warranty $false
    }
} else {
    Write-Host "         nvidia-smi: not found"
    Add-UniqueRow 'Видеокарта' 'nvidia-smi' 'NVIDIA' '' $null `
        'nvidia-smi не найден — обновите NVIDIA драйвер, чтобы попытаться прочитать S/N GTX' `
        'Установите актуальный драйвер NVIDIA и запустите снова' -Warranty $false
}

# ========== 8. Monitors ==========
Write-Host "  [8/17] Monitors..."
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
    $prod = if ($mon.ProductCodeID) { (& $decode $mon.ProductCodeID).Trim() } else { '' }
    $isInternal = ($mfr -eq 'BOE' -or $name -eq '' -and $prod)
    Add-UniqueRow 'Монитор' "Монитор #$monIdx — $(if ($name) { $name } else { "$mfr $prod" })" `
        $mfr $(if ($name) { $name } else { $prod }) $sn `
        $(@(
            if ($prod) { "ProductCode: $prod" }
            if ($mon.YearOfManufacture) { "год: $($mon.YearOfManufacture)" }
            if ($mon.WeekOfManufacture) { "неделя: $($mon.WeekOfManufacture)" }
            if (-not (Norm-Serial $sn)) { 'у матриц ноутбуков S/N часто пустой' }
        ) -join '; ') -Warranty $true
}

# ========== 9. Network (physical only) ==========
Write-Host "  [9/17] Network..."
$Nics = @(Get-Safe {
    Get-CimInstance Win32_NetworkAdapter | Where-Object {
        $_.MACAddress -and
        ($_.Manufacturer -ne 'Microsoft') -and
        ($_.Name -notmatch 'Virtual|WAN Miniport|Wi-Fi Direct|Bluetooth') -and
        ($_.PNPDeviceID -notmatch 'ROOT\\|MSFT_|SWD\\|BTH\\|VPN|Virtual|Hyper-V|VMware|Loopback|WAN Miniport|Wi-Fi Direct|TAP-|Bluetooth')
    }
} @())
$nicIdx = 0
foreach ($n in $Nics) {
    $nicIdx++
    Add-UniqueRow 'Сеть' "NIC #$nicIdx — $($n.Name)" $n.Manufacturer $n.Name $n.MACAddress `
        $(@(
            "MAC: $($n.MACAddress)"
            if ($n.NetConnectionID) { "Имя: $($n.NetConnectionID)" }
        ) -join '; ') -Warranty $true
}

# ========== 10. Battery ==========
Write-Host "  [10/17] Battery..."
foreach ($b in @(Get-Safe { Get-CimInstance Win32_Battery } @())) {
    Add-UniqueRow 'Батарея' "Battery — $($b.Name)" $b.Manufacturer $b.Name $b.DeviceID `
        $(if ($b.EstimatedChargeRemaining -ne $null) { "Заряд: $($b.EstimatedChargeRemaining)%" } else { '' }) -Warranty $true
}
foreach ($b in @(Get-Safe { Get-CimInstance -Namespace root\wmi -ClassName BatteryStaticData -ErrorAction SilentlyContinue } @())) {
    $sn = if ($b.SerialNumber) { [string]$b.SerialNumber } else { $null }
    $mfr = ''; $dev = ''
    if ($b.ManufactureName) { $mfr = -join ($b.ManufactureName | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) }
    if ($b.DeviceName) { $dev = -join ($b.DeviceName | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) }
    if (Norm-Serial $sn) {
        Add-UniqueRow 'Батарея' "BatteryStatic — $dev" $mfr $dev $sn '' -Warranty $true
    }
}

# ========== 11. TPM ==========
Write-Host "  [11/17] TPM..."
$Tpm = Get-Safe { Get-CimInstance -Namespace root\cimv2\Security\MicrosoftTpm -ClassName Win32_Tpm -ErrorAction SilentlyContinue }
if ($Tpm) {
    Add-UniqueRow 'TPM' 'Trusted Platform Module' $Tpm.ManufacturerIdTxt "Spec $($Tpm.SpecVersion)" $null `
        "Manufacturer: $($Tpm.ManufacturerIdTxt); Ver: $($Tpm.ManufacturerVersion); Enabled: $($Tpm.IsEnabled_InitialValue)" -Warranty $false
}

# ========== 12. Optical / Sound (compact) ==========
Write-Host "  [12/17] Optical / audio..."
foreach ($c in @(Get-Safe { Get-CimInstance Win32_CDROMDrive } @())) {
    Add-UniqueRow 'Оптический привод' $c.Name $c.Manufacturer $c.Name $c.SerialNumber $c.MediaType -Warranty $true
}

# ========== 13. USB with real serials only ==========
Write-Host "  [13/17] USB serial devices..."
foreach ($u in @(Get-Safe { Get-CimInstance Win32_PnPEntity | Where-Object { $_.PNPDeviceID -like 'USB\VID_*' } } @())) {
    $pnpSn = Get-UsbSerialFromPnp $u.PNPDeviceID
    if (-not (Norm-Serial $pnpSn)) { continue }
    if ($pnpSn -match '^(5&|6&|7&)') { continue }
    if ($u.Name -match 'Root Hub|Generic Hub|Composite') { continue }
    Add-UniqueRow 'USB' $u.Name $u.Manufacturer $u.Name $pnpSn "PNP: $($u.PNPDeviceID)" -Warranty $false
}

# ========== 14-17 skipped heavy noise (cache/slots/virtual) ==========
Write-Host "  [14/17] Skipping virtual/noise devices..."
Write-Host "  [15/17] Building warranty sheet..."
Write-Host "  [16/17] Writing reports..."
Write-Host "  [17/17] Done collecting."

$realCount = @($Rows | Where-Object HasRealSerial).Count
$warrantyRows = @($Rows | Where-Object { $_.Warranty -eq $true })
$totalCount = $Rows.Count

Write-Host ""
Write-Host "  RAM sticks: $stickCount / $totalRamGb GB"
Write-Host "  Rows: $totalCount | with ID: $realCount | warranty-focus: $($warrantyRows.Count)"

# ---- Warranty TXT (clean) ----
$w = New-Object System.Text.StringBuilder
[void]$w.AppendLine('==============================================================================')
[void]$w.AppendLine("  $AppTitle — ЛИСТ ДЛЯ СВЕРКИ С ГАРАНТИЕЙ")
[void]$w.AppendLine("  $BrandName — $BrandUrl")
[void]$w.AppendLine('==============================================================================')
[void]$w.AppendLine("ПК: $HostName | $($Cs.Manufacturer) $($Cs.Model) | $($CollectedAt.ToString('yyyy-MM-dd HH:mm:ss'))")
[void]$w.AppendLine("ОЗУ: $stickCount планки / $totalRamGb ГБ")
[void]$w.AppendLine('')
[void]$w.AppendLine('Что РЕАЛЬНО удалось прочитать как S/N или надёжный ID:')
[void]$w.AppendLine('------------------------------------------------------------------------------')

foreach ($r in $warrantyRows) {
    [void]$w.AppendLine("[$($r.Category)] $($r.Device)")
    if ($r.Manufacturer) { [void]$w.AppendLine("  Производитель: $($r.Manufacturer)") }
    if ($r.Model) { [void]$w.AppendLine("  Модель/Part:  $($r.Model)") }
    [void]$w.AppendLine("  S/N или ID:   $($r.Serial)")
    if ($r.Extra) { [void]$w.AppendLine("  Доп.:         $($r.Extra)") }
    [void]$w.AppendLine('')
}

[void]$w.AppendLine('------------------------------------------------------------------------------')
[void]$w.AppendLine('ПОЧЕМУ чего-то нет (это нормально для Windows):')
[void]$w.AppendLine('• CPU Intel/AMD — заводской Unique ID отключён; есть только ProcessorId.')
[void]$w.AppendLine('• Часть планок ОЗУ (часто Patriot/OEM) — S/N пустой в SMBIOS; смотрите Part# и наклейку.')
[void]$w.AppendLine('• GPU (Intel + старые NVIDIA mobile) — WMI/nvidia-smi часто N/A; смотрите наклейку.')
[void]$w.AppendLine('• Матрица ноутбука (BOE и т.п.) — EDID serial часто пустой.')
[void]$w.AppendLine('• Съёмная флешка в отчёт дисков попадает, но для гарантии ПК не нужна.')
[void]$w.AppendLine('')
[void]$w.AppendLine("Сделано: $BrandName — $BrandUrl")

$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($WarrantyPath, $w.ToString(), $utf8Bom)

# ---- Full TXT ----
$txt = New-Object System.Text.StringBuilder
[void]$txt.AppendLine('==============================================================================')
[void]$txt.AppendLine("  $AppTitle — полный отчёт")
[void]$txt.AppendLine("  $BrandName — $BrandUrl")
[void]$txt.AppendLine('==============================================================================')
[void]$txt.AppendLine("ПК:            $HostName")
[void]$txt.AppendLine("Пользователь:  $env:USERNAME")
[void]$txt.AppendLine("ОС:            $($Os.Caption) $($Os.Version) ($($Os.OSArchitecture))")
[void]$txt.AppendLine("Производитель: $($Cs.Manufacturer)")
[void]$txt.AppendLine("Модель:        $($Cs.Model)")
[void]$txt.AppendLine("Дата сбора:    $($CollectedAt.ToString('yyyy-MM-dd HH:mm:ss'))")
[void]$txt.AppendLine("Админ-права:   $(if ($IsAdmin) { 'да' } else { 'нет' })")
[void]$txt.AppendLine("ОЗУ планок:    $stickCount  ($totalRamGb ГБ)")
[void]$txt.AppendLine("Позиций:       $totalCount  |  с ID/S/N: $realCount")
[void]$txt.AppendLine("SMBIOS raw:    $($Smbios.Count) записей")
[void]$txt.AppendLine('==============================================================================')
[void]$txt.AppendLine('')
[void]$txt.AppendLine('>>> Для гарантии откройте файл *-GARANTIYA.txt <<<')
[void]$txt.AppendLine('')

foreach ($g in ($Rows | Group-Object Category)) {
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
[void]$txt.AppendLine("Сделано: $BrandName — $BrandUrl")
[System.IO.File]::WriteAllText($TxtPath, $txt.ToString(), $utf8Bom)

$Rows | Select-Object Category, Device, Manufacturer, Model, Serial, Extra, Warranty |
    Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

# ---- HTML ----
function Html-Enc([string]$s) {
    if ($null -eq $s) { return '' }
    [System.Net.WebUtility]::HtmlEncode($s)
}
function Rows-Html($items) {
    foreach ($r in $items) {
        $snClass = if ($r.HasRealSerial) { 'ok' } else { 'miss' }
        "<tr><td>$(Html-Enc $r.Category)</td><td>$(Html-Enc $r.Device)</td><td>$(Html-Enc $r.Manufacturer)</td><td>$(Html-Enc $r.Model)</td><td class=`"$snClass`">$(Html-Enc $r.Serial)</td><td class=`"extra`">$(Html-Enc $r.Extra)</td></tr>"
    }
}

$html = @"
<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>$AppTitle — $HostName — $BrandName</title>
<style>
:root{--bg:#0f1419;--panel:#1a222c;--line:#2a3542;--text:#e8eef4;--muted:#8b9aab;--accent:#3dd6c6;--ok:#3dd68c;--miss:#c4a35a}
body{margin:0;font-family:"Segoe UI",Candara,sans-serif;background:radial-gradient(1000px 500px at 10% -10%,#1a3a45,transparent),var(--bg);color:var(--text)}
header{padding:2rem 1.5rem 1rem;border-bottom:1px solid var(--line)}
.brand{font-size:1.7rem;font-weight:700;margin:0}.brand span{color:var(--accent)}
.tag{color:var(--muted);margin:.3rem 0 1rem}
.meta{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:.7rem;max-width:1100px}
.meta div{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:.6rem .8rem}
.meta .k{display:block;color:var(--muted);font-size:.72rem;text-transform:uppercase}
.meta .v{font-weight:600}
main{padding:1.2rem 1.5rem 3rem;max-width:1280px;margin:0 auto}
.note{background:rgba(61,214,198,.08);border:1px solid rgba(61,214,198,.25);border-radius:10px;padding:.85rem 1rem;color:var(--muted);font-size:.9rem}
.wrap{overflow-x:auto;border:1px solid var(--line);border-radius:12px;background:var(--panel);margin:0.6rem 0 1.2rem}
table{width:100%;border-collapse:collapse;font-size:.88rem}
th,td{padding:.55rem .7rem;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
th{color:var(--muted);font-size:.72rem;text-transform:uppercase;background:rgba(0,0,0,.2)}
td.ok{color:var(--ok);font-family:Consolas,monospace;font-weight:600}
td.miss{color:var(--miss);font-style:italic}
td.extra{color:var(--muted);font-size:.8rem;max-width:320px}
footer{color:var(--muted);font-size:.85rem;border-top:1px solid var(--line);padding-top:1rem}
a{color:var(--accent)}
h2{font-size:1.05rem;margin:1.2rem 0 .5rem}
</style></head><body>
<header>
<p class="brand">$BrandName <span>— $BrandUrl</span></p>
<p class="tag">$AppTitle — сверка серийников с гарантией</p>
<div class="meta">
<div><span class="k">ПК</span><span class="v">$(Html-Enc $HostName)</span></div>
<div><span class="k">Модель</span><span class="v">$(Html-Enc ($Cs.Manufacturer + ' ' + $Cs.Model))</span></div>
<div><span class="k">ОЗУ</span><span class="v">$stickCount планки / $totalRamGb ГБ</span></div>
<div><span class="k">С ID/S/N</span><span class="v">$realCount</span></div>
<div><span class="k">Дата</span><span class="v">$(Html-Enc $CollectedAt.ToString('yyyy-MM-dd HH:mm:ss'))</span></div>
</div></header>
<main>
<div class="note">
<strong>Важно:</strong> Windows не может прочитать то, чего нет в SMBIOS/драйвере.
CPU и часть ОЗУ/GPU часто без заводского S/N — это ограничение железа, не программы.
Отдельный файл <strong>*-GARANTIYA.txt</strong> — короткий список для продавца.
</div>
<h2>Для гарантии</h2>
<div class="wrap"><table><thead><tr><th>Категория</th><th>Устройство</th><th>Производитель</th><th>Модель</th><th>S/N или ID</th><th>Доп.</th></tr></thead>
<tbody>$($(Rows-Html $warrantyRows) -join "`n")</tbody></table></div>
<h2>Полный список</h2>
<div class="wrap"><table><thead><tr><th>Категория</th><th>Устройство</th><th>Производитель</th><th>Модель</th><th>S/N или ID</th><th>Доп.</th></tr></thead>
<tbody>$($(Rows-Html $Rows) -join "`n")</tbody></table></div>
<footer>Сделано: <strong>$BrandName</strong> — <a href="https://$BrandUrl">$BrandUrl</a></footer>
</main></body></html>
"@
[System.IO.File]::WriteAllText($HtmlPath, $html, $utf8Bom)

Write-Host ""
Write-Host "  GARANTIYA: $WarrantyPath"
Write-Host "  TXT:       $TxtPath"
Write-Host "  CSV:       $CsvPath"
Write-Host "  HTML:      $HtmlPath"
Write-Host ""
Write-Host "  $BrandName - $BrandUrl"

try { Start-Process $HtmlPath } catch { }
try { Start-Process notepad.exe $WarrantyPath } catch { }

exit 0
