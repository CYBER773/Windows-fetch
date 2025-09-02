<#
 WinSysFetch.ps1 - Fastfetch‑style Windows system summary
 --------------------------------------------------------
 Gathers and displays key hardware / OS facts similar to fastfetch / neofetch:
   • OS + version + uptime
   • CPU (model, cores, threads, base/max clocks, generation + suffix parse)
   • GPU(s) (name, VRAM, driver)
   • Memory (total, per‑DIMM details: size, type DDR#, speed, manufacturer)
   • Storage devices (model, size, SSD/HDD/NVMe, bus)
   • Volumes usage (C:, others)
   • Network IPv4 addresses

 Output: Colorized aligned text block OR JSON (-Json) for scripting.

 Usage examples:
   PS> .\WinSysFetch.ps1
   PS> .\WinSysFetch.ps1 -NoColor
   PS> .\WinSysFetch.ps1 -Json | ConvertTo-Json -Depth 4

 Requires: PowerShell 5+ (works best in PS7). Some info (Get-PhysicalDisk) needs admin + newer Storage module.
#>

[CmdletBinding()]
param(
    [switch]$NoColor,
    [switch]$Json
)

# --------------------------- Helpers ---------------------------
function Write-Color {
    param(
        [string]$Text,
        [ConsoleColor]$Color = [ConsoleColor]::Gray,
        [switch]$NoNewline
    )
    if ($script:NoColor) {
        if ($NoNewline) { Write-Host -NoNewline $Text } else { Write-Host $Text }
    } else {
        if ($NoNewline) { Write-Host -ForegroundColor $Color -NoNewline $Text } else { Write-Host -ForegroundColor $Color $Text }
    }
}

# Human‑readable size
function Format-Bytes {
    param([UInt64]$Bytes)
    if ($Bytes -lt 1KB) {return "$Bytes B"}
    elseif ($Bytes -lt 1MB) {return "{0:N2} KB" -f ($Bytes/1KB)}
    elseif ($Bytes -lt 1GB) {return "{0:N2} MB" -f ($Bytes/1MB)}
    elseif ($Bytes -lt 1TB) {return "{0:N2} GB" -f ($Bytes/1GB)}
    else {return "{0:N2} TB" -f ($Bytes/1TB)}
}

# Map SMBIOS memory type to DDR label (best‑effort)
$DDRTypeMap = @{ 20='DDR';21='DDR2';22='DDR2 FB-DIMM';24='DDR3';26='DDR4';27='LPDDR';29='LPDDR2';30='LPDDR3';31='LPDDR4';32='Logical';33='HBM';34='HBM2';35='DDR5';36='LPDDR5' }
function Get-DDRLabel {
    param($MemObj)
    $t = $null
    if ($MemObj.SMBIOSMemoryType -and $MemObj.SMBIOSMemoryType -ne 0) { $t = $MemObj.SMBIOSMemoryType }
    elseif ($MemObj.MemoryType -and $MemObj.MemoryType -ne 0) { $t = $MemObj.MemoryType }
    if ($t -and $DDRTypeMap.ContainsKey([int]$t)) { return $DDRTypeMap[[int]$t] }
    # heuristic from Speed strings
    if ($MemObj.Speed -ge 6400) {return 'DDR5?'}
    elseif ($MemObj.Speed -ge 3200) {return 'DDR4?'}
    elseif ($MemObj.Speed -ge 1600) {return 'DDR3?'}
    return 'Unknown'
}

# Parse Intel generation + suffix from CPU name string
function Parse-IntelCPUName {
    param([string]$Name)
    $out = [ordered]@{Generation=$null; Suffix=$null}
    if ($Name -match '(Core\s+i[3579]|Core\s+Ultra|Core\s*2)') {
        # capture model block like i7-12700H, i5-8250U, i9-13900K, etc.
        if ($Name -match '([iI][3579]|Ultra)\s*-?\s*([0-9]{3,5})([A-Za-z]{0,3})') {
            $digits = $Matches[2]; $suf = $Matches[3]
            # 5 digits => 10th gen+ (first 2 digits); 4 digits => first digit
            if ($digits.Length -ge 5) {$gen = [int]$digits.Substring(0,2)} else {$gen = [int]$digits.Substring(0,1)}
            $out.Generation = $gen
            if ($suf) { $out.Suffix = $suf.ToUpper() }
        }
    }
    return $out
}

# Parse AMD Ryzen series & guess architecture gen
$AMDZenMap = @{ '1'='Zen/Zen+'; '2'='Zen 2'; '3'='Zen 2/3 Mobile'; '4'='Zen 2 (Mobile)'; '5'='Zen 3'; '6'='Zen 3+/4 Mobile'; '7'='Zen 4'; '8'='Zen 5 (Est)'}
function Parse-AMDCPUName {
    param([string]$Name)
    $out = [ordered]@{Series=$null; ArchGuess=$null}
    if ($Name -match 'Ryzen\s+\w*\s*([1-9][0-9]{3})') {
        $series = $Matches[1]
        $out.Series = $series
        $lead = $series.Substring(0,1)
        if ($AMDZenMap.ContainsKey($lead)) { $out.ArchGuess = $AMDZenMap[$lead] }
    }
    return $out
}

# ASCII Windows logo (11‑ish) for left column art
$WinArt = @(
    "                    ",
    "  ____ _                 _            ",
    " / | | ___  _   _  __| | ___  ___  ",
    "| |  | |/ _ \\| | | |/ ' |/ _ \\ / _ ",
    "| || | | () | || | (| |  __/  __/",
    " \\||\\/ \\,_|\\,|\\|\\|",
    "                    "
)

# gather OS ------------------------------------------------------
$os = Get-CimInstance Win32_OperatingSystem
$cs = Get-CimInstance Win32_ComputerSystem
$boot = [Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
$uptimeSpan = (Get-Date) - $boot
$uptimeFmt = if ($uptimeSpan.TotalDays -ge 1) {"{0}d {1}h" -f [int]$uptimeSpan.TotalDays, $uptimeSpan.Hours} elseif ($uptimeSpan.TotalHours -ge 1) {"{0}h {1}m" -f [int]$uptimeSpan.TotalHours, $uptimeSpan.Minutes} else {"{0}m" -f [int]$uptimeSpan.TotalMinutes}

# CPU -------------------------------------------------------------
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
$cpuName = $cpu.Name.Trim()
$cpuVendor = $cpu.Manufacturer
$cpuCores = $cpu.NumberOfCores
$cpuThreads = $cpu.NumberOfLogicalProcessors
$cpuBaseMHz = $cpu.MaxClockSpeed   # reported in MHz
$cpuBaseGHz = [math]::Round($cpuBaseMHz/1000,2)
$cpuExtra = $null
if ($cpuVendor -match 'Intel') {
    $parsed = Parse-IntelCPUName $cpuName
    $cpuExtra = "Gen {0}{1}" -f $parsed.Generation, (if($parsed.Suffix){" ($($parsed.Suffix))"}else{''})
}
elseif ($cpuVendor -match 'AMD') {
    $parsed = Parse-AMDCPUName $cpuName
    if ($parsed.Series) {
        $cpuExtra = "Series $($parsed.Series) ($($parsed.ArchGuess))"
    }
}

# GPU(s) ----------------------------------------------------------
$gpus = Get-CimInstance Win32_VideoController | Sort-Object Name
$gpuInfo = foreach ($g in $gpus) {
    [pscustomobject]@{
        Name=$g.Name.Trim()
        VRAM_GB = [math]::Round(($g.AdapterRAM/1GB),2)
        DriverVersion=$g.DriverVersion
        DriverDate=([Management.ManagementDateTimeConverter]::ToDateTime($g.DriverDate)).ToString('yyyy-MM-dd')
    }
}

# Memory ----------------------------------------------------------
$memMods = Get-CimInstance Win32_PhysicalMemory
$memTotalBytes = ($memMods | Measure-Object -Property Capacity -Sum).Sum
$memTotalGB = [math]::Round($memTotalBytes/1GB,2)
$memModObjs = foreach($m in $memMods){
    [pscustomobject]@{
        Slot=$m.DeviceLocator
        SizeGB=[math]::Round($m.Capacity/1GB,2)
        Type=(Get-DDRLabel $m)
        SpeedMHz=$m.Speed
        Manufacturer=$m.Manufacturer
        PartNumber=$m.PartNumber.Trim()
    }
}

# Storage ---------------------------------------------------------
# Try modern Storage cmdlets first
$physicalDisks = @()
try {
    $physicalDisks = Get-PhysicalDisk -ErrorAction Stop | ForEach-Object {
        [pscustomobject]@{
            FriendlyName=$_.FriendlyName
            SizeGB=[math]::Round($_.Size/1GB,2)
            MediaType=if ($_.MediaType -ne 'Unspecified') {$_.MediaType} else {if($_.BusType -eq 'NVMe'){ 'SSD' } else {'Unknown'}}
            BusType=$_.BusType
            Serial=$_.SerialNumber
        }
    }
} catch {
    # fall back to Win32_DiskDrive
    $wmiDisks = Get-CimInstance Win32_DiskDrive
    $physicalDisks = $wmiDisks | ForEach-Object {
        $rot = $_.MediaType
        $isSSD = if ($rot -match 'Solid State' -or $_.Model -match 'SSD' -or $_.Model -match 'NVMe') { 'SSD' } else { 'HDD?' }
        [pscustomobject]@{
            FriendlyName=$_.Model
            SizeGB=[math]::Round($_.Size/1GB,2)
            MediaType=$isSSD
            BusType=$_.InterfaceType
            Serial=$_.SerialNumber
        }
    }
}

# Volumes usage (all fixed)
$volumes = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
    [pscustomobject]@{
        Drive=$_.DeviceID
        UsedGB=[math]::Round(($_.Size - $_.FreeSpace)/1GB,2)
        TotalGB=[math]::Round($_.Size/1GB,2)
        FS=$_.FileSystem
    }
}

# Network IPv4 addresses (up, not loopback)
$netIPs = Get-NetIPAddress -AddressFamily IPv4 -PrefixOrigin Manual, Dhcp -ErrorAction SilentlyContinue | Where-Object {$_.IPAddress -notlike '169.254*' -and $_.InterfaceAlias -notmatch 'Loopback'} | ForEach-Object {
    "{0} ({1})" -f $_.IPAddress,$_.InterfaceAlias
}

# Assemble object -------------------------------------------------
$result = [pscustomobject]@{
    ComputerName=$env:COMPUTERNAME
    User=$env:USERNAME
    OS=$os.Caption.Trim()
    OSVersion=$os.Version
    OSArch=$os.OSArchitecture
    Uptime=$uptimeFmt
    LastBoot=$boot
    CPU=[pscustomobject]@{
        Name=$cpuName
        Vendor=$cpuVendor
        Cores=$cpuCores
        Threads=$cpuThreads
        BaseGHz=$cpuBaseGHz
        ExtraInfo=$cpuExtra
    }
    GPU=$gpuInfo
    Memory=[pscustomobject]@{
        TotalGB=$memTotalGB
        Modules=$memModObjs
    }
    Storage=$physicalDisks
    Volumes=$volumes
    NetworkIPs=$netIPs
}

if ($Json) {
    $result # let caller pipe to ConvertTo-Json
    return
}

# --------------------------- Pretty Print ---------------------------
$script:NoColor = $NoColor.IsPresent

# Build right‑side lines
$lines = @()
$lines += "${env:USERNAME}@${env:COMPUTERNAME}"
$lines += ''.PadRight(40,'-')
$lines += "OS:        $($result.OS) $($result.OSArch)"
$lines += "Version:   $($result.OSVersion)"
$lines += "Uptime:    $($result.Uptime)"
$lines += "Boot:      $($result.LastBoot.ToString('yyyy-MM-dd HH:mm'))"
$lines += "CPU:       $cpuName"
$lines += "           Cores:$cpuCores  Threads:$cpuThreads  Base:$cpuBaseGHz GHz"
if ($cpuExtra) { $lines += "           $cpuExtra" }
foreach($gi in $gpuInfo){ $lines += "GPU:       $($gi.Name) [$($gi.VRAM_GB) GB] drv $($gi.DriverVersion)" }
$lines += "Memory:    $memTotalGB GB installed in $($memModObjs.Count) slot(s)"
foreach($mm in $memModObjs){ $lines += "           $($mm.Slot): $($mm.SizeGB)GB $($mm.Type) @$($mm.SpeedMHz)MHz" }
$lines += "Storage:   $($physicalDisks.Count) disk(s)"
foreach($pd in $physicalDisks){ $lines += "           $($pd.FriendlyName) $($pd.SizeGB)GB $($pd.MediaType) ($($pd.BusType))" }
$lines += "Volumes:"; foreach($v in $volumes){ $lines += "           $($v.Drive) $($v.UsedGB)/$($v.TotalGB)GB $($v.FS)" }
if ($netIPs.Count -gt 0){ $lines += "IP(s):     $($netIPs -join ', ')" }

# Determine max lines to align with art
$maxLines = [math]::Max($WinArt.Count, $lines.Count)
$artPadded = @(); $infoPadded=@()
for($i=0;$i -lt $maxLines;$i++){
    $artPadded += ($WinArt[$i]  ) 2>$null
    $infoPadded+= ($lines[$i]   ) 2>$null
}

# Print side‑by‑side
$colPad = ($WinArt | Measure-Object -Property Length -Maximum).Maximum + 2
for($i=0;$i -lt $maxLines;$i++){
    $a = if($i -lt $WinArt.Count){$WinArt[$i]}else{''}
    $b = if($i -lt $lines.Count){$lines[$i]}else{''}
    $pad = $colPad - $a.Length; if($pad -lt 1){$pad=1}
    if ($NoColor) {
        Write-Host ($a + (' ' * $pad) + $b)
    } else {
        # color art cyan, labels greenish
        Write-Host -NoNewline -ForegroundColor Cyan $a
        Write-Host -NoNewline (' ' * $pad)
        Write-Host $b
    }
}

# simple color test bar
if (-not $NoColor) {
    Write-Color "\nColor Test: " Cyan -NoNewline
    foreach($c in [enum]::GetValues([ConsoleColor])){ Write-Host -NoNewline -ForegroundColor $c '■' }
    Write-Host
}
