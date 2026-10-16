$desktop = [Environment]::GetFolderPath("Desktop")
$csvPath = Join-Path $desktop "hardware.csv"

function Convert-ToInventoryRows {
    param(
        [string]$Category,
        [object[]]$Items
    )

    for ($i = 0; $i -lt $Items.Count; $i++) {
        foreach ($p in $Items[$i].PSObject.Properties) {
            [PSCustomObject]@{
                Category = $Category
                Item     = $i + 1
                Property = $p.Name
                Value    = [string]$p.Value
            }
        }
    }
}

$ram = @(Get-CimInstance Win32_PhysicalMemory)

$ramSummary = [PSCustomObject]@{
    TotalGiB = [math]::Round(
        (($ram | Measure-Object -Property Capacity -Sum).Sum / 1GB), 2
    )
    ModuleCount = $ram.Count
    ConfiguredSpeedMTs = (
        ($ram | Select-Object -ExpandProperty ConfiguredClockSpeed -Unique) -join ", "
    )
}

$rows = @(
    Convert-ToInventoryRows "Motherboard" @(
        Get-CimInstance Win32_BaseBoard |
        Select-Object Manufacturer, Product, Version, SerialNumber
    )

    Convert-ToInventoryRows "BIOS" @(
        Get-CimInstance Win32_BIOS |
        Select-Object Manufacturer, SMBIOSBIOSVersion, ReleaseDate
    )

    Convert-ToInventoryRows "CPU" @(
        Get-CimInstance Win32_Processor |
        Select-Object Name, NumberOfCores,
            NumberOfLogicalProcessors, MaxClockSpeed
    )

    Convert-ToInventoryRows "RAM Summary" @($ramSummary)

    Convert-ToInventoryRows "RAM Module" @(
        $ram |
        Select-Object DeviceLocator, Manufacturer, PartNumber,
            @{N="CapacityGiB";E={$_.Capacity / 1GB}},
            Speed, ConfiguredClockSpeed
    )

    Convert-ToInventoryRows "GPU" @(
        Get-CimInstance Win32_VideoController |
        Select-Object Name, DriverVersion, VideoProcessor
    )

    Convert-ToInventoryRows "Storage" @(
        Get-CimInstance Win32_DiskDrive |
        Select-Object Model, SerialNumber,
            @{N="SizeGiB";E={[math]::Round($_.Size / 1GB, 2)}},
            InterfaceType
    )
)

# NVIDIA GPU
if (Get-Command nvidia-smi -ErrorAction SilentlyContinue) {
    $gpuInfo = & nvidia-smi `
        --query-gpu=name,memory.total,driver_version,power.limit `
        --format=csv,noheader,nounits 2>$null

    if ($LASTEXITCODE -eq 0 -and $gpuInfo) {
        $rows += Convert-ToInventoryRows "NVIDIA GPU" @(
            $gpuInfo | ConvertFrom-Csv -Header `
                Name, VRAM_MiB, DriverVersion, PowerLimit_W
        )
    }
}

$encoding = if ($PSVersionTable.PSVersion.Major -ge 7) {
    "utf8BOM"
} else {
    "UTF8"
}

# Treat model number or serial number as string
$excelRows = @(
    foreach ($row in $rows) {
        $copy = $row.PSObject.Copy()

        if ($copy.Property -in @(
            "Product", "SerialNumber", "PartNumber", "Model"
        )) {
            $copy.Value = '="' + $copy.Value.Replace('"', '""') + '"'
        }

        $copy
    }
)

$excelRows | Export-Csv $csvPath -NoTypeInformation -Encoding $encoding

Write-Host "Saved: $csvPath"

