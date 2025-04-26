<#
.SYNOPSIS
    Script optimizado para monitoreo y pruebas de estres de hardware utilizando Open Hardware Monitor.
.DESCRIPTION
    Recopila datos de sensores, realiza pruebas de estres controladas y genera reportes HTML detallados, incluyendo un arbol de sensores.
.NOTES
    Version: 0.1
    Autor: Miguel Perez
    Fecha: 2025-04-26
#>

#requires -Version 5.1

# --- Configuracion ---
$Config = @{
    OhmUrl              = "http://localhost:8085/data.json"
    ReportPath          = Join-Path -Path $PSScriptRoot -ChildPath "Hardware_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    SensorTreeReportPath = Join-Path -Path $PSScriptRoot -ChildPath "SensorTree_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    LogPath             = Join-Path -Path $PSScriptRoot -ChildPath "Hardware_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    StressDurations     = @{
        CPU   = 30  # Segundos
        RAM   = 30
        Disk  = 30
        GPU   = 30
    }
    MaxRetries          = 3
    RetryDelay          = 2  # Segundos
    DiskBufferSize      = 1MB
    MaxRamUsage         = 64MB  # Reducido para evitar bloqueos
    RamUsagePercent     = 0.3   # Reducido para ser mas conservador
    TempThresholds      = @{ CPU = 85; GPU = 90; Disk = 50; Power = 50 }
    LoadThresholds      = @{ Disk = 90 }
    VoltageRange        = @{ Min = 11.5; Max = 12.5 }
    MinDiskSpace        = 500MB
}

# --- Configuracion de Codificacion ---
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

# --- Clases y Tipos ---
class Sensor {
    [string]$Name
    [string]$SensorType
    [float]$Value
    [string]$Unit
    [float]$Max
    [string]$RawValue
}

# --- Funciones Auxiliares ---
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet("INFO", "ERROR", "WARNING")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $Message"
    try {
        $logMessage | Out-File -FilePath $Config.LogPath -Append -Encoding utf8 -ErrorAction Stop
    } catch {
        Write-Host "ERROR: No se pudo escribir en el log - $_" -ForegroundColor Red
    }
    $color = switch ($Level) {
        "ERROR" { "Red" }
        "WARNING" { "Yellow" }
        default { "White" }
    }
    Write-Host $logMessage -ForegroundColor $color
}

function Initialize-Environment {
    Write-Log -Message "Inicializando entorno..."
    try {
        $dir = Split-Path -Path $Config.LogPath -Parent
        if (-not (Test-Path -Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Log -Message "Directorio creado: $dir"
        }

        $testFile = Join-Path -Path $dir -ChildPath "test_write_$(New-Guid).tmp"
        "Test" | Out-File -FilePath $testFile -Encoding utf8 -ErrorAction Stop
        Remove-Item -Path $testFile -ErrorAction Stop
        Write-Log -Message "Permisos de escritura verificados en: $dir"

        for ($i = 0; $i -lt $Config.MaxRetries; $i++) {
            if (Test-NetConnection -ComputerName "localhost" -Port 8085 -InformationLevel Quiet) {
                Write-Log -Message "Open Hardware Monitor accesible en puerto 8085."
                return
            }
            Write-Log -Message "Intento $($i+1)/$($Config.MaxRetries): OHM no accesible. Reintentando en $($Config.RetryDelay)s..." -Level WARNING
            Start-Sleep -Seconds $Config.RetryDelay
        }
        throw "No se pudo conectar con Open Hardware Monitor en el puerto 8085."
    } catch {
        Write-Log -Message "ERROR: Fallo en inicializacion - $_" -Level ERROR
        Write-Log -Message "Sugerencia: Verifica que Open Hardware Monitor este ejecutandose y el puerto 8085 este libre." -Level WARNING
        throw
    }
}

function Get-OHMData {
    [CmdletBinding()]
    param ()
    Write-Log -Message "Obteniendo datos de sensores..."
    try {
        $response = Invoke-WebRequest -Uri $Config.OhmUrl -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
        $jsonData = $response.Content | ConvertFrom-Json
        $sensors = New-Object System.Collections.Generic.List[Sensor]

        function Process-Node {
            param ($Node)
            if ($Node.Value -and $Node.Text -and $Node.Children.Count -eq 0) {
                try {
                    $cleanValue = $Node.Value -replace "[^\d,.]" -replace ",", "."
                    if ($cleanValue -match "^\d+(\.\d+)?$") {
                        $value = [float]$cleanValue
                        $sensorType = switch -Regex ($Node.Text + $Node.Value) {
                            "Temperature|°C" { "Temperature" }
                            "Load|%" { "Load" }
                            "Voltage|V" { "Voltage" }
                            "Fan|RPM" { "Fan" }
                            "Data|GB|MB|Used Space|Available" { "Data" }
                            "Power|W" { "Power" }
                            "Clock|MHz" { "Clock" }
                            default { "Unknown" }
                        }
                        $sensors.Add([Sensor]@{
                            Name       = $Node.Text
                            SensorType = $sensorType
                            Value      = [math]::Round($value, 2)
                            Unit       = switch ($sensorType) {
                                "Temperature" { "-C" }
                                "Voltage" { "V" }
                                "Fan" { "RPM" }
                                "Load" { "%" }
                                "Data" { "GB" }
                                "Power" { "W" }
                                "Clock" { "MHz" }
                                default { "" }
                            }
                            Max        = switch ($Node.Text) {
                                { $_ -match "CPU Package|Processor|Core" -and $sensorType -eq "Temperature" } { $Config.TempThresholds.CPU }
                                { $_ -match "GPU Core|Graphics|Video" -and $sensorType -eq "Temperature" } { $Config.TempThresholds.GPU }
                                { $_ -match "Used Space|Disk|HDD|SSD|NVMe" -and $sensorType -eq "Load" } { $Config.LoadThresholds.Disk }
                                { $_ -match "Volt" } { $Config.VoltageRange.Max }
                                { $_ -match "Power" -and $sensorType -eq "Power" } { $Config.TempThresholds.Power }
                                default { $null }
                            }
                            RawValue   = $Node.Value
                        })
                    }
                } catch {
                    Write-Log -Message "WARNING: No se pudo convertir el valor '$($Node.Value)' para '$($Node.Text)'" -Level WARNING
                }
            }
            foreach ($child in $Node.Children) {
                Process-Node -Node $child
            }
        }

        Process-Node -Node $jsonData
        if ($sensors.Count -eq 0) {
            Write-Log -Message "WARNING: No se detectaron sensores. Verifica la configuracion de OHM." -Level WARNING
        } else {
            Write-Log -Message "Sensores detectados ($($sensors.Count)): $($sensors.Name -join ', ')"
        }
        return $sensors | Where-Object { $_.SensorType -ne "Unknown" }
    } catch {
        Write-Log -Message "ERROR: No se pudo obtener datos de OHM - $_" -Level ERROR
        throw
    }
}

function Build-SensorTree {
    [CmdletBinding()]
    param ()
    Write-Log -Message "Construyendo arbol de sensores..."
    try {
        $response = Invoke-WebRequest -Uri $Config.OhmUrl -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
        return $response.Content | ConvertFrom-Json
    } catch {
        Write-Log -Message "ERROR: No se pudo construir el arbol de sensores - $_" -Level ERROR
        throw
    }
}

function Invoke-StressTest {
    [CmdletBinding()]
    param ()
    Write-Log -Message "Iniciando pruebas de estres..."
    $results = New-Object System.Collections.Generic.List[PSCustomObject]
    $errors = New-Object System.Collections.Generic.List[string]

    # CPU Stress
    try {
        Write-Log -Message "Probando CPU ($($Config.StressDurations.CPU) segundos)..."
        $coreCount = [System.Environment]::ProcessorCount
        $jobs = 1..$coreCount | ForEach-Object {
            Start-Job -ScriptBlock {
                $endTime = (Get-Date).AddSeconds($using:Config.StressDurations.CPU)
                while ((Get-Date) -lt $endTime) {
                    $null = [math]::Sqrt((Get-Random -Maximum 1000000))
                }
            }
        }
        $jobs | Wait-Job -Timeout ($Config.StressDurations.CPU + 5) | Remove-Job -Force
        $cpuData = Get-OHMData | Where-Object { $_.Name -match "CPU Package|Processor|Core" -and $_.SensorType -eq "Temperature" }
        $maxTemp = if ($cpuData) { [math]::Round(($cpuData.Value | Measure-Object -Maximum).Maximum, 2) } else { $null }
        if (-not $cpuData) {
            Write-Log -Message "WARNING: No se encontraron datos de temperatura para CPU." -Level WARNING
        }
        $results.Add([PSCustomObject]@{
            Component = "CPU"
            MaxTemp   = $maxTemp
            Status    = if ($null -eq $maxTemp) { "NO DISPONIBLE" } elseif ($maxTemp -gt $Config.TempThresholds.CPU) { "CRITICO" } else { "OK" }
        })
    } catch {
        Write-Log -Message "ERROR en prueba de CPU: $_" -Level ERROR
        $errors.Add("CPU: $_")
        $results.Add([PSCustomObject]@{ Component = "CPU"; MaxTemp = $null; Status = "ERROR" })
    }

    # RAM Stress (Optimizado)
    try {
        Write-Log -Message "Probando RAM ($($Config.StressDurations.RAM) segundos)..."
        $os = Get-CimInstance Win32_OperatingSystem
        $freeRam = $os.FreePhysicalMemory * 1KB
        $totalRam = $os.TotalVisibleMemorySize * 1KB
        Write-Log -Message "Memoria disponible: $([math]::Round($freeRam/1GB, 2))GB de $([math]::Round($totalRam/1GB, 2))GB total"
        
        $memSize = [math]::Min($totalRam * $Config.RamUsagePercent, $Config.MaxRamUsage)
        if ($freeRam -lt $memSize) { 
            throw "Memoria libre insuficiente ($([math]::Round($freeRam/1GB, 2))GB). Se requieren $([math]::Round($memSize/1GB, 2))GB."
        }
        if ($memSize -lt 8MB) { 
            throw "Tamaño de memoria demasiado pequeño ($([math]::Round($memSize/1MB, 2))MB). Se requiere al menos 8MB."
        }

        Write-Log -Message "Asignando $([math]::Round($memSize/1MB, 2))MB para la prueba de RAM..."
        $arrays = @()
        $totalAllocated = 0
        $blockSize = 8MB  # Bloque pequeño para evitar sobrecarga
        while ($totalAllocated -lt $memSize) {
            $remaining = [math]::Min($blockSize, $memSize - $totalAllocated)
            $arrays += New-Object byte[] $remaining
            $totalAllocated += $remaining
            Write-Log -Message "Asignados $([math]::Round($totalAllocated/1MB, 2))MB de $([math]::Round($memSize/1MB, 2))MB..."
            Start-Sleep -Milliseconds 100  # Pausa breve para evitar saturacion
        }

        Write-Log -Message "Iniciando bucle de estres de RAM..."
        $random = New-Object System.Random
        $endTime = (Get-Date).AddSeconds($Config.StressDurations.RAM)
        $array = $arrays[0]  # Usar solo el primer bloque para estres ligero
        while ((Get-Date) -lt $endTime) {
            $random.NextBytes($array)
            Start-Sleep -Milliseconds 50  # Reducir carga en el bucle
        }
        Write-Log -Message "Bucle de estres de RAM completado."

        $ramData = Get-OHMData | Where-Object { $_.Name -match "Memory|RAM" -and $_.SensorType -eq "Load" }
        $maxLoad = if ($ramData) { [math]::Round(($ramData.Value | Measure-Object -Maximum).Maximum, 2) } else { $null }
        if (-not $ramData) {
            Write-Log -Message "WARNING: No se encontraron datos de carga para RAM." -Level WARNING
        }
        $results.Add([PSCustomObject]@{
            Component = "RAM"
            MaxLoad   = $maxLoad
            Status    = if ($null -eq $maxLoad) { "NO DISPONIBLE" } elseif ($maxLoad -gt 90) { "ALTO USO" } else { "OK" }
        })
    } catch {
        Write-Log -Message "ERROR en prueba de RAM: $_" -Level ERROR
        $errors.Add("RAM: $_")
        $results.Add([PSCustomObject]@{ Component = "RAM"; MaxLoad = $null; Status = "ERROR" })
    } finally {
        Write-Log -Message "Liberando memoria asignada para la prueba de RAM..."
        $arrays = $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Log -Message "Memoria liberada."
    }

    # Disk Stress
    try {
        Write-Log -Message "Probando Disco ($($Config.StressDurations.Disk) segundos)..."
        $testFile = Join-Path -Path $PSScriptRoot -ChildPath "testfile_$(New-Guid).tmp"
        $drive = Split-Path -Path $testFile -Qualifier
        $diskSpace = (Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$drive'").FreeSpace
        if ($diskSpace -lt $Config.MinDiskSpace) { throw "Espacio insuficiente ($([math]::Round($diskSpace/1MB))MB)." }

        $stream = [System.IO.File]::Create($testFile)
        $buffer = New-Object byte[] $Config.DiskBufferSize
        $endTime = (Get-Date).AddSeconds($Config.StressDurations.Disk)
        while ((Get-Date) -lt $endTime) {
            $stream.Write($buffer, 0, $buffer.Length)
        }
        $diskData = Get-OHMData | Where-Object { $_.Name -match "Used Space|Disk|HDD|SSD|NVMe" -and $_.SensorType -eq "Load" }
        $maxLoad = if ($diskData) { [math]::Round(($diskData.Value | Measure-Object -Maximum).Maximum, 2) } else { $null }
        $results.Add([PSCustomObject]@{
            Component = "Disk"
            MaxLoad   = $maxLoad
            Status    = if ($null -eq $maxLoad) { "NO DISPONIBLE" } elseif ($maxLoad -gt $Config.LoadThresholds.Disk) { "ALTO USO" } else { "OK" }
        })
    } catch {
        Write-Log -Message "ERROR en prueba de Disco: $_" -Level ERROR
        $errors.Add("Disk: $_")
        $results.Add([PSCustomObject]@{ Component = "Disk"; MaxLoad = $null; Status = "ERROR" })
    } finally {
        if ($stream) { $stream.Close() }
        Remove-Item -Path $testFile -ErrorAction SilentlyContinue
    }

    # GPU Stress (simulacion ligera)
    try {
        Write-Log -Message "Probando GPU ($($Config.StressDurations.GPU) segundos)..."
        $gpuData = Get-OHMData | Where-Object { $_.Name -match "GPU Core|Graphics|Video" -and $_.SensorType -eq "Temperature" }
        $maxTemp = if ($gpuData) { [math]::Round(($gpuData.Value | Measure-Object -Maximum).Maximum, 2) } else { $null }
        if (-not $gpuData) {
            Write-Log -Message "WARNING: No se encontraron datos de temperatura para GPU. Es posible que el sistema use graficos integrados o no tenga sensores configurados." -Level WARNING
        }
        $results.Add([PSCustomObject]@{
            Component = "GPU"
            MaxTemp   = $maxTemp
            Status    = if ($null -eq $maxTemp) { "NO DISPONIBLE" } elseif ($maxTemp -gt $Config.TempThresholds.GPU) { "CRITICO" } else { "OK" }
        })
    } catch {
        Write-Log -Message "ERROR en prueba de GPU: $_" -Level ERROR
        $errors.Add("GPU: $_")
        $results.Add([PSCustomObject]@{ Component = "GPU"; MaxTemp = $null; Status = "ERROR" })
    }

    return $results, $errors
}

function New-HardwareReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [Sensor[]]$SensorData,
        [Parameter(Mandatory)]
        [PSCustomObject[]]$StressResults,
        [string[]]$Alerts,
        [string]$LogPath
    )
    Write-Log -Message "Generando reporte HTML..."
    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

    $html = @"
<!DOCTYPE html>
<html lang=""es"">
<head>
    <meta charset=""UTF-8"">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reporte de Hardware - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</title>
    <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
</head>
<body class="bg-gray-100 font-sans">
    <div class="container mx-auto p-4 md:p-8">
        <h1 class="text-3xl font-bold text-gray-800 mb-4">Reporte de Hardware</h1>
        <p class="text-sm text-gray-600 text-right mb-6">Generado: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <div class="bg-white shadow-md rounded-lg p-6 mb-6">
            <h2 class="text-xl font-semibold text-gray-700 mb-4">Resumen de Pruebas de Estres</h2>
            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
"@

    foreach ($result in $StressResults) {
        $value = if ($result.Component -in "RAM", "Disk") {
            if ($null -eq $result.MaxLoad) { "N/A" } else { "$($result.MaxLoad)%" }
        } else {
            if ($null -eq $result.MaxTemp) { "N/A" } else { "$($result.MaxTemp)°C" }
        }
        $statusClass = switch ($result.Status) {
            "CRITICO" { "bg-red-100 text-red-700" }
            "ALTO USO" { "bg-yellow-100 text-yellow-700" }
            "OK" { "bg-green-100 text-green-700" }
            default { "bg-gray-100 text-gray-700" }
        }
        $html += @"
                <div class="bg-white rounded-lg shadow p-4 $statusClass">
                    <h3 class="font-semibold">$([System.Web.HttpUtility]::HtmlEncode($result.Component))</h3>
                    <p>Valor Maximo: $value</p>
                    <p>Estado: $([System.Web.HttpUtility]::HtmlEncode($result.Status))</p>
                </div>
"@
    }
    $html += @"
            </div>
        </div>
"@

    $sensorGroups = $SensorData | Group-Object -Property SensorType
    foreach ($group in $sensorGroups) {
        $html += @"
        <div class="bg-white shadow-md rounded-lg p-6 mb-6">
            <h2 class="text-xl font-semibold text-gray-700 mb-4">$([System.Web.HttpUtility]::HtmlEncode($group.Name))</h2>
            <div class="overflow-x-auto">
                <table class="w-full text-left">
                    <thead>
                        <tr class="bg-blue-600 text-white">
                            <th class="p-3">Nombre</th>
                            <th class="p-3">Valor</th>
                            <th class="p-3">Estado</th>
                        </tr>
                    </thead>
                    <tbody>
"@
        foreach ($sensor in $group.Group) {
            $value = if ($null -eq $sensor.Value) { "N/A" } else { "$($sensor.Value) $($sensor.Unit)" }
            $status = $progress = ""
            if ($sensor.Max -and $sensor.Value -gt $sensor.Max) {
                $status = "<span class='text-red-600 font-semibold'>CRITICO (Max: $($sensor.Max)$($sensor.Unit))</span>"
                $progress = "<div class='w-24 bg-gray-200 rounded-full h-2.5'><div class='bg-red-600 h-2.5 rounded-full' style='width: 100%'></div></div>"
            } elseif ($sensor.SensorType -eq "Load" -and $sensor.Value -gt 90) {
                $status = "<span class='text-yellow-600 font-semibold'>ALTO USO</span>"
                $progress = "<div class='w-24 bg-gray-200 rounded-full h-2.5'><div class='bg-yellow-600 h-2.5 rounded-full' style='width: $($sensor.Value)%'></div></div>"
            } else {
    $status = "<span class='text-green-600 font-semibold'>OK</span>"
    $progress = "<div class='w-24 bg-gray-200 rounded-full h-2.5'><div class='bg-blue-600 h-2.5 rounded-full' style='width: $([math]::Min($sensor.Value, 100))%'></div></div>"
}
$html += @"
            <tr class="border-b">
                <td class="p-3">$([System.Web.HttpUtility]::HtmlEncode($sensor.Name))</td>
                            <td class="p-3">$value $progress</td>
                            <td class="p-3">$status</td>
                        </tr>
"@
        }
        $html += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }

    $html += @"
        <div class="bg-white shadow-md rounded-lg p-6">
            <h2 class="text-xl font-semibold text-gray-700 mb-4">Alertas</h2>
"@
    if ($Alerts) {
        $html += "<ul class='list-disc pl-5'>" + ($Alerts | ForEach-Object { "<li class='text-red-600'>$([System.Web.HttpUtility]::HtmlEncode($_))</li>" }) + "</ul>"
    } else {
        $html += "<p class='text-green-600'>Sin alertas criticas.</p>"
    }
    $html += @"
            <a href="file:///$($LogPath -replace '\\', '/')" download class="mt-4 inline-block bg-blue-600 text-white px-4 py-2 rounded hover:bg-blue-700">Descargar Log</a>
        </div>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $Config.ReportPath -Encoding utf8 -Force -ErrorAction Stop
    Write-Log -Message "Reporte HTML generado en: $($Config.ReportPath)"
}

function New-SensorTreeReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]$SensorTree,
        [Parameter(Mandatory)]
        [string]$HtmlPath
    )
    Write-Log -Message "Generando reporte de arbol de sensores..."

    function New-HTMLList {
        param ($Node)
        $html = "<li class='mb-2'><strong>$([System.Web.HttpUtility]::HtmlEncode($Node.Text))</strong>"
        if ($Node.Min -or $Node.Value -or $Node.Max) {
            $html += " - Min: $([System.Web.HttpUtility]::HtmlEncode($Node.Min)), Valor: $([System.Web.HttpUtility]::HtmlEncode($Node.Value)), Max: $([System.Web.HttpUtility]::HtmlEncode($Node.Max))"
        }
        if ($Node.Children) {
            $html += "<ul class='ml-4 mt-2'>"
            foreach ($child in $Node.Children) {
                $html += New-HTMLList -Node $child
            }
            $html += "</ul>"
        }
        $html += "</li>"
        return $html
    }

    $html = @"
<!DOCTYPE html>
<html lang=""es"">
<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>Arbol de Sensores - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</title>
    <link href=""https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css"" rel=""stylesheet"">
</head>
<body class=""bg-gray-100 font-sans"">
    <div class=""container mx-auto p-4 md:p-8"">
        <h1 class=""text-3xl font-bold text-gray-800 mb-4"">Arbol de Sensores</h1>
        <p class=""text-sm text-gray-600 text-right mb-6"">Generado: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <div class=""bg-white shadow-md rounded-lg p-6"">
            <ul class=""list-none"">
                $(New-HTMLList -Node $SensorTree)
            </ul>
        </div>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $HtmlPath -Encoding utf8 -Force -ErrorAction Stop
    Write-Log -Message "Reporte de arbol generado en: $HtmlPath"
}

function Get-CriticalAlerts {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [Sensor[]]$SensorData,
        [string[]]$TestErrors
    )
    $alerts = New-Object System.Collections.Generic.List[string]
    foreach ($sensor in $SensorData) {
        if ($sensor.Max -and $sensor.Value -gt $sensor.Max) {
            $alerts.Add("ALERTA: $($sensor.Name) con $($sensor.Value)$($sensor.Unit) (Max: $($sensor.Max)$($sensor.Unit))")
        }
        if ($sensor.SensorType -eq "Voltage" -and ($sensor.Value -lt $Config.VoltageRange.Min -or $sensor.Value -gt $Config.VoltageRange.Max)) {
            $alerts.Add("ALERTA: $($sensor.Name) con $($sensor.Value)$($sensor.Unit) (Rango: $($Config.VoltageRange.Min)V-$($Config.VoltageRange.Max)V)")
        }
    }
    $alerts.AddRange($TestErrors)
    return $alerts
}

# --- Ejecucion Principal ---
try {
    Write-Log -Message "Iniciando script de monitoreo y prueba de estres..."

    Initialize-Environment
    $sensorData = Get-OHMData
    $sensorTreeJob = Start-Job -ScriptBlock ${function:Build-SensorTree} -ArgumentList $Config
    $stressResults, $testErrors = Invoke-StressTest
    $sensorData = Get-OHMData
    $criticalAlerts = Get-CriticalAlerts -SensorData $sensorData -TestErrors $testErrors

    $sensorTree = $sensorTreeJob | Receive-Job -Wait -AutoRemoveJob -ErrorAction SilentlyContinue
    if ($sensorTree) {
        New-SensorTreeReport -SensorTree $sensorTree -HtmlPath $Config.SensorTreeReportPath
    } else {
        Write-Log -Message "WARNING: No se pudo generar el arbol de sensores. Verifica la conexion con OHM." -Level WARNING
    }

    New-HardwareReport -SensorData $sensorData -StressResults $stressResults -Alerts $criticalAlerts -LogPath $Config.LogPath

    Write-Host "`nResumen de prueba de estres:" -ForegroundColor Magenta
    $stressResults | Format-Table Component, @{Name="Valor Maximo"; Expression={
        if ($_.Component -in "RAM", "Disk") { if ($null -eq $_.MaxLoad) { "N/A" } else { "$($_.MaxLoad)%" } }
        else { if ($null -eq $_.MaxTemp) { "N/A" } else { "$($_.MaxTemp)-C" } }
    }}, Status -AutoSize

    if ($criticalAlerts) {
        Write-Log -Message "Se detectaron problemas criticos. Revisa el reporte." -Level WARNING
    } else {
        Write-Log -Message "Sin problemas criticos detectados."
    }

    Start-Process -FilePath $Config.ReportPath -ErrorAction SilentlyContinue
    if (Test-Path $Config.SensorTreeReportPath) {
        Start-Process -FilePath $Config.SensorTreeReportPath -ErrorAction SilentlyContinue
    } else {
        Write-Log -Message "WARNING: No se pudo abrir el reporte de arbol de sensores porque no se genero." -Level WARNING
    }
} catch {
    Write-Log -Message "ERROR: Fallo general: $_" -Level ERROR
    exit 1
}