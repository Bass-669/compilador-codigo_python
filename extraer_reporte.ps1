$url = "http://10.192.44.199/mainindex.asp"
$maxWaitHomePage = 90
$maxWaitContent = 120
$registroFechas = "$PSScriptRoot\registro_dias.txt"
$directorioReportesBase = Join-Path $PSScriptRoot "Reportes_Tornos"
$directorioDatos = Join-Path $directorioReportesBase "datos"
New-Item -ItemType Directory -Path $directorioDatos -Force | Out-Null

function Obtener-DiasHabilesAnteriores {
    param([int]$cantidad = 7)
    $dias = @()
    $fecha = (Get-Date).AddDays(-1)
    while ($dias.Count -lt $cantidad) {
        $dias += $fecha.Date
        $fecha = $fecha.AddDays(-1)
    }
    return $dias | Sort-Object
}

function Cargar-FechasProcesadas {
    if (Test-Path $registroFechas) {
        return Get-Content $registroFechas | ForEach-Object { [datetime]::ParseExact($_, "yyyy-MM-dd", $null) }
    }
    return @()
}

function Guardar-FechaProcesada {
    param([datetime]$fecha)
    Add-Content -Path $registroFechas -Value $fecha.ToString("yyyy-MM-dd")
}

function Wait-For-ReportContent {
    param($ie, $frameName, $timeout, $expectedPattern)
    $startTime = Get-Date
    $frameContent = $null

    while (((Get-Date) - $startTime).TotalSeconds -lt $timeout) {
        try { $frameContent = $ie.Document.frames.Item($frameName) } catch {}
        if ($null -eq $frameContent) {
            0..2 | ForEach-Object {
                try {
                    if ($null -eq $frameContent) {
                        $frameContent = $ie.Document.frames.Item($_)
                    }
                } catch {}
            }
        }

        if ($frameContent -ne $null) {
            try {
                $text = $frameContent.document.body.innerText
                if (-not [string]::IsNullOrWhiteSpace($text) -and $text -match $expectedPattern) {
                    Write-Host "Contenido del frame '$frameName' cargado."
                    return $frameContent
                }
            } catch {}
        }
        Start-Sleep -Seconds 3
    }
    throw "No se cargo el contenido del frame '$frameName' despues de $timeout segundos."
}
function Generate-TornoReport {
    param($frameContent, $tornoId, $tornoNombre, $fechaInicio, $fechaFin, $directorioReportes)

    try {
        Write-Host "Configurando valores para $tornoNombre..."

        $scriptJS = @"
document.getElementsByName('workId')[0].value = '$tornoId';
document.getElementsByName('startDay')[0].value = '$($fechaInicio.Day.ToString("00"))';
document.getElementsByName('startMon')[0].value = '$($fechaInicio.Month.ToString("00"))';
document.getElementsByName('startYear')[0].value = '$($fechaInicio.Year)';
document.getElementsByName('startHour')[0].value = '00';
document.getElementsByName('startMin')[0].value = '00';
document.getElementsByName('startSec')[0].value = '00';
document.getElementsByName('endDay')[0].value = '$($fechaFin.Day.ToString("00"))';
document.getElementsByName('endMon')[0].value = '$($fechaFin.Month.ToString("00"))';
document.getElementsByName('endYear')[0].value = '$($fechaFin.Year)';
document.getElementsByName('endHour')[0].value = '23';
document.getElementsByName('endMin')[0].value = '59';
document.getElementsByName('endSec')[0].value = '59';
"@

        $frameContent.document.parentWindow.execScript($scriptJS, 'JavaScript')
        Start-Sleep -Seconds 2

        $boton = $frameContent.document.getElementsByName('btnButton1') | Select-Object -First 1
        if ($null -eq $boton) {
            throw [System.Exception]::new("No se encontro el boton 'Reporte'")
        }

        Write-Host "Clic en boton 'Reporte'..."
        $boton.click()
        Start-Sleep -Seconds 2
        # Esperar ventana emergente
        $shell = New-Object -ComObject Shell.Application
        $popupWindow = $null
        $timeout = 30
        $start = Get-Date
        while (-not $popupWindow -and ((Get-Date) - $start).TotalSeconds -lt $timeout) {
            Start-Sleep -Seconds 1
            foreach ($w in $shell.Windows()) {
                if ($w.LocationURL -match "reports.asp" -and $w.Document -ne $null) {
                    $popupWindow = $w
                    break
                }
            }
        }

        if ($null -eq $popupWindow) {
            throw "No se detecto la ventana emergente del reporte"
        }

        # Obtener el HTML completo en lugar del texto plano
        $htmlContent = $popupWindow.Document.documentElement.outerHTML
        # Crear un HTML más limpio y autocontenido
        $htmlFormateado = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Reporte $tornoNombre - $($fechaInicio.ToString('yyyy-MM-dd'))</title>
    <style>
        body { font-family: Consolas, monospace; white-space: pre; }
        .reporte { margin: 20px; }
        table { border-collapse: collapse; }
        td, th { padding: 2px 5px; border: 1px solid #ddd; }
    </style>
</head>
<body>
    <div class="reporte">
        <h2>REPORTE: $tornoNombre - $($fechaInicio.ToString('yyyy-MM-dd'))</h2>
        $($popupWindow.Document.body.innerHTML)
    </div>
</body>
</html>
"@

        $ruta = "$directorioReportes\Reporte_$($fechaInicio.ToString('dd-MM-yyyy'))_$tornoId.html"
        $htmlFormateado | Out-File -FilePath $ruta -Encoding UTF8
        Write-Host "Reporte HTML guardado: $ruta"

        try { $popupWindow.Quit() } catch {}
        return $true

    } catch {
        Write-Host "ERROR: $($_.Exception.Message)"
        return $false
    }
}

$diasProcesados = Cargar-FechasProcesadas
$diasObjetivo = Obtener-DiasHabilesAnteriores
$diasFaltantes = $diasObjetivo | Where-Object { $_ -notin $diasProcesados }

if ($diasFaltantes.Count -eq 0) {
    Write-Host "Todos los ultimos 7 dias habiles ya estan procesados. Nada que hacer."
    exit
}

$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $false
$ie.Silent = $true
$ie.Navigate($url)
Write-Host "Esperando carga de la pagina principal..."
while ($ie.Busy -or $ie.ReadyState -ne 4) { Start-Sleep -Seconds 1 }
Start-Sleep -Seconds 5

foreach ($fechaObjetivo in $diasFaltantes) {
    Write-Host "`n=== PROCESANDO FECHA: $($fechaObjetivo.ToString('yyyy-MM-dd')) ==="

    # Cargar frame de menú
    $frameMenu = $null
    @("menu", 1, 2) | ForEach-Object {
        try {
            if ($frameMenu -eq $null) {
                $frameMenu = $ie.Document.frames.Item($_)
            }
        } catch {}
    }

    if ($frameMenu -eq $null) {
        Write-Host "ERROR: No se encontro el frame del menu. Saltando fecha."
        continue
    }

    try {
        $frameMenu.document.parentWindow.execScript("openWindow('dynamic.asp?INCLUDEFILE=report/choice.htm&OperationNo=10100','content','');", 'JavaScript')
        Start-Sleep -Seconds 3
        $frameContent = Wait-For-ReportContent -ie $ie -frameName "content" -timeout $maxWaitContent -expectedPattern "Estaci|Reporte"
    } catch {
        Write-Host "ERROR: No se pudo cargar el formulario de reporte para $($fechaObjetivo.ToString('yyyy-MM-dd'))"
        continue
    }

    $fechaInicio = $fechaObjetivo
    $fechaFin = $fechaInicio.AddHours(23).AddMinutes(59).AddSeconds(59)
    $directorioReportes = $directorioReportesBase

    $tornos = @(
        @{ID="3011"; Nombre="3011 - TORNO 8 L1"},
        @{ID="3012"; Nombre="3012 - TORNO 8 L2"}
    )

    $fallo = $false
    foreach ($torno in $tornos) {
        Write-Host "`nGenerando reporte para: $($torno.Nombre)"
        $ok = Generate-TornoReport -frameContent $frameContent -tornoId $torno.ID -tornoNombre $torno.Nombre -fechaInicio $fechaInicio -fechaFin $fechaFin -directorioReportes $directorioReportes
        if (-not $ok) {
            $fallo = $true
            break
        }
    }

    if (-not $fallo) {
        Guardar-FechaProcesada -fecha $fechaObjetivo
    } else {
        Write-Host "WARNING: ERROR en uno de los tornos. No se marca esta fecha como completada."
    }
}

$filtroExe = Join-Path $PSScriptRoot "filtro.exe"
$directorioReportes = Join-Path $PSScriptRoot "Reportes_Tornos"
$directorioDatos = Join-Path $directorioReportes "datos"
Write-Host "Ruta del filtro.exe: $filtroExe"
Write-Host "Directorio de reportes: $directorioReportes"
Write-Host "Directorio de datos: $directorioDatos"

if (-not (Test-Path $directorioDatos)) {
    Write-Host "Carpeta datos no existe. Creando: $directorioDatos"
    New-Item -ItemType Directory -Path $directorioDatos | Out-Null
} else {
    Write-Host "Carpeta datos ya existe."
}

# Obtener todos los archivos HTML en Reportes_Tornos (excluyendo datos)
$htmlFiles = Get-ChildItem -Path $directorioReportes -Filter *.html -File | Where-Object { $_.DirectoryName -ne $directorioDatos }
Write-Host "Archivos HTML encontrados para procesar: $($htmlFiles.Count)"

if ($htmlFiles.Count -eq 0) {
    Write-Host "No hay archivos HTML para procesar. Terminando."
    exit
}

Write-Host "Ejecutando filtro.exe sobre la carpeta completa de reportes..."
Write-Host "Comando: $filtroExe `"$directorioReportes`""

try {
    $proc = Start-Process -FilePath $filtroExe -ArgumentList "`"$directorioReportes`"" -Wait -NoNewWindow -PassThru
    if ($proc.ExitCode -eq 0) {
        Write-Host "Filtro.exe terminó correctamente."
    } else {
        Write-Host "Filtro.exe terminó con código de error: $($proc.ExitCode)" -ForegroundColor Red
    }
} catch {
    Write-Host "Error al ejecutar filtro.exe: $_" -ForegroundColor Red
}