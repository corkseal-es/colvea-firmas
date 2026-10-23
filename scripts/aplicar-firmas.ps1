<#
.SYNOPSIS
    Genera y aplica la firma Colvea en Outlook (Exchange Online) para cada persona de comerciales.csv.

.DESCRIPTION
    Sin parámetros: solo genera las firmas en la carpeta "salida" para revisarlas. No toca ningún buzón.
    Con -Aplicar: además las pone en cada buzón y sustituye la firma que tuvieran.

    Requiere ser administrador de Exchange (no hace falta darse permisos sobre cada buzón).

.EXAMPLE
    .\aplicar-firmas.ps1                  # revisar primero
    .\aplicar-firmas.ps1 -Aplicar         # aplicar a todos
    .\aplicar-firmas.ps1 -Aplicar -Solo arnau@corkseal.es
#>
param(
    [string]$Csv = "$PSScriptRoot\comerciales.csv",
    [switch]$Aplicar,
    [string[]]$Solo,
    [switch]$DesactivarFirmasEnLaNube
)

$ErrorActionPreference = 'Stop'

# Valores por defecto si la fila deja vacía la dirección o el mapa
$DireccionDefecto = 'C/ Isaac Peral, 3B local 2, 08224 Terrassa'
$MapaDefecto      = 'https://maps.app.goo.gl/BUeytBcy65aFRbYp6'

$plantilla = [System.IO.File]::ReadAllText("$PSScriptRoot\plantilla-firma.html", [System.Text.Encoding]::UTF8)
$filas = Get-Content -Path $Csv -Encoding UTF8 | ConvertFrom-Csv -Delimiter ';'
if ($Solo) { $filas = $filas | Where-Object { $Solo -contains $_.Email } }
if (-not $filas) { throw "No hay filas que procesar en $Csv" }

function Enc([string]$s) { [System.Net.WebUtility]::HtmlEncode($s.Trim()) }

function Nueva-Firma($f) {
    $direccion = if ($f.Direccion) { $f.Direccion } else { $DireccionDefecto }
    $mapa      = if ($f.Mapa)      { $f.Mapa }      else { $MapaDefecto }
    $telLink   = ($f.Telefono -replace '[^\d+]', '')
    $plantilla.
        Replace('{{NOMBRE}}',    (Enc $f.Nombre)).
        Replace('{{CARGO}}',     (Enc $f.Cargo)).
        Replace('{{EMAIL}}',     (Enc $f.Email)).
        Replace('{{TELEFONO}}',  (Enc $f.Telefono)).
        Replace('{{TEL_LINK}}',  (Enc $telLink)).
        Replace('{{DIRECCION}}', (Enc $direccion)).
        Replace('{{MAPA}}',      (Enc $mapa))
}

# 1) Generar las firmas para revisarlas
$salida = Join-Path $PSScriptRoot 'salida'
New-Item -ItemType Directory -Force -Path $salida | Out-Null
$firmas = @{}
foreach ($f in $filas) {
    $html = Nueva-Firma $f
    $firmas[$f.Email] = $html
    $archivo = Join-Path $salida ("firma_" + ($f.Email -replace '@.*$', '') + ".html")
    $doc = "<!DOCTYPE html>`n<html lang=`"es`"><head><meta charset=`"UTF-8`"></head><body style=`"margin:0`">`n$html`n</body></html>"
    [System.IO.File]::WriteAllText($archivo, $doc, (New-Object System.Text.UTF8Encoding $true))
    Write-Host "Generada: $archivo"
}

if (-not $Aplicar) {
    Write-Host "`nRevisa las firmas en '$salida'. Para aplicarlas: .\aplicar-firmas.ps1 -Aplicar" -ForegroundColor Yellow
    return
}

# 2) Conectar a Exchange Online si no hay conexión
if (-not (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) -or -not (Get-ConnectionInformation)) {
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -ShowBanner:$false
}

# 3) Firmas en la nube (roaming): si están activas, Outlook web ignora la firma puesta por PowerShell
$org = Get-OrganizationConfig
if (-not $org.PostponeRoamingSignaturesUntilLater) {
    if ($DesactivarFirmasEnLaNube) {
        Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true
        Write-Host "Firmas en la nube desactivadas para la organización." -ForegroundColor Cyan
    } else {
        Write-Warning ("Las firmas en la nube están activas: Outlook web puede seguir mostrando la firma antigua. " +
                       "Vuelve a ejecutar con -DesactivarFirmasEnLaNube.")
    }
}

# 4) Aplicar a cada buzón (sustituye la firma existente)
foreach ($email in $firmas.Keys) {
    try {
        Set-MailboxMessageConfiguration -Identity $email `
            -SignatureHtml $firmas[$email] `
            -AutoAddSignature $true `
            -AutoAddSignatureOnReply $true `
            -DefaultFormat Html
        Write-Host "OK     $email" -ForegroundColor Green
    } catch {
        Write-Host "ERROR  $email -> $($_.Exception.Message)" -ForegroundColor Red
    }
}
