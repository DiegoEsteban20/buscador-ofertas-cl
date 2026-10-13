<#
.SYNOPSIS
    Buscador de Ofertas v2: revisa las tiendas de Fuentes.csv, arma un Excel con los
    descuentos y marca lo que BAJÓ de precio desde la última vez.

.EJEMPLOS
    .\BuscadorOfertas.ps1               # Busca ahora y abre el Excel
    .\BuscadorOfertas.ps1 -NoAbrir      # Busca sin abrir el Excel
    .\BuscadorOfertas.ps1 -Instalar     # Se ejecuta solo cada vez que inicias sesión en Windows
    .\BuscadorOfertas.ps1 -Desinstalar  # Quita la tarea programada
#>
param(
    [switch]$Instalar,
    [switch]$Desinstalar,
    [switch]$NoAbrir,
    [switch]$Forzar,                 # ignora la pausa mínima entre ejecuciones
    [int]$HorasEntreEjecuciones = 2, # evita repetir si reinicias el PC seguido
    [int]$TopPorCategoria = 5,
    [string]$Solo = ''               # prueba solo las fuentes cuya Categoría o Tienda contenga este texto
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Carpeta       = $PSScriptRoot
$ArchFuentes   = Join-Path $Carpeta 'Fuentes.csv'
$ArchHistorial = Join-Path $Carpeta 'Historial.csv'
$ArchLog       = Join-Path $Carpeta 'BuscadorOfertas.log'
$ArchUltima    = Join-Path $Carpeta 'ultima_ejecucion.txt'
$Escritorio    = [Environment]::GetFolderPath('Desktop')     # evita el problema de OneDrive
if ([string]::IsNullOrWhiteSpace($Escritorio)) { $Escritorio = $Carpeta }
$CarpetaExcel  = Join-Path $Escritorio 'Ofertas'
$NombreTarea   = 'Buscador de Ofertas'
$Inv           = [Globalization.CultureInfo]::InvariantCulture

function Log([string]$m) {
    $l = "{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m
    Add-Content -Path $ArchLog -Value $l -Encoding UTF8
    Write-Host $l
}

# ================================================================ Tarea programada
if ($Desinstalar) {
    Unregister-ScheduledTask -TaskName $NombreTarea -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Tarea '$NombreTarea' eliminada." -ForegroundColor Green
    return
}
if ($Instalar) {
    $usuario = "$env:USERDOMAIN\$env:USERNAME"
    $accion = New-ScheduledTaskAction -Execute 'powershell.exe' `
        -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`""
    $disparador = New-ScheduledTaskTrigger -AtLogOn -User $usuario
    $disparador.Delay = 'PT2M'   # espera 2 minutos para que haya internet
    $principal = New-ScheduledTaskPrincipal -UserId $usuario -LogonType Interactive -RunLevel Limited
    $config = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
    Register-ScheduledTask -TaskName $NombreTarea -Action $accion -Trigger $disparador `
        -Principal $principal -Settings $config -Description 'Busca ofertas al iniciar sesión y abre un Excel' -Force | Out-Null
    Write-Host "Listo: '$NombreTarea' se ejecutará cada vez que inicies sesión (2 min después)." -ForegroundColor Green
    return
}

if (-not $Forzar -and -not $Solo -and (Test-Path $ArchUltima)) {
    $ultima = [datetime]::ParseExact((Get-Content $ArchUltima -Raw).Trim(), 'yyyy-MM-dd HH:mm:ss', $null)
    if ((Get-Date) - $ultima -lt (New-TimeSpan -Hours $HorasEntreEjecuciones)) {
        Log "Última búsqueda hace menos de $HorasEntreEjecuciones h ($ultima). Usa -Forzar para repetir."
        return
    }
}

# ================================================================ Utilidades
$Headers = @{
    'User-Agent'      = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36'
    'Accept-Language' = 'es-CL,es;q=0.9'
    'Accept'          = 'text/html,application/json;q=0.9,*/*;q=0.8'
}

function Bajar([string]$url) {
    $r = Invoke-WebRequest -Uri $url -Headers $Headers -UseBasicParsing -TimeoutSec 40
    [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())   # UTF-8 para tildes y ñ
}

# JSON grande: Windows PowerShell 5.1 limita ConvertFrom-Json a ~2 MB, así que usa otro lector
function Leer-Json([string]$txt) {
    if ($PSVersionTable.PSVersion.Major -ge 6) { return ($txt | ConvertFrom-Json -AsHashtable) }
    Add-Type -AssemblyName System.Web.Extensions
    $js = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $js.MaxJsonLength = [int]::MaxValue
    $js.RecursionLimit = 1000
    return $js.DeserializeObject($txt)
}

function Numero([string]$t) {
    $t = ($t -replace '[^\d\.,]', '')
    if ($t -match '^\d+[\.,]\d{1,2}$') { return [double]::Parse(($t -replace ',', '.'), $Inv) }
    $t = $t -replace '[\.,]', ''
    if ($t -match '^\d+$') { return [double]$t }
    return $null
}

function Limpiar([string]$s) {
    $s = [regex]::Replace($s, '<[^>]+>', ' ')
    $s = [Net.WebUtility]::HtmlDecode($s)
    ($s -replace '\s+', ' ').Trim()
}

function Nueva-Oferta($f, $nombre, $precio, $normal, $link, $vence, $nota) {
    if ($null -eq $precio) { return }
    $precio = [double]$precio
    if (-not $normal -or [double]$normal -lt $precio) { $normal = $precio }
    $normal = [double]$normal
    $desc = 0
    if ($normal -gt 0) { $desc = [math]::Round(($normal - $precio) / $normal, 4) }
    [pscustomobject]@{
        Categoria = $f.Categoria; Producto = (Limpiar $nombre); Tienda = $f.Tienda
        Moneda = $f.Moneda; Precio = $precio; PrecioNormal = $normal
        Descuento = $desc; PrecioAnterior = $null; Estado = ''; Vence = $vence; Nota = $nota; Link = $link
    }
}

function Url-Pagina([string]$url, [int]$pagina) {
    if ($pagina -le 1) { return $url }
    if ($url -match '\{start\}') { return $url }
    $sep = if ($url.Contains('?')) { '&' } else { '?' }
    "$url${sep}page=$pagina"
}

# ---------------------------------------------------------------- Lectores por tipo
function Leer-Solotodo($f, [int]$pag) {
    $json = Leer-Json (Bajar $f.URL)
    foreach ($res in $json.results) {
        foreach ($pe in $res.product_entries) {
            $p = $pe.product
            $pc = @($pe.metadata.prices_per_currency)[0]
            if (-not $pc) { continue }
            Nueva-Oferta $f $p.name (Numero $pc.offer_price) (Numero $pc.normal_price) "https://www.solotodo.cl/products/$($p.id)-$($p.slug)" '' 'Precio más bajo entre tiendas (puede ser con transferencia)'
        }
    }
}

function Leer-SteamBusqueda($f, [int]$pag) {
    $start = ($pag - 1) * 50
    $json = Leer-Json (Bajar ($f.URL -replace '\{start\}', $start))
    $html = [string]$json.results_html
    $rx = '<a href="(https://store\.steampowered\.com/(?:app|sub|bundle)/[^"?]+)[^"]*".*?<span class="title">(.*?)</span>.*?discount_original_price">([^<]+)</div><div class="discount_final_price">([^<]+)</div>'
    foreach ($m in [regex]::Matches($html, $rx, 'Singleline')) {
        Nueva-Oferta $f $m.Groups[2].Value (Numero $m.Groups[4].Value) (Numero $m.Groups[3].Value) $m.Groups[1].Value '' ''
    }
}

function Leer-Falabella($f, [int]$pag) {
    $html = Bajar (Url-Pagina $f.URL $pag)
    $m = [regex]::Match($html, '<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', 'Singleline')
    if (-not $m.Success) { throw 'Falabella cambió su página (no se encontró __NEXT_DATA__)' }
    $data = Leer-Json $m.Groups[1].Value
    foreach ($r in $data.props.pageProps.results) {
        $normal = $null; $oferta = $null; $cmr = $null
        foreach ($p in $r.prices) {
            $v = Numero (@($p.price)[0])
            if ($p.crossed) { if (-not $normal -or $v -gt $normal) { $normal = $v } }
            elseif ($p.type -eq 'cmrPrice') { $cmr = $v }
            elseif (-not $oferta -or $v -lt $oferta) { $oferta = $v }
        }
        if (-not $oferta) { $oferta = $cmr; $cmr = $null }
        $nota = ''
        if ($cmr -and $cmr -lt $oferta) { $nota = "Con tarjeta CMR: $('{0:N0}' -f $cmr)" }
        if ($r.sellerName -and $r.sellerName -notmatch '^falabella$') { $nota = (($nota, "Vende: $($r.sellerName)") | Where-Object { $_ }) -join ' · ' }
        $nombre = $r.displayName
        if ($r.brand) { $nombre = "$($r.brand) - $nombre" }
        Nueva-Oferta $f $nombre $oferta $normal $r.url '' $nota
    }
}

function Leer-Vtex($f, [int]$pag) {
    $desde = ($pag - 1) * 50
    $url = "$($f.URL)&_from=$desde&_to=$($desde + 49)"
    $json = @(Leer-Json (Bajar $url))
    foreach ($p in $json) {
        $o = $null
        foreach ($it in $p.items) { foreach ($s in $it.sellers) { if ($s.commertialOffer.AvailableQuantity -gt 0) { $o = $s.commertialOffer; break } }; if ($o) { break } }
        if (-not $o) { continue }
        Nueva-Oferta $f $p.productName $o.Price $o.ListPrice $p.link '' ''
    }
}

function Leer-Nintendo($f, [int]$pag) {
    $html = Bajar $f.URL
    $base = ([regex]::Match($f.URL, '^https://[^/]+/[a-z]{2}-[a-z]{2}')).Value
    $vistos = @{}
    foreach ($trozo in ($html -split '"__typename":"Product"' | Select-Object -Skip 1)) {
        $n = [regex]::Match($trozo, '"name":"((?:[^"\\]|\\.)*)","nsuid":"(\d+)"')
        $p = [regex]::Match($trozo, '"prices":\{"__typename":"Prices","currency":"\w+","discounted":(true|false),"regularPrice":([\d\.]+),"finalPrice":([\d\.]+)')
        $u = [regex]::Match($trozo, '"urlKey":"([^"]+)"')
        if (-not ($n.Success -and $p.Success -and $u.Success)) { continue }
        if ($vistos.ContainsKey($n.Groups[2].Value)) { continue }
        $vistos[$n.Groups[2].Value] = 1
        $e = [regex]::Match($trozo, '"discountPriceEnd":"([^"]+)"')
        $vence = ''
        if ($e.Success) { $vence = ([datetime]::Parse($e.Groups[1].Value, $Inv, 'RoundtripKind')).ToLocalTime().ToString('yyyy-MM-dd HH:mm') }
        $nombre = [regex]::Unescape($n.Groups[1].Value)
        Nueva-Oferta $f $nombre ([double]::Parse($p.Groups[3].Value, $Inv)) ([double]::Parse($p.Groups[2].Value, $Inv)) "$base/store/products/$($u.Groups[1].Value)/" $vence ''
    }
}

function Leer-CheapShark($f, [int]$pag) {
    $tiendas = @{ '1' = 'Steam'; '2' = 'GamersGate'; '3' = 'Green Man Gaming'; '7' = 'GOG'; '11' = 'Humble'; '13' = 'Ubisoft'; '15' = 'Fanatical'; '25' = 'Epic'; '27' = 'Gamesplanet'; '30' = 'IndieGala' }
    $url = "$($f.URL)&pageNumber=$($pag - 1)"
    foreach ($d in @(Leer-Json (Bajar $url))) {
        $t = $tiendas[[string]$d.storeID]; if (-not $t) { $t = "tienda $($d.storeID)" }
        $nota = "En $t"
        if ($d.metacriticScore -and $d.metacriticScore -ne '0') { $nota += " · Metacritic $($d.metacriticScore)" }
        if ($d.steamRatingText) { $nota += " · Steam: $($d.steamRatingText)" }
        Nueva-Oferta $f $d.title (Numero $d.salePrice) (Numero $d.normalPrice) "https://www.cheapshark.com/redirect?dealID=$($d.dealID)" '' $nota
    }
}

function Leer-Epic($f, [int]$pag) {
    $json = Leer-Json (Bajar $f.URL)
    foreach ($x in $json.data.Catalog.searchStore.elements) {
        $promos = @()
        if ($x.promotions) { foreach ($g in $x.promotions.promotionalOffers) { $promos += @($g.promotionalOffers) } }
        $gratis = $promos | Where-Object { $_.discountSetting.discountPercentage -eq 0 } | Select-Object -First 1
        if (-not $gratis) { continue }
        $slug = $x.productSlug
        if (-not $slug -and $x.catalogNs -and $x.catalogNs.mappings) { $slug = @($x.catalogNs.mappings)[0].pageSlug }
        if (-not $slug -and $x.offerMappings) { $slug = @($x.offerMappings)[0].pageSlug }
        $slug = ($slug -replace '/home$', '')
        $normal = [double]$x.price.totalPrice.originalPrice / 100
        $vence = ([datetime]::Parse($gratis.endDate, $Inv, 'RoundtripKind')).ToLocalTime().ToString('yyyy-MM-dd HH:mm')
        Nueva-Oferta $f $x.title 0 $normal "https://store.epicgames.com/es-ES/p/$slug" $vence 'GRATIS en Epic Games (inicia sesión y reclámalo)'
    }
}

function Leer-GamerPower($f, [int]$pag) {
    foreach ($g in @(Leer-Json (Bajar $f.URL))) {
        if ($g.status -ne 'Active') { continue }
        $vence = ''
        if ($g.end_date -and $g.end_date -ne 'N/A') { $vence = ([datetime]::Parse($g.end_date, $Inv)).ToString('yyyy-MM-dd HH:mm') }
        $titulo = $g.title -replace '\s*Giveaway$', ''
        Nueva-Oferta $f $titulo 0 (Numero $g.worth) $g.open_giveaway_url $vence "GRATIS · $($g.platforms)"
    }
}

function Leer-Html($f, [int]$pag) {
    $html = Bajar (Url-Pagina $f.URL $pag)
    $u = [Uri]$f.URL
    $base = "$($u.Scheme)://$($u.Host)"
    $rxPrecio = if ($f.Moneda -eq 'USD') { '(?:US)?\$\s?(\d{1,4}[\.,]\d{2})(?!\d)' } else { '\$\s?(\d{1,3}(?:\.\d{3})+|\d{4,7})(?![\d,])' }
    $marcas = [regex]::Matches($html, $f.PatronLink)
    $orden = New-Object System.Collections.Generic.List[object]
    $visto = @{}
    foreach ($m in $marcas) { $l = $m.Groups[1].Value; if (-not $visto.ContainsKey($l)) { $visto[$l] = 1; $orden.Add(@{ Link = $l; Pos = $m.Index }) } }
    for ($i = 0; $i -lt $orden.Count; $i++) {
        $ini = $orden[$i].Pos
        $fin = if ($i + 1 -lt $orden.Count) { $orden[$i + 1].Pos } else { [math]::Min($html.Length, $ini + 6000) }
        $tarjeta = $html.Substring($ini, $fin - $ini)
        $link = $orden[$i].Link; if ($link.StartsWith('/')) { $link = $base + $link }
        $precios = @([regex]::Matches($tarjeta, $rxPrecio) | ForEach-Object { Numero $_.Groups[1].Value } | Where-Object { $_ -gt 0 } | Select-Object -First 2)
        if ($precios.Count -eq 0) { continue }
        $alt = [regex]::Match($tarjeta, 'alt="([^"]{4,150})"')
        $nombre = if ($alt.Success) { $alt.Groups[1].Value } else { ($link -split '/')[-1] -replace '[-_]', ' ' }
        Nueva-Oferta $f $nombre ($precios | Measure-Object -Minimum).Minimum ($precios | Measure-Object -Maximum).Maximum $link '' ''
    }
}

# ================================================================ Búsqueda
$fuentes = @(Import-Csv $ArchFuentes -Delimiter ';' -Encoding UTF8 | Where-Object { $_.Activo -match '^(SI|SÍ|S|1)$' })
if ($Solo) { $fuentes = @($fuentes | Where-Object { $_.Categoria -like "*$Solo*" -or $_.Tienda -like "*$Solo*" }) }
$ofertas = New-Object System.Collections.Generic.List[object]
$errores = New-Object System.Collections.Generic.List[object]
$vistosLink = @{}

foreach ($f in $fuentes) {
    $paginas = 1
    if ($f.Paginas -match '^\d+$') { $paginas = [int]$f.Paginas }
    $total = 0
    try {
        for ($pag = 1; $pag -le $paginas; $pag++) {
            $res = switch ($f.Tipo.ToLower()) {
                'solotodo'    { Leer-Solotodo $f $pag }
                'steam'       { Leer-SteamBusqueda $f $pag }
                'falabella'   { Leer-Falabella $f $pag }
                'vtex'        { Leer-Vtex $f $pag }
                'nintendo'    { Leer-Nintendo $f $pag }
                'cheapshark'  { Leer-CheapShark $f $pag }
                'epic'        { Leer-Epic $f $pag }
                'gamerpower'  { Leer-GamerPower $f $pag }
                'html'        { Leer-Html $f $pag }
                default       { throw "Tipo desconocido: $($f.Tipo)" }
            }
            $res = @($res | Where-Object { $_ })
            if ($f.Filtro) { $res = @($res | Where-Object { $_.Producto -match "(?i)$($f.Filtro)" }) }
            foreach ($o in $res) {
                if ($vistosLink.ContainsKey($o.Link)) { continue }   # evita repetidos entre búsquedas
                $vistosLink[$o.Link] = 1
                $ofertas.Add($o); $total++
            }
            if ($pag -lt $paginas) { Start-Sleep -Milliseconds 1200 }
        }
        if ($total -eq 0 -and $f.Tipo -notin @('epic', 'gamerpower')) { throw 'No se encontraron productos (la página pudo cambiar o bloquear la consulta)' }
        Log ("{0} / {1}: {2} productos" -f $f.Categoria, $f.Tienda, $total)
    } catch {
        $msg = $_.Exception.Message
        $errores.Add([pscustomobject]@{ Categoria = $f.Categoria; Tienda = $f.Tienda; Error = $msg; URL = $f.URL })
        Log ("ERROR {0} / {1}: {2}" -f $f.Categoria, $f.Tienda, $msg)
    }
    Start-Sleep -Milliseconds 1200   # no saturar las tiendas
}

# ================================================================ Comparar con la última vez
$ultimoPrecio = @{}
if (Test-Path $ArchHistorial) {
    foreach ($h in (Import-Csv $ArchHistorial -Delimiter ';' -Encoding UTF8)) {
        try { $ultimoPrecio[$h.Link] = [double]::Parse($h.Precio, $Inv) } catch { }
    }
}
foreach ($o in $ofertas) {
    if ($ultimoPrecio.ContainsKey($o.Link)) {
        $o.PrecioAnterior = $ultimoPrecio[$o.Link]
        if ($o.Precio -lt $o.PrecioAnterior) { $o.Estado = 'BAJÓ' }
        elseif ($o.Precio -gt $o.PrecioAnterior) { $o.Estado = 'subió' }
    } else { $o.Estado = 'nuevo' }
}
foreach ($o in $ofertas) { $ultimoPrecio[$o.Link] = $o.Precio }
$ultimoPrecio.GetEnumerator() | ForEach-Object {
    [pscustomobject]@{ Link = $_.Key; Precio = ([double]$_.Value).ToString($Inv) }
} | Export-Csv $ArchHistorial -Delimiter ';' -NoTypeInformation -Encoding UTF8

# ================================================================ Excel
$ordenadas = $ofertas | Sort-Object Categoria, @{ Expression = 'Descuento'; Descending = $true }
$mejores = foreach ($g in ($ofertas | Group-Object Categoria | Sort-Object Name)) {
    $g.Group | Where-Object { $_.Descuento -gt 0 -or $_.Estado -eq 'BAJÓ' } |
        Sort-Object @{ Expression = { if ($_.Estado -eq 'BAJÓ') { 0 } else { 1 } } }, @{ Expression = 'Descuento'; Descending = $true } |
        Select-Object -First $TopPorCategoria
}
$bajaron = @($ofertas | Where-Object Estado -eq 'BAJÓ' | Sort-Object @{ Expression = 'Descuento'; Descending = $true })

function Filas($lista) {
    foreach ($o in $lista) {
        [pscustomobject][ordered]@{
            'Categoría' = $o.Categoria; 'Producto' = $o.Producto; 'Tienda' = $o.Tienda; 'Moneda' = $o.Moneda
            'Precio' = $o.Precio; 'Precio normal' = $o.PrecioNormal; 'Descuento' = $o.Descuento
            'Precio anterior' = $o.PrecioAnterior; 'Estado' = $o.Estado; 'Vence' = $o.Vence; 'Nota' = $o.Nota; 'Link' = $o.Link
        }
    }
}

New-Item -ItemType Directory -Path $CarpetaExcel -Force | Out-Null
$archExcel = Join-Path $CarpetaExcel ("Ofertas_{0}.xlsx" -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'))

try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Log 'Instalando módulo ImportExcel (solo la primera vez)...'
        if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
        }
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module ImportExcel

    function Dar-Formato($pkg, [string]$hoja) {
        $ws = $pkg.Workbook.Worksheets[$hoja]
        if (-not $ws.Dimension) { return }
        $ult = $ws.Dimension.End.Row
        for ($r = 2; $r -le $ult; $r++) {
            $fmt = if ($ws.Cells[$r, 4].Value -eq 'USD') { '"US$" #,##0.00' } else { '"$" #,##0' }
            foreach ($c in 5, 6, 8) { $ws.Cells[$r, $c].Style.Numberformat.Format = $fmt }
            $ws.Cells[$r, 7].Style.Numberformat.Format = '0%'
            $celda = $ws.Cells[$r, 12]
            $url = [string]$celda.Value
            if ($celda.Hyperlink) { $url = $celda.Hyperlink.OriginalString; $celda.Hyperlink = $null }
            if ($url -match '^https?://') {
                $celda.Formula = 'HYPERLINK("' + ($url -replace '"', '%22') + '","Ver")'
                $celda.Style.Font.UnderLine = $true
                $celda.Style.Font.Color.SetColor([Drawing.Color]::FromArgb(5, 99, 193))
            }
            if ($ws.Cells[$r, 9].Value -eq 'BAJÓ') {
                $ws.Cells[$r, 9].Style.Font.Bold = $true
                $ws.Cells[$r, 9].Style.Font.Color.SetColor([Drawing.Color]::FromArgb(0, 97, 0))
            }
        }
        $anchos = @(20, 60, 22, 8, 12, 13, 11, 13, 9, 16, 34, 7)
        for ($i = 0; $i -lt $anchos.Count; $i++) { $ws.Column($i + 1).Width = $anchos[$i] }
        Add-ConditionalFormatting -Worksheet $ws -Range "G2:G$ult" -RuleType GreaterThanOrEqual -ConditionValue 0.4 -BackgroundColor ([Drawing.Color]::FromArgb(198, 239, 206)) -ForegroundColor ([Drawing.Color]::FromArgb(0, 97, 0))
        Add-ConditionalFormatting -Worksheet $ws -Range "G2:G$ult" -RuleType Between -ConditionValue 0.15 -ConditionValue2 0.3999 -BackgroundColor ([Drawing.Color]::FromArgb(255, 235, 156))
        Add-ConditionalFormatting -Worksheet $ws -Range "I2:I$ult" -RuleType Equal -ConditionValue '"BAJÓ"' -BackgroundColor ([Drawing.Color]::FromArgb(198, 239, 206))
    }

    $comun = @{ Path = $archExcel; AutoFilter = $true; FreezeTopRow = $true; BoldTopRow = $true; PassThru = $true }
    $pkg = @(Filas $mejores) | Export-Excel @comun -WorksheetName 'Mejores' -TableStyle Medium2
    Dar-Formato $pkg 'Mejores'; Close-ExcelPackage $pkg
    if ($bajaron.Count) {
        $pkg = @(Filas $bajaron) | Export-Excel @comun -WorksheetName 'Bajaron de precio' -TableStyle Medium7
        Dar-Formato $pkg 'Bajaron de precio'; Close-ExcelPackage $pkg
    }
    $pkg = @(Filas $ordenadas) | Export-Excel @comun -WorksheetName 'Todas' -TableStyle Light9
    Dar-Formato $pkg 'Todas'; Close-ExcelPackage $pkg
    if ($errores.Count) {
        $errores | Export-Excel -Path $archExcel -WorksheetName 'Fuentes con error' -AutoSize -BoldTopRow
    }
    Log "Excel creado: $archExcel  ($($ofertas.Count) productos, $($bajaron.Count) bajaron de precio, $($errores.Count) fuentes con error)"
} catch {
    $archExcel = [IO.Path]::ChangeExtension($archExcel, '.csv')
    Filas $ordenadas | Export-Csv $archExcel -Delimiter ';' -NoTypeInformation -Encoding UTF8
    Log "No se pudo crear el Excel ($($_.Exception.Message)). Se guardó como CSV: $archExcel"
}

# Deja solo los últimos 15 archivos para no llenar el Escritorio
Get-ChildItem $CarpetaExcel -Filter 'Ofertas_*' | Sort-Object LastWriteTime -Descending | Select-Object -Skip 15 | Remove-Item -ErrorAction SilentlyContinue

if (-not $Solo) { Set-Content -Path $ArchUltima -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }
if (-not $NoAbrir -and (Test-Path $archExcel)) { Invoke-Item $archExcel }
