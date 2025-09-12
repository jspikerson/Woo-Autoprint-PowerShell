param(
  [switch]$Backlog,
  [switch]$ReprintLast
)


# =========================
# CONFIG — Edit these
# =========================
$Store           = 'https://yourdomain.com' #Put your domain here
$StatusesToPrint = @('processing') #Only select orders that are processing
$PrinterName     = $null	# e.g. '"HP LaserJet"'; $null = default printer
$LogoPath        = $Null	# e.g. Join-Path $PSScriptRoot 'logo.png'
$PageSize        = 'Letter'	# maybe 'A4'?
$SumatraPath    = Join-Path $PSScriptRoot 'SumatraPDF.exe' #Change to point to SumatraPDF executable
$PrintSettings  = 'paper=letter,duplex,long-edge' # Alternative: $PrintSettings = 'paper=a4,scaling=noscale'

# Paths
$CredPath  = Join-Path $PSScriptRoot 'woo-cred.xml'
$StatePath = Join-Path $PSScriptRoot 'woo-print-state.json'
$OutDir    = Join-Path $PSScriptRoot 'printed'
New-Item -ItemType Directory -Path $OutDir -ErrorAction SilentlyContinue | Out-Null

# =========================
# Setup & Auth
# =========================
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# First run: prompt & save encrypted credential (ck_... / cs_...)
if (-not (Test-Path $CredPath)) {
  Write-Host "First run: enter your WooCommerce API credentials." -ForegroundColor Cyan
  Write-Host "Username = Consumer Key (ck_...), Password = Consumer Secret (cs_...)"
  $cred = Get-Credential -Message "WooCommerce API"
  $cred | Export-Clixml -Path $CredPath
  Write-Host "Saved encrypted credential to $CredPath"
}
$Cred = Import-Clixml -Path $CredPath
$Base = "$Store/wp-json/wc/v3"

# =========================
# Utilities
# =========================
function ConvertTo-QueryString {
  param([hashtable]$Params)
  if (-not $Params -or $Params.Count -eq 0) { return '' }
  ($Params.GetEnumerator() | Sort-Object Key | ForEach-Object {
    "$($_.Key)=$([uri]::EscapeDataString([string]$_.Value))"
  }) -join '&'
}

# =========================
# PDF generation & printing
# =========================
function Get-PdfEnginePath {
  $candidates = @(
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe"
  )
  foreach ($p in $candidates) { if (Test-Path $p) { return $p } }
  throw "No Edge/Chrome found. Install Microsoft Edge (recommended)."
}

function Convert-HtmlToPdf {
  param([Parameter(Mandatory)][string]$HtmlPath,
        [Parameter(Mandatory)][string]$PdfPath)
  $bin = Get-PdfEnginePath
  $fileUrl = "file:///" + ($HtmlPath -replace '\\','/')
  $args = @(
    '--headless',
    '--disable-gpu',
    "--print-to-pdf=$PdfPath",
    '--print-to-pdf-no-header',
    $fileUrl
  )
  Start-Process -FilePath $bin -ArgumentList $args -Wait -NoNewWindow | Out-Null
  if (-not (Test-Path $PdfPath)) { throw "PDF not created: $PdfPath" }
  return $PdfPath
}

function Print-Pdf {
  param([Parameter(Mandatory)][string]$PdfPath,
        [string]$PrinterName)

  # Use SumatraPDF if present (silent, no window)
  if (Test-Path $SumatraPath) {
    $args = @('-silent','-exit-when-done')

    if ($PrinterName) {
      $args += @('-print-to', $PrinterName)
    } else {
      $args += '-print-to-default'
    }

    if ($PrintSettings -and $PrintSettings.Trim()) {
      $args += @('-print-settings', $PrintSettings)
    }

    $args += $PdfPath

    try {
      # Wait so the process fully hands off to the spooler before we update LastId
      Start-Process -FilePath $SumatraPath -ArgumentList $args -Wait -NoNewWindow | Out-Null
      return $true
    } catch {
      Write-Warning "Silent print failed via SumatraPDF: $($_.Exception.Message)"
      return $false
    }
  }

  # Fallback (non-silent): default shell print (may pop briefly)
  try {
    if ($PrinterName) { Start-Process -FilePath $PdfPath -Verb PrintTo -ArgumentList $PrinterName | Out-Null }
    else { Start-Process -FilePath $PdfPath -Verb Print | Out-Null }
    Start-Sleep -Seconds 2
    return $true
  } catch {
    Write-Warning "Fallback print failed: $($_.Exception.Message)"
    return $false
  }
}


# =========================
# HTML helpers (addresses, options, notes)
# =========================
function Get-AddressHtml {
  param($a, [string]$Title)
  $enc = [System.Net.WebUtility]
  $lines = New-Object System.Collections.Generic.List[string]
  if ($Title) { $lines.Add("<div style='font-weight:600;margin-bottom:4px;'>$Title</div>") }
  $fullName = (($a.first_name, $a.last_name) -join ' ').Trim()
  foreach ($s in @(
      $fullName, $a.company, $a.address_1, $a.address_2,
      ("$($a.city) $($a.state) $($a.postcode)").Trim(), $a.country, $a.phone, $a.email
    )) {
    if ($s) { $lines.Add("<div>" + $enc::HtmlEncode([string]$s) + "</div>") }
  }
  return ($lines -join "`n")
}

# Extract Product Add-Ons pairs from _pao_ids and mask URLs
function Get-PAOOptionsHtml {
  param($LineItem)

  $enc = [System.Net.WebUtility]
  if (-not $LineItem.meta_data) { return '' }

  $pao = $LineItem.meta_data | Where-Object {
    ($_.display_key -and $_.display_key -ieq '_pao_ids') -or
    ($_.key         -and $_.key         -ieq '_pao_ids')
  } | Select-Object -First 1
  if (-not $pao) { return '' }

  $arr = $null
  if ($pao.PSObject.Properties.Name -contains 'display_value' -and $pao.display_value) { $arr = $pao.display_value }
  elseif ($pao.PSObject.Properties.Name -contains 'value') { $arr = $pao.value }

  if ($arr -is [string] -and $arr.Trim().StartsWith('[')) { try { $arr = $arr | ConvertFrom-Json } catch {} }
  if (-not $arr) { return '' }
  if (-not ($arr -is [System.Collections.IEnumerable]) -or ($arr -is [string])) { $arr = @($arr) }

  $reUrl = '(?i)^(?:https?://|www\.)\S+$'
  $mask  = 'Image uploaded by customer'

  $rows = foreach ($it in $arr) {
    $k = $it.key; if (-not $k) { if ($it.label){ $k=$it.label } elseif ($it.name){ $k=$it.name } }
    $v = $it.value; if (-not $v -and $it.raw_value) { $v = $it.raw_value }
    if ([string]::IsNullOrWhiteSpace([string]$k) -or $null -eq $v) { continue }

    $valStr = $null
    if ($v -is [string]) {
      $s = $v.Trim(); if ($s -match $reUrl) { $s = $mask }; $valStr = $s
    }
    elseif ($v -is [System.Collections.IEnumerable] -and -not ($v -is [string])) {
      $parts = foreach ($p in $v) {
        if ($p -is [string]) {
          $ps = $p.Trim(); if ($ps -match $reUrl) { $mask } else { $ps }
        }
        elseif ($p.PSObject) {
          $cand = $null
          foreach ($prop in @('value','raw_value','url','href','label','name')) {
            if ($p.PSObject.Properties.Name -contains $prop -and $p.$prop) { $cand = [string]$p.$prop; break }
          }
          if (-not $cand) { $cand = ($p | ConvertTo-Json -Compress -Depth 3) }
          if ($cand -match $reUrl) { $mask } else { $cand }
        }
        else { [string]$p }
      }
      $valStr = ($parts -join ', ')
    }
    else {
      if ($v.PSObject) {
        $cand = $null
        foreach ($prop in @('value','raw_value','url','href','label','name')) {
          if ($v.PSObject.Properties.Name -contains $prop -and $v.$prop) { $cand = [string]$v.$prop; break }
        }
        if (-not $cand) { $cand = ($v | ConvertTo-Json -Compress -Depth 3) }
        if ($cand -match $reUrl) { $cand = $mask }
        $valStr = $cand
      } else {
        $s = [string]$v; if ($s -match $reUrl) { $s = $mask }; $valStr = $s
      }
    }

    if ([string]::IsNullOrWhiteSpace([string]$valStr)) { continue }
    "<div><span class='optk'>$($enc::HtmlEncode([string]$k))</span>: <span class='optv'>$($enc::HtmlEncode([string]$valStr))</span></div>"
  }

  return ($rows -join "`n")
}

# Fetch notes authored by the customer from /orders/{id}/notes
function Get-CustomerNotes {
  param([Parameter(Mandatory)][int]$OrderId)
  $ck = $Cred.UserName
  $cs = $Cred.GetNetworkCredential().Password
  $uri = "$Base/orders/$OrderId/notes?per_page=100&order=asc&consumer_key=$([uri]::EscapeDataString($ck))&consumer_secret=$([uri]::EscapeDataString($cs))"
  $resp = Invoke-RestMethod -Method GET -Uri $uri -ErrorAction Stop
  return @($resp | Where-Object { $_.customer_note -eq $true } | Select-Object -ExpandProperty note)
}

# =========================
# HTML template builder
# =========================
function Get-PackingSlipHtml {
  param(
    [Parameter(Mandatory)]$Order,
    [string]$PageSize = 'Letter',
    [string]$LogoPath = $null,
    [string[]]$CustomerNotes = @()
  )
  $enc = [System.Net.WebUtility]
  $orderNo = $enc::HtmlEncode([string]$Order.number)
  $status  = $enc::HtmlEncode([string]$Order.status)
  $dateStr = try { ([datetime]$Order.date_created).ToString('yyyy-MM-dd HH:mm') } catch { '' }

  $logoHtml = ''
  if ($LogoPath -and (Test-Path $LogoPath)) {
    $logoUrl = "file:///" + ($LogoPath -replace '\\','/')
    $logoHtml = "<img src='$logoUrl' style='height:75px;object-fit:contain;' />"
  }

  $billingHtml  = Get-AddressHtml -a $Order.billing  -Title 'BILL TO'
  $shippingHtml = Get-AddressHtml -a $Order.shipping -Title 'SHIP TO'
  
	# Append shipping method(s) to the SHIP TO box
	$shipTitles = @()
	if ($Order.shipping_lines -and $Order.shipping_lines.Count -gt 0) {
	  foreach ($sl in $Order.shipping_lines) {
		if ($sl.PSObject.Properties.Name -contains 'method_title' -and $sl.method_title) {
		  $shipTitles += [string]$sl.method_title
		}
	  }
	}
	$shipTitles = $shipTitles | Where-Object { $_ } | Select-Object -Unique
	if ($shipTitles.Count -gt 0) {
	  $joined = ($shipTitles | ForEach-Object { $enc::HtmlEncode($_) }) -join ', '
	  $shippingHtml += "`n<div style='margin-top:6px;'><span style='font-weight:600'>Method:</span> $joined</div>"
	}


  # Build item rows (Qty | Item | Shipped | BO) and show _pao_ids under the name
  $rows = foreach ($li in $Order.line_items) {
    $qty  = [int]$li.quantity
    $name = $enc::HtmlEncode([string]$li.name)
    $optsHtml = Get-PAOOptionsHtml -LineItem $li
    $descHtml = if ($optsHtml) { "$name<div class='opts'>$optsHtml</div>" } else { $name }

@"
<tr>
  <td class='qty'>$qty</td>
  <td class='desc'>$descHtml</td>
  <td class='ship'>&nbsp;</td>
  <td class='bo'>&nbsp;</td>
</tr>
"@
  }
  $rowsHtml = ($rows -join "`n")

  # Customer notes block (checkout note + any customer-authored notes)
  $notesList = New-Object System.Collections.Generic.List[string]
  if ($Order.customer_note -and -not [string]::IsNullOrWhiteSpace([string]$Order.customer_note)) {
    $notesList.Add([string]$Order.customer_note)
  }
  foreach ($n in $CustomerNotes) {
    if ($n -and -not [string]::IsNullOrWhiteSpace([string]$n)) { $notesList.Add([string]$n) }
  }
  $seen = @{}
  $finalNotes = foreach ($n in $notesList) { $k = $n.Trim(); if (-not $seen[$k]) { $seen[$k]=$true; $k } }

  $notesBlock = ''
  if ($finalNotes.Count -gt 0) {
    $items = $finalNotes | ForEach-Object { "<div class='note'>• $($enc::HtmlEncode($_))</div>" }
    $notesBlock = "<div class='notes'><div class='notes-title'>Customer Notes</div>$($items -join "`n")</div>"
  }

@"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<style>
  @page { size: $PageSize; margin: 12mm; }
  html, body { padding:0; margin:0; }
  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, "Noto Sans", "Liberation Sans", sans-serif; font-size: 12px; color: #111; }
  h1 { font-size: 25px; margin: 0; }
  .header { display: grid; grid-template-columns: 1fr auto; gap: 12px; align-items: center; margin: 0 0 10px 0; }
  .meta { font-size: 12px; }
  .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin: 10px 0 16px; }
  .box { border: 1px solid #333; padding: 8px; border-radius: 6px; }
  table { width: 100%; border-collapse: collapse; }
  th, td { border: 1px solid #333; padding: 6px; vertical-align: top; }
  th { text-align: left; background: #f3f3f3; }

  .qty  { width: 10%; text-align: center; vertical-align: middle;}
  .desc { width: 60%; }
  .ship { width: 15%; text-align: center; }
  .bo   { width: 15%; text-align: center; }
  td.ship, td.bo { height: 24px; }

  .opts { margin-top: 4px; font-size: 11px; color: #333; }
  .opts .optk { font-weight: 600; }

  .notes { margin-top: 12px; border: 1px solid #333; border-radius: 6px; padding: 8px; }
  .notes-title { font-weight: 600; margin-bottom: 4px; }
  .note { margin-left: 8px; }

  tr { page-break-inside: avoid; }
  .footer { margin-top: 20px; font-size: 16px; color: #444; text-align: center;}
</style>
</head>
<body>
  <div class="header">
    <div>
      <h1>Packing Slip $orderNo</h1>
      <div class="meta">
        <div>Date: $dateStr</div>
        <div>Status: $status</div>
      </div>
    </div>
    <div>$logoHtml</div>
  </div>

  <div class="two-col">
    <div class="box">$billingHtml</div>
    <div class="box">$shippingHtml</div>
  </div>

  <table>
    <thead>
      <tr>
        <th class="qty">Qty</th>
        <th class="desc">Item</th>
        <th class="ship">Shipped</th>
        <th class="bo">BO</th>
      </tr>
    </thead>
    <tbody>
      $rowsHtml
    </tbody>
  </table>

  $notesBlock

  <div class="footer">Thank you for your order.</div>
</body>
</html>
"@
}

# =========================
# State & fetch
# =========================
function Save-PrintState($state) { $state | ConvertTo-Json | Set-Content -Encoding UTF8 $StatePath }

function Load-PrintState {
  param(
    [switch]$Backlog,        # existing behavior: start from 0 (print everything)
    [switch]$IncludeLast     # NEW: include the previously printed id in next run
  )

  # If we already have state and we're not doing a backlog run, return it (optionally stepped back by 1)
  if ((Test-Path $StatePath) -and -not $Backlog) {
    try {
      $s = Get-Content $StatePath | ConvertFrom-Json
      if ($IncludeLast -and $s -and ($s.PSObject.Properties.Name -contains 'LastId')) {
        $prev = [int]$s.LastId
        # Decrement by 1 so the > LastId filter will include the previous id
        return [pscustomobject]@{ LastId = [Math]::Max(0, $prev - 1) }
      }
      return $s
    } catch {}
  }

  # Seed: skip history on first run (start from newest existing)
  $ck = $Cred.UserName; $cs = $Cred.GetNetworkCredential().Password
  $seed = "$Base/orders?per_page=1&order=desc&orderby=id&status=$([string]::Join(',', $StatusesToPrint))&consumer_key=$ck&consumer_secret=$cs"
  $latest = $null
  try { $latest = Invoke-RestMethod -Method GET -Uri $seed -ErrorAction Stop } catch {}
  $maxId = if ($latest) { ($latest | Select-Object -ExpandProperty id -ErrorAction SilentlyContinue) } else { 0 }
  [pscustomobject]@{ LastId = [int]$maxId }
}

function Get-WooSinceId {
  param([int]$LastId, [string[]]$Statuses)

  $ck = $Cred.UserName; $cs = $Cred.GetNetworkCredential().Password
  $results = @(); $page = 1; $stop = $false
  do {
    $qs = ConvertTo-QueryString @{
      per_page = 100
      page     = $page
      order    = 'desc'
      orderby  = 'id'
      status   = ($Statuses -join ',')
    }
    $uri   = "$Base/orders?$qs&consumer_key=$([uri]::EscapeDataString($ck))&consumer_secret=$([uri]::EscapeDataString($cs))"
    $batch = Invoke-RestMethod -Method GET -Uri $uri -ErrorAction Stop
    if (-not $batch -or $batch.Count -eq 0) { break }

    $newOnPage = @($batch | Where-Object { [int]$_.id -gt $LastId })
    $results  += $newOnPage

    if ($newOnPage.Count -lt $batch.Count) { $stop = $true } # reached older items
    $page++
  } while (-not $stop)

  return $results | Sort-Object { [int]$_.id }   # print oldest → newest
}

# =========================
# MAIN
# =========================
$state = Load-PrintState -Backlog:$Backlog -IncludeLast:$ReprintLast
$new   = Get-WooSinceId -LastId $state.LastId -Statuses $StatusesToPrint

if ($new.Count -eq 0) {
  Write-Host "No new orders to print. LastId=$($state.LastId)"
  return
}

foreach ($o in $new) {
  $orderNo   = $o.number
  $custNotes = @()
  try { $custNotes = Get-CustomerNotes -OrderId ([int]$o.id) } catch { $custNotes = @() }

  $html = Get-PackingSlipHtml -Order $o -PageSize $PageSize -LogoPath $LogoPath -CustomerNotes $custNotes

  $htmlPath = Join-Path $OutDir ("order-$orderNo.html")
  $pdfPath  = Join-Path $OutDir ("order-$orderNo.pdf")

  $html | Set-Content -Encoding UTF8 $htmlPath
  Convert-HtmlToPdf -HtmlPath $htmlPath -PdfPath $pdfPath | Out-Null

  $ok = Print-Pdf -PdfPath $pdfPath -PrinterName $PrinterName
  if ($ok) { Write-Host "Printed order #$orderNo → $pdfPath" } else { Write-Host "Saved PDF (not printed): $pdfPath" }
}

$state.LastId = [int](($new | Measure-Object -Property id -Maximum).Maximum)
Save-PrintState $state
Write-Host "Updated LastId -> $($state.LastId)"
