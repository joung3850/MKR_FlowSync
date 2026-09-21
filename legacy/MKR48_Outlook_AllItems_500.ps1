#requires -version 5.1
# MKR48/26, MKR48/26-1, MKR49/26, MKR49/26-1, MKR56/26 Outlook synchronizer
# The workbook distributed with this script is intentionally blank until this script succeeds.
param([switch]$SelfTest)

[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding
$ErrorActionPreference = 'Stop'
$script:SuppressLog = $false

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$WorkbookPath = Join-Path $Root 'MKR_TEST.xlsx'
$LogPath = Join-Path $Root 'MKR_SYNC_LOG.txt'
$BackupDir = Join-Path $Root 'Backup'
$AttachmentDir = Join-Path $Root 'MKR_Attachments'
$RunAttachmentDir = Join-Path $AttachmentDir ('Run_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
$TargetSender = 'hnagasawa@milbon.com'
$TargetMkrKeys = @('MKR48/26', 'MKR48/26-1', 'MKR49/26', 'MKR49/26-1', 'MKR56/26')
$TargetSpecs = @(
    [pscustomobject]@{ Key='MKR48/26';   Sheet='MKR48_26' },
    [pscustomobject]@{ Key='MKR48/26-1'; Sheet='MKR48_26-1' },
    [pscustomobject]@{ Key='MKR49/26';   Sheet='MKR49_26' },
    [pscustomobject]@{ Key='MKR49/26-1'; Sheet='MKR49_26-1' },
    [pscustomobject]@{ Key='MKR56/26';   Sheet='MKR56_26' }
)
$MaxMessages = 500
$MinimumExpectedItems = 1
$DataStartRow = 10
$TemplateLastRow = 509
$StartDate = [datetime]'2026-08-01'
$EngineVersion = 'V11.1 Net New Order Quantity Excluding Included Back Orders'

function Write-Log($Text) {
    $safeText = Get-Text $Text
    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $safeText
    if($script:SuppressLog){ return }
    Write-Host $line
    try { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 } catch {}
}

function Get-Text($Value) {
    if($null -eq $Value){ return '' }
    if($Value -is [string]){ return $Value.Trim() }
    try { return $Value.ToString().Trim() } catch {}
    try { return [Convert]::ToString($Value, [Globalization.CultureInfo]::InvariantCulture).Trim() } catch {}
    return ''
}

function Release-Com($Object) {
    if($null -ne $Object){
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } catch {}
    }
}

function Get-SenderSmtp($Mail) {
    $sender = $null
    $exchangeUser = $null
    try {
        if((Get-Text $Mail.SenderEmailType) -eq 'EX'){
            $sender = $Mail.Sender
            if($null -ne $sender){
                $exchangeUser = $sender.GetExchangeUser()
                if($null -ne $exchangeUser -and (Get-Text $exchangeUser.PrimarySmtpAddress)){
                    return (Get-Text $exchangeUser.PrimarySmtpAddress).ToLowerInvariant()
                }
            }
        }
    } catch {
    } finally {
        Release-Com $exchangeUser
        Release-Com $sender
    }
    return (Get-Text $Mail.SenderEmailAddress).ToLowerInvariant()
}

function Get-CurrentBody($Body) {
    $bodyText = Get-Text $Body
    if([string]::IsNullOrWhiteSpace($bodyText)){ return '' }
    $cut = $bodyText.Length
    foreach($pattern in @(
        '(?im)^\s*-----Original Message-----',
        '(?im)^\s*From\s*:',
        '(?im)^\s*差出人\s*:',
        '(?im)^\s*送信者\s*:',
        '(?im)^\s*보낸 사람\s*:'
    )){
        $match = [regex]::Match($bodyText, $pattern)
        if($match.Success -and $match.Index -gt 20 -and $match.Index -lt $cut){ $cut = $match.Index }
    }
    return $bodyText.Substring(0, $cut).Trim()
}

function Get-MkrKeysFromText($Value) {
    $text = Get-Text $Value
    $found = New-Object System.Collections.Generic.List[string]
    if([string]::IsNullOrWhiteSpace($text)){ return @() }
    $pattern = '(?i)(?<![A-Z0-9])MKR\s*(48|49|56)\s*[/_\-]\s*26(?:\s*[/_\-]\s*1)?'
    foreach($match in [regex]::Matches($text, $pattern)){
        $endIndex = $match.Index + $match.Length
        if($endIndex -lt $text.Length -and [char]::IsDigit($text[$endIndex])){ continue }
        $number = $match.Groups[1].Value
        $hasSuffix = $match.Value -match '(?i)[/_\-]\s*1$'
        $key = 'MKR' + $number + '/26' + $(if($hasSuffix){ '-1' } else { '' })
        if($TargetMkrKeys -contains $key -and -not $found.Contains($key)){ [void]$found.Add($key) }
    }
    return @($found.ToArray())
}

function Test-TargetMkr($Value) {
    return @((Get-MkrKeysFromText $Value)).Count -gt 0
}

function Test-AnyMkrToken($Value) {
    $text = Get-Text $Value
    if(-not $text){ return $false }
    return $text -match '(?i)(?<![A-Z0-9])MKR\s*\d{1,3}\s*[/_\-]\s*\d{2}(?:\s*[/_\-]\s*\d+)?'
}

function Get-AttachmentKind($FileName) {
    $name = Get-Text $FileName
    if($name -match '(?i)sales[ _-]*note' -and $name -match '(?i)\.pdf$'){ return 'SalesNote' }
    if($name -notmatch '(?i)\.(?:xls|xlsx|xlsm)$'){ return '' }
    if($name -notmatch '(?i)(?:\(KRW\)|order|mkr)'){ return '' }
    if($name -match '(?i)^\s*rev[ _-]+'){ return 'RevisedOrder' }
    return 'OriginalOrder'
}

function Convert-ToNumber($Value, [ref]$Number) {
    if($null -eq $Value){ return $false }
    if($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or
       $Value -is [int64] -or $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]){
        $Number.Value = [double]$Value
        return $true
    }
    $clean = (Get-Text $Value).Replace(',', '')
    if($clean -notmatch '^[-+]?\d+(?:\.\d+)?$'){ return $false }
    [double]$parsed = 0
    if([double]::TryParse($clean, [Globalization.NumberStyles]::Any,
        [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)){
        $Number.Value = $parsed
        return $true
    }
    return $false
}

function Get-Code($Value) {
    $text = Get-Text $Value
    $match = [regex]::Match($text, '(?<!\d)(\d{6})(?!\d)')
    if($match.Success){ return $match.Groups[1].Value }
    [double]$numericCode = 0
    if((Convert-ToNumber $Value ([ref]$numericCode)) -and $numericCode -ge 10000 -and $numericCode -le 999999 -and $numericCode -eq [math]::Truncate($numericCode)){
        return ([int]$numericCode).ToString('000000')
    }
    return ''
}

function Test-CodeHeader($Value) {
    return (Get-Text $Value) -match '(?i)^(product\s*)?(item\s*)?(code|no\.?|number)$|product\s*code|item\s*(code|no)|商品.?コード|商品.?ｺｰﾄﾞ|商品.?cd|品番|품번'
}

function Test-NameHeader($Value) {
    return (Get-Text $Value) -match '(?i)product\s*name|item\s*name|goods\s*name|description|商品名|品名|품명|상품명|^name$'
}

function Test-QtyHeader($Value) {
    $text = Get-Text $Value
    if($text -match '(?i)price|amount|金額|単価|가격|금액'){ return $false }
    return $text -match '(?i)order\s*(qty|quantity)|q[\.\s-]?ty|(^|\s)(qty|quantity)(\s|$)|注文数|発注数|受注数|数量|수량'
}

function Test-NonItemText($Value) {
    return (Get-Text $Value) -match '(?i)^(code|name|description|qty|quantity|pce|pcs|ea)$|price|amount|total|currency|krw|thb|usd|数量|金額|単価|합계|금액'
}

function Add-ParsedItem($Map, $Code, $Name, [double]$Quantity) {
    $safeCode = Get-Text $Code
    $safeName = Get-Text $Name
    if(-not $safeCode -or -not $safeName -or $Quantity -lt 0){ return }
    if($Map.Contains($safeCode)){
        $Map[$safeCode].Quantity = [double]$Map[$safeCode].Quantity + $Quantity
        if(-not $Map[$safeCode].Name -and $safeName){ $Map[$safeCode].Name = $safeName }
    } else {
        $Map[$safeCode] = [ordered]@{ Code=$safeCode; Name=$safeName; Quantity=$Quantity }
    }
}

function Read-OrderAttachmentItems($Excel, $Path) {
    $Path = Get-Text $Path
    $book = $null
    $result = [ordered]@{}
    $usedFallback = $false
    try {
        $book = $Excel.Workbooks.Open($Path, 0, $true)
        for($sheetIndex = 1; $sheetIndex -le [int]$book.Worksheets.Count; $sheetIndex++){
            $sheet = $null
            $used = $null
            try {
                $sheet = $book.Worksheets.Item($sheetIndex)
                $used = $sheet.UsedRange
                $values = $used.Value2
                if($null -eq $values -or -not ($values -is [Array]) -or $values.Rank -ne 2){ continue }
                $rowCount = [math]::Min($values.GetLength(0), 1000)
                $columnCount = [math]::Min($values.GetLength(1), 80)
                $headerRow = 0
                $codeColumn = 0
                $nameColumn = 0
                $qtyColumn = 0

                for($r = 1; $r -le [math]::Min($rowCount, 80); $r++){
                    $candidateCode = 0
                    $candidateName = 0
                    $candidateQty = 0
                    for($c = 1; $c -le $columnCount; $c++){
                        $headerText = [regex]::Replace((Get-Text $values.GetValue($r, $c)), '\s+', ' ').Trim()
                        if(-not $candidateCode -and (Test-CodeHeader $headerText)){ $candidateCode = $c }
                        if(-not $candidateName -and (Test-NameHeader $headerText)){ $candidateName = $c }
                        if(-not $candidateQty -and (Test-QtyHeader $headerText)){ $candidateQty = $c }
                    }
                    if($candidateCode -and $candidateName -and $candidateQty){
                        $headerRow = $r
                        $codeColumn = $candidateCode
                        $nameColumn = $candidateName
                        $qtyColumn = $candidateQty
                        break
                    }
                }

                if($headerRow -gt 0){
                    $blankRun = 0
                    for($r = $headerRow + 1; $r -le $rowCount; $r++){
                        $code = Get-Code $values.GetValue($r, $codeColumn)
                        if(-not $code){
                            $blankRun++
                            if($blankRun -ge 30 -and $result.Count -gt 0){ break }
                            continue
                        }
                        $blankRun = 0
                        $name = Get-Text $values.GetValue($r, $nameColumn)
                        [double]$quantity = 0
                        if(-not $name -or -not (Convert-ToNumber $values.GetValue($r, $qtyColumn) ([ref]$quantity))){ continue }
                        Add-ParsedItem $result $code $name $quantity
                    }
                } else {
                    $usedFallback = $true
                    for($r = 1; $r -le $rowCount; $r++){
                        $foundCode = ''
                        $foundCodeColumn = 0
                        for($c = 1; $c -le $columnCount; $c++){
                            $foundCode = Get-Code $values.GetValue($r, $c)
                            if($foundCode){ $foundCodeColumn = $c; break }
                        }
                        if(-not $foundCode){ continue }
                        $name = ''
                        [double]$quantity = 0
                        $quantityFound = $false
                        for($c = $foundCodeColumn + 1; $c -le $columnCount; $c++){
                            $raw = $values.GetValue($r, $c)
                            $text = Get-Text $raw
                            if(-not $text){ continue }
                            [double]$number = 0
                            if(-not $name -and -not (Convert-ToNumber $raw ([ref]$number)) -and -not (Test-NonItemText $text)){
                                $name = $text
                                continue
                            }
                            if($name -and (Convert-ToNumber $raw ([ref]$number)) -and $number -ge 0 -and $number -le 100000000){
                                $quantity = $number
                                $quantityFound = $true
                                break
                            }
                        }
                        if($name -and $quantityFound){ Add-ParsedItem $result $foundCode $name $quantity }
                    }
                }
            } finally {
                Release-Com $used
                Release-Com $sheet
            }
        }
    } catch {
        Write-Log ('첨부 Excel 읽기 경고: ' + [IO.Path]::GetFileName($Path) + ' / ' + $_.Exception.Message)
    } finally {
        if($null -ne $book){ try { $book.Close($false) } catch {} }
        Release-Com $book
    }
    if($usedFallback){ Write-Log ('첨부 Excel 헤더를 찾지 못해 보조 규칙 사용: ' + [IO.Path]::GetFileName($Path)) }
    $items = New-Object System.Collections.Generic.List[object]
    foreach($code in $result.Keys){
        [void]$items.Add([pscustomobject]@{
            Code = Get-Text $result[$code].Code
            Name = Get-Text $result[$code].Name
            Quantity = [double]$result[$code].Quantity
        })
    }
    return @($items.ToArray())
}

function Get-BackorderSection($Text) {
    $safeText = Get-Text $Text
    if([string]::IsNullOrWhiteSpace($safeText)){ return '' }
    $marker = [regex]::Match($safeText, '(?i)バック\s*オーダー|Back\s*[- ]?\s*Order(?:\s*QTY)?|欠品時\s*(?:No\.?|番号)')
    if(-not $marker.Success){ return '' }
    return $safeText.Substring($marker.Index)
}

function Get-TextBeforeBackorderSection($Text) {
    $safeText = Get-Text $Text
    if([string]::IsNullOrWhiteSpace($safeText)){ return '' }
    $marker = [regex]::Match($safeText, '(?i)バック\s*オーダー|Back\s*[- ]?\s*Order(?:\s*QTY)?|欠品時\s*(?:No\.?|番号)')
    if(-not $marker.Success){ return $safeText }
    return $safeText.Substring(0, $marker.Index).Trim()
}

function Convert-ToCanonicalMkr($Value) {
    $match = [regex]::Match((Get-Text $Value), '(?i)MKR\s*(\d{1,3})\s*[/_\-]\s*(\d{2})(?:\s*[/_\-]\s*(\d+))?')
    if(-not $match.Success){ return '' }
    $suffix = if($match.Groups[3].Success){ '-' + $match.Groups[3].Value } else { '' }
    return 'MKR' + $match.Groups[1].Value + '/' + $match.Groups[2].Value + $suffix
}

function Convert-BackorderRow($Value) {
    $text = [regex]::Replace((Get-Text $Value), '\s+', ' ').Trim()
    if(-not $text){ return $null }
    $quantityPattern = '[+-]?\d[\d,]*(?:\.\d+)?'
    $rowPattern = ('(?i)^\s*(?<source>MKR\s*\d{{1,3}}\s*[/_\-]\s*\d{{2}}(?:\s*[/_\-]\s*\d+)?)\s+(?<code>\d{{6}})\s+(?<name>.+?)\s+(?<order>{0})\s+(?<shipped>{0})(?:\s+(?<remaining>[-－―—]|{0}))?\s*$' -f $quantityPattern)
    $match = [regex]::Match($text, $rowPattern)
    if(-not $match.Success){ return $null }

    [double]$orderQuantity = 0
    [double]$shippedQuantity = 0
    if(-not (Convert-ToNumber $match.Groups['order'].Value ([ref]$orderQuantity))){ return $null }
    if(-not (Convert-ToNumber $match.Groups['shipped'].Value ([ref]$shippedQuantity))){ return $null }
    if($orderQuantity -lt 0 -or $shippedQuantity -lt 0){ return $null }

    $remainingQuantity = [math]::Max(0, $orderQuantity - $shippedQuantity)
    $remainingText = Get-Text $match.Groups['remaining'].Value
    if($remainingText -and $remainingText -notmatch '^[-－―—]$'){
        [double]$parsedRemaining = 0
        if(Convert-ToNumber $remainingText ([ref]$parsedRemaining)){ $remainingQuantity = [math]::Max(0, $parsedRemaining) }
    }
    return [pscustomobject]@{
        SourceMkr = Convert-ToCanonicalMkr $match.Groups['source'].Value
        Code = Get-Text $match.Groups['code'].Value
        Name = Get-Text $match.Groups['name'].Value
        OrderQuantity = [double]$orderQuantity
        Quantity = [double]$shippedQuantity
        Remaining = [double]$remainingQuantity
    }
}

function Parse-BackorderText($Text) {
    $section = Get-BackorderSection $Text
    $result = [ordered]@{}
    if([string]::IsNullOrWhiteSpace($section)){ return @() }
    # Word's PDF converter separates table cells with CR/BEL instead of CRLF.
    # Treat all Word/PDF control separators as line breaks so table rows are not
    # returned as one long string and silently parsed as zero items.
    $lines = @($section -split '[\r\n\a\v\f]+' | ForEach-Object { [regex]::Replace($_.Trim(), '\s+', ' ') } | Where-Object { $_ })

    foreach($line in $lines){
        $item = Convert-BackorderRow $line
        if($null -eq $item){ continue }
        $key = (Get-Text $item.SourceMkr) + '|' + (Get-Text $item.Code)
        $result[$key] = $item
    }

    # Outlook/Word may expose each table cell on a separate line. Join only the
    # cells that belong to the same source-MKR row and parse the row again.
    for($i = 0; $i -lt $lines.Count; $i++){
        if($lines[$i] -notmatch '(?i)^\s*MKR\s*\d{1,3}\s*[/_\-]\s*\d{2}(?:\s*[/_\-]\s*\d+)?'){ continue }
        $parts = New-Object System.Collections.Generic.List[string]
        [void]$parts.Add($lines[$i])
        for($j = $i + 1; $j -lt [math]::Min($lines.Count, $i + 12); $j++){
            if($lines[$j] -match '(?i)^\s*MKR\s*\d{1,3}\s*[/_\-]\s*\d{2}(?:\s*[/_\-]\s*\d+)?'){ break }
            [void]$parts.Add($lines[$j])
            $item = Convert-BackorderRow ($parts -join ' ')
            if($null -ne $item){
                $key = (Get-Text $item.SourceMkr) + '|' + (Get-Text $item.Code)
                $result[$key] = $item
                break
            }
        }
    }
    return @($result.Values)
}

function Merge-BackorderLedger($Ledger, $Items, $SourceLabel, [datetime]$Received) {
    $merged = 0
    foreach($item in @($Items)){
        $sourceMkr = Convert-ToCanonicalMkr $item.SourceMkr
        $code = Get-Text $item.Code
        $name = Get-Text $item.Name
        if(-not $sourceMkr -or -not $code -or -not $name){ continue }
        $key = $sourceMkr + '|' + $code
        # Later mail/Sales Note rows replace an earlier copy of the same
        # source-MKR + item. This prevents quoted reply chains from double-counting.
        $Ledger[$key] = [pscustomobject]@{
            SourceMkr=$sourceMkr; Code=$code; Name=$name
            OrderQuantity=[double]$item.OrderQuantity
            Quantity=[double]$item.Quantity
            Remaining=[double]$item.Remaining
            Received=[datetime]$Received
            SourceLabel=(Get-Text $SourceLabel)
        }
        $merged++
    }
    if($merged -gt 0){ Write-Log ((Get-Text $SourceLabel) + ' 백오더 확인: ' + $merged + '개 행') }
}

function Merge-SalesNoteOrderEvidence($Map, $Items, $SourceLabel) {
    $merged = 0
    foreach($item in @($Items)){
        $code = Get-Text $item.Code
        if(-not $code -or $Map.Contains($code)){ continue }
        $orderProperty = $item.PSObject.Properties['OrderQuantity']
        if($null -eq $orderProperty -or $null -eq $item.OrderQuantity){ continue }
        [double]$orderQuantity = 0
        if(-not (Convert-ToNumber $item.OrderQuantity ([ref]$orderQuantity)) -or $orderQuantity -le 0){ continue }
        $Map[$code] = [double]$orderQuantity
        $merged++
    }
    if($merged -gt 0){ Write-Log ((Get-Text $SourceLabel) + ' 신규 주문수량 근거 확보: ' + $merged + '개 품목') }
}

function Apply-NetOrderExcludingBackorders($State, $SalesNoteOrderMap, $Ledger, $SourceLabel) {
    $tolerance = [double]0.0000001
    $backorderIncludedByCode = [ordered]@{}
    foreach($key in @($Ledger.Keys)){
        $item = $Ledger[$key]
        $code = Get-Text $item.Code
        if(-not $code){ continue }
        if(-not $backorderIncludedByCode.Contains($code)){ $backorderIncludedByCode[$code] = [double]0 }
        # Quantity is the backorder quantity actually shipped in this MKR (H:M Shipped QTY).
        # OrderQuantity is the historical backorder balance and must not be subtracted here.
        $backorderIncludedByCode[$code] = [double]$backorderIncludedByCode[$code] + [double]$item.Quantity
    }

    $excluded = 0
    $alreadyNet = 0
    $mismatched = 0
    foreach($code in @($State.Keys | Sort-Object)){
        $record = $State[$code]
        $rawOrderQty = if($record.Contains('RawOrderQty')){ [double]$record.RawOrderQty } else { [double]$record.OrderQty }
        $record.OrderQty = $rawOrderQty
        $backorderIncluded = if($backorderIncludedByCode.Contains($code)){ [double]$backorderIncludedByCode[$code] } else { [double]0 }
        if($backorderIncluded -le 0 -or -not $SalesNoteOrderMap.Contains($code)){ continue }

        $salesNoteOrderQty = [double]$SalesNoteOrderMap[$code]
        if([math]::Abs($rawOrderQty - ($salesNoteOrderQty + $backorderIncluded)) -le $tolerance){
            $record.OrderQty = [double]$salesNoteOrderQty
            $excluded++
            Write-Log ((Get-Text $SourceLabel) + ' 백오더 제외: ' + $code + ' / 원주문 ' + (Convert-ToNumberText $rawOrderQty) + ' - 포함 백오더 ' + (Convert-ToNumberText $backorderIncluded) + ' = 신규 주문 ' + (Convert-ToNumberText $salesNoteOrderQty))
        } elseif([math]::Abs($rawOrderQty - $salesNoteOrderQty) -le $tolerance){
            # The source order already excludes the backorder. Never subtract twice.
            $alreadyNet++
            Write-Log ((Get-Text $SourceLabel) + ' 백오더 기제외 확인: ' + $code + ' / 신규 주문 ' + (Convert-ToNumberText $rawOrderQty))
        } else {
            # Evidence does not reconcile, so preserve the first order instead of guessing.
            $mismatched++
            Write-Log ((Get-Text $SourceLabel) + ' 백오더 제외 판정 경고: ' + $code + ' / 원주문 ' + (Convert-ToNumberText $rawOrderQty) + ' / Sales Note 주문 ' + (Convert-ToNumberText $salesNoteOrderQty) + ' / 포함 백오더 ' + (Convert-ToNumberText $backorderIncluded) + ' / 원주문 유지')
        }
    }
    Write-Log ((Get-Text $SourceLabel) + ' 신규 Order QTY 판정: 백오더 제외 ' + $excluded + '개 / 이미 제외 ' + $alreadyNet + '개 / 불일치 유지 ' + $mismatched + '개')
}

function Apply-BackorderLedger($State, $Ledger, $SourceLabel) {
    $applied = 0
    foreach($key in @($Ledger.Keys | Sort-Object)){
        $item = $Ledger[$key]
        $code = Get-Text $item.Code
        if(-not $code){ continue }
        if(-not $State.Contains($code)){
            $State[$code] = [ordered]@{
                Code=$code; Name=(Get-Text $item.Name); OrderQty=[double]0; ShippedQty=[double]0
            }
        }
        $State[$code].OrderQty = [double]$State[$code].OrderQty + [double]$item.OrderQuantity
        $State[$code].ShippedQty = [double]$State[$code].ShippedQty + [double]$item.Quantity
        if(-not (Get-Text $State[$code].Name)){ $State[$code].Name = Get-Text $item.Name }
        $calculatedRemaining = [math]::Max(0, [double]$item.OrderQuantity - [double]$item.Quantity)
        if([math]::Abs($calculatedRemaining - [double]$item.Remaining) -gt 0.0000001){
            Write-Log ('백오더 잔량 검증 경고: ' + (Get-Text $item.SourceMkr) + ' / ' + $code + ' / 표기 ' + (Convert-ToNumberText $item.Remaining) + ' / 계산 ' + (Convert-ToNumberText $calculatedRemaining))
        }
        $applied++
    }
    if($applied -gt 0){ Write-Log ((Get-Text $SourceLabel) + ' 백오더 누적 반영: ' + $applied + '개 원본 행') }
}

function Parse-UpdateText($Text) {
    $Text = Get-TextBeforeBackorderSection $Text
    $result = [ordered]@{}
    if([string]::IsNullOrWhiteSpace($Text)){ return @() }
    $lines = @($Text -split '[\r\n\a\v\f]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

    foreach($line in $lines){
        $clean = [regex]::Replace($line, '(?i)\b(PCE|PCS|EA)\b', ' ')
        $codeMatch = [regex]::Match($clean, '^\s*(\d{6})\s+(.+)$')
        if(-not $codeMatch.Success){ continue }
        $code = $codeMatch.Groups[1].Value
        $tail = $codeMatch.Groups[2].Value.Trim()
        $tokens = @($tail -split '\s+' | Where-Object { $_ })
        $nameTokens = New-Object System.Collections.Generic.List[string]
        $numbers = New-Object System.Collections.Generic.List[double]
        $numericStarted = $false
        foreach($token in $tokens){
            [double]$number = 0
            if(Convert-ToNumber $token ([ref]$number)){
                $numericStarted = $true
                [void]$numbers.Add($number)
            } elseif(-not $numericStarted){
                [void]$nameTokens.Add($token)
            }
        }
        if($nameTokens.Count -eq 0 -or $numbers.Count -eq 0){ continue }
        $order = if($numbers.Count -ge 2){ [double]$numbers[0] } else { $null }
        $shipped = if($numbers.Count -ge 2){ $numbers[1] } else { $numbers[0] }
        $result[$code] = [pscustomobject]@{ Code=$code; Name=($nameTokens -join ' '); OrderQuantity=$order; Quantity=[double]$shipped }
    }

    for($i = 0; $i -lt $lines.Count; $i++){
        if($lines[$i] -notmatch '^\d{6}$'){ continue }
        $code = $lines[$i]
        if($result.Contains($code)){ continue }
        $name = ''
        $numbers = New-Object System.Collections.Generic.List[double]
        for($j = $i + 1; $j -lt [math]::Min($lines.Count, $i + 16); $j++){
            $value = $lines[$j]
            if($value -match '^\d{6}$'){ break }
            if((Test-CodeHeader $value) -or (Test-NameHeader $value) -or (Test-QtyHeader $value)){ continue }
            [double]$number = 0
            if(-not $name -and -not (Convert-ToNumber $value ([ref]$number)) -and -not (Test-NonItemText $value)){
                $name = $value
                continue
            }
            if($name -and (Convert-ToNumber $value ([ref]$number))){ [void]$numbers.Add($number) }
        }
        if(-not $name -or $numbers.Count -eq 0){ continue }
        $order = if($numbers.Count -ge 2){ [double]$numbers[0] } else { $null }
        $shipped = if($numbers.Count -ge 2){ $numbers[1] } else { $numbers[0] }
        $result[$code] = [pscustomobject]@{ Code=$code; Name=$name; OrderQuantity=$order; Quantity=[double]$shipped }
    }
    return @($result.Values)
}

function Read-PdfTextWithWord($Path) {
    $Path = Get-Text $Path
    $word = $null
    $document = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $document = $word.Documents.Open($Path, $false, $true, $false)
        return Get-Text $document.Content.Text
    } catch {
        Write-Log ('Sales Note PDF 텍스트 읽기 경고: ' + [IO.Path]::GetFileName($Path) + ' / ' + $_.Exception.Message)
        return ''
    } finally {
        if($null -ne $document){ try { $document.Close($false) } catch {} }
        if($null -ne $word){ try { $word.Quit() } catch {} }
        Release-Com $document
        Release-Com $word
    }
}

function Apply-OrderUpdates($State, $Items, $SourceLabel) {
    $SourceLabel = Get-Text $SourceLabel
    $applied = 0
    foreach($item in @($Items)){
        $code = Get-Text $item.Code
        $name = Get-Text $item.Name
        [double]$orderQuantity = 0
        if(-not $code -or -not $name -or -not (Convert-ToNumber $item.Quantity ([ref]$orderQuantity))){ continue }
        if(-not $State.Contains($code)){
            $State[$code] = [ordered]@{ Code=$code; Name=$name; OrderQty=[double]0; ShippedQty=[double]0 }
        }
        $State[$code].OrderQty = [double]$orderQuantity
        $State[$code].Name = $name
        # A revised order changes OrderQty only.  The initial full-shipment
        # fallback remains in place until a mail table or Sales Note supplies
        # an explicit shipped quantity for this item.
        $applied++
    }
    if($applied -gt 0){ Write-Log ($SourceLabel + ' 주문수량 반영: ' + $applied + '개 품목') }
}

function Apply-ShippedUpdates($State, $Items, $SourceLabel, [switch]$PreserveOrderQuantity, [switch]$AllowNewItems) {
    $SourceLabel = Get-Text $SourceLabel
    $applied = 0
    foreach($item in @($Items)){
        $code = Get-Text $item.Code
        if(-not $State.Contains($code)){
            if(-not $AllowNewItems){
                Write-Log ('기준 주문에 없는 품번은 제외: ' + $code + ' / ' + $SourceLabel)
                continue
            }
            $newOrderQty = [double]$item.Quantity
            $newOrderProperty = $item.PSObject.Properties['OrderQuantity']
            if($null -ne $newOrderProperty -and $null -ne $item.OrderQuantity -and [double]$item.OrderQuantity -gt 0){ $newOrderQty = [double]$item.OrderQuantity }
            $State[$code] = [ordered]@{ Code=$code; Name=(Get-Text $item.Name); OrderQty=$newOrderQty; ShippedQty=[double]0 }
        }
        $orderProperty = $item.PSObject.Properties['OrderQuantity']
        if(-not $PreserveOrderQuantity -and $null -ne $orderProperty -and $null -ne $item.OrderQuantity -and [double]$item.OrderQuantity -gt 0){
            $State[$code].OrderQty = [double]$item.OrderQuantity
        }
        $State[$code].ShippedQty = [double]$item.Quantity
        if($item.Name){ $State[$code].Name = Get-Text $item.Name }
        $applied++
    }
    if($applied -gt 0){ Write-Log ($SourceLabel + ' 수량 반영: ' + $applied + '개 품목') }
}

function Get-DateFromBody($Body, $Keyword, [datetime]$ReferenceDate) {
    try {
        $bodyText = Get-Text $Body
        $keywordText = Get-Text $Keyword
        $reference = [datetime]$ReferenceDate
        foreach($rawLine in ($bodyText -split "`r?`n")){
            $line = Get-Text $rawLine
            if(-not $line -or $line -notmatch $keywordText){ continue }
            $keywordMatch = [regex]::Match($line, $keywordText)
            $dateText = $line
            if($keywordMatch.Success -and ($keywordMatch.Index + $keywordMatch.Length) -lt $line.Length){
                $dateText = $line.Substring($keywordMatch.Index + $keywordMatch.Length)
            }
            $numeric = [regex]::Matches($dateText, '(?<!\d)(20\d{2})\s*[./\-年]\s*(\d{1,2})\s*[./\-月]\s*(\d{1,2})(?:일)?')
            if($numeric.Count -gt 0){
                $m = $numeric[$numeric.Count - 1]
                $year = [int]$m.Groups[1].Value
                try { return Get-Date -Year $year -Month ([int]$m.Groups[2].Value) -Day ([int]$m.Groups[3].Value) -Hour 0 -Minute 0 -Second 0 } catch {}
            }
            $slashDate = [regex]::Matches($dateText, '(?<!\d)(\d{1,2})\s*[./\-]\s*(\d{1,2})\s*[./\-]\s*(20\d{2})(?!\d)')
            if($slashDate.Count -gt 0){
                $m = $slashDate[$slashDate.Count - 1]
                try { return Get-Date -Year ([int]$m.Groups[3].Value) -Month ([int]$m.Groups[1].Value) -Day ([int]$m.Groups[2].Value) -Hour 0 -Minute 0 -Second 0 } catch {}
            }
            $english = [regex]::Matches($dateText, '(?i)(\d{1,2})\s*[- ]\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*[- ,]\s*(20\d{2})|(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2}),?\s+(20\d{2})')
            if($english.Count -gt 0){
                $m = $english[$english.Count - 1]
                $monthMap = @{Jan=1;Feb=2;Mar=3;Apr=4;May=5;Jun=6;Jul=7;Aug=8;Sep=9;Oct=10;Nov=11;Dec=12}
                try {
                    if($m.Groups[1].Success){
                        return Get-Date -Year ([int]$m.Groups[3].Value) -Month $monthMap[$m.Groups[2].Value.Substring(0,3)] -Day ([int]$m.Groups[1].Value) -Hour 0 -Minute 0 -Second 0
                    }
                    return Get-Date -Year ([int]$m.Groups[6].Value) -Month $monthMap[$m.Groups[4].Value.Substring(0,3)] -Day ([int]$m.Groups[5].Value) -Hour 0 -Minute 0 -Second 0
                } catch {}
            }
            $fallback = [regex]::Matches($line, '(?<!\d)(\d{1,2})\s*[./\-]\s*(\d{1,2})(?!\d)')
            if($fallback.Count -gt 0){
                $m = $fallback[$fallback.Count - 1]
                try { return Get-Date -Year $reference.Year -Month ([int]$m.Groups[1].Value) -Day ([int]$m.Groups[2].Value) -Hour 0 -Minute 0 -Second 0 } catch {}
            }
        }
    } catch {
        try { Write-Log ('날짜 분석 경고: ' + (Get-Text $_.Exception.Message)) } catch {}
    }
    return $null
}

function Set-ExcelDateCell($Sheet, $Address, [object]$Value) {
    $Address = [string](Get-Text $Address)
    if($Value -isnot [datetime]){ return }
    $cell = $null
    try {
        $cell = $Sheet.Range($Address)
        # Some Office/PowerShell 5.1 combinations fail while marshaling a
        # Double OLE date into Excel.  ISO text is unambiguous and avoids the
        # same Double -> String COM conversion failure seen in the run log.
        $cell.NumberFormat = '@'
        $cell.Value2 = [string](([datetime]$Value).ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture))
    } finally {
        Release-Com $cell
    }
}

function Get-OutputLastRow($Sheet, [int]$MinimumLastRow, [int]$RequiredLastRow) {
    $used = $null
    try {
        $lastUsedRow = 0
        $used = $Sheet.UsedRange
        if($null -ne $used){ $lastUsedRow = [int]$used.Row + [int]$used.Rows.Count - 1 }
        return [math]::Max($MinimumLastRow, [math]::Max($RequiredLastRow, $lastUsedRow))
    } catch {
        return [math]::Max($MinimumLastRow, $RequiredLastRow)
    } finally {
        Release-Com $used
    }
}

function Ensure-OutputFormatRows($ExcelApp, $Sheet, [int]$LastRow) {
    if($LastRow -le $TemplateLastRow){ return }
    # Do not use the clipboard/PasteSpecial path here: Excel COM can expose numeric
    # optional arguments as System.Double and abort the whole transaction. Data rows
    # are still written beyond the template area; the existing template formatting
    # remains untouched.
    Write-Log ('입력 행 자동 확장: ' + $LastRow + '행까지 기록합니다.')
}

function Convert-ToNumberText($Value) {
    [double]$number = 0
    if(-not (Convert-ToNumber $Value ([ref]$number))){ return '' }
    if([math]::Abs($number - [math]::Round($number)) -lt 0.0000001){
        return [string](([long][math]::Round($number)).ToString([Globalization.CultureInfo]::InvariantCulture))
    }
    return [string]($number.ToString('0.##########', [Globalization.CultureInfo]::InvariantCulture))
}

function Write-ExcelItemsBulk($Sheet, [int]$StartRow, $State) {
    $dataRange = $null
    $textRange = $null
    $numberRange = $null
    $stage = '입력값 준비'
    try {
        $codes = @($State.Keys | Sort-Object)
        $count = [int]$codes.Count
        if($count -le 0){ return 0 }

        # Every element crossing the PowerShell 5.1 -> Excel COM boundary is
        # deliberately a String.  Range.Formula converts the quantity strings
        # to Excel numeric constants while item codes remain text.
        $textValues = New-Object 'object[,]' $count, 2
        $numberValues = New-Object 'object[,]' $count, 3
        for($i = 0; $i -lt $count; $i++){
            $record = $State[$codes[$i]]
            $code = [string](Get-Text $record.Code)
            $name = [string](Get-Text $record.Name)
            [double]$orderQty = 0
            [double]$shippedQty = 0
            if(-not $code -or -not $name){ throw ('필수 품목정보 누락: 순번 ' + ($i + 1)) }
            if(-not (Convert-ToNumber $record.OrderQty ([ref]$orderQty))){ throw ('주문수량 숫자 변환 실패: ' + $code) }
            if(-not (Convert-ToNumber $record.ShippedQty ([ref]$shippedQty))){ throw ('출하수량 숫자 변환 실패: ' + $code) }
            $shortage = [math]::Max(0, $orderQty - $shippedQty)
            $textValues[$i, 0] = [string]$code
            $textValues[$i, 1] = [string]$name
            $numberValues[$i, 0] = [string](Convert-ToNumberText $orderQty)
            $numberValues[$i, 1] = [string](Convert-ToNumberText $shippedQty)
            $numberValues[$i, 2] = [string](Convert-ToNumberText $shortage)
        }

        $lastRow = [int]($StartRow + $count - 1)
        $textAddress = [string]('A{0}:B{1}' -f $StartRow, $lastRow)
        $numberAddress = [string]('C{0}:E{1}' -f $StartRow, $lastRow)
        $dataAddress = [string]('A{0}:E{1}' -f $StartRow, $lastRow)

        $stage = 'Excel 범위 확인'
        $dataRange = $Sheet.Range($dataAddress)
        $textRange = $Sheet.Range($textAddress)
        $numberRange = $Sheet.Range($numberAddress)

        $stage = '품번·품명 일괄 기록'
        $textRange.NumberFormat = '@'
        $textRange.Value2 = $textValues

        $stage = '수량 일괄 기록'
        $numberRange.NumberFormat = '0'
        $numberRange.Formula = $numberValues
        return $count
    } catch {
        Write-Log ('Excel 일괄 반영 오류 [' + $stage + ']: ' + (Get-Text $_.Exception.Message))
        if(Get-Text $_.ScriptStackTrace){ Write-Log ('오류 위치: ' + (Get-Text $_.ScriptStackTrace)) }
        throw
    } finally {
        Release-Com $numberRange
        Release-Com $textRange
        Release-Com $dataRange
    }
}

function Write-ExcelBackordersBulk($Sheet, [int]$StartRow, $Ledger) {
    $dataRange = $null
    $textRange = $null
    $numberRange = $null
    $stage = '백오더 입력값 준비'
    try {
        $keys = @($Ledger.Keys | Sort-Object)
        $count = [int]$keys.Count
        if($count -le 0){ return 0 }

        $textValues = New-Object 'object[,]' $count, 3
        $numberValues = New-Object 'object[,]' $count, 3
        for($i = 0; $i -lt $count; $i++){
            $record = $Ledger[$keys[$i]]
            $sourceMkr = [string](Convert-ToCanonicalMkr $record.SourceMkr)
            $code = [string](Get-Text $record.Code)
            $name = [string](Get-Text $record.Name)
            [double]$orderQty = 0
            [double]$shippedQty = 0
            if(-not $sourceMkr -or -not $code -or -not $name){ throw ('백오더 필수정보 누락: 순번 ' + ($i + 1)) }
            if(-not (Convert-ToNumber $record.OrderQuantity ([ref]$orderQty))){ throw ('백오더 주문수량 숫자 변환 실패: ' + $sourceMkr + ' / ' + $code) }
            if(-not (Convert-ToNumber $record.Quantity ([ref]$shippedQty))){ throw ('백오더 선적수량 숫자 변환 실패: ' + $sourceMkr + ' / ' + $code) }
            $remaining = [math]::Max(0, $orderQty - $shippedQty)
            $textValues[$i, 0] = $sourceMkr
            $textValues[$i, 1] = $code
            $textValues[$i, 2] = $name
            $numberValues[$i, 0] = [string](Convert-ToNumberText $orderQty)
            $numberValues[$i, 1] = [string](Convert-ToNumberText $shippedQty)
            $numberValues[$i, 2] = [string](Convert-ToNumberText $remaining)
        }

        $lastRow = [int]($StartRow + $count - 1)
        $textAddress = [string]('H{0}:J{1}' -f $StartRow, $lastRow)
        $numberAddress = [string]('K{0}:M{1}' -f $StartRow, $lastRow)
        $dataAddress = [string]('H{0}:M{1}' -f $StartRow, $lastRow)

        $stage = '백오더 Excel 범위 확인'
        $dataRange = $Sheet.Range($dataAddress)
        $textRange = $Sheet.Range($textAddress)
        $numberRange = $Sheet.Range($numberAddress)

        $stage = '백오더 번호·품번·품명 일괄 기록'
        $textRange.NumberFormat = '@'
        $textRange.Value2 = $textValues

        $stage = '백오더 수량 일괄 기록'
        $numberRange.NumberFormat = '0'
        $numberRange.Formula = $numberValues
        return $count
    } catch {
        Write-Log ('백오더 Excel 일괄 반영 오류 [' + $stage + ']: ' + (Get-Text $_.Exception.Message))
        if(Get-Text $_.ScriptStackTrace){ Write-Log ('오류 위치: ' + (Get-Text $_.ScriptStackTrace)) }
        throw
    } finally {
        Release-Com $numberRange
        Release-Com $textRange
        Release-Com $dataRange
    }
}

if($SelfTest){
    $script:SuppressLog = $true
    try {
        $fixture = @'
なお、下記のバックオーダーをMKR56/26へ追加しております。
欠品時No. Code Name Back Order QTY Shipped QTY Back Order Remaining
MKR48/26 760010 ORDEVE The Professional Haircolor 9,1L 4,680 1,032 3,648
'@
        $parsedBackorders = @(Parse-BackorderText $fixture)
        if($parsedBackorders.Count -ne 1){ throw ('백오더 행 분석 수 불일치: ' + $parsedBackorders.Count + ' / 예상 1') }
        $testLedger = [ordered]@{}
        Merge-BackorderLedger $testLedger $parsedBackorders 'SelfTest 1' ([datetime]'2026-08-17')
        Merge-BackorderLedger $testLedger $parsedBackorders 'SelfTest repeated reply' ([datetime]'2026-08-18')
        if($testLedger.Count -ne 1){ throw ('백오더 중복 제거 실패: ' + $testLedger.Count + ' / 예상 1') }
        $backorderRow = $testLedger['MKR48/26|760010']
        if($null -eq $backorderRow -or [double]$backorderRow.OrderQuantity -ne 4680 -or [double]$backorderRow.Quantity -ne 1032 -or [double]$backorderRow.Remaining -ne 3648){
            throw 'MKR56/26 백오더 분리 실패'
        }

        $testState = [ordered]@{
            '760010' = [ordered]@{ Code='760010'; Name='ORDEVE The Professional Haircolor 9,1L'; RawOrderQty=[double]4800; OrderQty=[double]4800; ShippedQty=[double]4800 }
            '762016' = [ordered]@{ Code='762016'; Name='SelfTest unchanged item'; RawOrderQty=[double]3600; OrderQty=[double]3600; ShippedQty=[double]3600 }
            '417134' = [ordered]@{ Code='417134'; Name='SelfTest changed item'; RawOrderQty=[double]18000; OrderQty=[double]18000; ShippedQty=[double]18000 }
        }
        $wordPdfFixture = "760010`r`aORDEVE The Professional Haircolor 9,1L`r`a4,800`r`a120`r`a4,680`r`a762016`r`aSelfTest unchanged item`r`a3,600`r`a3,600`r`a0`r`a417134`r`aSelfTest changed item`r`a18,000`r`a17,100`r`a900"
        $wordPdfItems = @(Parse-UpdateText $wordPdfFixture)
        if($wordPdfItems.Count -ne 3){ throw ('Word PDF 표 구분문자 분석 실패: ' + $wordPdfItems.Count + ' / 예상 3') }
        Apply-ShippedUpdates $testState $wordPdfItems 'SelfTest first Sales Note snapshot' -PreserveOrderQuantity
        $mainShortage = [math]::Max(0, [double]$testState['760010'].OrderQty - [double]$testState['760010'].ShippedQty)
        if([double]$testState['760010'].OrderQty -ne 4800 -or [double]$testState['760010'].ShippedQty -ne 120 -or $mainShortage -ne 4680){
            throw 'MKR48/26 최초 주문·최초 Sales Note 보존 실패'
        }
        if([double]$testState['762016'].ShippedQty -ne 3600){ throw '762016 최초 Sales Note 선적수량 실패' }
        $changedShortage = [math]::Max(0, [double]$testState['417134'].OrderQty - [double]$testState['417134'].ShippedQty)
        if([double]$testState['417134'].ShippedQty -ne 17100 -or $changedShortage -ne 900){ throw '417134 최초 Sales Note 차액 계산 실패' }

        # A later MKR backorder shipment is displayed only in H:M of the
        # destination MKR. It must never change the original MKR snapshot.
        if([double]$testState['760010'].ShippedQty -ne 120 -or $mainShortage -ne 4680){ throw '후속 MKR 백오더가 원본 MKR 수량을 변경했습니다.' }

        $netOrderEvidence = [ordered]@{ '760010' = [double]4800 }
        $rawIncludesBackorder = [ordered]@{
            '760010' = [ordered]@{ Code='760010'; Name='Net order test'; RawOrderQty=[double]5832; OrderQty=[double]5832; ShippedQty=[double]4800 }
        }
        Apply-NetOrderExcludingBackorders $rawIncludesBackorder $netOrderEvidence $testLedger 'SelfTest raw order includes backorder'
        if([double]$rawIncludesBackorder['760010'].OrderQty -ne 4800){ throw '원주문 5,832에서 포함 백오더 1,032 제외 실패' }

        $rawAlreadyExcludesBackorder = [ordered]@{
            '760010' = [ordered]@{ Code='760010'; Name='No double deduction test'; RawOrderQty=[double]4800; OrderQty=[double]4800; ShippedQty=[double]4800 }
        }
        Apply-NetOrderExcludingBackorders $rawAlreadyExcludesBackorder $netOrderEvidence $testLedger 'SelfTest raw order already excludes backorder'
        if([double]$rawAlreadyExcludesBackorder['760010'].OrderQty -ne 4800){ throw '이미 백오더를 제외한 신규 주문수량이 이중 차감되었습니다.' }

        Write-Host '[OK] Net new order, no double deduction, separate backorder, and MKR56 self-test passed.' -ForegroundColor Green
        exit 0
    } catch {
        Write-Host ('[ERROR] Backorder self-test failed: ' + (Get-Text $_.Exception.Message)) -ForegroundColor Red
        exit 3
    }
}

try {
Write-Log ($EngineVersion + ' 시작: 주문/선적/부족수량 분리 엔진')
if(-not (Test-Path -LiteralPath $WorkbookPath)){ throw "대상 Excel 파일이 없습니다: $WorkbookPath" }
if(Test-Path -LiteralPath (Join-Path $Root ('~$' + [IO.Path]::GetFileName($WorkbookPath)))){
    throw '대상 Excel 파일이 열려 있습니다. Excel을 닫고 다시 실행해 주세요.'
}

New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
New-Item -ItemType Directory -Path $RunAttachmentDir -Force | Out-Null
$backupPath = Join-Path $BackupDir ('MKR_Multi_BlankOrPrevious_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsx')
Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath -Force
Write-Log ('실행 전 백업 완료: ' + $backupPath)

$mailRecords = New-Object System.Collections.Generic.List[object]
$attachmentEvents = New-Object System.Collections.Generic.List[object]
$outlook = $null
$namespace = $null
$inbox = $null
$items = $null
$scanned = 0
$targetMessages = 0
$savedAttachments = 0

try {
    Write-Log 'Classic Outlook 연결 확인 중...'
    try { $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
    catch { throw '실행 중인 Classic Outlook을 찾지 못했습니다. Classic Outlook을 먼저 열어 주세요.' }
    $namespace = $outlook.Session
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    try {
        $utcStart = $StartDate.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
        $jetStart = $StartDate.ToString('g', [Globalization.CultureInfo]::CurrentCulture)
        $unfilteredCount = [int]$items.Count
        $filterQueries = @(
            ("[ReceivedTime] >= '$jetStart' AND [SenderName] = 'Nagasawa Hikari (長澤 光里)'"),
            ("[ReceivedTime] >= '$jetStart' AND [SenderName] = 'Nagasawa Hikari'"),
            ("@SQL=(`"urn:schemas:httpmail:datereceived`" >= '{0}') AND (`"urn:schemas:httpmail:subject`" LIKE '%MKR%26%')" -f $utcStart),
            ("@SQL=(`"urn:schemas:httpmail:datereceived`" >= '{0}') AND (`"http://schemas.microsoft.com/mapi/proptag/0x0C1A001F`" LIKE '%Nagasawa%')" -f $utcStart)
        )
        $filtered = $null
        $bestFilteredCount = [int]::MaxValue
        foreach($filterQuery in $filterQueries){
            $candidate = $null
            try {
                $candidate = $items.Restrict($filterQuery)
                $candidateCount = if($null -ne $candidate){ [int]$candidate.Count } else { 0 }
                if($candidateCount -gt 0 -and $candidateCount -lt $unfilteredCount -and $candidateCount -lt $bestFilteredCount){
                    Release-Com $filtered
                    $filtered = $candidate
                    $bestFilteredCount = $candidateCount
                    $candidate = $null
                }
            } catch {
                Write-Log ('메일 사전필터 방식 전환: ' + (Get-Text $_.Exception.Message))
            } finally {
                Release-Com $candidate
            }
        }
        if($null -ne $filtered){
            Release-Com $items
            $items = $filtered
            Write-Log ('날짜·발신자/MKR 사전필터 결과: ' + [int]$items.Count + '건')
        } else {
            Write-Log '사전필터가 메일 범위를 줄이지 못해 받은편지함 최신 500건 안에서 복구 모드로 확인합니다.'
        }
    } catch {
        Write-Log ('사전필터 실패. 받은편지함 최신 500건 안에서만 확인: ' + (Get-Text $_.Exception.Message))
    }
    $items.Sort('[ReceivedTime]', $true)
    $scanCount = [math]::Min([int]$items.Count, $MaxMessages)
    Write-Log ('검색 시작: 최대 ' + $scanCount + '건 / 절대 상한 ' + $MaxMessages + '건')

    for($index = $scanCount; $index -ge 1; $index--){
        $mail = $null
        try {
            $mail = $items.Item($index)
            $scanned++
            if($null -eq $mail -or [int]$mail.Class -ne 43){ continue }
            [datetime]$received = [datetime]::MinValue
            try { $received = [datetime]$mail.ReceivedTime } catch { continue }
            if($received -lt $StartDate){ continue }
            if((Get-SenderSmtp $mail) -ne $TargetSender){ continue }
            $subject = Get-Text $mail.Subject
            $body = Get-CurrentBody $mail.Body
            # The destination MKR belongs to the current message subject.
            # Older MKR numbers inside a backorder table identify the source
            # order and must not cause the same message to update that old sheet.
            $subjectTargetKeys = @(Get-MkrKeysFromText $subject)
            $mailTargetKeys = if($subjectTargetKeys.Count -gt 0){ $subjectTargetKeys } else { @(Get-MkrKeysFromText $body) }
            if($mailTargetKeys.Count -eq 0){ continue }
            $targetMessages++
            foreach($mailTargetKey in $mailTargetKeys){
                [void]$mailRecords.Add([pscustomobject]@{
                    TargetKey=$mailTargetKey
                    Received=$received
                    Subject=$subject
                    Body=$body
                })
            }

            $attachmentCount = 0
            try { $attachmentCount = [int]$mail.Attachments.Count } catch { $attachmentCount = 0 }
            for($a = 1; $a -le $attachmentCount; $a++){
                $attachment = $null
                try {
                    $attachment = $mail.Attachments.Item($a)
                    $name = Get-Text $attachment.FileName
                    $kind = Get-AttachmentKind $name
                    if(-not $kind){ continue }
                    $attachmentKeys = @(Get-MkrKeysFromText $name)
                    $attachmentTargetKey = ''
                    foreach($candidateKey in $mailTargetKeys){
                        if($attachmentKeys -contains $candidateKey){ $attachmentTargetKey = $candidateKey; break }
                    }
                    if(-not $attachmentTargetKey -and $attachmentKeys.Count -gt 0){
                        Write-Log ('첨부 건너뜀(MKR 불일치): 메일 ' + ($mailTargetKeys -join ', ') + ' / 첨부 ' + $name)
                        continue
                    }
                    if(-not $attachmentTargetKey -and (Test-AnyMkrToken $name)){
                        Write-Log ('첨부 건너뜀(대상 외 MKR): 메일 ' + ($mailTargetKeys -join ', ') + ' / 첨부 ' + $name)
                        continue
                    }
                    if(-not $attachmentTargetKey -and $mailTargetKeys.Count -eq 1){ $attachmentTargetKey = $mailTargetKeys[0] }
                    if(-not $attachmentTargetKey){ continue }
                    $safeName = $name -replace '[\\/:*?"<>|]', '_'
                    $destination = Join-Path $RunAttachmentDir (('{0}_{1:D3}_{2:D2}_{3}' -f $received.ToString('yyyyMMdd_HHmmss'), $scanned, $a, $safeName))
                    $attachment.SaveAsFile($destination)
                    $savedAttachments++
                    [void]$attachmentEvents.Add([pscustomobject]@{
                        TargetKey=$attachmentTargetKey
                        Received=$received
                        Kind=$kind
                        Path=$destination
                        FileName=$name
                    })
                    Write-Log ('첨부 저장: ' + $attachmentTargetKey + ' / ' + $kind + ' / ' + $name)
                } catch {
                    Write-Log ('첨부 저장 경고: ' + (Get-Text $_.Exception.Message))
                } finally {
                    Release-Com $attachment
                }
            }
            Write-Log ('대상 메일: ' + $received.ToString('yyyy-MM-dd HH:mm') + ' | ' + ($mailTargetKeys -join ', ') + ' | ' + $subject)
        } catch {
            Write-Log ('메일 처리 경고: ' + (Get-Text $_.Exception.Message))
        } finally {
            Release-Com $mail
        }
    }
} finally {
    Release-Com $items
    Release-Com $inbox
    Release-Com $namespace
    Release-Com $outlook
}

if($scanned -gt $MaxMessages){ throw '내부 오류: 메일 검색 상한 500건을 초과했습니다.' }
if($targetMessages -eq 0){ throw '대상 MKR 메일을 찾지 못했습니다. Excel은 변경하지 않았습니다.' }

$excel = $null
$workbook = $null
$pendingPath = Join-Path $Root '.MKR_TEST_pending.xlsx'
if(Test-Path -LiteralPath $pendingPath){ Remove-Item -LiteralPath $pendingPath -Force }
$commitReady = $false
$targetResults = [ordered]@{}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach($spec in $TargetSpecs){
        $targetKey = Get-Text $spec.Key
        $targetRecords = @($mailRecords | Where-Object { $_.TargetKey -eq $targetKey } | Sort-Object Received)
        $targetEvents = @($attachmentEvents | Where-Object { $_.TargetKey -eq $targetKey } | Sort-Object Received)
        if($targetRecords.Count -eq 0 -and $targetEvents.Count -eq 0){
            Write-Log ('[' + $targetKey + '] 대상 메일/첨부가 없어 해당 시트는 기존 상태를 유지합니다.')
            continue
        }

        $parsedSalesNoteEvents = New-Object System.Collections.Generic.List[object]
        foreach($event in @($targetEvents | Where-Object { $_.Kind -eq 'SalesNote' } | Sort-Object Received)){
            $pdfText = Read-PdfTextWithWord $event.Path
            $parsed = if($pdfText){ @(Parse-UpdateText $pdfText) } else { @() }
            $parsedBackorders = if($pdfText){ @(Parse-BackorderText $pdfText) } else { @() }
            $revisionRank = if((Get-Text $event.FileName) -match '(?i)(?:^|[\s_\-\(])rev(?:ised|ision)?(?:[\s_\-\)\.]|$)'){ 1 } else { 0 }
            $salesNoteType = if($revisionRank -gt 0){ '변경 Sales Note' } else { 'Sales Note' }
            Write-Log ('[' + $targetKey + '] ' + $salesNoteType + ' 분석: ' + $event.FileName + ' / 일반 ' + $parsed.Count + '개 / 백오더 ' + $parsedBackorders.Count + '개')
            [void]$parsedSalesNoteEvents.Add([pscustomobject]@{
                Received=$event.Received; Label=$event.FileName; Items=$parsed
                BackorderItems=$parsedBackorders; Kind='SalesNote'; RevisionRank=$revisionRank
            })
        }

        # Order QTY is fixed by the earliest successfully parsed original
        # order. If several attachments arrived in that same first message,
        # use the most complete one; later original copies and all Rev orders
        # are deliberately ignored for the historical snapshot.
        $parsedOriginalEvents = New-Object System.Collections.Generic.List[object]
        foreach($event in @($targetEvents | Where-Object { $_.Kind -eq 'OriginalOrder' } | Sort-Object Received)){
            $parsed = @(Read-OrderAttachmentItems $excel $event.Path)
            Write-Log ('[' + $targetKey + '] 원본 주문 첨부 분석: ' + $event.FileName + ' / ' + $parsed.Count + '개 품목')
            if($parsed.Count -gt 0){
                [void]$parsedOriginalEvents.Add([pscustomobject]@{ Received=[datetime]$event.Received; Label=$event.FileName; Items=$parsed })
            }
        }
        if($parsedOriginalEvents.Count -eq 0){
            Write-Log ('[' + $targetKey + '] 최초 원본 주문을 읽지 못해 해당 시트는 기존 상태를 유지합니다. Rev 주문서나 Sales Note로 Order QTY를 대신 만들지 않습니다.')
            continue
        }
        $firstOriginalReceived = [datetime](($parsedOriginalEvents | Measure-Object -Property Received -Minimum).Minimum)
        $firstOrderEvent = @($parsedOriginalEvents | Where-Object { ([datetime]$_.Received) -eq $firstOriginalReceived } | Sort-Object @{ Expression={ @($_.Items).Count }; Descending=$true }, Label | Select-Object -First 1)[0]
        $bestBaseline = @($firstOrderEvent.Items)
        $baselineReceived = [datetime]$firstOrderEvent.Received
        $baselineSource = '최초 주문 첨부 ' + (Get-Text $firstOrderEvent.Label)

        $state = [ordered]@{}
        foreach($item in $bestBaseline){
            $baselineOrderQty = [double]$item.Quantity
            $orderProperty = $item.PSObject.Properties['OrderQuantity']
            if($null -ne $orderProperty -and $null -ne $item.OrderQuantity -and [double]$item.OrderQuantity -gt 0){ $baselineOrderQty = [double]$item.OrderQuantity }
            $itemCode = Get-Text $item.Code
            $itemName = Get-Text $item.Name
            if(-not $itemCode -or -not $itemName){ continue }
            # The original order is also the safe default shipment snapshot.
            # Only an explicit mail/Sales Note quantity update may reduce it.
            # Therefore an unchanged item has ShortageQty = 0, not OrderQty.
            $state[$itemCode] = [ordered]@{ Code=$itemCode; Name=$itemName; RawOrderQty=$baselineOrderQty; OrderQty=$baselineOrderQty; ShippedQty=$baselineOrderQty }
        }
        if($state.Count -eq 0){
            Write-Log ('[' + $targetKey + '] 기준 주문 품목이 0개라 해당 시트는 기존 상태를 유지합니다.')
            continue
        }
        if($state.Count -gt ($TemplateLastRow - $DataStartRow + 1)){
            Write-Log ('[' + $targetKey + '] 품목 ' + $state.Count + '개: 기본 행 범위를 초과했지만 입력 행을 자동 확장합니다.')
        }
        $orderDate = $baselineReceived.Date
        $latestShipDate = $null
        $latestEta = $null
        $latestEtd = $null

        # Shipped QTY is the first ordinary Sales Note snapshot. Revised Sales
        # Notes and later mail tables never rewrite A:E. The mail body from the
        # same first Sales Note message is used only as a parser fallback.
        $salesNoteOrderMap = [ordered]@{}
        $firstSalesNote = @($parsedSalesNoteEvents | Where-Object { [int]$_.RevisionRank -eq 0 } | Sort-Object Received, Label | Select-Object -First 1)
        if($firstSalesNote.Count -gt 0){
            $firstSalesNoteEvent = $firstSalesNote[0]
            Merge-SalesNoteOrderEvidence $salesNoteOrderMap @($firstSalesNoteEvent.Items) ('[' + $targetKey + '] 최초 Sales Note ' + $firstSalesNoteEvent.Label)
            Apply-ShippedUpdates $state @($firstSalesNoteEvent.Items) ('[' + $targetKey + '] 최초 Sales Note ' + $firstSalesNoteEvent.Label) -PreserveOrderQuantity
            $sameMailRecords = @($targetRecords | Where-Object { [math]::Abs((([datetime]$_.Received) - ([datetime]$firstSalesNoteEvent.Received)).TotalSeconds) -le 2 } | Sort-Object Received)
            foreach($sameRecord in $sameMailRecords){
                $sameMailItems = @(Parse-UpdateText (Get-Text $sameRecord.Body))
                if($sameMailItems.Count -gt 0){
                    Merge-SalesNoteOrderEvidence $salesNoteOrderMap $sameMailItems ('[' + $targetKey + '] 최초 Sales Note 메일본문 보완')
                    Apply-ShippedUpdates $state $sameMailItems ('[' + $targetKey + '] 최초 Sales Note 메일본문 보완') -PreserveOrderQuantity
                }
                $firstRecordReceived = [datetime]$sameRecord.Received
                $foundEta = Get-DateFromBody (Get-Text $sameRecord.Body) '(?i)\bETA\b|到着予定|着港|도착예정' $firstRecordReceived
                $foundEtd = Get-DateFromBody (Get-Text $sameRecord.Body) '(?i)\bETD\b|出港|船積|출항' $firstRecordReceived
                $foundShip = Get-DateFromBody (Get-Text $sameRecord.Body) '(?i)\b(?:Ship|Shipment|Shipping)\s*Date\b|出荷日|発送日|출하일|선적일' $firstRecordReceived
                if($foundEta -is [datetime]){ $latestEta = $foundEta }
                if($foundEtd -is [datetime]){ $latestEtd = $foundEtd }
                if($foundShip -is [datetime]){ $latestShipDate = $foundShip }
            }
            Write-Log ('[' + $targetKey + '] A:E 기준 고정: ' + $baselineSource + ' + 최초 Sales Note ' + (Get-Text $firstSalesNoteEvent.Label))
        } else {
            Write-Log ('[' + $targetKey + '] 일반 Sales Note를 찾지 못했습니다. A:E는 최초 주문 기준 전량 선적 상태로 유지하고 Rev 자료는 적용하지 않습니다.')
        }

        # Backorders are owned by the destination MKR sheet and are written to
        # H:M only. They never alter the original order snapshot in A:E.
        $backorderLedger = [ordered]@{}
        foreach($record in $targetRecords){
            try {
                $recordReceived = [datetime]$record.Received
                $mailSourceLabel = '[' + $targetKey + '] 메일 본문 ' + $recordReceived.ToString('yyyy-MM-dd HH:mm')
                Merge-BackorderLedger $backorderLedger @(Parse-BackorderText (Get-Text $record.Body)) $mailSourceLabel $recordReceived
            } catch { Write-Log ('[' + $targetKey + '] 메일 본문 처리 경고: ' + (Get-Text $_.Exception.Message) + ' / 해당 메일은 건너뜁니다.') }
        }

        # Keep one latest snapshot per source-MKR + item. This prevents quoted
        # reply chains from adding the same backorder repeatedly.
        foreach($event in @($parsedSalesNoteEvents | Sort-Object Received, RevisionRank)){
            $salesNoteLabel = '[' + $targetKey + '] Sales Note ' + (Get-Text $event.Label)
            Merge-BackorderLedger $backorderLedger @($event.BackorderItems) $salesNoteLabel ([datetime]$event.Received)
        }
        foreach($ledgerKey in @($backorderLedger.Keys)){
            if((Get-Text $backorderLedger[$ledgerKey].SourceMkr) -eq $targetKey){
                [void]$backorderLedger.Remove($ledgerKey)
                Write-Log ('[' + $targetKey + '] 현재 MKR과 동일한 백오더 원본행 제외: ' + $ledgerKey)
            }
        }

        # A:E Order QTY represents only this MKR's new order. Subtract the H:M
        # included backorder shipment only when first-order and first-Sales-Note
        # evidence reconciles exactly; otherwise keep the original quantity.
        Apply-NetOrderExcludingBackorders $state $salesNoteOrderMap $backorderLedger ('[' + $targetKey + ']')
        foreach($code in @($state.Keys)){
            if([double]$state[$code].OrderQty -lt 0){
                throw ('[' + $targetKey + '] 백오더 제외 후 Order QTY가 음수입니다: ' + $code)
            }
            if([double]$state[$code].ShippedQty -gt [double]$state[$code].OrderQty){
                throw ('[' + $targetKey + '] 최초 Sales Note 선적수량이 백오더 제외 신규 주문수량을 초과합니다: ' + $code)
            }
        }

        $shortageItemCount = 0
        foreach($code in @($state.Keys | Sort-Object)){
            if(([double]$state[$code].OrderQty - [double]$state[$code].ShippedQty) -gt 0){ $shortageItemCount++ }
        }
        $targetResults[$targetKey] = [pscustomobject]@{
            Key=$targetKey; Sheet=(Get-Text $spec.Sheet); State=$state; OrderDate=$orderDate
            ShipDate=$latestShipDate; Eta=$latestEta; Etd=$latestEtd; ShortageCount=$shortageItemCount
            Backorders=$backorderLedger
        }
        Write-Log ('[' + $targetKey + '] 최초 주문·최초 Sales Note 확정: ' + $state.Count + '개 품목 / 부족 ' + $shortageItemCount + '개 / H:M 백오더 ' + $backorderLedger.Count + '개')
    }

    if($targetResults.Count -eq 0){ throw '대상 MKR 중 읽을 수 있는 최초 주문 자료가 없습니다. Excel은 변경하지 않았습니다.' }
    $workbook = $excel.Workbooks.Open($WorkbookPath)
    foreach($spec in $TargetSpecs){
        $targetKey = Get-Text $spec.Key
        if(-not $targetResults.Contains($targetKey)){ continue }
        $sheet = $null
        try {
            $result = $targetResults[$targetKey]
            $sheet = $workbook.Worksheets.Item((Get-Text $spec.Sheet))
            $requiredLastRow = $DataStartRow + [math]::Max([int]$result.State.Count, [int]$result.Backorders.Count) - 1
            $clearLastRow = Get-OutputLastRow $sheet $TemplateLastRow $requiredLastRow
            Ensure-OutputFormatRows $excel $sheet $clearLastRow
            $sheet.Range(([string]('A{0}:E{1}' -f $DataStartRow, $clearLastRow))).ClearContents()
            $sheet.Range(([string]('H{0}:M{1}' -f $DataStartRow, $clearLastRow))).ClearContents()
            $writtenCount = Write-ExcelItemsBulk $sheet $DataStartRow $result.State
            if($writtenCount -ne $result.State.Count){
                throw ('Excel 반영 품목 수 불일치: 예정 ' + $result.State.Count + '개 / 반영 ' + $writtenCount + '개')
            }
            $writtenBackorders = Write-ExcelBackordersBulk $sheet $DataStartRow $result.Backorders
            if($writtenBackorders -ne $result.Backorders.Count){
                throw ('Excel 백오더 반영 수 불일치: 예정 ' + $result.Backorders.Count + '개 / 반영 ' + $writtenBackorders + '개')
            }
            $sheet.Range('D3').Value2 = [string]$targetKey
            $sheet.Range('K4:K7').ClearContents()
            Set-ExcelDateCell $sheet 'K4' $result.OrderDate
            Set-ExcelDateCell $sheet 'K5' $result.ShipDate
            Set-ExcelDateCell $sheet 'K6' $result.Eta
            Set-ExcelDateCell $sheet 'K7' $result.Etd
            Write-Log ('[' + $targetKey + '] 시트 반영 완료: ' + (Get-Text $spec.Sheet) + ' / A:E ' + $result.State.Count + '개 / H:M 백오더 ' + $result.Backorders.Count + '개')
        } finally { Release-Com $sheet }
    }
    $workbook.SaveCopyAs($pendingPath)
    $commitReady = $true
    Write-Log ('임시 저장 완료: 검색 ' + $scanned + '건 / 대상메일 ' + $targetMessages + '건 / 반영 시트 ' + $targetResults.Count + '개 / 첨부 ' + $savedAttachments + '개')
} catch {
    Write-Log ('[ERROR] ' + (Get-Text $_.Exception.Message))
    throw
} finally {
    if($null -ne $workbook){ try { $workbook.Close($false) } catch {} }
    if($null -ne $excel){ try { $excel.Quit() } catch {} }
    Release-Com $workbook
    Release-Com $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if(-not $commitReady -or -not (Test-Path -LiteralPath $pendingPath)){ throw '임시 Excel 파일이 생성되지 않아 기존 파일을 변경하지 않았습니다.' }

$archive = $null
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($pendingPath)
    $entryNames = @($archive.Entries | ForEach-Object { $_.FullName })
    foreach($requiredEntry in @('[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml', 'xl/worksheets/sheet4.xml', 'xl/worksheets/sheet5.xml')){
        if($entryNames -notcontains $requiredEntry){ throw ('XLSX 필수 항목 누락: ' + $requiredEntry) }
    }
} finally { if($null -ne $archive){ $archive.Dispose() } }

Move-Item -LiteralPath $pendingPath -Destination $WorkbookPath -Force
Write-Log ('완료: 검색 ' + $scanned + '건(상한 500) / 대상메일 ' + $targetMessages + '건 / 반영 시트 ' + $targetResults.Count + '개 / 첨부 ' + $savedAttachments + '개')
Start-Process -FilePath $WorkbookPath
} catch {
    $finalMessage = Get-Text $_.Exception.Message
    try { Write-Log ('[STOPPED] ' + $finalMessage) } catch {}
    $pendingVariable = Get-Variable -Name pendingPath -ErrorAction SilentlyContinue
    if($null -ne $pendingVariable -and (Test-Path -LiteralPath $pendingVariable.Value)){
        try { Remove-Item -LiteralPath $pendingVariable.Value -Force } catch {}
    }
    Write-Host ''
    Write-Host ('[STOPPED] ' + $finalMessage) -ForegroundColor Yellow
    Write-Host '기존 Excel은 교체하지 않았습니다. MKR_SYNC_LOG.txt에 진단 내용을 저장했습니다.'
    exit 1
}
exit 0

# The legacy single-sheet block below is retained only as an inert compatibility record.
# It is never executed; the multi-MKR block above is the only active path.
if($false){
try {
if(-not (Test-Path -LiteralPath $WorkbookPath)){ throw "대상 Excel 파일이 없습니다: $WorkbookPath" }
if(Test-Path -LiteralPath (Join-Path $Root ('~$' + [IO.Path]::GetFileName($WorkbookPath)))){
    throw '대상 Excel 파일이 열려 있습니다. Excel을 닫고 다시 실행해 주세요.'
}

New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
New-Item -ItemType Directory -Path $RunAttachmentDir -Force | Out-Null
$backupPath = Join-Path $BackupDir ('MKR48_BlankOrPrevious_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsx')
Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath -Force
Write-Log ('실행 전 백업 완료: ' + $backupPath)

$mailRecords = New-Object System.Collections.Generic.List[object]
$attachmentEvents = New-Object System.Collections.Generic.List[object]
$outlook = $null
$namespace = $null
$inbox = $null
$items = $null
$scanned = 0
$targetMessages = 0
$savedAttachments = 0

try {
    Write-Log 'Classic Outlook 연결 확인 중...'
    try { $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
    catch { throw '실행 중인 Classic Outlook을 찾지 못했습니다. Classic Outlook을 먼저 열어 주세요.' }
    $namespace = $outlook.Session
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    try {
        $utcStart = $StartDate.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
        $jetStart = $StartDate.ToString('g', [Globalization.CultureInfo]::CurrentCulture)
        $unfilteredCount = [int]$items.Count
        $filterQueries = @(
            ("[ReceivedTime] >= '$jetStart' AND [SenderName] = 'Nagasawa Hikari (長澤 光里)'"),
            ("[ReceivedTime] >= '$jetStart' AND [SenderName] = 'Nagasawa Hikari'"),
            ("@SQL=(`"urn:schemas:httpmail:datereceived`" >= '{0}') AND (`"urn:schemas:httpmail:subject`" LIKE '%MKR48%')" -f $utcStart),
            ("@SQL=(`"urn:schemas:httpmail:datereceived`" >= '{0}') AND (`"urn:schemas:httpmail:subject`" ci_phrasematch 'MKR48')" -f $utcStart),
            ("@SQL=(`"urn:schemas:httpmail:datereceived`" >= '{0}') AND (`"http://schemas.microsoft.com/mapi/proptag/0x0C1A001F`" LIKE '%Nagasawa%')" -f $utcStart)
        )
        $filtered = $null
        $bestFilteredCount = [int]::MaxValue
        foreach($filterQuery in $filterQueries){
            $candidate = $null
            try {
                $candidate = $items.Restrict($filterQuery)
                $candidateCount = if($null -ne $candidate){ [int]$candidate.Count } else { 0 }
                if($candidateCount -gt 0 -and $candidateCount -lt $unfilteredCount -and $candidateCount -lt $bestFilteredCount){
                    Release-Com $filtered
                    $filtered = $candidate
                    $bestFilteredCount = $candidateCount
                    $candidate = $null
                }
            } catch {
                Write-Log ('제목 사전필터 방식 전환: ' + $_.Exception.Message)
            } finally {
                Release-Com $candidate
            }
        }
        if($null -ne $filtered){
            Release-Com $items
            $items = $filtered
            Write-Log ('날짜·발신자/MKR48 사전필터 결과: ' + [int]$items.Count + '건')
        } else {
            Write-Log '사전필터가 메일 범위를 줄이지 못해 받은편지함 최신 500건 안에서 복구 모드로 확인합니다.'
        }
    } catch {
        Write-Log ('사전필터 실패. 받은편지함 최신 500건 안에서만 확인: ' + $_.Exception.Message)
    }
    $items.Sort('[ReceivedTime]', $true)
    $scanCount = [math]::Min([int]$items.Count, $MaxMessages)
    Write-Log ('검색 시작: 최대 ' + $scanCount + '건 / 절대 상한 ' + $MaxMessages + '건')

    for($index = $scanCount; $index -ge 1; $index--){
        $mail = $null
        try {
            $mail = $items.Item($index)
            $scanned++
            if($null -eq $mail -or [int]$mail.Class -ne 43){ continue }
            $received = [datetime]$mail.ReceivedTime
            if($received -lt $StartDate){ continue }
            if((Get-SenderSmtp $mail) -ne $TargetSender){ continue }
            $subject = Get-Text $mail.Subject
            $body = Get-CurrentBody (Get-Text $mail.Body)
            if(-not (Test-TargetMkr ($subject + "`n" + $body))){ continue }
            $targetMessages++
            [void]$mailRecords.Add([pscustomobject]@{ Received=$received; Subject=$subject; Body=$body })

            for($a = 1; $a -le [int]$mail.Attachments.Count; $a++){
                $attachment = $null
                try {
                    $attachment = $mail.Attachments.Item($a)
                    $name = Get-Text $attachment.FileName
                    $isOriginal = $name -match '(?i)^MKR48_26-1A\(KRW\)\.xls[x]?$'
                    $isRevised = $name -match '(?i)^Rev[ _-]+MKR48_26-1A\(KRW\)\.xls[x]?$'
                    $isSalesNote = $name -match '(?i)(MKR48[_/\-]26[_/\-]1.*SALES[ _-]*NOTE|SALES[ _-]*NOTE.*MKR48[_/\-]26[_/\-]1).*\.pdf$'
                    if(-not $isOriginal -and -not $isRevised -and -not $isSalesNote){ continue }
                    $safeName = $name -replace '[\\/:*?"<>|]', '_'
                    $destination = Join-Path $RunAttachmentDir (('{0}_{1:D3}_{2:D2}_{3}' -f $received.ToString('yyyyMMdd_HHmmss'), $scanned, $a, $safeName))
                    $attachment.SaveAsFile($destination)
                    $savedAttachments++
                    $kind = if($isOriginal){ 'OriginalOrder' } elseif($isRevised){ 'RevisedOrder' } else { 'SalesNote' }
                    [void]$attachmentEvents.Add([pscustomobject]@{
                        Received=$received
                        Kind=$kind
                        Path=$destination
                        FileName=$name
                    })
                    Write-Log ('첨부 저장: ' + $kind + ' / ' + $name)
                } catch {
                    Write-Log ('첨부 저장 경고: ' + $_.Exception.Message)
                } finally {
                    Release-Com $attachment
                }
            }
            Write-Log ('대상 메일: ' + $received.ToString('yyyy-MM-dd HH:mm') + ' | ' + $subject)
        } catch {
            Write-Log ('메일 처리 경고: ' + $_.Exception.Message)
        } finally {
            Release-Com $mail
        }
    }
} finally {
    Release-Com $items
    Release-Com $inbox
    Release-Com $namespace
    Release-Com $outlook
}

if($scanned -gt $MaxMessages){ throw '내부 오류: 메일 검색 상한 500건을 초과했습니다.' }
if($targetMessages -eq 0){ throw 'MKR48/26-1 대상 메일을 찾지 못했습니다. 빈 초안은 변경하지 않았습니다.' }

$excel = $null
$workbook = $null
$sheet = $null
$pendingPath = Join-Path $Root '.MKR48_26-1_pending.xlsx'
if(Test-Path -LiteralPath $pendingPath){ Remove-Item -LiteralPath $pendingPath -Force }
$commitReady = $false
$state = [ordered]@{}
$baselineCount = 0
$baselineReceived = [datetime]::MaxValue
$orderDate = $null
$latestShipDate = $null
$latestEta = $null
$latestEtd = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $parsedRevisedEvents = New-Object System.Collections.Generic.List[object]
    foreach($event in @($attachmentEvents | Where-Object { $_.Kind -eq 'RevisedOrder' } | Sort-Object Received)){
        $parsed = @(Read-OrderAttachmentItems $excel $event.Path)
        Write-Log ('수정 주문 첨부 분석: ' + $event.FileName + ' / ' + $parsed.Count + '개 품목')
        if($parsed.Count -gt 0){
            [void]$parsedRevisedEvents.Add([pscustomobject]@{ Received=$event.Received; Label=$event.FileName; Items=$parsed; Kind='RevisedOrder' })
        }
    }

    $parsedSalesNoteEvents = New-Object System.Collections.Generic.List[object]
    foreach($event in @($attachmentEvents | Where-Object { $_.Kind -eq 'SalesNote' } | Sort-Object Received)){
        $pdfText = Read-PdfTextWithWord $event.Path
        $parsed = if($pdfText){ @(Parse-UpdateText $pdfText) } else { @() }
        Write-Log ('Sales Note 분석: ' + $event.FileName + ' / ' + $parsed.Count + '개 품목')
        if($parsed.Count -gt 0){
            [void]$parsedSalesNoteEvents.Add([pscustomobject]@{ Received=$event.Received; Label=$event.FileName; Items=$parsed; Kind='SalesNote' })
        }
    }

    $bestBaseline = @()
    $baselineSource = ''
    foreach($event in @($attachmentEvents | Where-Object { $_.Kind -eq 'OriginalOrder' } | Sort-Object Received)){
        $parsed = @(Read-OrderAttachmentItems $excel $event.Path)
        Write-Log ('최초 주문 첨부 분석: ' + $event.FileName + ' / ' + $parsed.Count + '개 품목')
        if($parsed.Count -gt $bestBaseline.Count){
            $bestBaseline = $parsed
            $baselineReceived = [datetime]$event.Received
            $baselineSource = '최초 주문 첨부 ' + $event.FileName
        }
    }
    if($bestBaseline.Count -eq 0){
        foreach($candidate in @(@($parsedRevisedEvents) + @($parsedSalesNoteEvents))){
            $candidateItems = @($candidate.Items)
            if($candidateItems.Count -gt $bestBaseline.Count){
                $bestBaseline = $candidateItems
                $baselineReceived = [datetime]$candidate.Received
                $baselineSource = '복구 기준 ' + (Get-Text $candidate.Label)
            }
        }
        if($bestBaseline.Count -gt 0){
            Write-Log ('최초 주문 첨부 누락/분석 실패. 가장 완전한 후속 자료로 자동 복구: ' + $baselineSource)
        }
    }
    if($bestBaseline.Count -eq 0){
        throw '주문서·수정 주문서·Sales Note 어디에서도 품목을 읽지 못했습니다. 첨부 저장 폴더와 진단 로그를 확인해 주세요.'
    }

    foreach($item in $bestBaseline){
        $baselineOrderQty = [double]$item.Quantity
        $orderProperty = $item.PSObject.Properties['OrderQuantity']
        if($null -ne $orderProperty -and $null -ne $item.OrderQuantity -and [double]$item.OrderQuantity -gt 0){
            $baselineOrderQty = [double]$item.OrderQuantity
        }
        $safeItemCode = Get-Text $item.Code
        $safeItemName = Get-Text $item.Name
        if(-not $safeItemCode -or -not $safeItemName){ continue }
        $state[$safeItemCode] = [ordered]@{
            Code = $safeItemCode
            Name = $safeItemName
            OrderQty = $baselineOrderQty
            ShippedQty = [double]$item.Quantity
        }
    }
    $baselineCount = $state.Count
    if($baselineCount -lt $MinimumExpectedItems){
        throw ('전체 품목 검증 실패: ' + $baselineCount + '개만 확인되었습니다. 4개 변경품목만 기록하지 않도록 저장을 중단합니다.')
    }
    $orderDate = $baselineReceived.Date
    Write-Log ('전체 품목 기준 확정: ' + $baselineCount + '개 품목 / ' + $baselineSource)

    foreach($event in @($parsedRevisedEvents | Sort-Object Received)){
        Apply-ShippedUpdates $state @($event.Items) ('수정 주문 첨부 ' + $event.Label)
    }

    foreach($event in @($parsedSalesNoteEvents | Sort-Object Received)){
        Apply-ShippedUpdates $state @($event.Items) ('Sales Note ' + $event.Label)
    }

    foreach($record in @($mailRecords | Sort-Object Received)){
        try {
            $recordSubject = Get-Text $record.Subject
            $recordBody = Get-Text $record.Body
            $recordReceived = $record.Received
            if($recordReceived -isnot [datetime]){ $recordReceived = [datetime]$recordReceived }
            $subjectDate = [regex]::Match($recordSubject, '(?<!\d)(20\d{2})(\d{2})(\d{2})(?!\d)')
            if($subjectDate.Success){
                try {
                    $candidateOrderDate = Get-Date -Year ([int]$subjectDate.Groups[1].Value) -Month ([int]$subjectDate.Groups[2].Value) -Day ([int]$subjectDate.Groups[3].Value) -Hour 0 -Minute 0 -Second 0
                    if($null -eq $orderDate -or $candidateOrderDate -lt $orderDate){ $orderDate = $candidateOrderDate }
                } catch {}
            }
            Apply-ShippedUpdates $state @(Parse-UpdateText $recordBody) ('메일 본문 ' + $recordReceived.ToString('yyyy-MM-dd HH:mm'))
            $foundOrder = $null
            $foundEta = $null
            $foundEtd = $null
            $foundShip = $null
            try {
                $foundOrder = Get-DateFromBody $recordBody '(?i)\bOrder\s*Date\b|注文日|発注日|발주일|주문일' $recordReceived
                $foundEta = Get-DateFromBody $recordBody '(?i)\bETA\b|到着予定|着港|도착예정' $recordReceived
                $foundEtd = Get-DateFromBody $recordBody '(?i)\bETD\b|出港|船積|출항' $recordReceived
                $foundShip = Get-DateFromBody $recordBody '(?i)\b(?:Ship|Shipment|Shipping)\s*Date\b|出荷日|発送日|출하일|선적일' $recordReceived
            } catch {
                Write-Log ('메일 본문 날짜 분석 경고: ' + (Get-Text $_.Exception.Message))
            }
            if($foundOrder -is [datetime] -and ($null -eq $orderDate -or $foundOrder -lt $orderDate)){ $orderDate = $foundOrder }
            if($foundEta -is [datetime]){ $latestEta = $foundEta }
            if($foundEtd -is [datetime]){ $latestEtd = $foundEtd }
            if($foundShip -is [datetime]){ $latestShipDate = $foundShip }
        } catch {
            Write-Log ('메일 본문 처리 경고: ' + (Get-Text $_.Exception.Message) + ' / 해당 메일은 건너뜁니다.')
        }
    }

    if($state.Count -ne $baselineCount){ throw '전체 품목 검증 실패: 최초 주문 품목 수와 최종 품목 수가 다릅니다.' }
    if($state.Count -gt ($TemplateLastRow - $DataStartRow + 1)){ Write-Log ('기본 행 범위를 초과했지만 입력 행을 자동 확장합니다: ' + $state.Count + '개') }

    $workbook = $excel.Workbooks.Open($WorkbookPath)
    $sheet = $workbook.Worksheets.Item('MKR48_26-1')
    $requiredLastRow = $DataStartRow + $state.Count - 1
    $clearLastRow = Get-OutputLastRow $sheet $TemplateLastRow $requiredLastRow
    $sheet.Range(('A{0}:E{1}' -f $DataStartRow, $clearLastRow)).ClearContents()
    $row = $DataStartRow
    $shortageItemCount = 0
    foreach($code in @($state.Keys | Sort-Object)){
        $record = $state[$code]
        $shortage = [math]::Max(0, [double]$record.OrderQty - [double]$record.ShippedQty)
        if($shortage -gt 0){ $shortageItemCount++ }
        $sheet.Range(('A{0}' -f $row)).Value2 = Get-Text $record.Code
        $sheet.Range(('B{0}' -f $row)).Value2 = Get-Text $record.Name
        $sheet.Range(('C{0}' -f $row)).Value2 = [double]$record.OrderQty
        $sheet.Range(('D{0}' -f $row)).Value2 = [double]$record.ShippedQty
        $sheet.Range(('E{0}' -f $row)).Value2 = [double]$shortage
        $row++
    }
    $sheet.Range('D3').Value2 = $TargetMkr
    $sheet.Range('K4:K7').ClearContents()
    Set-ExcelDateCell $sheet 'K4' $orderDate
    Set-ExcelDateCell $sheet 'K5' $latestShipDate
    Set-ExcelDateCell $sheet 'K6' $latestEta
    Set-ExcelDateCell $sheet 'K7' $latestEtd

    $workbook.SaveCopyAs($pendingPath)
    $commitReady = $true
    Write-Log ('임시 저장 완료: 검색 ' + $scanned + '건 / 대상메일 ' + $targetMessages + '건 / 전체품목 ' + $state.Count + '개 / 수량변경 ' + $shortageItemCount + '개')
} catch {
    Write-Log ('[ERROR] ' + $_.Exception.Message)
    throw
} finally {
    if($null -ne $workbook){ try { $workbook.Close($false) } catch {} }
    if($null -ne $excel){ try { $excel.Quit() } catch {} }
    Release-Com $sheet
    Release-Com $workbook
    Release-Com $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if(-not $commitReady -or -not (Test-Path -LiteralPath $pendingPath)){
    throw '임시 Excel 파일이 생성되지 않아 기존 파일을 변경하지 않았습니다.'
}

$archive = $null
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($pendingPath)
    $entryNames = @($archive.Entries | ForEach-Object { $_.FullName })
    foreach($requiredEntry in @('[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml')){
        if($entryNames -notcontains $requiredEntry){ throw ('XLSX 필수 항목 누락: ' + $requiredEntry) }
    }
} finally {
    if($null -ne $archive){ $archive.Dispose() }
}

Move-Item -LiteralPath $pendingPath -Destination $WorkbookPath -Force
Write-Log ('완료: 검색 ' + $scanned + '건(상한 500) / 대상메일 ' + $targetMessages + '건 / 전체품목 ' + $state.Count + '개 / 수량변경 ' + $shortageItemCount + '개 / 첨부 ' + $savedAttachments + '개')
Start-Process -FilePath $WorkbookPath
} catch {
    $finalMessage = $_.Exception.Message
    try { Write-Log ('[STOPPED] ' + $finalMessage) } catch {}
    $pendingVariable = Get-Variable -Name pendingPath -ErrorAction SilentlyContinue
    if($null -ne $pendingVariable -and (Test-Path -LiteralPath $pendingVariable.Value)){
        try { Remove-Item -LiteralPath $pendingVariable.Value -Force } catch {}
    }
    Write-Host ''
    Write-Host ('[STOPPED] ' + $finalMessage) -ForegroundColor Yellow
    Write-Host '기존 Excel은 교체하지 않았습니다. MKR48_SYNC_LOG.txt에 진단 내용을 저장했습니다.'
    exit 1
}
exit 0
}
