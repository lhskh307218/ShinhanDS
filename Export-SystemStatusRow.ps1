<#
.SYNOPSIS
    VRtools(RVTools) 추출 xlsx의 vInfo / vDisk 탭을 읽어 시스템현황 규격의 값을 계산하고,
    기존 서비스현황 파일의 시스템현황 탭과 비교한 결과를 별도 엑셀 파일로 생성합니다.

.DESCRIPTION
    - 폐쇄망용: 인터넷/외부 모듈/Excel 설치 없이 동작 (순수 .NET zip + XML 파싱)
    - PowerShell 5.1 이상 지원
    - 원본(서비스현황) 파일은 읽기만 하고 절대 수정하지 않습니다.

    [동작]
    1. RVTools 파일의 vInfo/vDisk 를 바탕으로 outputex.xlsx 형태의 결과 파일 생성
       (B~N열에 연속 출력: Lookup용, 망구분, 내/외부, 구성형태, 노드, 그룹사, 구분,
        호스트명, IP, CPU(Core), MEM(GB), DISK(SSD), DISK(SAS))
    2. 서비스현황 파일의 시스템현황 탭에서 같은 VM(Lookup용/C열 기준)을 찾아 값 비교
    3. 달라진 셀은 결과 파일에서 노란색 음영으로 표시
       (시스템현황에 없는 신규 VM은 행 전체가 노란색)

    [값 계산 규칙] (docs\prompt.txt 기준)
    - DISK(SSD)/DISK(SAS): 같은 VM의 vDisk 행들을 Path 기준으로 분류
      (primera 포함 → SSD, unity 포함 → SAS) 후 Capacity MiB 합산 → GB 변환(/1024)
    - Network #1 열을 '_'로 분리 → 그룹사 / 용도(T=개발, P=운영) / 망(I,A=내부, E=외부)
    - 호스트명: DNS Name에서 첫 번째 '.' 앞부분 / 노드: Host에서 첫 번째 '.' 앞부분
    - 구성형태는 항상 "클라우드"
    - Memory는 MiB→GB 변환 (그대로 쓰려면 -RawMemory)

.PARAMETER VrtoolsPath
    vInfo / vDisk 탭이 들어있는 VRtools xlsx 파일 경로

.PARAMETER ServicePath
    기존 서비스현황 xlsx 파일 경로 (시스템현황 탭과 비교, 읽기 전용).
    생략하면 비교 없이 결과 파일만 생성합니다. (별칭: -TargetPath)

.PARAMETER TargetSheet
    서비스현황 파일에서 비교할 탭 이름. 기본값: 시스템현황

.PARAMETER OutFile
    결과 xlsx 파일 경로 (생략 시 자동: 시스템현황_비교_일시.xlsx)

.PARAMETER RawMemory
    지정하면 Memory 값을 변환 없이 그대로 출력 (기본: MiB → GB 변환)

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Export-SystemStatusRow.ps1 -VrtoolsPath ".\RVTools_export_all_....xlsx" -ServicePath ".\신한DS_클라우드_서비스현황_....xlsx"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$VrtoolsPath,

    [Alias('TargetPath')]
    [string]$ServicePath,

    [string]$TargetSheet = '시스템현황',

    [string]$OutFile,

    [switch]$RawMemory
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# 0. 준비
# ---------------------------------------------------------------------------
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
} catch {
    throw "System.IO.Compression.FileSystem 어셈블리를 로드할 수 없습니다. .NET Framework 4.5 이상이 필요합니다."
}

if (-not (Test-Path -LiteralPath $VrtoolsPath)) {
    $here = (Get-Location).Path
    throw "파일을 찾을 수 없습니다: '$VrtoolsPath' (현재 폴더: $here)`n-VrtoolsPath 에는 실제 VRtools xlsx 파일 이름을 넣어주세요."
}
$xlsxFile = (Resolve-Path -LiteralPath $VrtoolsPath).ProviderPath

$serviceFile = $null
if (-not [string]::IsNullOrEmpty($ServicePath)) {
    if (-not (Test-Path -LiteralPath $ServicePath)) {
        $here = (Get-Location).Path
        throw "서비스현황 파일을 찾을 수 없습니다: '$ServicePath' (현재 폴더: $here)"
    }
    $serviceFile = (Resolve-Path -LiteralPath $ServicePath).ProviderPath
}

$RelNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# ---------------------------------------------------------------------------
# 1. xlsx 읽기 공통 함수 (순수 .NET: zip 해제 + XML)
# ---------------------------------------------------------------------------
function Read-ZipXml {
    param($Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { return $null }
    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        $text = $reader.ReadToEnd()
        $reader.Close()
    } finally {
        $stream.Dispose()
    }
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($text)
    return $doc
}

function Get-ColumnLetter {
    param([string]$CellRef)
    if ($CellRef -match '^([A-Z]+)\d+$') { return $Matches[1] }
    return $null
}

function Get-SheetEntryPath {
    param($Zip, [string]$SheetName)
    $wb   = Read-ZipXml -Zip $Zip -EntryName 'xl/workbook.xml'
    $rels = Read-ZipXml -Zip $Zip -EntryName 'xl/_rels/workbook.xml.rels'
    if ($null -eq $wb -or $null -eq $rels) { throw "xlsx 내부 구조를 읽을 수 없습니다." }

    $relMap = @{}
    foreach ($r in $rels.SelectNodes("//*[local-name()='Relationship']")) {
        $relMap[$r.GetAttribute('Id')] = $r.GetAttribute('Target')
    }

    $sheetNodes = @($wb.SelectNodes("//*[local-name()='sheet']"))
    $names = New-Object System.Collections.ArrayList
    foreach ($s in $sheetNodes) { [void]$names.Add($s.GetAttribute('name')) }

    $target = $null
    foreach ($s in $sheetNodes) {
        if ($s.GetAttribute('name') -ieq $SheetName) {
            $target = $relMap[$s.GetAttribute('id', $RelNs)]; break
        }
    }
    if ($null -eq $target) {
        foreach ($s in $sheetNodes) {
            if ($s.GetAttribute('name') -like "*$SheetName*") {
                $target = $relMap[$s.GetAttribute('id', $RelNs)]; break
            }
        }
    }
    if ([string]::IsNullOrEmpty($target)) {
        return @{ Path = $null; Names = $names }
    }
    $target = $target.TrimStart('/')
    if ($target -notlike 'xl/*') { $target = 'xl/' + $target }
    return @{ Path = $target; Names = $names }
}

function Read-SharedStrings {
    param($Zip)
    $list = New-Object System.Collections.ArrayList
    $sst = Read-ZipXml -Zip $Zip -EntryName 'xl/sharedStrings.xml'
    if ($null -ne $sst) {
        foreach ($si in $sst.DocumentElement.SelectNodes("*[local-name()='si']")) {
            $texts = $si.SelectNodes(".//*[local-name()='t'][not(ancestor::*[local-name()='rPh'])]")
            $sb = New-Object System.Text.StringBuilder
            foreach ($t in $texts) { [void]$sb.Append($t.InnerText) }
            [void]$list.Add($sb.ToString())
        }
    }
    return ,$list
}

function Get-CellText {
    param($CellNode, $SharedStrings)
    $type = $CellNode.GetAttribute('t')
    if ($type -eq 'inlineStr') {
        $sb = New-Object System.Text.StringBuilder
        foreach ($t in $CellNode.SelectNodes(".//*[local-name()='t']")) { [void]$sb.Append($t.InnerText) }
        return $sb.ToString()
    }
    $v = $CellNode.SelectSingleNode("*[local-name()='v']")
    if ($null -eq $v) { return '' }
    $value = $v.InnerText
    if ($type -eq 's') {
        $idx = 0
        if ([int]::TryParse($value, [ref]$idx) -and $idx -lt $SharedStrings.Count) {
            return $SharedStrings[$idx]
        }
    }
    return $value
}

# 시트를 행 배열로 반환. 각 행은 @{ 'A'='값'; ...; '_r'=행번호 }
function Read-XlsxSheet {
    param([string]$Path, [string]$SheetName)
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $info = Get-SheetEntryPath -Zip $zip -SheetName $SheetName
        if ($null -eq $info.Path) { return $null }
        $sharedStrings = Read-SharedStrings -Zip $zip
        $sheetDoc = Read-ZipXml -Zip $zip -EntryName $info.Path
        $rows = New-Object System.Collections.ArrayList
        foreach ($rowNode in $sheetDoc.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
            $cells = @{}
            foreach ($c in $rowNode.SelectNodes("*[local-name()='c']")) {
                $col = Get-ColumnLetter -CellRef $c.GetAttribute('r')
                if ($null -eq $col) { continue }
                $cells[$col] = Get-CellText -CellNode $c -SharedStrings $sharedStrings
            }
            if ($cells.Count -gt 0) {
                $rn = 0
                [void][int]::TryParse($rowNode.GetAttribute('r'), [ref]$rn)
                $cells['_r'] = $rn
                [void]$rows.Add($cells)
            }
        }
        return ,$rows
    } finally {
        $zip.Dispose()
    }
}

function ConvertTo-Records {
    param($Rows)
    if ($null -eq $Rows -or $Rows.Count -eq 0) { return @() }
    $headerMap = @{}
    foreach ($col in $Rows[0].Keys) {
        if ($col -eq '_r') { continue }
        $name = ([string]$Rows[0][$col]).Trim()
        if ($name -ne '') { $headerMap[$col] = $name }
    }
    $records = New-Object System.Collections.ArrayList
    for ($i = 1; $i -lt $Rows.Count; $i++) {
        $rec = @{}
        $hasValue = $false
        foreach ($col in $headerMap.Keys) {
            $val = ''
            if ($Rows[$i].ContainsKey($col)) { $val = [string]$Rows[$i][$col] }
            $rec[$headerMap[$col]] = $val
            if ($val.Trim() -ne '') { $hasValue = $true }
        }
        if ($hasValue) { [void]$records.Add($rec) }
    }
    return ,$records
}

function Find-Header {
    param($Record, [string[]]$Candidates, [string]$SheetLabel)
    foreach ($cand in $Candidates) {
        foreach ($key in $Record.Keys) {
            if ($key -ieq $cand) { return $key }
        }
    }
    throw "[$SheetLabel] 탭에서 열을 찾을 수 없습니다: $($Candidates -join ' / ')"
}

function Format-Number {
    param($Value)
    if ($null -eq $Value -or [string]$Value -eq '') { return '' }
    $d = 0.0
    $styles  = [System.Globalization.NumberStyles]::Any
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    if ([double]::TryParse([string]$Value, $styles, $culture, [ref]$d)) {
        return ([math]::Round($d, 2)).ToString($culture)
    }
    return [string]$Value
}

function Test-IsNumber {
    param($Value)
    $d = 0.0
    $styles  = [System.Globalization.NumberStyles]::Any
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    return [double]::TryParse([string]$Value, $styles, $culture, [ref]$d)
}

# 두 값이 같은지 비교 (숫자면 수치로, 문자면 대소문자 무시)
function Test-SameValue {
    param($A, $B)
    $sa = ([string]$A).Trim()
    $sb = ([string]$B).Trim()
    $da = 0.0; $db = 0.0
    $styles  = [System.Globalization.NumberStyles]::Any
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    if ([double]::TryParse($sa, $styles, $culture, [ref]$da) -and
        [double]::TryParse($sb, $styles, $culture, [ref]$db)) {
        return ([math]::Round($da, 4) -eq [math]::Round($db, 4))
    }
    return [string]::Equals($sa, $sb, [System.StringComparison]::OrdinalIgnoreCase)
}

# ---------------------------------------------------------------------------
# 2. 결과 xlsx 생성 함수 (outputex.xlsx 형태, 변경 셀 노란색)
# ---------------------------------------------------------------------------
function ConvertTo-XmlText {
    param([string]$Text)
    $t = $Text -replace '&', '&amp;'
    $t = $t -replace '<', '&lt;'
    $t = $t -replace '>', '&gt;'
    return $t
}

function New-ComparisonWorkbook {
    param(
        [string]$Path,
        $DataRows    # 배열: 각 항목 = @{ Values=문자열13개배열; IsNum=bool13개배열; Changed=bool13개배열 }
    )

    $cols = @('B','C','D','E','F','G','H','I','J','K','L','M','N')
    $headers = @('Lookup용', '망구분', '내/외부', '구성형태', '노드', '그룹사', '구분',
                 '호스트명', 'IP', "CPU`n(Core)", "MEM`n(GB)", "DISK`n(SSD)", "DISK`n(SAS)")
    $headerStyle = @(1, 2, 2, 2, 2, 3, 3, 4, 4, 5, 5, 5, 5)   # outputex.xlsx 헤더 음영 배치

    # --- sheet1.xml ---
    $lastRow = 2 + $DataRows.Count
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$sb.Append("<dimension ref=`"B2:N$lastRow`"/>")
    [void]$sb.Append('<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="16.5"/>')
    [void]$sb.Append('<sheetData>')

    # 헤더 행 (2행, 높이 31.5)
    [void]$sb.Append('<row r="2" spans="2:14" ht="31.5">')
    for ($i = 0; $i -lt $cols.Count; $i++) {
        $txt = (ConvertTo-XmlText -Text $headers[$i]) -replace "`n", '&#10;'
        [void]$sb.Append("<c r=`"$($cols[$i])2`" s=`"$($headerStyle[$i])`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$txt</t></is></c>")
    }
    [void]$sb.Append('</row>')

    # 데이터 행 (3행부터)
    $r = 3
    foreach ($row in $DataRows) {
        [void]$sb.Append("<row r=`"$r`" spans=`"2:14`">")
        for ($i = 0; $i -lt $cols.Count; $i++) {
            $val = [string]$row.Values[$i]
            $style = 6
            if ($row.Changed[$i]) { $style = 7 }   # 변경/신규 → 노란색
            if ($val -eq '') {
                [void]$sb.Append("<c r=`"$($cols[$i])$r`" s=`"$style`"/>")
            } elseif ($row.IsNum[$i]) {
                [void]$sb.Append("<c r=`"$($cols[$i])$r`" s=`"$style`"><v>$val</v></c>")
            } else {
                $txt = ConvertTo-XmlText -Text $val
                [void]$sb.Append("<c r=`"$($cols[$i])$r`" s=`"$style`" t=`"inlineStr`"><is><t>$txt</t></is></c>")
            }
        }
        [void]$sb.Append('</row>')
        $r++
    }
    [void]$sb.Append('</sheetData>')
    [void]$sb.Append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    [void]$sb.Append('</worksheet>')
    $sheetXml = $sb.ToString()

    # --- 고정 파트들 (outputex.xlsx 서식 재현 + 변경표시용 노란색 스타일 추가) ---
    $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>'

    $rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>'

    $workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
        '</workbook>'

    $wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>'

    $stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="2">' +
        '<font><sz val="11"/><color theme="1"/><name val="맑은 고딕"/><family val="2"/><charset val="129"/><scheme val="minor"/></font>' +
        '<font><b/><sz val="9"/><color theme="1"/><name val="원신한 Light"/><family val="3"/><charset val="129"/></font>' +
        '</fonts>' +
        '<fills count="8">' +
        '<fill><patternFill patternType="none"/></fill>' +
        '<fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD6DCE4"/><bgColor indexed="64"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/><bgColor indexed="64"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFEDEDED"/><bgColor indexed="64"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>' +
        '</fills>' +
        '<borders count="2">' +
        '<border><left/><right/><top/><bottom/><diagonal/></border>' +
        '<border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>' +
        '</borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs>' +
        '<cellXfs count="8">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' +
        '<xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' +
        '<xf numFmtId="0" fontId="1" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' +
        '<xf numFmtId="0" fontId="1" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' +
        '<xf numFmtId="0" fontId="1" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="center"/></xf>' +
        '<xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0" applyFill="1" applyBorder="1"><alignment vertical="center"/></xf>' +
        '</cellXfs>' +
        '<cellStyles count="1"><cellStyle name="표준" xfId="0" builtinId="0"/></cellStyles>' +
        '</styleSheet>'

    # --- zip 생성 ---
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        foreach ($pair in @(
            @('[Content_Types].xml',        $contentTypes),
            @('_rels/.rels',                $rootRels),
            @('xl/workbook.xml',            $workbook),
            @('xl/_rels/workbook.xml.rels', $wbRels),
            @('xl/styles.xml',              $stylesXml),
            @('xl/worksheets/sheet1.xml',   $sheetXml)
        )) {
            $entry = $zip.CreateEntry($pair[0])
            $stream = $entry.Open()
            try {
                $writer = New-Object System.IO.StreamWriter($stream, $utf8)
                $writer.Write([string]$pair[1])
                $writer.Flush()
                $writer.Close()
            } finally {
                $stream.Dispose()
            }
        }
    } finally {
        $zip.Dispose()
    }
}

# ---------------------------------------------------------------------------
# 3. vInfo / vDisk 읽기
# ---------------------------------------------------------------------------
Write-Host "VRtools 파일 읽는 중: $xlsxFile" -ForegroundColor Cyan

$vInfoRows = Read-XlsxSheet -Path $xlsxFile -SheetName 'vInfo'
if ($null -eq $vInfoRows) { throw "vInfo 탭을 찾을 수 없습니다." }
$vDiskRows = Read-XlsxSheet -Path $xlsxFile -SheetName 'vDisk'
if ($null -eq $vDiskRows) { throw "vDisk 탭을 찾을 수 없습니다." }

$vInfo = ConvertTo-Records -Rows $vInfoRows
$vDisk = ConvertTo-Records -Rows $vDiskRows
if ($vInfo.Count -eq 0) { throw "vInfo 탭에 데이터 행이 없습니다." }

$hVm      = Find-Header -Record $vInfo[0] -Candidates @('VM')                            -SheetLabel 'vInfo'
$hDns     = Find-Header -Record $vInfo[0] -Candidates @('DNS Name')                      -SheetLabel 'vInfo'
$hCpu     = Find-Header -Record $vInfo[0] -Candidates @('CPUs')                          -SheetLabel 'vInfo'
$hMem     = Find-Header -Record $vInfo[0] -Candidates @('Memory')                        -SheetLabel 'vInfo'
$hHost    = Find-Header -Record $vInfo[0] -Candidates @('Host')                          -SheetLabel 'vInfo'
$hIp      = Find-Header -Record $vInfo[0] -Candidates @('Primary IP Address')            -SheetLabel 'vInfo'
$hNet     = Find-Header -Record $vInfo[0] -Candidates @('Network #1', 'Network')         -SheetLabel 'vInfo'
$hDVm     = Find-Header -Record $vDisk[0] -Candidates @('VM')                            -SheetLabel 'vDisk'
$hDSize   = Find-Header -Record $vDisk[0] -Candidates @('Capacity MiB', 'Capacity MB')   -SheetLabel 'vDisk'
$hDPath   = Find-Header -Record $vDisk[0] -Candidates @('Path')                          -SheetLabel 'vDisk'

# ---------------------------------------------------------------------------
# 4. vDisk 집계: VM별 SSD(primera) / SAS(unity) 용량 합산 (MiB)
# ---------------------------------------------------------------------------
$diskMap = @{}
foreach ($d in $vDisk) {
    $dvm = ([string]$d[$hDVm]).Trim()
    if ($dvm -eq '' -or $dvm -like '*{{*') { continue }
    $key = $dvm.ToUpperInvariant()
    if (-not $diskMap.ContainsKey($key)) {
        $diskMap[$key] = @{ Ssd = 0.0; Sas = 0.0; Unknown = New-Object System.Collections.ArrayList }
    }
    $sizeMiB = 0.0
    $styles  = [System.Globalization.NumberStyles]::Any
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    if (-not [double]::TryParse([string]$d[$hDSize], $styles, $culture, [ref]$sizeMiB)) {
        Write-Warning "vDisk: '$dvm' 디스크 용량을 숫자로 해석할 수 없어 제외합니다: '$($d[$hDSize])'"
        continue
    }
    $path = [string]$d[$hDPath]
    if ($path -match 'primera') {
        $diskMap[$key].Ssd += $sizeMiB
    } elseif ($path -match 'unity') {
        $diskMap[$key].Sas += $sizeMiB
    } else {
        [void]$diskMap[$key].Unknown.Add($path)
    }
}

# ---------------------------------------------------------------------------
# 5. 서비스현황(시스템현황 탭) 기존 데이터 읽기 (읽기 전용, 비교용)
#    시스템현황 열 배치: C=VM이름 D=망구분 E=내/외부 F=구성형태 G=노드
#                       L=그룹사 O=구분 R=호스트명 S=IP V=CPU W=MEM X=SSD Y=SAS
# ---------------------------------------------------------------------------
$svcCols = @('C','D','E','F','G','L','O','R','S','V','W','X','Y')
$svcMap = @{}          # VM이름(대문자) → 13개 값 배열
$svcOrder = New-Object System.Collections.ArrayList

if ($null -ne $serviceFile) {
    Write-Host "서비스현황 파일 읽는 중(수정 안 함): $serviceFile" -ForegroundColor Cyan
    $svcRows = Read-XlsxSheet -Path $serviceFile -SheetName $TargetSheet
    if ($null -eq $svcRows) {
        $zipTmp = [System.IO.Compression.ZipFile]::OpenRead($serviceFile)
        try { $info = Get-SheetEntryPath -Zip $zipTmp -SheetName $TargetSheet } finally { $zipTmp.Dispose() }
        throw "서비스현황 파일에서 '$TargetSheet' 탭을 찾을 수 없습니다. 존재하는 탭: $($info.Names -join ', ')"
    }
    foreach ($row in $svcRows) {
        if ([int]$row['_r'] -le 5) { continue }               # 헤더 영역(1~5행) 제외
        $key = ''
        if ($row.ContainsKey('C')) { $key = ([string]$row['C']).Trim() }
        if ($key -eq '' -or $key -like '*{{*' -or $key -eq 'VM이름') { continue }   # 빈행/템플릿/범례행 제외
        $ukey = $key.ToUpperInvariant()
        if ($svcMap.ContainsKey($ukey)) {
            Write-Warning "시스템현황: VM '$key' 행이 중복되어 첫 번째 행(${TargetSheet} $($svcMap[$ukey].Row)행)만 비교에 사용합니다."
            continue
        }
        $vals = New-Object System.Collections.ArrayList
        foreach ($col in $svcCols) {
            $v = ''
            if ($row.ContainsKey($col)) { $v = [string]$row[$col] }
            [void]$vals.Add($v)
        }
        $svcMap[$ukey] = @{ Values = $vals; Row = [int]$row['_r'] }
        [void]$svcOrder.Add($ukey)
    }
    Write-Host "시스템현황 기존 데이터: $($svcMap.Count)건" -ForegroundColor Cyan
} else {
    Write-Host "서비스현황 파일이 지정되지 않아 비교 없이 결과만 생성합니다. (-ServicePath 지정 시 변경 셀 표시)" -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# 6. VM별 값 계산 + 비교
# ---------------------------------------------------------------------------
$useTokens = @('T', 'P')        # T=개발, P=운영
$netTokens = @('I', 'A', 'E')   # I,A=내부, E=외부
$fieldNames = @('VM이름','망구분','내/외부','구성형태','노드','그룹사','구분','호스트명','IP','CPU(Core)','MEM(GB)','DISK(SSD)','DISK(SAS)')
$numericField = @($false,$false,$false,$false,$false,$false,$false,$false,$false,$true,$true,$true,$true)

$dataRows = New-Object System.Collections.ArrayList
$summary  = New-Object System.Collections.ArrayList
$rvVmSet  = @{}

foreach ($vm in $vInfo) {
    $vmName = ([string]$vm[$hVm]).Trim()
    if ($vmName -eq '' -or $vmName -like '*{{*') { continue }
    $rvVmSet[$vmName.ToUpperInvariant()] = $true

    # --- Network #1 분해 ---
    $group = ''; $useToken = ''; $netToken = ''
    $network = ([string]$vm[$hNet]).Trim()
    if ($network -ne '') {
        $tokens = $network -split '_'
        $group = $tokens[0]
        $limit = [Math]::Min(2, $tokens.Count - 1)
        for ($i = 1; $i -le $limit; $i++) {
            $tk = ([string]$tokens[$i]).Trim().ToUpperInvariant()
            if ($useTokens -contains $tk) { $useToken = $tk }
            elseif ($netTokens -contains $tk) { $netToken = $tk }
        }
    } else {
        Write-Warning "vInfo: '$vmName' 의 Network 값이 비어 있습니다. 망구분/내외부/그룹사가 공란으로 출력됩니다."
    }

    $mangubun = ''; $gubun = ''
    switch ($useToken) {
        'T' { $mangubun = '개발망'; $gubun = '개발' }
        'P' { $mangubun = '운영망'; $gubun = '운영' }
        default {
            if ($network -ne '') { Write-Warning "vInfo: '$vmName' 의 용도 코드(T/P)를 찾을 수 없습니다: '$network'" }
        }
    }
    $naeoebu = ''
    switch ($netToken) {
        'I' { $naeoebu = '내부' }
        'A' { $naeoebu = '내부' }
        'E' { $naeoebu = '외부' }
        default {
            if ($network -ne '') { Write-Warning "vInfo: '$vmName' 의 망 코드(I/A/E)를 찾을 수 없습니다: '$network'" }
        }
    }

    $hostname = (([string]$vm[$hDns]).Trim() -split '\.')[0]
    $nodeName = (([string]$vm[$hHost]).Trim() -split '\.')[0]

    $cpu = Format-Number -Value ([string]$vm[$hCpu]).Trim()
    $memRaw = ([string]$vm[$hMem]).Trim()
    $mem = ''
    if ($memRaw -ne '') {
        if ($RawMemory) {
            $mem = Format-Number -Value $memRaw
        } else {
            $memMiB = 0.0
            $styles  = [System.Globalization.NumberStyles]::Any
            $culture = [System.Globalization.CultureInfo]::InvariantCulture
            if ([double]::TryParse($memRaw, $styles, $culture, [ref]$memMiB)) {
                $mem = Format-Number -Value ($memMiB / 1024)   # MiB → GB
            } else {
                $mem = $memRaw
            }
        }
    }

    $ssdGb = '0'; $sasGb = '0'
    $dKey = $vmName.ToUpperInvariant()
    if ($diskMap.ContainsKey($dKey)) {
        $ssdGb = Format-Number -Value ($diskMap[$dKey].Ssd / 1024)
        $sasGb = Format-Number -Value ($diskMap[$dKey].Sas / 1024)
        foreach ($u in $diskMap[$dKey].Unknown) {
            Write-Warning "vDisk: '$vmName' 의 디스크 Path 에 primera/unity 가 없어 합산에서 제외했습니다: $u"
        }
    } else {
        Write-Warning "vDisk: '$vmName' 의 디스크 정보가 없습니다. DISK(SSD)/DISK(SAS)=0 으로 출력됩니다."
    }

    $ip = ([string]$vm[$hIp]).Trim()

    # 결과 13개 값 (출력 열 B~N 순서 = 시스템현황 C,D,E,F,G,L,O,R,S,V,W,X,Y 순서)
    $values = @($vmName, $mangubun, $naeoebu, '클라우드', $nodeName, $group, $gubun,
                $hostname, $ip, $cpu, $mem, $ssdGb, $sasGb)
    $isNum = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $values.Count; $i++) {
        [void]$isNum.Add(($numericField[$i] -and (Test-IsNumber $values[$i])))
    }

    # --- 비교 ---
    $changed = New-Object System.Collections.ArrayList
    $changedFields = New-Object System.Collections.ArrayList
    $status = '비교안함'
    if ($null -ne $serviceFile) {
        if ($svcMap.ContainsKey($dKey)) {
            for ($i = 0; $i -lt $values.Count; $i++) {
                $same = Test-SameValue -A $values[$i] -B $svcMap[$dKey].Values[$i]
                [void]$changed.Add((-not $same))
                if (-not $same) { [void]$changedFields.Add($fieldNames[$i]) }
            }
            if ($changedFields.Count -eq 0) { $status = '동일' }
            else { $status = '변경' }
        } else {
            for ($i = 0; $i -lt $values.Count; $i++) { [void]$changed.Add($true) }
            $status = '신규'
        }
    } else {
        for ($i = 0; $i -lt $values.Count; $i++) { [void]$changed.Add($false) }
    }

    [void]$dataRows.Add(@{ Values = $values; IsNum = $isNum; Changed = $changed })

    $chgText = ''
    if ($status -eq '변경') { $chgText = $changedFields -join ', ' }
    elseif ($status -eq '신규') { $chgText = '(시스템현황에 없는 VM)' }
    [void]$summary.Add([pscustomobject]@{
        'VM이름'   = $vmName
        '상태'     = $status
        '변경항목' = $chgText
        '망구분'   = $mangubun
        '내/외부'  = $naeoebu
        '노드'     = $nodeName
        '그룹사'   = $group
        '구분'     = $gubun
        '호스트명' = $hostname
        'IP'       = $ip
        'CPU'      = $cpu
        'MEM'      = $mem
        'SSD'      = $ssdGb
        'SAS'      = $sasGb
    })
}

if ($dataRows.Count -eq 0) { throw "처리할 VM 데이터가 없습니다." }

# 시스템현황에만 있고 RVTools에 없는 VM 안내
if ($null -ne $serviceFile) {
    $missing = New-Object System.Collections.ArrayList
    foreach ($k in $svcOrder) {
        if (-not $rvVmSet.ContainsKey($k)) { [void]$missing.Add($k) }
    }
    if ($missing.Count -gt 0) {
        Write-Warning "시스템현황에는 있으나 RVTools에 없는 VM ($($missing.Count)건): $($missing -join ', ')"
    }
}

# ---------------------------------------------------------------------------
# 7. 결과 xlsx 생성 + 콘솔 요약
# ---------------------------------------------------------------------------
if ([string]::IsNullOrEmpty($OutFile)) {
    $dir = Split-Path -Parent $xlsxFile
    if ($null -ne $serviceFile) { $dir = Split-Path -Parent $serviceFile }
    $OutFile = Join-Path $dir ("시스템현황_비교_{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}

New-ComparisonWorkbook -Path $OutFile -DataRows $dataRows

Write-Host ""
Write-Host "=== 비교 결과 ($($dataRows.Count)건) ===" -ForegroundColor Green
$summary | Format-Table -AutoSize | Out-String -Width 400 | Write-Host

$chgCount = @($summary | Where-Object { $_.'상태' -eq '변경' }).Count
$newCount = @($summary | Where-Object { $_.'상태' -eq '신규' }).Count
$sameCount = @($summary | Where-Object { $_.'상태' -eq '동일' }).Count
if ($null -ne $serviceFile) {
    Write-Host "동일 $sameCount / 변경 $chgCount / 신규 $newCount" -ForegroundColor Green
}
Write-Host ""
Write-Host "결과 파일 생성 완료: $OutFile" -ForegroundColor Green
Write-Host "  - 기존 시스템현황과 달라진 셀은 노란색 음영으로 표시했습니다." -ForegroundColor Yellow
Write-Host "  - 원본 파일(RVTools/서비스현황)은 수정하지 않았습니다." -ForegroundColor Yellow
