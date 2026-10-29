<#
    시스템현황 비교 실행기
    - 같은 폴더의 '시스템현황_실행.bat' 을 더블클릭하면 이 파일이 실행됩니다.
    - RVTools 파일과 서비스현황 파일을 비교하여 "비교 결과 xlsx"를 새로 만듭니다.
      (원본 파일은 수정하지 않습니다)
    - 아래 [설정] 부분만 고쳐서 사용하세요. (메모장으로 열어 수정 후 저장)
#>

# ==========================[ 설정 - 여기만 수정 ]==========================
# 파일명만 적으면 이 파일이 있는 폴더에서 찾습니다. (전체 경로도 가능)
# ''(빈 값)으로 두면 아래 자동 탐색 패턴으로 폴더에서 최신 파일을 자동 선택합니다.

$VrtoolsFile = ''      # 예: 'RVTools_export_all_2026-08-27_09.20.19.xlsx'
$ServiceFile = ''      # 예: '신한DS_클라우드_서비스현황_V1.08_2608(0827수정).xlsx'
$OutputFile  = ''      # 결과 파일명. 빈 값이면 자동: 시스템현황_비교_일시.xlsx

$VrtoolsPattern = 'RVTools*.xlsx'       # 자동 탐색 패턴 (VRtools 추출 파일)
$ServicePattern = '*서비스현황*.xlsx'   # 자동 탐색 패턴 (서비스현황 파일)

$TargetSheetName = '시스템현황'         # 서비스현황 파일에서 비교할 탭 이름
# ========================================================================

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot

try {
    # 본 스크립트 위치 확인
    $main = $null
    foreach ($name in @('Export-SystemStatusRow.ps1', 'SystemStatus.ps1')) {
        $p = Join-Path $here $name
        if (Test-Path -LiteralPath $p) { $main = $p; break }
    }
    if ($null -eq $main) {
        throw "같은 폴더에 Export-SystemStatusRow.ps1 이 없습니다. 본 스크립트를 이 폴더에 함께 넣어주세요. (폴더: $here)"
    }

    function Resolve-InputFile {
        param([string]$Name, [string]$Pattern, [string]$Label)
        if ($Name -ne '') {
            $p = $Name
            if (-not [System.IO.Path]::IsPathRooted($p)) { $p = Join-Path $here $p }
            if (-not (Test-Path -LiteralPath $p)) {
                throw "$Label 파일을 찾을 수 없습니다: $p`nRun-SystemStatus.ps1 상단 설정의 파일명을 확인해주세요."
            }
            return (Resolve-Path -LiteralPath $p).ProviderPath
        }
        # 자동 탐색: 패턴에 맞는 파일 중 가장 최근 파일 (엑셀 임시파일 ~$*, 이전 비교결과 제외)
        $cands = @(Get-ChildItem -Path $here -Filter $Pattern -File |
                   Where-Object { $_.Name -notlike '~$*' -and $_.Name -notlike '시스템현황_비교_*' } |
                   Sort-Object LastWriteTime -Descending)
        if ($cands.Count -eq 0) {
            throw "$Label 파일을 찾지 못했습니다.`n이 폴더에 '$Pattern' 형태의 파일을 넣거나, Run-SystemStatus.ps1 상단 설정에 파일명을 직접 적어주세요. (폴더: $here)"
        }
        if ($cands.Count -gt 1) {
            Write-Host "$Label 파일이 $($cands.Count)개 있어 가장 최근 파일을 사용합니다: $($cands[0].Name)" -ForegroundColor Yellow
        }
        return $cands[0].FullName
    }

    $src = Resolve-InputFile -Name $VrtoolsFile -Pattern $VrtoolsPattern -Label 'VRtools'
    $svc = Resolve-InputFile -Name $ServiceFile -Pattern $ServicePattern -Label '서비스현황'

    Write-Host "VRtools 파일   : $src" -ForegroundColor Cyan
    Write-Host "서비스현황 파일: $svc (읽기 전용)" -ForegroundColor Cyan
    Write-Host ""

    if ($OutputFile -ne '') {
        $outPath = $OutputFile
        if (-not [System.IO.Path]::IsPathRooted($outPath)) { $outPath = Join-Path $here $outPath }
        & $main -VrtoolsPath $src -ServicePath $svc -TargetSheet $TargetSheetName -OutFile $outPath
    } else {
        & $main -VrtoolsPath $src -ServicePath $svc -TargetSheet $TargetSheetName
    }
} catch {
    Write-Host ""
    Write-Host "[오류] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
