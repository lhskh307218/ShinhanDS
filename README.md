# 시스템현황 비교 도구

VRtools(RVTools)로 추출한 VM 정보를 읽어, 신한DS 클라우드 서비스현황 파일의 "시스템현황" 탭 형식에 맞게 값을 가공하고, 기존 값과 달라진 부분을 노란색으로 표시한 비교용 엑셀 파일을 만들어주는 도구입니다.

**폐쇄망(인터넷 차단) 환경 전용으로 제작**되었습니다. Excel 설치, 외부 PowerShell 모듈(ImportExcel 등) 모두 필요 없고, Windows에 기본 포함된 PowerShell과 .NET만으로 동작합니다.

## 목적

- vInfo/vDisk 원본 데이터를 매번 손으로 가공(디스크 종류 분류, 망구분 추출, 호스트명 추출 등)하는 반복 작업을 없애기 위함
- 서비스현황 파일을 열어 값을 직접 덮어쓰다가 함수(수식) 셀을 망가뜨리는 사고를 방지하기 위함 → **원본 파일은 읽기만 하고 절대 수정하지 않습니다.** 결과는 항상 새 엑셀 파일로 따로 생성됩니다.
- 매번 바뀌는 VM 값(CPU/MEM/디스크 증설, 신규 VM 추가 등)을 사람이 한 줄씩 대조하지 않고, 달라진 셀만 한눈에 보이게 하기 위함

## 구성 파일

| 파일 | 역할 |
|---|---|
| [Export-SystemStatusRow.ps1](Export-SystemStatusRow.ps1) | 실제 로직이 들어있는 메인 스크립트. RVTools 읽기 → 값 가공 → 서비스현황과 비교 → 결과 xlsx 생성까지 전부 수행 |
| [Run-SystemStatus.ps1](Run-SystemStatus.ps1) | 실행기 겸 설정 파일. 파일명을 매번 입력하지 않도록 자동 탐색 기능을 붙여 메인 스크립트를 호출 |
| [시스템현황_실행.bat](시스템현황_실행.bat) | 더블클릭용 배치 파일. 내부적으로 Run-SystemStatus.ps1을 실행 |
| [시스템현황_비교도구.zip](시스템현황_비교도구.zip) | 위 3개 파일을 묶은 압축본 (폐쇄망으로 옮길 때 사용) |

세 파일(ps1 2개 + bat 1개)은 **반드시 같은 폴더**에 있어야 합니다.

## 실행 방법

### 방법 1. 더블클릭 (권장)

1. 세 파일(`Export-SystemStatusRow.ps1`, `Run-SystemStatus.ps1`, `시스템현황_실행.bat`)이 있는 폴더에 RVTools 추출 파일과 서비스현황 파일을 넣습니다.
2. 서비스현황 파일이 엑셀에서 열려 있다면 닫습니다. (읽기만 하므로 열려있어도 대부분 동작하지만, 안전하게 닫고 실행하는 것을 권장합니다)
3. **`시스템현황_실행.bat`을 더블클릭**합니다.

파일명은 신경 쓰지 않아도 됩니다. `Run-SystemStatus.ps1`이 폴더 안에서 `RVTools*.xlsx` 패턴과 `*서비스현황*.xlsx` 패턴에 맞는 파일을 자동으로 찾고, 같은 패턴 파일이 여러 개면 가장 최근에 수정된 파일을 사용합니다.

특정 파일을 지정해서 쓰고 싶으면 `Run-SystemStatus.ps1`을 메모장으로 열어 상단 설정 부분만 고치면 됩니다.

```powershell
$VrtoolsFile = ''      # 예: 'RVTools_export_all_2026-08-27_09.20.19.xlsx'
$ServiceFile = ''      # 예: '신한DS_클라우드_서비스현황_V1.08_2608(0827수정).xlsx'
$OutputFile  = ''      # 결과 파일명. 빈 값이면 자동: 시스템현황_비교_일시.xlsx
$TargetSheetName = '시스템현황'   # 서비스현황 파일에서 비교할 탭 이름
```

### 방법 2. 명령어 직접 실행

```powershell
powershell -ExecutionPolicy Bypass -File .\Export-SystemStatusRow.ps1 -VrtoolsPath ".\RVTools_export_all_....xlsx" -ServicePath ".\신한DS_클라우드_서비스현황_....xlsx"
```

파일명에 괄호·공백·한글이 있으면 반드시 따옴표로 감싸야 합니다.

주요 옵션:

| 옵션 | 설명 |
|---|---|
| `-VrtoolsPath` (필수) | vInfo/vDisk 탭이 있는 RVTools 추출 파일 |
| `-ServicePath` | 비교 대상 서비스현황 파일. 생략하면 비교 없이 결과만 생성 |
| `-TargetSheet` | 서비스현황 파일에서 비교할 탭 이름 (기본값: `시스템현황`) |
| `-OutFile` | 결과 파일 경로 (생략 시 자동 생성) |
| `-RawMemory` | Memory 값을 MiB→GB로 변환하지 않고 원본 그대로 출력 |

## 결과물

실행 폴더(또는 서비스현황 파일 폴더)에 **`시스템현황_비교_YYYYMMDD_HHMMSS.xlsx`** 파일이 새로 생성됩니다. (원본 RVTools/서비스현황 파일은 전혀 수정되지 않습니다)

- Sheet1에 VM별 한 줄씩, 다음 13개 값이 B~N열에 빈 칸 없이 연속으로 출력됩니다.

  `Lookup용(VM이름) · 망구분 · 내/외부 · 구성형태 · 노드 · 그룹사 · 구분 · 호스트명 · IP · CPU(Core) · MEM(GB) · DISK(SSD) · DISK(SAS)`

- `-ServicePath`를 지정한 경우, 서비스현황 탭의 기존 값과 셀 단위로 비교하여:
  - **값이 달라진 셀만 노란색 음영**으로 표시
  - 서비스현황에 없는 **신규 VM은 행 전체가 노란색**
  - 변경이 없으면 음영 없이 그대로 출력
- 콘솔에는 VM별 비교 결과 표(동일/변경/신규, 변경된 항목 목록)와 전체 건수 요약이 출력됩니다.
- 서비스현황에는 있지만 RVTools에는 없는 VM(삭제되었을 가능성)이 있으면 경고 메시지로 따로 알려줍니다.

## 값 가공 규칙

| 결과 항목 | 계산 방식 |
|---|---|
| DISK(SSD) / DISK(SAS) | 같은 VM의 vDisk 행들을 Path 기준으로 분류 (경로에 `primera` 포함 → SSD, `unity` 포함 → SAS) 후 Capacity MiB를 합산하여 GB로 변환 |
| 망구분 / 내·외부 / 그룹사 | vInfo의 `Network #1` 값을 `_`로 분리. 토큰 중 `T`=개발, `P`=운영 / `I`,`A`=내부, `E`=외부로 판별, 첫 토큰은 그룹사 |
| 호스트명 | `DNS Name`에서 첫 번째 `.` 앞부분 |
| 노드 | `Host`에서 첫 번째 `.` 앞부분 |
| 구성형태 | 항상 `클라우드` 고정 |
| MEM(GB) | vInfo의 Memory(MiB) 값을 1024로 나눠 GB로 변환 (`-RawMemory` 지정 시 변환 없이 원본 사용) |

## 환경 요구사항

- Windows PowerShell 5.1 이상 (Windows 10/11 기본 내장)
- .NET Framework 4.5 이상 (대부분의 Windows에 기본 포함)
- 인터넷 연결, Excel 설치, 추가 모듈 설치 **불필요**
