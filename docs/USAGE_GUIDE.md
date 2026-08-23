# HWPX_wiz 사용 안내서

이 문서는 `HWPX_wiz`를 처음 사용하는 사람이 문서의 종류를 판별하고,
적절한 도구를 선택하고, 변환 결과를 확인하고, 문제가 생겼을 때 원인을
좁힐 수 있도록 작성한 상세 운영 안내서다.

이 저장소는 한글 문서 자체를 보관하는 저장소가 아니다. 공식 Kordoc과
공식 `hwpx` skill을 일관된 방식으로 호출하기 위한 Windows용 통합 계층이며,
구형 Word `.doc`를 `.docx`로 바꾸는 보조 변환기를 함께 제공한다.

> 기준 경로 예시: `D:\Code\Projects\HWPX_wiz`
> 기준 Kordoc 버전: `4.9.0`
> 대상 환경: Windows PowerShell, Microsoft Word desktop, Node.js/npm, `uv`

---

## 1. 먼저 도구를 선택한다

문서의 확장자와 목적에 따라 진입점이 달라진다.

| 목적 | 권장 진입점 | 비고 |
|---|---|---|
| HWP/HWPX/DOCX/PDF/XLSX 내용 읽기·비교·구조 추출 | 공식 Kordoc MCP | Codex에서 Kordoc 도구를 우선 사용 |
| 새 HWPX 작성 또는 기존 HWPX 편집 | `hwpx` skill | 편집 가능한 HWPX와 검증 절차를 함께 사용 |
| 구형 `.doc`를 `.docx`로 변환 | `tools/doc-to-docx/convert-doc-to-docx.ps1` | Microsoft Word COM 필요 |
| 구형 `.doc`를 `.docx`와 Markdown으로 변환 | `tools/doc-to-docx/convert-doc-to-md.ps1` | Word와 로컬 Kordoc CLI 필요 |
| 프로젝트 `inbox` 전체를 일괄 변환 | 루트의 `.bat` 파일 | 모든 인자를 무시하고 `inbox`만 처리 |
| 저장소 코드·테스트·잠금 파일 검증 | `tools/verify.ps1` | 설치는 하지 않고 검증만 수행 |

### 중요한 경계

- 구형 `.doc`를 Kordoc에 직접 넘기지 않는다. 먼저 Word로 `.docx`를 만든다.
- `Kordoc_helper` 애플리케이션이나 그 자체 MCP는 사용하지 않는다.
- 루트의 `convert-doc-to-docx.bat`, `convert-doc-to-md.bat`은 경로 인자를
  받는 실행기가 아니다. 인자를 전달해도 무시하고 프로젝트의 `inbox`만
  처리한다.
- 특정 파일이나 폴더를 변환할 때는 반드시 PowerShell 스크립트에 `-Path`를
  지정한다.
- 원본 문서는 삭제하거나 덮어쓰지 않는다.

---

## 2. 저장소 폴더와 파일 역할

저장소 루트는 다음과 같이 사용한다.

```text
D:\Code\Projects\HWPX_wiz\
├─ inbox\                 # 임시 입력 문서; 실제 파일은 Git에 포함하지 않음
├─ output\                # 별도 보고서·HWPX·결과를 모을 때 사용
├─ tools\
│  ├─ doc-to-docx\
│  │  ├─ convert-doc-to-docx.ps1
│  │  ├─ convert-doc-to-md.ps1
│  │  └─ README.md
│  ├─ kordoc\
│  │  ├─ package.json
│  │  ├─ package-lock.json
│  │  └─ node_modules\
│  └─ verify.ps1
├─ tests\
├─ pyproject.toml
├─ uv.lock
└─ README.md
```

`inbox/*`와 `output/*`는 실제 문서 작업 영역이므로 Git 추적 대상에서
제외된다. 개인정보나 업무 문서를 소스 파일, 테스트 fixture, 커밋에 넣지
않는다.

---

## 3. 최초 설정

### 3.1 저장소 루트로 이동

PowerShell을 열고 다음을 실행한다.

```powershell
Set-Location -LiteralPath "D:\Code\Projects\HWPX_wiz"
```

경로가 다르면 본인의 실제 clone 경로로 바꾼다.

### 3.2 Python 환경 준비

Python 패키지는 전역 `pip`로 설치하지 않는다. 이 프로젝트는 `uv.lock`을
기준으로 환경을 만든다.

```powershell
uv sync --locked
```

정상적으로 완료되면 다음 파일이 있어야 한다.

```text
.venv\Scripts\python.exe
```

Python 코드를 직접 실행할 때도 전역 `python`이나 `pip` 대신 다음 형식을
사용한다.

```powershell
uv run python .\path\to\script.py
```

### 3.3 로컬 Kordoc CLI 설치

Markdown 통합 변환기는 저장소에 고정된 로컬 Kordoc CLI를 사용한다.
`npx`로 매번 임의 버전을 실행하지 않는다.

```powershell
npm ci --prefix .\tools\kordoc
```

설치 후 버전을 확인한다.

```powershell
node .\tools\kordoc\node_modules\kordoc\dist\cli.js --version
```

출력 버전은 `4.9.0`이어야 한다. 설치가 끝나면 다음 파일이 있어야 한다.

```text
tools\kordoc\node_modules\.bin\kordoc.cmd
```

### 3.4 저장소 검증 도구 설치

검증기는 누락된 도구를 자동 설치하지 않는다. 최초 한 번 사용자 범위로
Pester와 PSScriptAnalyzer를 설치한다.

```powershell
Install-Module Pester -MinimumVersion 6.1.0 -Scope CurrentUser
Install-Module PSScriptAnalyzer -MinimumVersion 1.25.0 -Scope CurrentUser
```

설치 후 다음 명령으로 버전을 확인할 수 있다.

```powershell
Get-Module -ListAvailable Pester |
  Sort-Object Version -Descending |
  Select-Object -First 1 Name, Version

Get-Module -ListAvailable PSScriptAnalyzer |
  Sort-Object Version -Descending |
  Select-Object -First 1 Name, Version
```

### 3.5 Microsoft Word 준비

구형 `.doc` 변환은 Microsoft Word desktop의 COM 자동화를 사용한다.
따라서 다음 조건이 필요하다.

- Windows에서 실행할 것
- Microsoft Word desktop이 설치되어 있을 것
- `Word.Application` COM 등록이 정상일 것
- 변환 중 자동화에 사용할 수 있는 사용자 세션일 것

Word Online이나 단순히 `.docx` 파일을 열 수 있는 다른 프로그램만 설치된
환경으로는 이 변환기를 사용할 수 없다.

### 3.6 공식 MCP와 `hwpx` skill 확인

문서 읽기·분석은 공식 Kordoc MCP를 우선 사용한다.

```powershell
codex mcp list
```

목록에서 `kordoc` 항목의 명령은 검증된 버전인
`npx -y kordoc@4.9.0 mcp`를 사용해야 한다.

새 HWPX 작성이나 기존 HWPX 편집은 설치된 공식 `hwpx` skill을 사용한다.
skill이 없다면 프로젝트 규칙에 따라 별도 clone으로 설치하고, 설치된
버전을 확인한다.

```powershell
git clone --branch main `
  https://github.com/jkf87/hwpx-skill.git `
  "C:\Users\Hank\.agents\skills\hwpx"

git -C "C:\Users\Hank\.agents\skills\hwpx" describe --tags --exact-match HEAD
```

이미 clone이 있다면 작업 트리를 먼저 확인하고, upstream 갱신은
`git pull --ff-only`만 사용한다.

---

## 4. 가장 빠른 사용법: `inbox`

`inbox` 방식은 파일 하나를 특정하지 않고 프로젝트 폴더에 들어온 구형
문서를 일괄 처리할 때 사용한다.

### 4.1 DOCX만 만들기

1. `.doc` 파일을 `inbox`에 복사한다.
2. 저장소 루트에서 배치 파일을 실행한다.

```powershell
.\convert-doc-to-docx.bat
```

또는 파일 탐색기에서 `convert-doc-to-docx.bat`을 더블클릭할 수 있다.

기본 동작:

- `inbox` 바로 아래와 하위 폴더의 `.doc`를 찾는다.
- `.docx`가 없으면 원본 옆에 만든다.
- 기존 `.docx`가 있으면 기본적으로 건너뛴다.
- 원본 `.doc`는 그대로 둔다.
- 변환 결과는 다음과 같이 같은 폴더에 놓인다.

```text
inbox\보고서.doc
inbox\보고서.docx
```

### 4.2 DOCX와 Markdown을 함께 만들기

```powershell
.\convert-doc-to-md.bat
```

결과는 다음과 같다.

```text
inbox\보고서.doc
inbox\보고서.docx
inbox\보고서.md
```

이 실행은 먼저 Word로 `.docx`를 준비한 다음, 로컬 Kordoc CLI로 `.md`를
생성한다.

### 4.3 배치 파일에 경로를 전달하지 않는다

다음 명령은 `D:\문서`를 변환하지 않는다.

```powershell
.\convert-doc-to-docx.bat "D:\문서"
```

루트 배치 파일은 인자를 의도적으로 무시한다. 이 경우에도 `inbox`가
처리된다. 특정 경로가 목적이면 아래의 PowerShell 직접 실행법을 사용한다.

---

## 5. 특정 파일 또는 폴더 변환

### 5.1 특정 DOC 하나를 DOCX로 변환

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path "D:\문서\보고서.doc"
```

결과:

```text
D:\문서\보고서.doc
D:\문서\보고서.docx
```

### 5. 특정 DOC 하나를 Markdown까지 변환

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-md.ps1" `
  -Path "D:\문서\보고서.doc"
```

결과:

```text
D:\문서\보고서.doc
D:\문서\보고서.docx
D:\문서\보고서.md
```

### 5.3 폴더 바로 아래의 DOC만 변환

`-Recurse`가 없으면 지정한 폴더의 바로 아래 파일만 처리한다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path "D:\문서\구형자료"
```

### 5.4 하위 폴더까지 재귀 처리

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path "D:\문서\구형자료" `
  -Recurse
```

Markdown까지 만들려면 통합 실행기에 같은 옵션을 준다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-md.ps1" `
  -Path "D:\문서\구형자료" `
  -Recurse
```

폴더를 지정하면 확장자가 `.doc`인 파일만 처리한다. `.docx`, `.docm` 등은
입력으로 다시 처리하지 않는다.

### 5.5 특수문자가 있는 경로

PowerShell에서는 경로를 큰따옴표로 감싼다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-md.ps1" `
  -Path "D:\문서\보고서 & 최종본.doc"
```

`&`, 공백, 괄호, 한글이 들어간 경로도 이 방식으로 전달한다. 경로를
따옴표 없이 쓰면 `&`가 PowerShell 명령 연산자로 해석될 수 있다.

---

## 6. PowerShell 옵션 전체

### 6.1 DOC → DOCX 변환기

스크립트:

```text
tools\doc-to-docx\convert-doc-to-docx.ps1
```

문법:

```text
convert-doc-to-docx.ps1
  -Path <파일 또는 폴더>
  [-Recurse]
  [-Overwrite]
  [-LogPath <로그 파일>]
  [-DocumentKey <문서 암호>]
```

`-Password`는 `-DocumentKey`의 별칭이다.

| 옵션 | 필수 | 설명 |
|---|---:|---|
| `-Path` | 예 | `.doc` 파일 또는 `.doc`가 있는 폴더 |
| `-Recurse` | 아니오 | 폴더 지정 시 하위 폴더까지 포함 |
| `-Overwrite` | 아니오 | 기존 `.docx`를 성공 결과로 교체 |
| `-LogPath` | 아니오 | 처리 로그를 추가 기록할 파일 |
| `-DocumentKey` | 아니오 | 암호가 걸린 문서를 열 때 사용할 암호 |
| `-Password` | 아니오 | `-DocumentKey`의 별칭 |

예:

```powershell
# 기존 DOCX가 없을 때만 생성
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path "D:\문서\보고서.doc"

# 하위 폴더 포함, 기존 DOCX 교체, 로그 기록
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path "D:\문서" `
  -Recurse `
  -Overwrite `
  -LogPath "D:\문서로그\doc-conversion.log"

# 암호 문서
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path "D:\보안문서\계약서.doc" `
  -DocumentKey "문서암호"
```

`-LogPath`의 부모 폴더는 미리 존재해야 한다.

```powershell
New-Item -ItemType Directory -Force -Path "D:\문서로그"
```

로그 파일 경로를 원본 `.doc` 또는 결과 `.docx`와 동일하게 지정할 수
없다. 원본이나 결과를 로그로 덮어쓰는 상황을 방지하기 위해 변환 전에
실패한다.

### 6.2 DOC → DOCX → Markdown 통합 변환기

스크립트:

```text
tools\doc-to-docx\convert-doc-to-md.ps1
```

문법:

```text
convert-doc-to-md.ps1
  -Path <파일 또는 폴더>
  [-Recurse]
  [-Overwrite]
  [-KordocCliPath <kordoc.cmd 경로>]
```

| 옵션 | 필수 | 설명 |
|---|---:|---|
| `-Path` | 예 | `.doc` 파일 또는 `.doc`가 있는 폴더 |
| `-Recurse` | 아니오 | 폴더 지정 시 하위 폴더까지 포함 |
| `-Overwrite` | 아니오 | 기존 DOCX/Markdown을 갱신 |
| `-KordocCliPath` | 아니오 | 기본 로컬 CLI 대신 사용할 Kordoc 실행 파일 |

기본 Kordoc 경로는 다음이다.

```text
tools\kordoc\node_modules\.bin\kordoc.cmd
```

특수한 로컬 설치를 명시해야 할 때만 `-KordocCliPath`를 사용한다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-md.ps1" `
  -Path "D:\문서\보고서.doc" `
  -KordocCliPath ".\tools\kordoc\node_modules\.bin\kordoc.cmd"
```

일반적인 사용에서는 이 옵션을 생략하는 편이 안전하다. 기본 경로와
설치 버전이 저장소의 lockfile 기준선에 맞춰져 있기 때문이다.

### 6.3 도움말

PowerShell에서 전체 주석 기반 도움말을 확인한다.

```powershell
Get-Help ".\tools\doc-to-docx\convert-doc-to-docx.ps1" -Full
Get-Help ".\tools\doc-to-docx\convert-doc-to-md.ps1" -Full
```

간단한 매개 변수 문법만 보려면 다음을 사용한다.

```powershell
Get-Help ".\tools\doc-to-docx\convert-doc-to-docx.ps1" -Parameter *
Get-Help ".\tools\doc-to-docx\convert-doc-to-md.ps1" -Parameter *
```

---

## 7. 결과를 해석하는 방법

두 변환기는 진행 로그와 최종 요약을 출력하고, 마지막에는 구조화된
PowerShell 레코드도 출력한다.

### 7.1 DOCX 변환 결과

레코드의 주요 필드는 다음과 같다.

| 필드 | 의미 |
|---|---|
| `Status` | `Converted`, `Skipped`, `Failed` |
| `Source` | 원본 `.doc`의 절대 경로 |
| `Target` | 결과 `.docx`의 절대 경로 |
| `Reason` | 건너뛴 이유 등 |
| `Error` | 실패 시 오류 메시지 |

최종 요약은 다음 형태다.

```text
CONVERSION SUMMARY
  Converted : 1
  Skipped   : 0
  Failed    : 0
```

### 7.2 Markdown 통합 변환 결과

DOCX 경로를 나타내는 `Docx` 필드가 추가될 수 있다.

| 필드 | 의미 |
|---|---|
| `Status` | `Converted`, `Skipped`, `Failed` |
| `Source` | 원본 `.doc` |
| `Docx` | 중간 `.docx` |
| `Target` | 최종 `.md` |
| `Reason` | 최신 상태라 건너뛴 이유 등 |
| `Error` | 실패 시 오류 메시지 |

`Failed`가 하나라도 있으면 프로세스 종료 코드는 `1`이다. 모든 항목이
성공하거나 건너뛰면 종료 코드는 `0`이다. 입력 폴더에 `.doc`가 하나도
없어도 오류가 아니라 경고 후 종료 코드 `0`이다.

PowerShell에서 종료 코드를 직접 확인하려면:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-md.ps1" `
  -Path ".\inbox"

$LASTEXITCODE
```

---

## 8. 기존 결과와 `-Overwrite` 정책

### 8.1 기본 모드

안전하게 기존 결과를 보존하는 것이 기본이다.

| 상황 | DOCX 변환기 | Markdown 통합 변환기 |
|---|---|---|
| 결과가 없음 | 변환 | DOCX를 만든 뒤 Markdown 변환 |
| 결과가 있고 최신 | 건너뜀 | 최신 Markdown이면 건너뜀 |
| 결과가 있고 오래됨 | 기본 모드에서는 건너뜀 | 오래된 DOCX/Markdown이면 오류로 알리고 `-Overwrite` 요구 |
| 일부 결과만 있음 | 없는 결과만 처리 | 필요한 중간 결과를 사용하거나 생성 |

Markdown 변환기는 오래된 결과를 성공으로 오인해 조용히 건너뛰지 않는다.
`Existing DOCX is older` 또는 `Existing Markdown is older`와 비슷한 메시지가
나오면 원본을 확인한 뒤 `-Overwrite`를 명시한다.

### 8.2 덮어쓰기 모드

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-md.ps1" `
  -Path "D:\문서\보고서.doc" `
  -Overwrite
```

`-Overwrite`도 원본 `.doc` 자체를 수정하지 않는다. 변환기는 다음 순서로
동작한다.

1. 기존 결과가 있으면 rollback 또는 backup 임시 파일을 만든다.
2. Word 결과를 숨겨진 GUID 기반 임시 `.docx`에 저장한다.
3. Kordoc 결과를 숨겨진 GUID 기반 임시 `.md`에 저장한다.
4. 전체 단계가 성공하면 임시 파일을 최종 파일로 원자적으로 교체한다.
5. 성공 후 backup과 staging 파일을 정리한다.
6. 실패하면 기존 DOCX/Markdown을 복원하고 새로 만든 중간 결과를 제거한다.

따라서 변환 중 실패했다고 해서 기존 결과가 빈 파일이나 반쯤 생성된
파일로 바뀌는 것을 전제로 사용하지 않는다. 다만 저장 장치 오류나
권한·바이러스 백신 잠금처럼 운영체제가 파일 교체 자체를 거부하는
경우에는 오류 메시지와 잔류 파일을 확인해야 한다.

---

## 9. 암호 문서와 실패 처리

### 9.1 암호 문서

암호가 알려져 있다면 `-DocumentKey` 또는 `-Password`를 지정한다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path "D:\문서\암호문서.doc" `
  -Password "정확한 암호"
```

암호가 틀리면 Word의 대화형 암호 입력창을 기다리지 않고 실패 레코드와
비정상 종료 코드를 반환해야 한다. 원본 해시는 변경되지 않으며, 성공하지
못한 `.docx`는 남기지 않는다.

암호를 문서, 로그, 셸 기록, Git에 평문으로 남기지 않도록 주의한다.
공유 환경에서는 명령행 대신 안전한 별도 입력 방식을 사용하고, 현재
스크립트 계약이 허용하는 범위 안에서만 자동화한다.

### 9.2 손상된 DOC

손상되었거나 실제 Word 문서가 아닌 파일은 변환이 실패한다.

- `Status = Failed`를 확인한다.
- 원본 `.doc`의 해시와 존재 여부를 확인한다.
- 동일한 이름의 불완전한 `.docx`가 생성되지 않았는지 확인한다.
- Word에서 수동으로 열 수 있는지 확인한다.

### 9.3 로그 경로 오류

다음 조건에서는 Word를 시작하기 전에 실패한다.

- `-LogPath`의 부모 폴더가 존재하지 않음
- 로그 경로가 입력 `.doc`와 충돌함
- 로그 경로가 결과 `.docx`와 충돌함

예:

```powershell
# 먼저 부모 폴더를 만든다.
New-Item -ItemType Directory -Force -Path "D:\변환로그"

powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path "D:\문서\보고서.doc" `
  -LogPath "D:\변환로그\보고서.log"
```

---

## 10. Word 자동화 안전 동작

DOC 변환기는 Word COM을 자동화할 때 다음 정책을 적용한다.

- Word 창을 표시하지 않는다.
- 문서를 읽기 전용으로 연다.
- 매크로 실행을 강제로 비활성화한다.
- Word 경고 대화상자를 끈다.
- 문서 열 때 자동 링크 업데이트를 끈다.
- 최근 문서 목록에 자동으로 추가하지 않는다.
- 변환 성공 전까지 최종 `.docx`를 교체하지 않는다.
- 변환기가 소유한 Word 프로세스만 종료 대상으로 삼는다.
- 변환기가 소유하지 않은 별도 Word 프로세스는 강제 종료하지 않는다.

변환 중에는 자동화 대상 문서가 잠길 수 있다. 같은 원본을 Word에서
동시에 편집하지 말고, 변환이 끝난 뒤 결과를 확인한다.

---

## 11. HWP/HWPX 문서 작업

### 11.1 읽기·분석·비교

HWP, HWPX, DOCX, PDF, XLSX의 내용을 읽거나 구조를 추출할 때는 Codex의
공식 Kordoc MCP를 우선 사용한다.

요청 예:

```text
이 파일의 표와 문단 구조를 공식 Kordoc MCP로 추출하고,
원본 순서를 유지한 Markdown 요약을 만들어줘.
```

```text
두 HWPX 파일을 공식 Kordoc MCP로 비교해서
추가·삭제·변경된 표와 문단을 구분해줘.
```

파일 경로를 요청에 명시하면 원본 파일과 같은 폴더를 기준으로 작업한다.
중앙 결과 폴더가 필요하면 `output`에 저장하도록 명시한다.

### 11.2 HWPX 생성·편집

편집 가능한 HWPX를 만들거나 수정할 때는 `hwpx` skill을 명시한다.

요청 예:

```text
hwpx skill을 사용해서 이 초안을 편집 가능한 공문 HWPX로 만들고,
결과를 D:\Code\Projects\HWPX_wiz\output에 저장한 뒤
namespace와 layout 검증까지 수행해줘.
```

Python 보조 코드가 필요하면 다음 원칙을 지킨다.

```powershell
uv run python .\scripts\build_hwpx.py
```

전역 `pip install`, 임의의 Python 환경, 프로젝트 밖의 자체 HWPX 변환기를
사용하지 않는다. 최종 HWPX는 한컴오피스에서 직접 열어 페이지 나눔,
표, 글꼴, 이미지와 빈 페이지를 눈으로 확인한다.

### 11.3 결과 저장 원칙

- 원본 HWP/HWPX/DOCX/PDF/XLSX는 덮어쓰지 않는다.
- 별도 생성물은 사용자가 지정한 `output`에 둔다.
- 임시 입력은 `inbox`에 둘 수 있다.
- 실제 업무 문서와 결과 파일은 Git에 추가하지 않는다.
- 문서에 개인정보나 비밀정보가 있으면 테스트 fixture로 복사하지 않는다.

---

## 12. 저장소 검증

### 12.1 표준 검증 명령

개발·정비 후 저장소 루트에서 다음 한 명령을 실행한다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\verify.ps1"
```

검증기는 다음을 순서대로 확인한다.

1. Pester `6.1.0` 이상
2. PSScriptAnalyzer `1.25.0` 이상
3. PATH의 `uv`
4. PATH의 Node.js와 npm
5. 등록된 `Word.Application` COM
6. `tools\kordoc\package.json`, lockfile, 로컬 CLI
7. Kordoc 선언·lockfile·실행 파일 버전 `4.9.0`
8. `npm ci --dry-run --ignore-scripts`
9. 전체 Pester 테스트
10. `tools`와 `tests`의 PSScriptAnalyzer
11. `uv lock --check`

마지막에 다음 문장이 나오면 통과다.

```text
Repository verification passed.
```

이 검증 명령은 누락된 도구를 설치하지 않는다. 실패 메시지의 설치 명령을
확인하고 필요한 준비를 한 뒤 다시 실행한다.

### 12.2 사전 조건만 확인

Word, Pester 전체 테스트까지 실행하지 않고 환경만 점검할 때:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\verify.ps1" `
  -PrerequisiteCheckOnly
```

정상이면 다음을 출력한다.

```text
Prerequisite check passed.
```

### 12.3 개별 검사

실패 원인을 좁힐 때만 구성 요소를 나누어 실행한다.

```powershell
Import-Module Pester -MinimumVersion 6.1.0 -Force
Invoke-Pester -Path .\tests -Output Detailed
```

```powershell
Import-Module PSScriptAnalyzer -MinimumVersion 1.25.0 -Force
Invoke-ScriptAnalyzer -Path .\tools -Recurse
Invoke-ScriptAnalyzer -Path .\tests -Recurse
```

```powershell
uv lock --check
```

전체 게이트를 통과하지 않은 상태에서 “검증 완료”라고 보고하지 않는다.

### 12.4 실제 표면 수동 확인

검증 게이트가 통과해도 문서 변환이 실제로 사용 가능한지 다음 표면을
확인하는 것이 좋다.

```powershell
Get-Help ".\tools\doc-to-docx\convert-doc-to-docx.ps1" -Full
Get-Help ".\tools\doc-to-docx\convert-doc-to-md.ps1" -Full
```

확인할 항목:

- 실제 `.doc` 하나가 `.docx`로 열리는가
- Markdown 통합 변환 결과가 생성되는가
- 특수문자·한글 경로가 동작하는가
- 손상 문서가 실패로 보고되는가
- 틀린 암호가 대화상자 없이 실패하는가
- 기존 결과가 실패 시 보존되는가
- 임시 `.backup`, `.rollback`, staging 파일이 남지 않는가
- 변환 종료 후 소유하지 않은 Word를 종료하지 않는가

---

## 13. 문제 해결 표

| 메시지·증상 | 원인 | 조치 |
|---|---|---|
| `Path does not exist` | 입력 경로가 없음 | 경로를 확인하고 큰따옴표로 감싼다 |
| `Input file must have a .doc extension` | `.doc`가 아닌 파일을 직접 전달 | 구형 `.doc`만 전달한다 |
| `.docx`가 생기지 않고 `Failed` | Word가 문서를 열지 못함, 손상, 암호 오류 | Word에서 원본을 수동 확인하고 오류 필드를 읽는다 |
| Word 암호 입력창이 나타남 | 암호가 필요하거나 자동화 조건이 맞지 않음 | `-DocumentKey`/`-Password`를 지정하고 Word 세션을 확인한다 |
| `Existing DOCX is older` | 기존 중간 결과가 원본보다 오래됨 | 확인 후 `-Overwrite`를 사용한다 |
| `Existing Markdown is older` | Markdown이 DOC 또는 DOCX보다 오래됨 | 확인 후 `-Overwrite`를 사용한다 |
| `Local Kordoc CLI is missing` | npm 설치가 안 됨 | `npm ci --prefix .\tools\kordoc` 실행 |
| Kordoc 버전 gate 실패 | package, lockfile, 설치 CLI 버전 불일치 | 로컬 설치를 lockfile 기준으로 재설치한다 |
| `node`/`npm`을 찾을 수 없음 | PATH에 Node.js가 없음 | Node.js 설치 후 PowerShell을 다시 연다 |
| `Word.Application` COM 오류 | Microsoft Word desktop 미설치 또는 COM 미등록 | 지원되는 Windows Word 환경에서 실행 |
| `LogPath parent does not exist` | 로그 부모 폴더가 없음 | `New-Item -ItemType Directory`로 먼저 만든다 |
| `.bat`에 준 경로가 무시됨 | 루트 BAT는 의도적으로 무인자 `inbox` 전용 | `.ps1 -Path`를 직접 사용한다 |
| 결과가 없는데 종료 코드 0 | 입력 폴더에 `.doc`가 없음 | 입력 위치와 `-Recurse` 여부를 확인한다 |
| Pester 모듈 오류 | Pester가 없거나 버전이 낮음 | 사용자 범위로 최소 버전을 설치한다 |
| PSScriptAnalyzer 오류 | analyzer가 없거나 경고가 있음 | 오류 파일과 규칙을 확인하고 코드 수정 후 재실행한다 |
| `uv lock --check` 실패 | 잠금 파일과 프로젝트 정의가 불일치 | `uv` 환경과 `uv.lock`을 확인한다 |

### 실패 시 우선 확인 순서

1. 명령을 저장소 루트에서 실행했는지 확인한다.
2. 입력 경로와 파일 확장자를 확인한다.
3. `$LASTEXITCODE`를 확인한다.
4. `Status`, `Source`, `Target`, `Error` 레코드를 확인한다.
5. Word와 Kordoc을 분리해 DOCX 중간 결과가 정상인지 확인한다.
6. 기존 결과가 있다면 시간을 비교하고 필요할 때만 `-Overwrite`를 사용한다.
7. 마지막으로 `tools\verify.ps1`를 다시 실행한다.

---

## 14. 고급: 파이프라인과 dot-source 호출

PowerShell 호출자는 스크립트를 dot-source하여 구조화된 레코드를 받을 수
있다. 이 방식은 다른 PowerShell 자동화에서 상태를 직접 검사할 때만
사용한다.

```powershell
$records = . ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path "D:\문서\보고서.doc"

$records | Format-Table Status, Source, Target, Reason, Error -AutoSize

if (@($records | Where-Object Status -eq "Failed").Count -gt 0) {
    throw "문서 변환 실패"
}
```

Markdown 통합 변환은 다음과 같이 `Docx` 필드도 확인할 수 있다.

```powershell
$records = . ".\tools\doc-to-docx\convert-doc-to-md.ps1" `
  -Path "D:\문서\보고서.doc"

$records |
  Select-Object Status, Source, Docx, Target, Reason, Error |
  Format-Table -AutoSize
```

dot-source를 사용할 때도 원본을 덮어쓰지 않으며, 호출자의 현재 위치를
변환기가 임의로 바꾸지 않는 계약을 유지한다.

---

## 15. 선택 사항: Herdr + OMO 작업 흐름

문서 변환만 할 때는 이 절이 필요 없다. PLAN/BUILD/VERIFY pane을 사용해
저장소 작업을 관리할 때만 참고한다.

### 15.1 Combo pane 시작

Herdr 안에서 저장소를 열고:

```powershell
Set-Location -LiteralPath "D:\Code\Projects\HWPX_wiz"
herdr
```

Herdr가 관리하는 pane에서:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\herdr\start-combo.ps1" `
  -Focus
```

구성만 확인하고 pane을 만들지 않으려면:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\tools\herdr\start-combo.ps1" `
  -ValidateOnly
```

구성은 다음 세 lane이다.

```text
PLAN / Sol    : 요구사항과 GOAL.md 작성
BUILD / Luna  : 구현과 테스트
VERIFY / Sol  : 독립 검증과 요구사항 감사
```

launcher는 명령을 입력해 두지만 자동으로 autonomous turn을 시작하지
않는다. 원하는 pane에서 Enter를 눌러 실행한다.

### 15.2 권장 순서

1. PLAN에서 요구사항과 acceptance criterion을 `GOAL.md`로 정리한다.
2. BUILD에서 `AGENTS.md`와 `GOAL.md`를 읽고 구현한다.
3. BUILD에서 실제 테스트와 `tools\verify.ps1`를 실행한다.
4. VERIFY에서 diff, 테스트 출력, 실제 동작을 독립적으로 감사한다.
5. FAIL 또는 PARTIAL이면 BUILD로 돌아가 수정하고 VERIFY를 반복한다.

문서 변환 자체는 이 pane workflow 없이도 PowerShell 명령만으로 실행할 수
있다.

---

## 16. 운영 전 체크리스트

### 처음 사용하는 경우

- [ ] 저장소 루트에서 실행하고 있는가?
- [ ] `uv sync --locked`를 완료했는가?
- [ ] `npm ci --prefix .\tools\kordoc`를 완료했는가?
- [ ] `node ...\cli.js --version`이 `4.9.0`인가?
- [ ] Pester와 PSScriptAnalyzer 최소 버전이 설치되어 있는가?
- [ ] Microsoft Word desktop이 설치되어 있는가?
- [ ] Kordoc MCP와 `hwpx` skill이 공식 경로인가?

### DOC 변환 전

- [ ] 원본이 실제 Word 97–2003 `.doc`인가?
- [ ] 원본을 Word에서 닫았는가?
- [ ] 결과를 둘 위치에 쓰기 권한이 있는가?
- [ ] 기존 `.docx` 또는 `.md`가 있으면 갱신이 정말 필요한가?
- [ ] 필요할 때만 `-Overwrite`를 사용했는가?
- [ ] 로그를 쓸 경우 부모 폴더가 존재하는가?
- [ ] 암호가 필요한 경우 안전하게 전달할 방법이 준비되었는가?

### 변환 후

- [ ] 종료 코드가 `0`인가?
- [ ] 요약의 `Failed`가 `0`인가?
- [ ] `.docx`가 Word에서 열리는가?
- [ ] `.md`가 예상 문단·표·텍스트를 포함하는가?
- [ ] 원본 파일의 수정 시각과 해시가 의도치 않게 바뀌지 않았는가?
- [ ] 숨은 staging, backup, rollback 파일이 남지 않았는가?
- [ ] 실제 문서와 결과가 Git에 추가되지 않았는가?
- [ ] 코드나 설정을 바꿨다면 `tools\verify.ps1`를 통과했는가?

---

## 17. 관련 문서

- 프로젝트 개요와 기본 명령: [`README.md`](../README.md)
- 변환기 간단 안내: [`tools/doc-to-docx/README.md`](../tools/doc-to-docx/README.md)
- 공식 upstream과 버전 정책: [`UPSTREAMS.md`](UPSTREAMS.md)
- Herdr + OMO workflow: [`HERDR_OMO_WORKFLOW.md`](HERDR_OMO_WORKFLOW.md)
- 정비 역사 기록: [`MAINTENANCE_HANDOFF.md`](MAINTENANCE_HANDOFF.md)
- 다음 개선 순서: [`ROADMAP.md`](ROADMAP.md)
- 저장소 검증 스크립트: [`../tools/verify.ps1`](../tools/verify.ps1)
