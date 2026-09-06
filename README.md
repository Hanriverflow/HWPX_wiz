# HWPX_wiz

한글 문서 작업은 활발히 업데이트되는 공식 upstream 두 개를 중심으로 운영합니다.

- [Kordoc](https://github.com/chrisryugj/kordoc): HWP/HWPX/DOCX/PDF/XLSX 등의 읽기, 구조 추출, 비교, 양식 분석과 변환
- [hwpx-skill](https://github.com/jkf87/hwpx-skill): 편집 가능한 HWPX 생성·수정과 품질 검증
- 이 프로젝트의 DOC 변환기: Kordoc이 직접 지원하지 않는 구형 Word `.doc`를 `.docx`로 선행 변환

기존 `Kordoc_helper` 애플리케이션과 자체 MCP는 사용하지 않습니다.

정비 이력은 [`docs/MAINTENANCE_HANDOFF.md`](docs/MAINTENANCE_HANDOFF.md), 다음 개선 순서는 [`docs/ROADMAP.md`](docs/ROADMAP.md)를 참고합니다.
Herdr에서 PLAN/BUILD/VERIFY OMO 세션을 운영하는 방법은 [`docs/HERDR_OMO_WORKFLOW.md`](docs/HERDR_OMO_WORKFLOW.md)를 따릅니다.

> **처음 사용하는 분이라면:** [HWPX_wiz 실전 HTML 가이드 바로 보기](https://hanriverflow.github.io/HWPX_wiz/)
> · [HTML 원본 파일](HWPX_wiz_easy_guide.html)

## 프로젝트 성격과 원본 프로젝트 고지

**HWPX_wiz는 다른 개발자가 만든 두 원본 프로젝트를 fork/기반으로
활용하여, Windows 문서 작업의 사용성과 접근성을 높인 통합 프로젝트입니다.**
이 저장소는 해당 원본 프로젝트 자체이거나 원 저작자를 대신하는 공식 배포본이
아닙니다.

- **[Kordoc](https://github.com/chrisryugj/kordoc)** — `chrisryugj/kordoc`의
  원본 프로젝트를 기반으로 문서 읽기·구조 추출·비교·변환 작업을 연결합니다.
- **[hwpx-skill](https://github.com/jkf87/hwpx-skill)** — `jkf87/hwpx-skill`의
  원본 프로젝트를 기반으로 편집 가능한 HWPX 작성·수정·검증을 연결합니다.

각 원본 프로젝트의 저작권, 라이선스와 원 저작자 표기는 해당 upstream
프로젝트를 따릅니다. HWPX_wiz의 역할은 두 upstream을 Windows용 DOC 변환기,
검증 도구와 운영 문서로 통합하여 사용성을 높이는 것입니다.
HWPX_wiz 자체 통합 계층은 [MIT License](LICENSE)로 배포합니다.

## 역할 분담

| 작업 | 사용할 도구 |
|---|---|
| HWP/HWPX/DOCX/PDF/XLSX 내용 파악·비교 | Codex의 공식 Kordoc MCP |
| 새 HWPX 작성, 기존 HWPX 편집·검증 | Codex의 `hwpx` skill + 이 프로젝트의 `uv` 환경 |
| Kordoc `generate`/`fill`/`patch`/`validate` | upstream 보조 기능. 명시적 호환성 검토 전에는 편집 가능한 HWPX 작업의 기본 경로로 사용하지 않음 |
| `inbox`의 구형 `.doc` 변환 | 인자를 무시하는 `convert-doc-to-docx.bat` 또는 `convert-doc-to-md.bat` |
| 특정 파일·폴더의 구형 `.doc` 변환 | `tools/doc-to-docx/convert-doc-to-docx.ps1` 또는 `convert-doc-to-md.ps1`에 `-Path` 지정 |
| 경로가 지정된 문서 | 원본과 같은 폴더에 결과 저장 |
| 경로 없는 DOC 임시 작업 | `inbox`에서 찾고 결과도 해당 DOC 옆에 저장 |
| 별도 생성 문서·보고서·로그 | 필요할 때 `output` 사용 |

Kordoc CLI에 HWPX 쓰기·검증 명령이 있어도 이 프로젝트의 기본 역할은
Kordoc=MCP 읽기·구조 추출·비교·일반 Markdown 변환, hwpx skill=편집 가능한
HWPX 생성·수정·namespace/layout/한컴 검증이다. 겹치는 Kordoc 명령은 별도
작업에서 결과 호환성을 검증하고 명시적으로 선택할 때만 사용한다.

## 처음 한 번 또는 환경 복원

Python은 전역 설치 대신 `uv`로만 관리합니다.

```powershell
cd D:\Code\Projects\HWPX_wiz
uv sync --locked
```

현재 환경은 Python 3.12, `python-hwpx`, `lxml`, `pywin32`를 잠금 파일 기준으로 설치합니다.
hwpx skill의 Python 스크립트도 저장소 루트에서 `uv run python`으로 실행해
이 잠금 환경의 `hwpx`, `lxml`, `win32com`을 사용합니다.

공식 Kordoc MCP 등록 상태는 다음처럼 확인합니다.

```powershell
codex mcp list
```

`kordoc` 항목은 이 저장소 lock과 같은 CLI를 사용하도록
`node.exe <clone 경로>\tools\kordoc\node_modules\kordoc\dist\cli.js mcp`로 등록합니다. 새 Codex 작업에서
Kordoc 도구를 사용하면 됩니다.

공식 `hwpx` skill이 없다면 다음 위치에 설치합니다.

```powershell
git clone --branch main https://github.com/jkf87/hwpx-skill.git C:\Users\Hank\.agents\skills\hwpx
git -C C:\Users\Hank\.agents\skills\hwpx describe --tags --exact-match HEAD
```

## 다른 PC에서 동일하게 사용하기

이 저장소는 다른 PC에 clone한 뒤 로컬 실행 환경만 준비하면 같은 방식으로
사용할 수 있습니다. 저장소에 포함된 스크립트, 테스트, `uv.lock`,
`tools/kordoc/package-lock.json`과 운영 문서는 공유되지만 `.venv`,
`node_modules`, Codex MCP 등록, `hwpx` skill과 Herdr/OMO 설치는 PC별로
다시 준비해야 합니다.

### Windows 기준 설치 순서

다른 PC에서 다음 명령을 실행합니다. 설치 경로는 해당 PC에 맞게 바꿔도
되지만 이후 명령은 clone한 저장소 루트에서 실행해야 합니다.

```powershell
git clone https://github.com/Hanriverflow/HWPX_wiz.git `
  "D:\Code\Projects\HWPX_wiz"
Set-Location -LiteralPath "D:\Code\Projects\HWPX_wiz"

uv sync --locked
npm ci --prefix .\tools\kordoc

Install-Module Pester -MinimumVersion 6.1.0 -Scope CurrentUser
Install-Module PSScriptAnalyzer -MinimumVersion 1.25.0 -Scope CurrentUser

node .\tools\kordoc\node_modules\kordoc\dist\cli.js --version
$kordocCli = (Resolve-Path `
  .\tools\kordoc\node_modules\kordoc\dist\cli.js).Path
codex mcp add kordoc -- node.exe $kordocCli mcp
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\verify.ps1
```

마지막 검증에서 다음 문장이 나오면 기본 환경 준비가 끝난 것입니다.

```text
Repository verification passed.
```

### PC마다 별도로 준비해야 하는 항목

- 구형 `.doc` 변환: Microsoft Word desktop과 `Word.Application` COM
- 문서 읽기·분석: Codex에 공식 Kordoc MCP 등록
- HWPX 생성·편집: 사용자 계정의 공식 `hwpx` skill clone
- Herdr workflow: Herdr, OMO 설치와 모델 인증
- 실제 문서: 각 PC의 `inbox` 또는 `output`에 별도로 배치

`C:\Users\Hank\.agents\skills\hwpx`처럼 사용자 이름이 들어간 경로는
예시이므로 다른 PC의 사용자 계정에 맞게 바꿉니다. MCP 등록 상태는 다음으로
확인합니다.

```powershell
codex mcp list
```

### 운영체제 호환성

저장소의 DOC 변환과 전체 검증 게이트는 **Windows 기준**입니다. Word COM
자동화를 사용하는 `.doc` → `.docx` 및 `.doc` → Markdown 변환은 Microsoft
Word desktop이 없는 macOS/Linux에서 동일하게 실행할 수 없습니다.

HWPX 분석이나 `uv` 기반 작업은 다른 운영체제에서 일부 사용할 수 있지만,
이 저장소가 보장하는 검증 기준선은 Windows입니다. 다른 운영체제에서는
Kordoc MCP와 `hwpx` skill의 별도 지원 조건을 먼저 확인해야 합니다.

상세 옵션, overwrite 정책, 암호 문서, 문제 해결과 수동 확인 절차는
[`docs/USAGE_GUIDE.md`](docs/USAGE_GUIDE.md)를 참고합니다.

## 일상 작업 흐름

### 1. 구형 DOC를 DOCX 또는 Markdown으로 변환

Codex에 `.doc` 파일 경로를 지정해 Markdown 변환을 요청하면 원본 폴더에 `.docx` 중간본과 `.md` 결과가 함께 생성됩니다.

```text
D:\문서\보고서.doc
D:\문서\보고서.docx
D:\문서\보고서.md
```

통합 실행기를 직접 사용할 수도 있습니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tools\doc-to-docx\convert-doc-to-md.ps1" `
  -Path "D:\문서\보고서.doc"
```

루트의 `convert-doc-to-docx.bat`과 `convert-doc-to-md.bat`은 인자를 무시하고 프로젝트 `inbox`만 처리합니다. 특정 파일이나 폴더를 지정할 때는 `tools/doc-to-docx/convert-doc-to-docx.ps1` 또는 `tools/doc-to-docx/convert-doc-to-md.ps1`에 `-Path`를 사용합니다. `.docx`와 `.md`는 각 원본 `.doc` 옆에 생성됩니다.

DOCX까지만 필요한 경우에는 같은 방식의 DOCX 변환기를 사용합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path "C:\문서\구형보고서.doc"
```

기존 결과를 명시적으로 갱신하려면 PowerShell 통합 실행기에 `-Overwrite`를 사용합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tools\doc-to-docx\convert-doc-to-md.ps1" `
  -Path "D:\문서\보고서.doc" -Overwrite
```

세부 옵션은 `tools/doc-to-docx/README.md`를 참고합니다.

### 2. HWP를 편집 가능한 HWPX로 변환

Windows에서는 hwpx skill의 한컴오피스 Automation 변환기를 우선 사용합니다.
결과를 원본과 같은 폴더에 두려면 `-OutputDirectory`를 생략합니다.

```powershell
$hwpxSkill = Join-Path $HOME ".agents\skills\hwpx"
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File "$hwpxSkill\scripts\convert_hwp_hancom.ps1" `
  -InputPath "D:\문서\보고서.hwp"
```

한컴 Automation을 사용할 수 없을 때는 같은 skill의 rhwp 기반 Python
fallback을 이 저장소 uv 환경으로 실행합니다.

```powershell
uv run python "$hwpxSkill\scripts\convert_hwp.py" `
  "D:\문서\보고서.hwp" `
  -o "D:\문서\보고서.hwpx"
```

### 3. 일반 문서를 읽고 구조를 비교

문서의 현재 경로를 Codex에 알려주면 공식 Kordoc MCP로 HWP/HWPX/DOCX/PDF/
XLSX를 읽고 구조를 추출·비교합니다. `inbox`는 별도 경로가 없는 임시
작업용입니다.

예: `inbox의 계약서.hwpx를 Kordoc으로 읽고 조항별 위험을 비교해줘.`

MCP가 아닌 로컬 CLI로 Markdown 파일이 필요한 경우에도 저장소 lock과 같은
진입점을 사용합니다.

```powershell
node .\tools\kordoc\node_modules\kordoc\dist\cli.js `
  --silent `
  -o "D:\문서\보고서.md" `
  "D:\문서\보고서.pdf"
```

### 4. HWPX를 생성·편집하고 검증

Codex에 `hwpx` skill을 사용하도록 명시하고 결과를 `output`에 저장하도록
요청합니다.

예: `hwpx skill로 이 초안을 편집 가능한 공문 HWPX로 만들어 output에 저장하고 검증해줘.`

skill의 Python 명령은 저장소 루트에서 항상 다음 형태로 실행합니다.

```powershell
uv run python "$hwpxSkill\scripts\validate.py" `
  ".\output\결과.hwpx" `
  --layout
```

## 저장소 관리

이 저장소에는 개인 통합 실행기, 프로젝트 규칙, 운영 문서와 잠금 파일만
보관합니다. Kordoc과 hwpx-skill의 소스는 복사하거나 submodule로 포함하지 않고,
공식 배포본과 별도 clone으로 사용합니다. upstream을 직접 수정해야 할 때만 해당
저장소를 개인 계정으로 fork합니다.

검증된 버전 기준선, 업데이트 브랜치 운영과 호환성 확인 절차는
[`docs/UPSTREAMS.md`](docs/UPSTREAMS.md)를 따릅니다.

## 업데이트 원칙

Kordoc은 호환성 검증을 마친 정확한 버전을 사용합니다. 버전 확인:

```powershell
node .\tools\kordoc\node_modules\kordoc\dist\cli.js --version
```

설치된 `hwpx` skill은 먼저 로컬 변경 여부를 확인한 뒤 fast-forward로만 갱신합니다.

```powershell
git -C C:\Users\Hank\.agents\skills\hwpx status -sb
git -C C:\Users\Hank\.agents\skills\hwpx pull --ff-only
```

### 새 버전 확인과 승인된 업데이트

새 버전은 자동으로 설치하지 않는다. 먼저 읽기 전용 확인을 실행한다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\update-upstreams.ps1
```

이 명령은 다음을 확인하고 보고한다.

- npm의 현재 Kordoc 고정 버전, 같은 major의 최신 버전, 전체 최신 버전
- 별도 `hwpx-skill` clone의 현재 tag·commit·branch·변경 여부
- `hwpx-skill` 공식 `main`의 최신 commit

새 버전이 있어도 기본 모드에서는 파일과 외부 clone을 변경하지 않는다.
업데이트를 검토한 뒤에만 다음 명령으로 적용을 시작한다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\update-upstreams.ps1 -Apply
```

`-Apply`는 구성 요소별로 다시 확인을 요청한다. `-Yes`는 이미 검토한
버전을 자동화 환경에서 적용할 때만 `-Apply`와 함께 사용한다.

```powershell
# Kordoc 4.x 후보만 확인·적용
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\update-upstreams.ps1 `
  -Component Kordoc -KordocVersion 4.13.1 -Apply

# hwpx-skill만 fast-forward로 적용
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\update-upstreams.ps1 `
  -Component HwpxSkill -Apply
```

이 프로세스의 안전장치:

- Kordoc major 업데이트는 자동 선택하지 않는다.
- `hwpx-skill`에 로컬 변경이 있거나 `main`이 아니면 중단한다.
- `hwpx-skill`은 `git pull --ff-only`만 사용한다.
- Kordoc 적용 후 `tools/verify.ps1`와 관련 문서·MCP pin을 검토해야 한다.
- 확인만 자동화하려면 `-FailOnUpdate`를 사용한다. 업데이트가 있으면
  종료 코드 `10`을 반환한다.

Python 패키지를 새 버전으로 올릴 때만 잠금 파일을 갱신합니다.

```powershell
cd D:\Code\Projects\HWPX_wiz
uv lock --upgrade-package python-hwpx
uv sync --locked
```

업데이트 직후에는 대표 HWPX 하나를 열기·저장·검증해 호환성을 확인하는 것이 좋습니다.

## 저장소 검증

개발·정비 후에는 저장소 루트에서 다음 한 명령을 기본 검증 게이트로 실행합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\verify.ps1
```

검증기는 Pester 6.1.0 이상, PSScriptAnalyzer 1.25.0 이상, PATH의 `uv`와
`uv sync --locked`로 만든 Python 환경을 확인합니다. 필요한 도구는 자동으로
설치하지 않습니다. 사용자가 다음 명령으로 사용자 범위 모듈과 잠금 환경을
준비해야 합니다.

```powershell
Install-Module Pester -MinimumVersion 6.1.0 -Scope CurrentUser
Install-Module PSScriptAnalyzer -MinimumVersion 1.25.0 -Scope CurrentUser
uv sync --locked
```

게이트 실패 원인을 좁힐 때만 저장소 루트에서 개별 검사를 실행합니다.

```powershell
Import-Module Pester -MinimumVersion 6.1.0 -Force
Invoke-Pester -Path .\tests -Output Detailed
Import-Module PSScriptAnalyzer -MinimumVersion 1.25.0 -Force
Invoke-ScriptAnalyzer -Path .\tools\doc-to-docx -Recurse
uv lock --check
```

GitHub Actions는 Word 없이 실행 가능한 `-Tier Static` 계층만 검증합니다.
Word COM을 포함한 기본 `-Tier Full` 게이트는 로컬 Windows에서 실행해야 하며,
한컴오피스의 설치 또는 사용 가능 상태를 대신 해결하지 않습니다.
hwpx skill template의 구조·layout·한컴 smoke까지 확인하려면 로컬에서
`.\tools\verify.ps1 -IncludeHwpx`를 사용합니다.

## 안전 메모

- DOC 변환기는 Word 자동화에서 매크로와 자동 링크 업데이트를 끕니다.
- 원본 문서는 삭제하지 않습니다.
- 경로를 지정한 변환 결과는 원본 문서와 같은 폴더에 저장합니다.
- 기존 `.docx`와 `.md`는 `-Overwrite`를 명시하지 않으면 덮어쓰지 않습니다.
- `-Overwrite`도 임시 파일에 먼저 변환한 뒤 성공한 결과만 원자적으로 교체합니다.
- `inbox`와 `output`의 실제 문서는 Git 추적 대상에서 제외됩니다.
- Kordoc은 공식 npm 패키지와 MCP만 사용하며 별도 `Kordoc_helper` 체크아웃을 두지 않습니다.
