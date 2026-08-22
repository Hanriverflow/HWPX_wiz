# HWPX_wiz

한글 문서 작업은 활발히 업데이트되는 공식 upstream 두 개를 중심으로 운영합니다.

- [Kordoc](https://github.com/chrisryugj/kordoc): HWP/HWPX/DOCX/PDF/XLSX 등의 읽기, 구조 추출, 비교, 양식 분석과 변환
- [hwpx-skill](https://github.com/jkf87/hwpx-skill): 편집 가능한 HWPX 생성·수정과 품질 검증
- 이 프로젝트의 DOC 변환기: Kordoc이 직접 지원하지 않는 구형 Word `.doc`를 `.docx`로 선행 변환

기존 `Kordoc_helper` 애플리케이션과 자체 MCP는 사용하지 않습니다.

## 역할 분담

| 작업 | 사용할 도구 |
|---|---|
| HWP/HWPX/DOCX/PDF/XLSX 내용 파악·비교 | Codex의 공식 Kordoc MCP |
| 새 HWPX 작성, 기존 HWPX 편집·검증 | Codex의 `hwpx` skill + 이 프로젝트의 `uv` 환경 |
| 구형 `.doc` → Markdown | `convert-doc-to-md.bat`으로 Word 변환 후 Kordoc 실행 |
| 경로가 지정된 문서 | 원본과 같은 폴더에 결과 저장 |
| 경로 없는 DOC 임시 작업 | `inbox`에서 찾고 결과도 해당 DOC 옆에 저장 |
| 별도 생성 문서·보고서·로그 | 필요할 때 `output` 사용 |

## 처음 한 번 또는 환경 복원

Python은 전역 설치 대신 `uv`로만 관리합니다.

```powershell
cd C:\CODE\Project\HWPX_wiz
uv sync --locked
```

현재 환경은 Python 3.12, `python-hwpx`, `lxml`, `pywin32`를 잠금 파일 기준으로 설치합니다.

공식 Kordoc MCP 등록 상태는 다음처럼 확인합니다.

```powershell
codex mcp list
```

`kordoc` 항목의 명령은 `cmd /c npx -y kordoc@4 mcp`입니다. 새 Codex 작업에서 Kordoc 도구를 사용하면 됩니다.

## 일상 작업 흐름

### 1. 일반 문서 읽기·분석

문서의 현재 경로를 Codex에 알려주면 원본 위치에서 바로 작업합니다. `inbox`는 별도 경로가 없는 임시 작업용입니다.

예: `inbox의 계약서.hwpx를 Kordoc으로 읽고 조항별 위험을 비교해줘.`

### 2. HWPX 생성·편집

Codex에 `hwpx` skill을 사용하도록 명시하고 결과를 `output`에 저장하도록 요청합니다.

예: `hwpx skill로 이 초안을 편집 가능한 공문 HWPX로 만들어 output에 저장하고 검증해줘.`

Python 명령이 필요하면 항상 다음 형태를 사용합니다.

```powershell
uv run python <script.py>
```

### 3. 구형 DOC를 Markdown으로 변환

Codex에 `.doc` 파일 경로를 지정해 Markdown 변환을 요청하면 원본 폴더에 `.docx` 중간본과 `.md` 결과가 함께 생성됩니다.

```text
D:\문서\보고서.doc
D:\문서\보고서.docx
D:\문서\보고서.md
```

통합 실행기를 직접 사용할 수도 있습니다.

```bat
convert-doc-to-md.bat "D:\문서\보고서.doc"
```

경로를 생략하고 더블클릭하면 `inbox` 아래의 `.doc`를 처리하되, `.docx`와 `.md`는 각 원본 `.doc` 옆에 생성합니다.

DOCX까지만 필요한 경우에는 기존 실행기를 사용합니다.

```bat
convert-doc-to-docx.bat "C:\문서\구형보고서.doc"
```

기존 결과를 명시적으로 갱신하려면 PowerShell 통합 실행기에 `-Overwrite`를 사용합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tools\doc-to-docx\convert-doc-to-md.ps1" `
  -Path "D:\문서\보고서.doc" -Overwrite
```

세부 옵션은 `tools/doc-to-docx/README.md`를 참고합니다.

## 저장소 관리

이 저장소에는 개인 통합 실행기, 프로젝트 규칙, 운영 문서와 잠금 파일만
보관합니다. Kordoc과 hwpx-skill의 소스는 복사하거나 submodule로 포함하지 않고,
공식 배포본과 별도 clone으로 사용합니다. upstream을 직접 수정해야 할 때만 해당
저장소를 개인 계정으로 fork합니다.

검증된 버전 기준선, 업데이트 브랜치 운영과 호환성 확인 절차는
[`docs/UPSTREAMS.md`](docs/UPSTREAMS.md)를 따릅니다.

## 업데이트 원칙

Kordoc은 `kordoc@4`로 메이저 버전을 고정해 실행할 때 최신 4.x를 사용합니다. 버전 확인:

```powershell
cmd /c npx -y kordoc@4 --version
```

설치된 `hwpx` skill은 먼저 로컬 변경 여부를 확인한 뒤 fast-forward로만 갱신합니다.

```powershell
git -C C:\Users\Hank\.agents\skills\hwpx status -sb
git -C C:\Users\Hank\.agents\skills\hwpx pull --ff-only
```

Python 패키지를 새 버전으로 올릴 때만 잠금 파일을 갱신합니다.

```powershell
cd C:\CODE\Project\HWPX_wiz
uv lock --upgrade-package python-hwpx
uv sync --locked
```

업데이트 직후에는 대표 HWPX 하나를 열기·저장·검증해 호환성을 확인하는 것이 좋습니다.

## 안전 메모

- DOC 변환기는 Word 자동화에서 매크로와 자동 링크 업데이트를 끕니다.
- 원본 문서는 삭제하지 않습니다.
- 경로를 지정한 변환 결과는 원본 문서와 같은 폴더에 저장합니다.
- 기존 `.docx`와 `.md`는 `-Overwrite`를 명시하지 않으면 덮어쓰지 않습니다.
- `inbox`와 `output`의 실제 문서는 Git 추적 대상에서 제외됩니다.
- `Kordoc_helper` 원본 저장소는 기존 미커밋 변경 보존을 위해 별도 확인 전까지 유지합니다.
