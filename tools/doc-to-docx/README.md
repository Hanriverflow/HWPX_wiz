# DOC → DOCX 변환기

구형 Microsoft Word `.doc` 파일을 현재 도구들이 다루기 쉬운 `.docx`로 바꾸는 Windows 전용 보조 도구입니다. Microsoft Word가 설치되어 있어야 합니다. Markdown 통합 변환에는 Node.js/npm과 최초 실행 시 npm 레지스트리 접근도 필요합니다.

Markdown까지 한 번에 만들려면 PowerShell 통합 실행기를 사용합니다. `.docx`와 `.md`는 모두 원본 `.doc`와 같은 폴더에 생성됩니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tools\doc-to-docx\convert-doc-to-md.ps1" `
  -Path "D:\문서\구형보고서.doc"
```

기존 결과를 덮어써야 할 때만 PowerShell에서 `-Overwrite`를 명시합니다.
덮어쓰기는 임시 파일 변환이 성공한 뒤 기존 결과를 원자적으로 교체합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tools\doc-to-docx\convert-doc-to-md.ps1" `
  -Path "D:\문서\구형보고서.doc" -Overwrite
```

## 가장 간단한 사용법

1. `.doc` 파일을 프로젝트 루트의 `inbox` 폴더에 넣습니다.
2. `convert-doc-to-docx.bat`을 더블클릭합니다.
3. 원본 옆에 생성된 `.docx`를 공식 Kordoc CLI/MCP로 읽거나 변환합니다.

원본 `.doc`는 수정하거나 삭제하지 않습니다. 하위 폴더도 함께 처리하며, 이미 `.docx`가 있으면 기본적으로 건너뜁니다.

## 특정 파일 또는 폴더 변환

배치 파일은 인자를 받지 않고 `inbox`만 처리합니다. 특정 경로는 PowerShell 스크립트에 `-Path`로 전달합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path "C:\문서\구형보고서.doc"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path "C:\문서\구형자료" -Recurse
```

## 세부 옵션

PowerShell에서 직접 실행하면 덮어쓰기, 재귀 처리, 로그 파일과 출력 형식을 선택할 수 있습니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path ".\inbox" -Recurse -Overwrite -LogPath ".\output\doc-conversion.log"
```

- `-Recurse`: 하위 폴더 포함
- `-Overwrite`: 기존 `.docx` 덮어쓰기
- `-LogPath`: 처리 로그 저장. 두 변환기 모두 와일드카드와 `.doc`/`.docx`
  충돌을 거부하며, Markdown 통합 변환기는 `.md` 충돌도 거부
- `-OutputFormat Json`: 에이전트용 한 줄 JSON 레코드 배열 출력
- `-TimeoutSeconds`: 문서 한 건의 Word 변환 제한 시간. 기본 300초
- `-DocumentKey`: `String` 또는 `SecureString` 암호. 미지정 시 `HWPX_WIZ_DOC_PASSWORD` 사용

변환 중 Word 알림, 문서 매크로, 자동 링크 업데이트는 비활성화됩니다. 암호가 걸렸거나 손상된 문서는 실패 목록에 기록될 수 있습니다.

## 도움말과 검증

이 구성 요소 문서 디렉터리 `tools\doc-to-docx`에서 다음처럼 저장소 루트로
이동한 뒤 기본 검증 게이트를 실행합니다.

```powershell
cd ..\..
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\verify.ps1
```

검증기는 도구를 설치하지 않습니다. Pester 6.1.0 이상과 PSScriptAnalyzer
1.25.0 이상은 사용자가 `Install-Module ... -Scope CurrentUser`로 설치하고,
PATH에 `uv`를 둔 뒤 `uv sync --locked`로 Python 환경을 준비해야 합니다.
Word 설치 여부는 이 게이트가 해결하지 않습니다.

변환기 관련 실패 원인을 좁힐 때만 저장소 루트에서 다음 개별 검사를 실행합니다.

```powershell
Import-Module Pester -MinimumVersion 6.1.0 -Force
Invoke-Pester -Path .\tests -Output Detailed
Import-Module PSScriptAnalyzer -MinimumVersion 1.25.0 -Force
Invoke-ScriptAnalyzer -Path .\tools\doc-to-docx -Recurse
uv lock --check
```
