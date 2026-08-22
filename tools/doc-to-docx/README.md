# DOC → DOCX 변환기

구형 Microsoft Word `.doc` 파일을 현재 도구들이 다루기 쉬운 `.docx`로 바꾸는 Windows 전용 보조 도구입니다. Microsoft Word가 설치되어 있어야 합니다.

Markdown까지 한 번에 만들려면 프로젝트 루트의 `convert-doc-to-md.bat`을 사용합니다. `.docx`와 `.md`는 모두 원본 `.doc`와 같은 폴더에 생성됩니다.

```bat
convert-doc-to-md.bat "D:\문서\구형보고서.doc"
```

기존 결과를 덮어써야 할 때만 PowerShell에서 `-Overwrite`를 명시합니다.

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

```bat
convert-doc-to-docx.bat "C:\문서\구형보고서.doc"
convert-doc-to-docx.bat "C:\문서\구형자료"
```

## 세부 옵션

PowerShell에서 직접 실행하면 덮어쓰기, 재귀 처리, 로그 파일을 선택할 수 있습니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tools\doc-to-docx\convert-doc-to-docx.ps1" `
  -Path ".\inbox" -Recurse -Overwrite -LogPath ".\output\doc-conversion.log"
```

- `-Recurse`: 하위 폴더 포함
- `-Overwrite`: 기존 `.docx` 덮어쓰기
- `-LogPath`: 처리 로그 저장

변환 중 Word 알림, 문서 매크로, 자동 링크 업데이트는 비활성화됩니다. 암호가 걸렸거나 손상된 문서는 실패 목록에 기록될 수 있습니다.
