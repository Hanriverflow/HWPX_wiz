# HWPX_wiz 개선 로드맵

기준: 2026-08-23
전제: 공식 Kordoc `4.9.0`, hwpx-skill `v1.17.0`, DOC 변환기 안전 정비는 끝난 상태다. 2026-08-22 인수인계에 적힌 Pester 16/16은 변환기 중심의 역사적 기준선이다. 현재 저장소 전체 기준선은 변환기 테스트와 Herdr/OMO workflow 테스트를 합친 Pester 40/40이다. 이 문서는 **다음에 손댈 가치가 있는 항목만** 적는다.

우선순위는 영향 × 확실성이다. P0는 지금 작업 트리를 닫기 전에 하는 일, P1은 다음 정비 세션, P2는 실제 문서 작업이 늘어난 뒤다.

## 현재 기준선

이미 끝난 것:

- 공식 Kordoc MCP와 공식 `hwpx` skill 설치
- `Kordoc_helper` 제거
- BAT를 무인자 `inbox` 전용으로 축소
- DOCX/Markdown 스테이징, stale Markdown 거부, Kordoc 버전 고정
- Kordoc 실패 시 기존 DOCX rollback
- Word PID를 COM HWND + 프로세스 핸들로 고정
- 역사적 변환기 기준선: Pester 16/16
- 현재 전체 저장소 기준선: Pester 40/40, PSScriptAnalyzer 0

아직 제품이 아닌 점:

- 변경이 커밋되지 않았다.
- CI가 없다.
- 대표 실무 문서 픽스처가 저장소에 없다.
- 문서 계약 일부가 BAT 경로 인자 시절을 가리킨다.

## P0: 작업 트리를 닫기 전에

### 1. 커밋 단위를 나눈다

현재 추적 수정과 미추적 파일은 아래 다섯 경계로 나눈다. 각 경로는 한 경계에만 속한다. `.omo`의 ignored runtime state는 소스 변경이 아니므로 목록에서 제외한다.

1. 변환기 구현, `fix(converter): harden DOC conversion safety`
   - `convert-doc-to-docx.bat`
   - `convert-doc-to-md.bat`
   - `tools/doc-to-docx/convert-doc-to-docx.ps1`
   - `tools/doc-to-docx/convert-doc-to-md.ps1`
2. 변환기 회귀 테스트, `test(converter): cover launcher and rollback regressions`
   - `tests/BatchLaunchers.Tests.ps1`
   - `tests/ConverterSafety.Tests.ps1`
3. Herdr workflow, `build(herdr): add PLAN BUILD VERIFY workflow`
   - `.gitignore`
   - `.omo/omo.jsonc`
   - `.omo/prompts/build.md`
   - `.omo/prompts/plan.md`
   - `.omo/prompts/verify.md`
   - `docs/HERDR_OMO_WORKFLOW.md`
   - `tests/HerdrCombo.Tests.ps1`
   - `tools/herdr/start-combo.ps1`
4. 정비 문서, `docs: refresh maintenance baseline`
   - `docs/MAINTENANCE_HANDOFF.md`
   - `docs/ROADMAP.md`
5. BAT 계약 문서, `docs: align BAT and explicit-path contracts`
   - `README.md`
   - `docs/UPSTREAMS.md`
   - `tools/doc-to-docx/README.md`

이 경계는 제안일 뿐이다. 현재 커밋이나 staging은 만들지 않았다.

커밋 전 기본 검증은 저장소 루트에서 다음 명령으로 실행한다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\verify.ps1
```

### 2. 문서 계약을 BAT 무인자 모델에 맞춘다

남아 있는 불일치:

- `AGENTS.md`가 DOC→Markdown 기본 경로로 `convert-doc-to-md.bat`를 적는다. 명시적 경로는 `.ps1 -Path`다.
- `docs/UPSTREAMS.md` 호환성 게이트 3번이 같은 BAT를 가리킨다.
- `README.md` 역할 표의 “실행기” 열이 아직 BAT를 경로 실행기처럼 읽힌다.

수정 범위는 문장 교체만으로 충분하다. 동작 변경은 필요 없다.

### 3. 정비 스냅샷과 상시 문서를 나눈다

`docs/MAINTENANCE_HANDOFF.md`는 2026-08-22 정비 일지다. 상시 운영 문서처럼 읽히면 다음 작업자가 오래된 차단 항목을 현재 결함으로 오해한다.

- 일지는 그대로 두고 첫머리에 “역사 기록”임을 더 분명히 한다.
- 현재 사용법의 단일 진입점은 `README.md` + 이 로드맵으로 둔다.

## P1: 다음 정비 세션

### 4. 저장소에서 한 명령으로 검증한다

로컬 저장소의 기본 검증 게이트는 저장소 루트에서 실행하는 한 명령으로 정리됐다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\verify.ps1
```

이 스크립트는 Pester 6.1.0 이상, PSScriptAnalyzer 1.25.0 이상, PATH의 `uv`와
`uv sync --locked`로 만든 Python 환경을 요구한다. 설치 책임은 사용자에게 있고,
누락 시 사용자 범위 설치 명령을 출력하고 실패한다. 호스팅 CI, Word 또는
한컴오피스의 사용 가능 상태를 제공하거나 해결하지 않는다.

실패 원인을 좁힐 때만 아래 구성 요소 명령을 저장소 루트에서 따로 실행한다.

```powershell
Import-Module Pester -MinimumVersion 6.1.0 -Force
Invoke-Pester -Path .\tests -Output Detailed
Import-Module PSScriptAnalyzer -MinimumVersion 1.25.0 -Force
Invoke-ScriptAnalyzer -Path .\tools -Recurse
Invoke-ScriptAnalyzer -Path .\tests -Recurse
uv lock --check
```

### 5. Kordoc을 npm lock으로 고정한다

변환기와 MCP는 `kordoc@4.9.0`을 쓰지만 매번 `npx -y`가 레지스트리에서 받는다. 버전 문자열만으로는 integrity가 없다.

최소안:

- 저장소에 `tools/kordoc/` 또는 루트 `package.json` + lockfile을 두고 `kordoc@4.9.0`을 정확히 설치한다.
- 변환기는 `npx -y kordoc@4.9.0` 대신 lock된 로컬 바이너리를 호출한다.
- `docs/UPSTREAMS.md`의 갱신 절차에 lockfile 커밋을 넣는다.

### 6. 통합 변환 실패 픽스처를 더 좁힌다

지금 커버:

- 기존 DOCX/Markdown 보존
- 신규 DOCX 제거
- stale Markdown 거부
- 독립 Word 프로세스 생존

아직 없는 것:

- 암호 DOC / 손상 DOC의 종료 코드와 원본 미변경
- Word가 이미 실행 중일 때 COM이 기존 인스턴스를 재사용하는 경우
- `LogPath` 부모 디렉터리가 없을 때의 실패 메시지
- 공백·한글·`&`가 동시에 있는 실제 파일명 (디렉터리만 있음)

실무 DOC 한두 개를 `inbox`가 아니라 `tests/fixtures/`에 익명화해 넣는 것이 가짜 RTF보다 가치가 크다. 원본 업무 문서는 넣지 않는다.

### 7. 변환기 출력 계약을 정리한다

`convert-doc-to-docx.ps1` 주석은 파이프라인 객체를 약속하지만, 요약은 `Out-Host` + `Format-Table`이다. 에이전트가 결과를 파싱하기 어렵다.

최소안:

- 사람용 요약은 정보 스트림에 유지한다.
- 마지막에 `Converted`/`Skipped`/`Failed` 객체만 성공 스트림으로 내보낸다.
- `exit` 대신 `$exitCode`를 돌려 `-File` 호출과 도트소싱을 구분한다.

## P2: 실제 문서 작업량이 늘어난 뒤

### 8. HWPX 생성 스모크를 저장소에 남긴다

지금은 hwpx skill과 Hancom 검증이 머신 설치에 기대 있다. 프로젝트에는 대표 HWPX가 없다.

- `output/`이 아니라 `tests/fixtures/`에 작은 공개 HWPX 하나
- `uv run python .../validate.py --layout`를 `tools/verify.ps1`의 선택 단계로 연결
- Hancom이 없으면 skip. 실패로 만들지 않는다.

### 9. DOC → HWPX 경로를 문서로만 고정한다

코드로 파이프라인을 새로 짜지 않는다. 이미 역할이 나뉘어 있다.

1. `.doc` → `.docx`/`.md` : 이 저장소 변환기
2. 구조 확인 : Kordoc MCP
3. 편집 가능 HWPX : `hwpx` skill

README에 이 세 단계를 한 시퀀스로 적고, “DOC를 HWPX로 바로 변환하는 스크립트”는 만들지 않는다. 중복 변환기는 이전 helper와 같은 함정이다.

### 10. 로그와 작업 산출물 위치를 통일한다

`LogPath`는 호출자가 만들어야 하고, `output/`은 정책만 있다. 실제 변환 결과는 원본 옆에 생긴다.

- 기본 로그를 `output/logs/`로 둘지, 아예 만들지 않을지 한쪽으로 정한다.
- `inbox`/`output` `.gitkeep` 계약은 유지한다.
- 변환기가 `output/`으로 결과를 옮기게 바꾸지 않는다. `AGENTS.md`의 “원본 옆 저장”이 우선이다.

## 하지 말 것

- `Kordoc_helper`나 로컬 MCP 재도입
- BAT에 경로 인자를 다시 받기. cmd 재해석을 이길 수 없다
- 변환기 전체의 C#/Python 재작성
- Word 없는 환경에서 COM을 흉내 내는 큰 mock 계층
- 사용자 Word 프로세스를 스냅샷 차이로 강제 종료하는 로직 복원
- 머신 절대 경로를 더 늘리는 설치 문서

## 추천 진행 순서

```text
P0 커밋 분할 → P0 문서 계약 정리
    → P1 verify.ps1
    → P1 Kordoc lockfile
    → P1 실패/파일명 픽스처
    → P1 출력 계약
    → P2 는 실제 HWPX 작업이 반복될 때만
```

한 세션에 P0 + `verify.ps1`이면 충분하다. P2를 같은 세션에 넣지 않는다.
