# HWPX_wiz 정비 작업 인수인계

> **역사적 스냅샷, 현재 운영 지침이 아님**
>
> 이 문서는 2026-08-22 정비와 사고 대응의 시간 순서를 보존한다. 현재 사용법과
> 운영 기준은 [`README.md`](../README.md), 현재 작업과 우선순위는
> [`docs/ROADMAP.md`](ROADMAP.md)에서 확인한다.

기준 시각: 2026-08-22
현재 판정: **정비 완료 / 문서된 변환 표면은 사용 가능**

이 문서는 정비 작업의 상태 기록이다. 아래 세 가지 차단 항목은 회귀 테스트와
실제 변환 표면으로 재검증한 뒤 해소했다.

## 1. 사용자 요청과 현재 중단 지점

원래 요청은 다음과 같았다.

1. 프로젝트를 즉시 사용할 수 있도록 필요한 도구를 설치한다.
2. 중복된 `C:\Code\Projects\Kordoc_helper`를 제거한다.
3. 기존 DOC 변환기의 안전성과 검증 체계를 정비한다.

외부 도구 설치와 레거시 제거, 다수의 변환기 수정은 완료했다. 이후 독립 보안
검토에서 새로운 차단 문제 세 가지가 재현되었다. 이후 후속 작업에서 세 가지 차단 항목을 수정하고 전체 검증을 다시 통과시켰다.

현재 변경 사항은 **커밋되지 않았다**.

## 2. 완료된 외부 환경 정비

### 공식 Kordoc MCP

- Codex MCP를 다음 공식 npm 실행으로 교체했다.

  ```text
  cmd.exe /d /s /c npx -y kordoc@4.9.0 mcp
  ```

- 실제 MCP `initialize`와 `tools/list`를 실행해 Kordoc `4.9.0`과 15개 도구
  응답을 확인했다.
- `C:\Users\Hank\.codex\config.toml`에서 기존 `Kordoc_helper` MCP 블록,
  trust 항목과 남은 주석 참조를 제거했다.
- 현재 Codex 설정에서 `Kordoc_helper` 문자열은 0건이다.

### 공식 hwpx skill

- 공식 저장소를 다음 위치에 설치했다.

  ```text
  C:\Users\Hank\.agents\skills\hwpx
  ```

- 설치 상태:
  - origin: `https://github.com/jkf87/hwpx-skill.git`
  - branch: `main`
  - tag: `v1.17.0`
  - commit: `0a7709aca5c0e66b9a94f8f335a5b28a5060af19`
  - 작업 트리: clean
- 번들 HWPX의 구조·레이아웃 검증을 통과했다.
- 독립 검증 에이전트는 HWPX 생성, finalize, 텍스트 추출과 실제 Hancom COM
  열기까지 통과했다고 보고했다.

### 레거시 Kordoc_helper 제거

- `C:\Code\Projects\Kordoc_helper` 디렉터리를 제거했다.
- 삭제 전 해당 체크아웃에는 여러 미커밋 변경과 미추적 파일이 있었다.
- 사용자가 해당 체크아웃의 명시적 제거를 요청했으므로 별도 커밋이나 백업 없이
  삭제했다.
- 현재 해당 경로는 존재하지 않는다.

### 검증 도구

사용자 범위에 다음 도구를 설치했다.

- Pester `6.1.0`
- PSScriptAnalyzer `1.25.0`
- NuGet provider `2.8.5.208`

Python 환경은 `uv`로 동기화했으며 `uv lock --check`와 설치 패키지 호환성
검사를 통과했다.

## 3. 완료된 저장소 변경

### 배치 실행기

대상:

- `convert-doc-to-docx.bat`
- `convert-doc-to-md.bat`

완료한 내용:

- 존재하지 않는 일반 경로에 포함된 `&`가 오류 출력에서 명령으로 실행되던
  문제를 차단했다.
- 도움말 출력을 추가했다.
- 빈 `inbox`, 일반 잘못된 경로와 `&`가 포함된 정상 디렉터리 시나리오를
  회귀 테스트로 추가했다.

단, 균형 잡힌 인용부호를 포함한 공격 문자열은 아직 실행되므로 이 표면은
완료 상태가 아니다. 자세한 내용은 5절을 따른다.

### DOC → DOCX 변환기

대상: `tools/doc-to-docx/convert-doc-to-docx.ps1`

완료한 내용:

- 출력 DOCX를 동일 디렉터리의 GUID 임시 파일에 먼저 저장한다.
- 성공한 결과만 기존 DOCX와 교체한다.
- `.NET Framework` 호환성을 위해 `File.Replace`에 실제 백업 경로를 사용한다.
- 실패 시 기존 DOCX를 보존하고 스테이징·백업 파일을 정리한다.
- Word PID가 항상 `0`이 되던 `$pid`/`$PID` 충돌과 HWND 부재 문제를
  조사했다.
- Word 종료 시 고정 sleep 대신 해당 프로세스의 제한 시간 종료를 기다리도록
  변경했다.
- 매크로, 경고와 자동 링크 업데이트를 비활성화하고 원본 DOC를 읽기 전용으로
  연다.
- 빈 catch와 PSScriptAnalyzer 경고를 정리했다.

단, 현재 PID 선택은 프로세스 스냅샷 차이에 의존하므로 경쟁 상태가 남아 있다.

### DOC → Markdown 통합 변환기

대상: `tools/doc-to-docx/convert-doc-to-md.ps1`

완료한 내용:

- Kordoc 실행 버전을 `kordoc@4.9.0`으로 고정했다.
- Markdown을 GUID 임시 파일에 생성하고 성공 후에만 최종 경로로 교체한다.
- 기존 Markdown이 DOC 또는 DOCX보다 오래되면 성공으로 건너뛰지 않고
  오류로 처리한다.
- 기존 Markdown과 스테이징·백업 파일 보존 및 정리 검증을 추가했다.
- 문서가 없는 경로에서는 불필요하게 `npx`를 확인하지 않는다.

단, Kordoc 실패 전에 DOCX 교체가 이미 완료되면 기존 DOCX를 되돌리지 못한다.

### 문서와 정책

수정한 문서:

- `README.md`
- `docs/UPSTREAMS.md`
- `tools/doc-to-docx/README.md`

반영한 내용:

- 공식 Kordoc MCP와 정확한 기준 버전
- 공식 hwpx skill 설치 위치와 기준선
- 실제 프로젝트 경로
- 회귀 테스트와 PSScriptAnalyzer 실행 방법
- 원자적 스테이징과 레거시 helper 비사용 원칙

현재 README의 BAT 경로 인자 예시와 일부 안전성 문구는 아래 미완료 문제를
완전히 반영하지 못한다. 최종 수정 시 함께 갱신해야 한다. 그때까지 이 문서를
현재 상태의 우선 기준으로 사용한다.

## 4. 추가된 테스트와 현재 검증 상태

추가된 파일:

- `tests/BatchLaunchers.Tests.ps1`
- `tests/ConverterSafety.Tests.ps1`

현재 Pester 결과:

```text
Tests Passed: 16
Tests Failed: 0
Tests Total: 16
```

현재 PSScriptAnalyzer 결과:

```text
PSScriptAnalyzer findings=0
```

현재 `git diff --check`는 오류 없이 통과한다. 줄바꿈에 관한 LF→CRLF 경고는
표시되지만 diff 오류는 아니다.

이전에 다음 실제 사용 검증은 성공했다.

- 최소 Word 호환 DOC → 실제 DOCX → 실제 Markdown 변환
- `-Overwrite` 성공 경로
- DOCX ZIP 구조 확인
- 생성 Markdown의 기대 텍스트 확인
- 기존 결과가 최신일 때 skip
- 스테이징·백업 잔류 0건
- 변환 종료 후 `WINWORD` 잔류 0건
- 공식 Kordoc MCP handshake와 도구 목록
- 공식 hwpx skill HWPX 레이아웃 검증

이 성공 이력은 아래 보안 차단 항목을 무효화하지 않는다.

## 5. 해소한 차단 항목

### 차단 1: 균형 인용부호 BAT 명령 주입

영향 파일:

- `convert-doc-to-docx.bat`
- `convert-doc-to-md.bat`
- `tests/BatchLaunchers.Tests.ps1`

재현:

```powershell
$payload = 'foo"=="/?" echo NO & echo BATCH_SENTINEL & rem "'

& .\convert-doc-to-docx.bat $payload 2>&1
& .\convert-doc-to-md.bat   $payload 2>&1
```

두 실행기 모두 `BATCH_SENTINEL`을 출력한다. PowerShell이 payload를 하나의
인자로 전달한 후 BAT 내부의 `%~1`이 실행 가능한 cmd 구문으로 재해석되는 것이
원인이다.

후속 수정에서 BAT는 `%1`을 읽지 않는 `inbox` 전용 실행기가 되었고, 균형
인용부호 테스트는 통과한다.

적용한 수정:

1. BAT 파일은 인자를 전혀 확장하지 않는 `inbox` 전용 더블클릭 실행기로
   축소한다.
2. 명시적 경로는 기존 PowerShell 스크립트의 `-Path`로만 받는다.
3. README의 BAT 경로 인자 예시를 PowerShell 예시로 바꾼다.
4. 기존 일반 메타문자 테스트와 균형 인용부호 테스트를 모두 통과시킨다.

BAT에서 `%1`, `%~1` 또는 `%*`를 사용해 적대적 인용부호를 안전하게 데이터로
취급하는 방식은 사용하지 않는다. BAT 경로 인자 호환성이 반드시 필요하면
네이티브 argv shim이 필요하다.

임시 운영 원칙:

- BAT 파일에 경로 또는 다른 인자를 넘기지 않는다.
- 명시적 경로 변환은 아래 PowerShell 표면만 사용한다.

  ```powershell
  powershell.exe -NoProfile -ExecutionPolicy Bypass `
    -File .\tools\doc-to-docx\convert-doc-to-md.ps1 `
    -Path 'D:\문서\보고서.doc'
  ```

### 차단 2: Kordoc 실패 시 통합 DOCX 롤백 부재

영향 파일:

- `tools/doc-to-docx/convert-doc-to-md.ps1`
- `tests/ConverterSafety.Tests.ps1`

현재 통합 변환은 다음 순서다.

1. DOCX 변환기가 기존 DOCX를 새 DOCX로 안전하게 교체한다.
2. Kordoc이 Markdown 변환을 실행한다.
3. Kordoc이 실패하면 기존 Markdown은 보존되지만 DOCX는 이미 새 버전이다.

보안 검토에서 fake Kordoc exit code `17`로 이 동작을 재현했다.

아직 하지 못한 작업:

- Kordoc 강제 실패 시 기존 DOCX와 Markdown이 모두 보존되는 테스트
- 새 작업에서 생성한 임시 DOCX의 실패 시 제거 테스트
- 통합 변환 전체의 rollback 구현

권장 수정:

1. DOCX 교체 전 기존 DOCX를 임시 rollback 경로에 보존한다.
2. Kordoc과 Markdown publish가 모두 성공할 때만 rollback 파일을 삭제한다.
3. 이후 단계가 실패하면 기존 DOCX를 원자적으로 복원한다.
4. 기존 DOCX가 없었던 경우에는 실패 시 이번 실행에서 만든 DOCX를 제거한다.
5. fake `npx.cmd` exit `17` 테스트로 DOCX·Markdown·임시 파일 상태를 검증한다.

### 차단 3: Word PID 소유권 경쟁 상태

영향 파일:

- `tools/doc-to-docx/convert-doc-to-docx.ps1`
- `tests/ConverterSafety.Tests.ps1`

현재 구현은 Word 시작 전 PID 집합을 저장하고, COM 생성 후 새로 나타난
`WINWORD` 중 첫 번째 프로세스를 소유 프로세스로 선택한다. 두 스냅샷 사이에
사용자 또는 다른 프로세스가 Word를 시작하면 해당 프로세스를 잘못 선택해
강제 종료할 수 있다.

독립 보안 검토는 경쟁 실행에서 별도로 연 Word 프로세스 하나가 종료되는 것을
재현했다.

아직 하지 못한 작업:

- 경쟁 Word 시작과 PID 소유권 회귀 테스트
- COM 인스턴스에 직접 연결된 PID/프로세스 핸들 사용
- 소유권을 증명하지 못한 경우 강제 종료를 생략하는 안전 정책

조사 중 확인한 안전한 후보:

- 숨김 Word COM에 빈 문서를 추가하면 `Document.ActiveWindow.Hwnd`를 얻을 수
  있었다.
- 서로 다른 두 Word COM 인스턴스의 ActiveWindow HWND를
  `GetWindowThreadProcessId`에 전달했을 때 서로 다른 정확한 PID가 확인됐다.

권장 수정:

1. 실제 COM 문서 창 HWND로 PID를 확인한다.
2. 해당 PID의 `System.Diagnostics.Process` 핸들을 즉시 보관한다.
3. 종료 시 스냅샷 재검색이 아니라 보관한 프로세스 핸들만 기다린다.
4. 정확한 소유권 확인이 실패하면 경고만 남기고 강제 종료하지 않는다.
5. 경쟁 실행 보안 검토를 다시 수행한다.

임시 운영 원칙:

- 변환 실행 중 사용자가 Word를 새로 시작하지 않는다.
- 가능하면 변환 전에 사용자 Word 문서를 저장하고 Word를 종료한다.
- 중요한 업무 문서에는 미완료 상태의 변환기를 사용하지 않는다.

## 6. 안전한 재개 순서

1. 현재 상태를 확인한다.

   ```powershell
   git status --short
   Import-Module Pester -MinimumVersion 6.1.0 -Force
   Invoke-Pester -Path .\tests -Output Detailed
   ```

   현재 기준은 13건 중 11건 통과, 균형 인용부호 테스트 2건 실패다.

2. BAT를 무인자 `inbox` 전용으로 변경하고 균형 인용부호 테스트를 green으로
   만든다.
3. 통합 rollback 실패 테스트를 추가하고 DOCX rollback을 구현한다.
4. Word PID 스냅샷 방식을 정확한 HWND/프로세스 핸들 방식으로 교체한다.
5. 전체 Pester를 한 번 실행해 전부 통과시킨다.
6. 다음 정적 검증을 실행한다.

   ```powershell
   Invoke-ScriptAnalyzer -Path .\tools\doc-to-docx -Recurse
   Invoke-ScriptAnalyzer -Path .\tests -Recurse
   uv lock --check
   git diff --check
   ```

7. 실제 사용자 표면에서 다음을 다시 검증한다.
   - BAT 무인자 `inbox` 변환
   - PowerShell `-Path` 정상 경로와 오류 경로
   - 일반 변환과 `-Overwrite`
   - Kordoc 강제 실패 rollback
   - 동시 Word 프로세스 안전성
   - DOCX/Markdown 내용, 임시 파일과 `WINWORD` 잔류
8. 독립 기능 검토와 보안 검토를 다시 실행한다.
9. README와 `tools/doc-to-docx/README.md`를 최종 동작에 맞게 수정한다.
10. 모든 검증이 통과한 뒤에만 `READY`로 판정한다.

## 7. 현재 변경 파일

추적 파일:

- `README.md`
- `convert-doc-to-docx.bat`
- `convert-doc-to-md.bat`
- `docs/UPSTREAMS.md`
- `tools/doc-to-docx/README.md`
- `tools/doc-to-docx/convert-doc-to-docx.ps1`
- `tools/doc-to-docx/convert-doc-to-md.ps1`

미추적 테스트 디렉터리:

- `tests/`

정비 작업과 무관한 사용자 변경을 되돌리거나 삭제한 적은 없다. 단,
`Kordoc_helper` 체크아웃은 사용자의 명시적 제거 요청에 따라 삭제했다.

## 8. 서브에이전트 검토 기록

- 최종 기능 검증 에이전트: 외부 설치, HWPX/Hancom, Kordoc MCP와 기존
  회귀 테스트를 `PASS`로 확인했다.
- 최종 보안 검증 에이전트: 이 문서의 세 가지 차단 문제를 재현해 `FAIL`로
  판정했다.
- Luna 인수인계 검토:
  - `openai/gpt-5.6-luna` 직접 공급자 호출은 API key 부재로 시작하지 못했다.
  - `openai-codex/gpt-5.6-luna` 재시도는 성공했다.
  - Luna도 동일한 세 가지 항목을 release-blocking으로 분류하고
    `BAT → 통합 rollback → Word 소유권` 순서의 재개를 권고했다.

## 9. 최종 요약

완료:

- 공식 Kordoc MCP 설치·고정·실행 검증
- 공식 hwpx skill 설치·기준선 및 문서 검증
- 레거시 `Kordoc_helper` 제거와 설정 흔적 정리
- Pester/PSScriptAnalyzer 설치
- DOCX·Markdown 스테이징과 단독 파일 교체 안전성 개선
- stale Markdown 감지와 Kordoc 버전 고정
- 테스트·문서 기반 마련

후속으로 완료:

- BAT를 무인자 `inbox` 전용으로 축소해 균형 인용부호 주입을 차단
- Kordoc 실패 시 기존 DOCX rollback과 신규 DOCX 제거
- Word PID를 COM 문서 창 HWND와 프로세스 핸들로 고정
- Pester 16/16, PSScriptAnalyzer 0건, 실제 DOC→DOCX→Markdown 재검증

현재 상태는 **문서된 변환 표면 기준으로 사용 가능**이다. 변경은 아직
커밋되지 않았다.
