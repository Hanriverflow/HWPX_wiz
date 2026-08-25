# HWPX_wiz 개선 로드맵

기준일: 2026-08-25 · 기준 커밋: `f84eb50` (`main`)

> **실행 상태 (2026-08-25)**
>
> - [x] P0-1~P0-6 완료
> - [x] P1-7~P1-13 완료 — Kordoc `4.9.2`, 로컬 MCP, Static/Full,
>   JSON·LogPath·timeout·보안 암호 경로
> - [x] P2-14~P2-16 완료 — HWPX smoke, uv bridge, 4단계 문서 흐름
> - [ ] P2-17~P2-20 조건부 보류 — 반복 수요·5개 이상 DOC 폴더·온보딩
>   피드백·중앙 로그 요구가 아직 확인되지 않음
> - [x] P3 기록 확인 — BAT 코드페이지와 Actions SHA pin은 후속 기록으로
>   유지하고, Herdr 모델 규칙·현재 `python-hwpx 6.3.0` 상태·Kordoc/hwpx
>   역할 경계는 운영 문서에 반영
>
> 아래 본문은 기준 커밋 당시의 분석과 결정 근거다. 현재 운영 명령과 버전은
> `README.md`, `docs/USAGE_GUIDE.md`, `docs/UPSTREAMS.md`를 따른다.

이 문서는 2026-08-23판 로드맵을 대체한다. 이전 항목의 처리 결과는 마지막 절에
정리했다. 우선순위는 **영향 × 확실성 ÷ 비용**이다. P0는 다음 작업 세션에서
바로 닫을 것, P1은 그다음 정비 세션, P2는 실제 문서 작업량이 늘어난 뒤,
P3는 기록만 남기는 항목이다.

## 0. 이번 분석에서 실제로 확인한 값

문서가 아니라 실행·조회로 확인한 값만 적는다. 모든 진단과 우선순위는 이
표를 근거로 한다.

| 항목 | 확인 결과 |
|---|---|
| `tools/verify.ps1` | 통과. Pester **44/44** (83초), PSScriptAnalyzer 0건, `uv lock --check` 통과 |
| Kordoc 로컬 lock | `4.9.1` (package.json = lockfile = CLI). npm 최신은 **`4.9.2`** (2026-08-24 배포), 미검증 |
| Kordoc MCP (Codex) | `npx.cmd -y kordoc@4.9.1 mcp` — 레지스트리 경유, 로컬 lock과 별개 경로 |
| hwpx-skill clone | `main@96a2633` = `v1.17.0-3-g96a2633`, clean, `origin/main`과 동일. tag가 아니라 tag 뒤 3커밋이다 |
| Python 환경 | uv 0.11.28, Python 3.12, python-hwpx 6.3.0, lxml 6.1.2, pywin32. hwpx-skill 스크립트가 import하는 외부 모듈(`hwpx`, `lxml`, `win32com`)은 모두 충족 |
| Office | Word desktop(Office16) 있음. **한컴오피스 2022 (Hwp.exe 12.0.0.4204, `C:\Tools\HWP2022`) 있음**, `HWPFrame.HwpObject` COM 등록됨 |
| GitHub Pages | 배포 워크플로 정상 (최근 2회 성공) |
| 작업 트리 | `.omc/`, `.gjc/` 미추적 런타임 디렉터리 (gitignore 누락) |

## 1. 진단

### 1.1 프로젝트 성격

- 세 부분으로 이루어진다. (a) DOC→DOCX→Markdown 변환기(PowerShell, Word COM +
  로컬 Kordoc CLI), (b) 검증·업데이트 도구(`verify.ps1`, `update-upstreams.ps1`),
  (c) upstream 연결 규칙과 운영 문서(README, USAGE_GUIDE, HTML 가이드,
  Herdr/OMO workflow).
- PowerShell은 `tools/` 약 1,800줄 + `tests/` 약 800줄, Pester 44건. 저장소 안에
  Python 코드는 없고 uv 환경은 hwpx skill 실행을 위한 의존성 보관용이다.
- 2026-08-22 시작, 커밋 10개. 저장소 크기에 비해 안전 설계(스테이징→원자
  교체, rollback, Word 프로세스 소유권 확인)와 검증 게이트가 잘 갖춰져 있다.
  이 로드맵은 그 기반을 무너뜨리지 않는 범위에서 빈틈을 메우는 순서다.

### 1.2 발견한 문제

심각도 순. 각 항목은 뒤의 P0–P3 번호로 연결된다.

**A. [버그·안전] `convert-doc-to-md.ps1`가 Kordoc을 npm `.cmd` shim으로 호출해 `&` 경로에서 깨진다 → P0-1**

- `tools/doc-to-docx/convert-doc-to-md.ps1:200`은
  `tools/kordoc/node_modules/.bin/kordoc.cmd`를 `& $kordocCli --silent -o $staged $docx`로
  호출한다. Windows PowerShell 5.1은 공백이 없는 인자를 따옴표 없이 넘기고,
  `.cmd`는 cmd.exe가 해석하므로 `&`에서 인자가 잘리고 뒷부분이 별도 명령으로
  실행된다.
- 재현: 인자를 echo하는 가짜 `kordoc.cmd`에 `…\amp&test\x.docx`를 넘기면
  `ARGS=--silent -o …\amp`까지만 전달되고 `test\.x….md …`가 명령으로 실행돼
  "지정된 경로를 찾을 수 없습니다"로 종료 코드 1이 났다.
- 영향: 폴더명이나 파일명에 공백 없이 `&`가 있으면 Markdown 변환이 실패하고,
  경로 일부가 명령으로 실행되는 주입 표면이 된다. 정비 때 BAT에서 막은 것과
  같은 부류다. 기존 테스트는 `&`를 공백과 함께 쓰거나(따옴표 처리됨) 빈
  디렉터리만 검사해서 잡지 못한다. `verify.ps1:228-229`는 이미 `node.exe` +
  `dist/cli.js` 직접 호출을 쓰고 있어 같은 방식으로 고치면 된다.

**B. [테스트 위생] 실행기 테스트가 실제 `inbox`를 변환한다 → P0-2**

- `tests/BatchLaunchers.Tests.ps1:6-47`의 세 It(메타문자 인자, 균형 인용부호
  인자, 무인자) × 두 BAT = 6건이 모두 실제 루트 BAT를 실행한다. BAT는 인자를
  전혀 읽지 않고 `%~dp0inbox`를 처리하므로, 주입 테스트를 포함한 6건 전부가
  사용자가 `inbox`에 둔 `.doc`를 실제로 변환한다(기존 결과는 skip되지만 신규
  DOCX/MD가 생긴다). 검증 게이트가 사용자 데이터를 만지면 안 된다.

**C. [문서 드리프트] 버전·기준선 문자열이 흩어져 이미 어긋났다 → P0-4, P1-8, P2-19**

- `HWPX_wiz_easy_guide.html`은 Kordoc `4.9.0`을 5곳(156, 279, 299, 577, 631행)에서
  말한다. 저장소는 `4.9.1`이고, 이 파일은 GitHub Pages로 공개된다.
- Kordoc pin이 저장소 8개 파일(README, UPSTREAMS, USAGE_GUIDE, HTML 가이드,
  `verify.ps1`, `Verify.Tests.ps1`, `package.json`, `package-lock.json`) + Codex
  `config.toml`에 하드코딩돼 있다.
  `update-upstreams.ps1 -Apply`는 그중 4개만 갱신한다(npm이 `package.json`·
  `package-lock.json`, 텍스트 치환이 `verify.ps1`·`Verify.Tests.ps1`, `:382-395`).
  README·UPSTREAMS·USAGE_GUIDE·MCP 등록은 경고(`:367-370`)로만 넘기고, **HTML
  가이드는 경고 목록에도 없다** — 바로 그 파일이 어긋났다. 드리프트는
  구조적으로 재발한다.
- `docs/UPSTREAMS.md` 기준선 "`v1.17.0` / `96a2633…`"은 tag와 commit이 맞지
  않는다(`v1.17.0` = `0a7709a`, `96a2633`은 그 뒤 3커밋).
- `docs/USAGE_GUIDE.md:267` 제목 번호가 `### 5.`다(`5.2`여야 한다).
- 이전 `docs/ROADMAP.md`는 이미 끝난 P0/P1을 미완료로 적고 있었다.

**D. [저장소 위생] → P0-3, P0-5, P0-6**

- `.gitignore`에 `.omc/`, `.gjc/`가 없다.
- `.gitattributes`가 없다. `git ls-files --eol` 기준으로 두 BAT는 인덱스와
  작업 트리 모두 LF다(`core.autocrlf=true`인데도 그렇다). 지금은 테스트가
  통과하지만 cmd.exe는 CRLF를 전제로 하며, README가 "다른 PC에서 동일하게"
  절을 두고 있어 체크아웃마다 결과가 달라질 수 있는 상태다.
- `LICENSE`가 없다. 공개 저장소이고 README가 upstream 라이선스를 언급하면서
  자기 라이선스는 없다. Kordoc은 MIT, hwpx-skill도 MIT 계열이다.
- `CLAUDE.md`가 없다. 규칙은 `AGENTS.md`(Codex 규약)에만 있어서 Claude Code
  세션은 같은 규칙을 자동으로 읽지 않는다.

**E. [upstream 추적] → P1-7, P1-9**

- Kordoc `4.9.2`가 나왔고 `update-upstreams.ps1`이 이미 보고한다. 절차는 있는데
  실행되지 않았다.
- MCP는 `npx -y`로 매번 레지스트리를 본다. 변환기(lock)와 MCP(registry)의
  Kordoc이 서로 다른 버전이 될 수 있고, 오프라인에서는 MCP가 뜨지 않는다.

**F. [제품 공백] "HWP/HWPX 작업 통합"을 표방하지만 DOC 경로만 도구화돼 있다 → P2-14, P2-15, P2-16, P2-17**

- HWP(바이너리)→HWPX, HWP/HWPX/PDF/XLSX→Markdown 경로는 실행기도 문서도
  없다. 로컬 lock Kordoc CLI는 이미 HWP/HWPX/PDF/XLSX/DOCX/이미지→Markdown을
  지원한다(`--help` 확인).
- HWPX 생성·검증 스모크가 없다. 한컴오피스가 설치돼 있어
  `validate.py --hancom`까지 로컬에서 돌릴 수 있는데 게이트에 연결돼 있지 않다.
- uv 환경(python-hwpx, lxml, pywin32)이 hwpx skill 실행용이라는 사실이 어디에도
  명시돼 있지 않다. skill의 `SKILL.md`는 `python3` + `pip install`을 말하고
  저장소 규칙은 `uv run python`을 요구한다. 연결 지점이 없다.

**G. [성능] → P2-18**

- `convert-doc-to-md.ps1:161`은 파일마다 `powershell.exe` 자식 프로세스와 Word
  인스턴스를 새로 띄운다. N개 파일이면 Word 기동 N회, `Add-Type` 컴파일 N회.
  테스트 83초의 대부분도 Word 기동이다.

**H. [운영성·계약] → P1-11, P1-12, P1-13, P2-20, P3**

- 결과가 `Format-Table` 텍스트 + 성공 스트림 레코드라 `-File`로 호출하는
  에이전트(Codex/Claude)는 표를 파싱해야 한다. JSON 출력이 없다.
- 문서당 타임아웃이 없다. 억제되지 않은 대화상자나 손상 파일로 Word가
  멈추면 무한 대기한다.
- `-DocumentKey`/`-Password`는 명령행 평문이다(프로세스 목록·셸 기록 노출).
  통합 md 변환기에는 암호 옵션 자체가 없어 암호 DOC은 Markdown까지 못 간다.
- md 변환기에 `-LogPath`가 없다(DOCX 변환기에는 있다).
- BAT의 `chcp 65001`이 종료 후 복원되지 않아 호출한 콘솔의 코드페이지를
  바꾼다(사소).

**I. [CI 부재] → P1-10**

- `verify.ps1:93-99`는 Word COM을 필수 전제로 삼는다. Word 없는 머신이나
  GitHub Actions에서는 정적 검사(PSSA, `uv lock`, BAT·updater·verifier 단위
  테스트)조차 돌릴 수 없다. 테스트에 계층(tag)이 없다.

## 2. P0 — 다음 세션에서 닫는다

### P0-1. Kordoc 호출을 `.cmd` shim에서 `node.exe` 직접 호출로 바꾼다

- `Resolve-KordocCli`가 `tools/kordoc/node_modules/kordoc/dist/cli.js` 경로를
  돌려주고, 호출은 `& $node $cliJs --silent -o $staged $docx`로 바꾼다.
  `verify.ps1`의 `--version` 호출과 같은 방식이다.
- `-KordocCliPath`는 유지하되 `.js` 진입점을 받는 것으로 계약을 바꾼다.
  `ConverterSafety.Tests.ps1`의 가짜 `kordoc.cmd` 세 개(`:218`, `:264`, `:300`)는
  가짜 `.js`로 바꾼다. `verify.ps1:104-110`이 전제 조건으로 `.bin/kordoc.cmd`
  존재를 요구하므로 이것도 `dist/cli.js` 존재 검사로 바꾼다.
- 회귀 테스트: 공백 없이 `&`가 든 디렉터리(`a&b`)에 실제 DOC 픽스처를 넣고
  Markdown까지 변환되는지, 그리고 가짜 `.js`가 받은 인자가 전체 경로인지
  확인한다.
- 완료 기준: 새 테스트 통과, `Should -Not -Match "node_modules[\\/]+\.bin"`로
  shim 경로가 다시 들어오지 못하게 고정.

### P0-2. 실행기 테스트가 실제 `inbox`를 만지지 않게 한다

- BAT는 `%~dp0` 기준으로 동작하므로 테스트에서 BAT를 `TestDrive\repo\`로
  복사하고, 같은 위치에 인자를 파일로 기록하는 스텁
  `tools\doc-to-docx\convert-doc-to-*.ps1`과 빈 `inbox`를 만든다. 세 It(주입
  2건 + 무인자 1건) 모두 이 복사본을 실행하도록 바꾼다. 한 건만 옮기면 주입
  테스트가 계속 실제 `inbox`를 건드린다.
- 스텁이 받은 인자가 `-Path <TestDrive>\repo\inbox -Recurse`인지, 주입
  페이로드의 sentinel이 출력에 없는지 검사한다. Word를 띄우지 않고 BAT 배선만
  검증하므로 빠르고 부작용이 없다.
- BAT에 환경변수 override를 넣는 방식은 택하지 않는다. "무인자 inbox 전용"
  계약을 흐린다.

### P0-3. 저장소 위생 파일

- `.gitignore`: `.omc/`, `.gjc/` 추가.
- `.gitattributes` 신설: `* text=auto`와 `*.bat text eol=crlf` 두 줄이면 된다.
  `.ps1`은 LF로도 문제없으므로 강제하지 않는다(전체 CRLF 정규화 diff를
  피한다). 추가 뒤 `git add --renormalize .`로 BAT만 바뀌는지 `git diff --stat`으로
  확인하고 `verify.ps1`를 다시 돌린다.

### P0-4. 문서 드리프트 정리 (문장 교체만)

- `HWPX_wiz_easy_guide.html` 5곳 `4.9.0` → `4.9.1`.
- `docs/UPSTREAMS.md` 기준선을 `main@96a2633 (v1.17.0 + 3 commits)`처럼 tag와
  commit을 분리해 적는다.
- `docs/USAGE_GUIDE.md:267` → `### 5.2`.
- 이 문서로 `docs/ROADMAP.md`를 교체한다(이번 작업에서 완료).

### P0-5. `LICENSE` 추가

- MIT 권장(두 upstream과 같은 계열, 개인 통합 계층에 적합). README "원본
  프로젝트 고지" 절 아래에 저장소 자체 라이선스 한 줄을 추가한다.

### P0-6. `CLAUDE.md` 추가

- 내용은 `@AGENTS.md` 한 줄(import)이면 충분하다. Claude Code 세션이 Codex와
  같은 규칙(uv 전용, 원본 보존, BAT 무인자, `-Path` 계약)을 읽게 된다.
- 규칙은 계속 `AGENTS.md` 한 곳에서만 고친다.

P0 커밋 분할 제안:

1. `fix(converter): call Kordoc through node instead of the npm cmd shim`
   — P0-1 구현 + 테스트
2. `test(launcher): verify BAT wiring against a TestDrive copy` — P0-2
3. `chore: add gitattributes, LICENSE, CLAUDE.md, ignore runtime dirs`
   — P0-3, P0-5, P0-6
4. `docs: sync Kordoc pin and hwpx-skill baseline wording` — P0-4

## 3. P1 — 다음 정비 세션

### P1-7. Kordoc `4.9.2` 검토·적용

- `docs/UPSTREAMS.md` 절차 그대로: `chore/update-kordoc-2026-08-xx` 브랜치,
  `update-upstreams.ps1 -Component Kordoc -KordocVersion 4.9.2 -Apply`,
  `verify.ps1`, 대표 HWPX·PDF 읽기, MCP 등록 갱신, 문서 pin 갱신.
- P1-8을 먼저 하면 문서 pin 갱신이 테스트로 강제되므로 순서는 8 → 7이 낫다.

### P1-8. Kordoc 버전 pin의 단일 출처

- 출처는 `tools/kordoc/package.json` 하나로 한다. `verify.ps1:24`의
  `$script:RequiredKordocVersion` 하드코딩과 `Verify.Tests.ps1:92`의 리터럴을
  package.json에서 읽도록 바꾼다(verify는 package.json = lockfile = CLI 삼자
  일치만 검사).
- `tests/Docs.Tests.ps1` 신설: README, `docs/*.md`, HTML 가이드에서
  `kordoc@x.y.z` / `Kordoc x.y.z` / `4.9.x` 패턴을 찾아 package.json 버전과
  다르면 실패. `MAINTENANCE_HANDOFF.md`와 이 로드맵은 날짜가 박힌 스냅샷이므로
  제외한다.
- `update-upstreams.ps1`의 "문서를 검토하라" 경고는 유지하되, 이제 검토를
  빠뜨리면 `verify.ps1`가 잡는다.

### P1-9. Kordoc MCP를 로컬 lock에 맞춘다

- Codex 등록을 `npx.cmd -y kordoc@4.9.1 mcp`에서
  `node.exe <repo>\tools\kordoc\node_modules\kordoc\dist\cli.js mcp`로 바꾼다.
  변환기와 MCP가 같은 바이너리를 쓰고 오프라인에서도 뜬다.
- 절대 경로가 들어가므로 README "다른 PC에서" 절에 `codex mcp add` 명령을
  clone 경로 치환 형태로 적는다. `verify.ps1`에 "codex config의 kordoc 명령이
  로컬 cli.js를 가리키는가"를 경고 수준으로만 추가한다(설정 파일은 저장소
  밖이므로 실패로 만들지 않는다).

### P1-10. 검증 계층화와 GitHub Actions

- Pester tag를 도입한다: `Static`(Verify/UpstreamUpdate/HerdrCombo/Docs,
  BAT 배선 테스트), `Office`(ConverterSafety/ConverterIntegration).
- `verify.ps1 -Tier Static`은 Word COM 전제를 건너뛰고 PSSA, `uv lock --check`,
  `npm ci --dry-run`, Static 태그 테스트만 돌린다. 기본(-Tier Full)은 지금과
  같다.
- `.github/workflows/verify-static.yml`: `windows-latest`, `astral-sh/setup-uv`,
  `Install-Module Pester, PSScriptAnalyzer -Scope CurrentUser`,
  `npm ci --prefix tools/kordoc`, `verify.ps1 -Tier Static`. Word·한컴 계층은
  로컬 전용으로 남긴다. README의 "호스팅 CI를 제공하지 않는다" 문장을
  "정적 계층만 CI"로 바꾼다.

### P1-11. 변환기 출력 계약

- 두 변환기에 `-OutputFormat Table|Json` (기본 `Table`)을 추가한다. `Json`이면
  사람용 표를 생략하고 레코드 배열을 `ConvertTo-Json -Compress`로 성공
  스트림에 한 줄 출력한다. 에이전트 호출 예시를 USAGE_GUIDE 14절에 추가한다.
- md 변환기에 `-LogPath`를 추가해 DOCX 변환기와 옵션을 맞춘다(자식 변환기에
  그대로 전달).

### P1-12. 문서당 타임아웃

- `-TimeoutSeconds`(기본 300)를 추가한다. 구현은 소유한 Word 프로세스
  핸들(`$wordProcess`)을 감시하는 타이머로, 문서 하나가 시간을 넘기면 소유
  프로세스를 종료해 COM 호출이 예외로 빠지게 하고 `Failed` 레코드에
  `Reason = "Timeout"`을 남긴다. 소유하지 않은(Borrowed) Word는 절대 죽이지
  않는 기존 정책을 유지한다.
- 테스트: Word를 흉내 내는 스텁 COM 계층은 만들지 않는다(하지 말 것 참고).
  타임아웃 분기는 아주 짧은 타임아웃(1초)과 실제 픽스처로 재현하고, 최소한
  타이머 발화 → Failed 기록 → 소유 프로세스 잔류 0건만 검증한다.

### P1-13. 암호 전달 경로

- `-DocumentKey`를 `SecureString`으로도 받게 하고, 환경변수
  `HWPX_WIZ_DOC_PASSWORD`를 대안 입력으로 허용한다. 명령행 평문은 유지하되
  USAGE_GUIDE 9.1에 노출 경고를 그대로 둔다.
- md 변환기에 같은 옵션을 추가해 암호 DOC도 Markdown까지 갈 수 있게 한다.

## 4. P2 — 실제 문서 작업이 반복될 때

### P2-14. HWPX 생성·검증 스모크를 게이트에 연결

- `verify.ps1 -IncludeHwpx`(선택 단계): hwpx skill clone이 있으면
  `uv run python <skill>\scripts\validate.py <skill>\assets\report-template.hwpx --layout`,
  한컴 COM이 등록돼 있으면 `--hancom`까지 실행한다. skill이 없으면 skip, 한컴이
  없으면 `--hancom`만 skip. 실패로 만들지 않는다.
- 저장소에 HWPX 픽스처를 추가하지 않는다. skill 번들 자산을 쓰면 upstream
  갱신 시 자동으로 최신 템플릿을 검증하게 된다.

### P2-15. hwpx skill ↔ uv 환경 연결을 명시

- `AGENTS.md`에 한 줄: "hwpx skill 스크립트는 이 저장소 루트에서
  `uv run python <skill 경로>\scripts\<script>.py`로 실행한다. `pip install`을
  하지 않는다."
- `tools/hwpx/run.ps1` 얇은 래퍼(선택): skill 경로 해석 + `uv run python`
  호출만 한다. 로직을 넣지 않는다.
- `verify.ps1 -PrerequisiteCheckOnly`에 skill의 외부 import(`hwpx`, `lxml`,
  `win32com`)가 uv 환경에서 import되는지 한 줄 검사로 추가한다. 지금은 모두
  충족하므로 회귀 방지용이다.

### P2-16. HWP→HWPX, 임의 문서→Markdown 경로를 문서로 고정

- README "일상 작업 흐름"에 4단계 시퀀스를 한 블록으로 적는다:
  ① `.doc` → `.docx`/`.md`(이 저장소 변환기) ② `.hwp` → `.hwpx`(hwpx skill의
  `convert_hwp_hancom.ps1`, 한컴 없으면 rhwp 폴백) ③ 구조 확인·비교(Kordoc MCP)
  ④ 편집 가능 HWPX 생성·검증(hwpx skill). DOC을 HWPX로 바로 바꾸는 스크립트는
  만들지 않는다.
- USAGE_GUIDE 11절에 로컬 lock Kordoc CLI 직접 호출 예시
  (`node .\tools\kordoc\node_modules\kordoc\dist\cli.js -o 결과.md 문서.hwpx`)를
  추가한다. MCP 없이도 같은 버전으로 같은 결과를 얻는 경로다.

### P2-17. (조건부) `inbox` 범용 Markdown 실행기

- 비개발 사용자가 더블클릭으로 HWP/HWPX/PDF→Markdown을 반복 요구할 때만
  만든다. `convert-to-md.bat`(무인자, `inbox` 전용) + `tools/kordoc/convert-to-md.ps1`.
  `.doc`은 기존 통합 변환기로 위임하고, 그 외 Kordoc 지원 포맷은 CLI 직접
  호출. 원본 옆 저장, 스테이징→원자 교체, 기존 결과 skip, `-Overwrite` 정책을
  DOC 변환기와 동일하게 맞춘다.
- Kordoc이 이미 하는 일을 다시 구현하지 않는다. 래퍼는 경로·정책·요약만 담당한다.

### P2-18. 폴더 변환 시 Word 1회 기동

- `convert-doc-to-md.ps1`가 파일마다 자식 PowerShell을 띄우는 구조를,
  DOCX 변환기를 폴더 단위로 한 번 dot-source 호출해 레코드를 받고 파일별
  Kordoc만 반복하는 구조로 바꾼다. rollback 파일은 파일별로 그대로 유지한다.
- 기대 효과: N개 파일에서 Word 기동 N→1, 테스트 시간도 눈에 띄게 준다.
  P0-1과 P1-11(레코드 계약)이 끝난 뒤에 손대는 것이 안전하다.

### P2-19. 문서 통합

- README를 "무엇인지 + 빠른 시작 + 링크"로 줄이고, 상세는 USAGE_GUIDE 한
  곳으로 모은다. 업데이트 절차는 UPSTREAMS.md에만 둔다(현재 README와 중복).
- HTML 가이드는 손으로 유지하는 사본이다. 두 선택지 중 하나를 고른다:
  (a) Pages 워크플로에서 USAGE_GUIDE.md로부터 생성(pandoc 또는 marked +
  기존 CSS), (b) 손 유지 + P1-8의 버전 일치 테스트로 최소 드리프트만
  차단. 작업량이 늘기 전까지는 (b).
- Pages 워크플로가 `docs/*.md`도 같이 배포하게 하면 README 링크가 웹에서도
  동작한다.

### P2-20. 로그 위치 정책

- 기본 로그 없음을 유지한다. `-LogPath` 지정 시 부모 폴더가 없으면 지금처럼
  실패하되, `-CreateLogDirectory` 스위치로 생성을 허용한다. `output/logs/`를
  기본으로 만드는 안은 채택하지 않는다(원본 옆 저장 원칙과 충돌).

## 5. P3 — 기록만

- BAT `chcp 65001` 복원: 시작 시 `chcp` 출력에서 원래 코드페이지를 저장해 두고
  종료 직전 `chcp <원래값> >nul`로 되돌린다.
- GitHub Actions 액션을 SHA로 고정.
- `HerdrCombo.Tests.ps1`의 모델 ID 리터럴(`gpt-5.6-sol/luna`)은 모델 교체 시
  `.omo/omo.jsonc`와 함께 바꿔야 한다. 테스트가 잡아주므로 별도 작업은 없다.
- `python-hwpx` 갱신(`uv lock --upgrade-package python-hwpx`)을 Kordoc 갱신과
  같은 주기로 `update-upstreams.ps1`에 읽기 전용 보고만 추가.
- 역할 분담 재검토: Kordoc CLI에도 `generate`/`fill`/`patch`/`validate`가 있어
  hwpx skill과 기능이 겹친다. skill의 `references/kordoc-integration.md`를
  기준으로 "어느 쪽을 쓸지"를 README 역할 표에 한 줄로 못 박는다. 코드 변경
  없음.

## 6. 하지 말 것 (유지)

- `Kordoc_helper`나 자체 MCP 재도입
- BAT에 경로 인자나 환경변수 override 추가. cmd 재해석을 이길 수 없다
- 변환기 전체의 C#/Python 재작성
- Word·한컴 없는 환경에서 COM을 흉내 내는 mock 계층
- 스냅샷 차이로 사용자 Word 프로세스를 강제 종료하는 로직 복원
- 머신 절대 경로를 늘리는 설치 문서(P1-9의 MCP 등록은 예외이며 치환 형태로만 적는다)
- DOC→HWPX 직접 변환 스크립트(중복 변환기)
- 실제 업무 문서를 `tests/fixtures/`에 넣는 것

## 7. 추천 진행 순서

```text
세션 A (P0)   P0-1 shim 제거 → P0-2 inbox 격리 → P0-3 위생 파일
              → P0-4 문서 정리 → P0-5 LICENSE → P0-6 CLAUDE.md
              → verify.ps1 통과 → 커밋 4개
세션 B (P1)   P1-8 pin 단일화 → P1-7 Kordoc 4.9.2 → P1-9 MCP 로컬화
세션 C (P1)   P1-10 계층화 + CI → P1-11 JSON 출력 → P1-12 타임아웃 → P1-13 암호
P2            실제 HWPX 작업이 반복될 때만. 14 → 15 → 16 → 19 순, 17·18·20은 필요가 확인된 뒤
```

한 세션에 P0 전체면 충분하다. P0에 P1을 섞지 않는다.

## 8. 완료 기준

- 모든 세션 끝에 `tools/verify.ps1`가 통과한다(현재 44/44 → 항목마다 회귀
  테스트가 늘어난다).
- P0 종료 시: `&` 무공백 경로 Markdown 변환 테스트 통과, 실행기 테스트가
  `inbox`를 읽지 않음, `git status`에 런타임 디렉터리가 뜨지 않음, HTML
  가이드와 저장소의 Kordoc 버전 일치, `LICENSE`·`CLAUDE.md` 존재.
- P1 종료 시: Kordoc 버전 리터럴이 `tools/kordoc/package.json` 한 곳,
  `verify.ps1 -Tier Static`이 Word 없는 GitHub Actions에서 통과, 변환기
  `-OutputFormat Json` 사용 예시가 USAGE_GUIDE에 있음.

## 9. 이전 로드맵(2026-08-23) 대비 처리 결과

| 이전 항목 | 상태 | 비고 |
|---|---|---|
| P0-1 커밋 분할 | 완료 | 커밋 10개로 반영됨 |
| P0-2 BAT 무인자 문서 계약 | 완료 | README·AGENTS·UPSTREAMS 모두 `-Path` 계약 |
| P0-3 정비 스냅샷/상시 문서 분리 | 완료 | MAINTENANCE_HANDOFF에 역사 기록 머리말 |
| P1-4 한 명령 검증 | 완료 | `tools/verify.ps1`, 44/44 |
| P1-5 Kordoc npm lock | 완료 | `tools/kordoc/package-lock.json`, 변환기는 로컬 CLI 사용. 단 **`.cmd` shim 경유가 새 결함(A)** |
| P1-6 실패 픽스처 | 대부분 완료 | 암호·손상·LogPath·특수 이름 커버. 무공백 `&` 경로만 미커버 → P0-1 |
| P1-7 출력 계약 | 부분 | 레코드는 성공 스트림으로 나가고 dot-source도 구분됨. JSON 없음 → P1-11 |
| P2-8 HWPX 스모크 | 미착수 | → P2-14 (한컴 설치 확인됨) |
| P2-9 DOC→HWPX 문서 시퀀스 | 미착수 | → P2-16 |
| P2-10 로그 위치 | 미착수 | → P2-20 |
