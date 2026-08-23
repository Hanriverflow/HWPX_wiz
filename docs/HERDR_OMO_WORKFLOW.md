# Herdr + OMO 작업 흐름

이 저장소는 Herdr를 세션을 유지하는 outer runtime으로, OMO Native를 각 pane
내부의 agent orchestrator로 사용한다.

```text
PLAN / Sol high
  └─ GOAL.md
       ↓
BUILD / Luna xhigh
  └─ ulw-loop + task/team
       ↓
VERIFY / Sol xhigh
  └─ 독립 요구사항 감사
```

구현 파일을 직접 수정하는 기본 pane은 BUILD 하나뿐이다. PLAN은 `GOAL.md`만
작성하고 VERIFY는 read-only permission preset으로 실행된다. 병렬 구현이 꼭
필요하면 OMO worker별 Git worktree를 사용한다.

## 프로젝트 설정

프로젝트 전용 OMO 설정은 `.omo/omo.jsonc`에 있다. 이 파일이 사용자 범위
`~/.omo/omo.jsonc`보다 높은 우선순위로 병합되므로 다른 프로젝트의 model
routing은 바꾸지 않는다.

| 역할 | 모델 |
|---|---|
| PLAN | `openai-codex/gpt-5.6-sol`, high |
| BUILD | `openai-codex/gpt-5.6-luna`, xhigh |
| VERIFY | `openai-codex/gpt-5.6-sol`, xhigh |
| `quick`, `explore`, `librarian` | Luna low |
| `deep`, 구현 category | Luna xhigh |
| `ultrabrain`, `momus` | Sol xhigh |
| `writing`, `metis` | Sol high |

Herdr에는 OMO 전용 integration target이 없으므로 `omp` integration을 대신
설치하지 않는다. OMO는 일반 pane process로 실행하고, pane 제어는 설치된
Herdr skill과 `herdr pane` 명령을 사용한다.

## 시작

먼저 Herdr 안에서 저장소를 연다.

```powershell
cd D:\Code\Projects\HWPX_wiz
herdr
```

Herdr-managed pane에서 다음을 실행한다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\herdr\start-combo.ps1 -Focus
```

스크립트는 `OMO Combo` tab을 만들고 아래 layout을 구성한 뒤 각 shell prompt에
맞는 OMO 명령을 입력해 둔다. 자동 실행하지 않으므로 원하는 lane부터 해당
pane에서 Enter를 눌러 시작한다.

```text
┌──────────────────────────────────────┐
│ PLAN - SOL                           │
├───────────────────┬──────────────────┤
│ BUILD - LUNA      │ VERIFY - SOL     │
└───────────────────┴──────────────────┘
```

같은 label의 tab이 이미 있으면 기존 tab을 재사용한다. 구성만 확인하고 pane을
만들지 않으려면 다음을 실행한다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\herdr\start-combo.ps1 -ValidateOnly
```

현재 OMO beta는 빈 interactive session 시작 시에도 내부 도구 permission을
요청할 수 있다. 그래서 launcher가 세 agent를 동시에 자동 실행하지 않는다.
PLAN과 VERIFY의 permission preset을 약화해 이 prompt를 우회하지 말고, 필요한
lane만 시작해 요청 내용을 확인한 뒤 승인한다.

## 운용 순서

### PLAN

PLAN pane에서 `/plan`을 실행한 뒤 작업 요구사항을 전달한다. PLAN은 구현하지
않고 저장소 루트의 `GOAL.md`만 작성한다. `GOAL.md`는 작업별 임시 계약이므로
Git에서 제외된다.

### BUILD

`GOAL.md`를 검토한 뒤 BUILD pane에서 실행한다.

```text
/skill:ulw-loop
/build
```

필요하면 이어서 다음처럼 구체적인 실행 지시를 보낸다.

```text
Execute GOAL.md through verified completion.
Read AGENTS.md and GOAL.md first.
Keep implementation ownership in this pane.
Use isolated worktrees for parallel implementation.
For every delegated task state TASK, DELIVERABLE, SCOPE, VERIFY, and STOP WHEN.
```

BUILD가 특정 실패 사례만 때우거나 scope를 줄이면 `GOAL.md`를 다시 읽고 첫
미충족 acceptance criterion부터 계속하라고 짧게 steer한다.

### VERIFY

BUILD 완료 후 VERIFY pane에서 `/verify`를 실행하고 다음처럼 독립 감사를
요청한다.

```text
Audit every GOAL.md acceptance criterion against the complete diff,
relevant tests, and actual command output. Classify each criterion as
PASS, PARTIAL, FAIL, or NOT VERIFIED. Do not edit files.
```

FAIL 또는 PARTIAL은 BUILD pane으로 되돌려 수정하고 다시 VERIFY한다.

## 사전 점검

```powershell
herdr --version
omo --version
omo doctor
omo auth check --model openai-codex/gpt-5.6-sol
omo auth check --model openai-codex/gpt-5.6-luna
```

현재 workflow는 Fast mode 상속을 전제로 하지 않는다. Luna Fast 지원 여부와
관계없이 BUILD 및 child routing은 명시된 reasoning level로 동작해야 한다.
