# HWPX 작성 경로와 검증

## 연결 계약

HWPX_wiz는 Kordoc으로 문서를 읽고, 별도 설치한 `hwpx` skill로 편집 가능한
HWPX를 작성한다. 프로젝트 `AGENTS.md`와 `CLAUDE.md`는 이 역할 구분을 안내한다.
실제 공문 양식을 받으면 스킬의 `references/reference-official-letter.md`를 먼저
읽는다. 원고의 사실과 참고 양식의 시각 요소는 별도로 취급한다.

공문 재작성 경험을 반영한 스킬은 개인 fork의 수정 브랜치에 보관한다.
원본 코드를 HWPX_wiz에 복사하지 않고, [`skill-lock.json`](../tools/hwpx/skill-lock.json)에
저장소·브랜치·커밋과 공문 지침의 LF 정규화 SHA-256을 기록한다.
현재 수정은 지침 4개 파일이며 생성기 자체는 upstream v1.18.0을 사용한다.

새 PC에서는 **대상 폴더가 없는 경우에만** 다음과 같이 설치한다.

```powershell
$hwpxSkill = Join-Path $HOME '.agents\skills\hwpx'
git clone --branch codex/reference-official-letter `
  https://github.com/Hanriverflow/hwpx-skill.git $hwpxSkill
git -C $hwpxSkill rev-parse HEAD
uv sync --locked
```

기존 설치가 있으면 clone하거나 덮어쓰지 않는다. Git 상태·링크 대상·원격을 먼저
확인하고 로컬 변경을 보존한 채 위 lock의 커밋을 사용할 방법을 검토한다.
이 PC의 `.agents/skills/hwpx`는 정본 디렉터리를 가리키는 junction이므로 별도
복사본을 수정할 필요가 없다. 사용자별 절대 경로는 저장소 계약에 넣지 않는다.

## 실행 환경

저장소 루트의 `uv run python`을 사용한다. `pyproject.toml`과 `uv.lock`은
`python-hwpx`, `lxml`, `pywin32`, `jsonschema`, `Pillow`를 관리한다.
한컴 Automation과 Poppler는 별도 설치 조건이다. 자동 설치나 전역 pip는 쓰지 않는다.

```powershell
$hwpxSkill = Join-Path $HOME '.agents\skills\hwpx'
uv run --locked python -X utf8 tools\hwpx\check_integration.py `
  --skill-root $hwpxSkill --hancom required --render
```

이 검사는 다음을 수행한다.

1. 실제 링크 대상, 고정 커밋, 공문 참고 지침과 진입점 연결 확인
2. 개인정보 없는 작은 공문을 `one_shot.py`로 생성
3. 구조·내용·참조 검사 및 게시/리포트 저장 성공 확인
4. 독립 한컴 열림과 페이지 PNG 렌더, 최종 HWPX 해시 대조

근거 파일은 실행마다 새 `output/hwpx-integration-*` 폴더에 남는다.
자동 검사는 `visual_review: not_run`으로 기록한다. 사람이 모든 PNG를 보고
잘림·표 분할·글꼴과 참고 양식 충실도를 확인해야 시각 검토 완료라고 할 수 있다.
이 작은 공문은 실행 연결 검사이며 모든 참고 양식의 자동 재현을 입증하지 않는다.

전체 프로젝트 게이트와 함께 실행하려면:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\verify.ps1 -IncludeHwpx
```

설치 경로가 다르면 `-HwpxSkillPath`를 지정한다. `-IncludeHwpx`를 명시한 경우
스킬 누락·커밋/지침 불일치·생성 실패·한컴/렌더 실패는 종료 코드 실패다.
일반 `-Tier Static` CI는 외부 스킬·한컴을 요구하지 않으며 통합 검사 코드의
실패 처리 테스트만 수행한다. 한컴 없이 생성만 점검하려면 직접 실행기에
`--hancom off`를 쓰고 `--render`를 생략하되 한컴 검증 완료로 보고하지 않는다.

## 갱신과 보존

개인 수정 브랜치에서는 기존 `update-upstreams.ps1 -Component HwpxSkill -Apply`가
중단하는 것이 정상이다. 이 도구는 깨끗한 공식 `main`의 fast-forward만 지원한다.
수정 브랜치를 무시하고 강제로 `main`으로 바꾸거나 reset하지 않는다.
새 upstream을 통합하려면 수정 브랜치에서 별도 검토하고, 검증 후 스킬을 푸시한
다음 HWPX_wiz의 lock과 기준선을 함께 갱신한다. 공식 upstream에 대한 PR은
별도 요청이 있을 때만 제출한다.

고객 원고, 참고 PDF, 직인, 거래내역, 계좌·연락처 및 생성물은 Git에 넣지 않는다.
