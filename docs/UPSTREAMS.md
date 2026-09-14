# Upstream 관리 정책

`HWPX_wiz`는 Kordoc과 hwpx-skill의 소스를 합치는 저장소가 아니라, 두 도구를
일관된 방식으로 사용하는 개인 통합·자동화 계층입니다.

## 저장소 경계

| 구성 요소 | 관리 방식 |
|---|---|
| `HWPX_wiz` | 실행기, 프로젝트 규칙, Python 잠금 파일과 운영 문서를 직접 관리 |
| [Kordoc](https://github.com/chrisryugj/kordoc) | 공식 npm 패키지와 MCP를 통해 사용 |
| [hwpx-skill](https://github.com/jkf87/hwpx-skill) | 별도 로컬 clone을 공식 upstream에서 fast-forward로 갱신 |

- upstream 소스를 이 저장소로 복사하거나 Git submodule로 포함하지 않습니다.
- upstream을 수정해야 할 때만 해당 저장소를 개인 계정으로 fork합니다.
- fork의 `main`은 공식 upstream과 동일하게 유지하고, 개인 수정은 별도
  수정 브랜치에서 관리합니다. 현재는 `codex/reference-official-letter`입니다.

## 검증된 기준선

2026-09-14에 다음 상태를 확인했습니다.

| 구성 요소 | 기준선 |
|---|---|
| Kordoc | `4.13.1` (`kordoc@4`가 해석한 최신 4.x) |
| hwpx-skill | `main@34b34f99ee29d930efdfd0c3fc428909e03a9f81` (`v1.18.0`) |
| hwpx-skill 개인 지침 | `Hanriverflow/hwpx-skill`, `codex/reference-official-letter@40b5d7e06bb68b87099974ff76cd0d23ae1c367d` |
| Python 패키지 | `uv.lock` (`python-hwpx 6.3.0`; 2026-08-25 outdated 보고 없음) |

실제 활성 스킬은 위 개인 지침 브랜치를 사용합니다. 원본 생성기는 그대로 두고
공문 양식 분석·재작성 지침을 추가했습니다. 고정 커밋·지침 해시, 새 PC 복원 및
실행 검증은 [`HWPX_INTEGRATION.md`](HWPX_INTEGRATION.md)와
[`skill-lock.json`](../tools/hwpx/skill-lock.json)을 따릅니다. 수정 브랜치에서는
아래 자동 적용기의 `main` 전용 안전장치가 의도적으로 적용을 거부합니다.

현재 MCP와 DOC→Markdown 변환기는 모두
`tools\kordoc\node_modules\kordoc\dist\cli.js`를 사용해 검증 기준선을
고정합니다. 새 버전은 아래 호환성 게이트를 통과한 뒤 로컬 lock,
MCP 설정과 이 문서의 기준선을 함께 갱신합니다.

## 업데이트 절차

### 새 버전 확인과 승인

새 버전 확인은 전용 스크립트로 수행합니다. 기본 실행은 npm registry와
공식 `hwpx-skill` 원격을 읽기만 하며, 로컬 파일과 외부 clone을 변경하지
않습니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\update-upstreams.ps1
```

보고서에서 후보 버전과 현재 작업 트리 상태를 검토한 뒤에만 적용합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\update-upstreams.ps1 -Apply
```

`-Apply`는 구성 요소별로 다시 확인을 요청합니다. 이미 승인한 후보를
자동화 환경에서 적용할 때만 `-Yes`를 함께 사용합니다.

```powershell
# 특정 Kordoc 후보를 명시
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\update-upstreams.ps1 `
  -Component Kordoc -KordocVersion 4.13.1 -Apply

# hwpx-skill만 확인·적용
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\update-upstreams.ps1 `
  -Component HwpxSkill -Apply
```

스크립트는 Kordoc의 같은 major 최신 버전을 기본 후보로 선택하고, major
업데이트는 자동 선택하지 않습니다. `hwpx-skill`에 변경 사항이 있거나
`main`이 아니면 `pull --ff-only` 전에 중단합니다. Kordoc 적용 후에는
`tools/verify.ps1`, Codex MCP 등록, README와 이 문서의 버전 pin을 함께
검토합니다. 확인을 자동화하되 적용은 막으려면 `-FailOnUpdate`를 사용하고,
업데이트가 발견되면 종료 코드 `10`을 처리합니다.

upstream은 한 번에 하나씩 업데이트합니다.

1. `HWPX_wiz`와 대상 upstream의 작업 트리가 깨끗한지 확인합니다.
2. `chore/update-<component>-<date>` 브랜치를 만듭니다.
3. 업데이트 전 버전 또는 커밋을 기록합니다.
4. Kordoc은 정확한 후보 버전을 명시해 실행하고, hwpx-skill은 `pull --ff-only`만
   사용합니다.
5. 아래 호환성 게이트를 통과한 뒤 기준선과 관련 잠금 파일을 갱신합니다.
6. 의존성별로 독립된 커밋을 만들고 `main`에 반영합니다.

hwpx-skill 갱신:

```powershell
git -C C:\Users\Hank\.agents\skills\hwpx status -sb
git -C C:\Users\Hank\.agents\skills\hwpx pull --ff-only
git -C C:\Users\Hank\.agents\skills\hwpx describe --tags --exact-match HEAD
git -C C:\Users\Hank\.agents\skills\hwpx rev-parse HEAD
```

Python 패키지 갱신:

```powershell
uv lock --upgrade-package python-hwpx
uv sync --locked
```

로컬 Kordoc 변환기 갱신:

```powershell
npm ci --prefix .\tools\kordoc
node .\tools\kordoc\node_modules\kordoc\dist\cli.js --version
```

변환기와 MCP는 `tools\kordoc\package-lock.json`의 정확한 `kordoc@4.13.1`만
사용합니다. 두 경로 모두 registry-backed `npx` 호출 없이 로컬
`node_modules\kordoc\dist\cli.js`를 직접 실행합니다.

## 호환성 게이트

저장소 자체의 기본 검증 게이트는 저장소 루트에서 먼저 실행합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\verify.ps1
```

업데이트 후 최소한 다음 작업을 실제 문서 표면에서 확인합니다.

1. Kordoc으로 대표 HWPX 또는 PDF를 읽고 구조를 추출합니다.
2. 인자를 무시하는 `convert-doc-to-docx.bat`과 `convert-doc-to-md.bat`이 프로젝트 `inbox`를 처리하는지 확인합니다. 특정 파일이나 폴더는 `tools/doc-to-docx/convert-doc-to-docx.ps1` 또는 `convert-doc-to-md.ps1`에 `-Path`로 전달해 변환합니다.
3. hwpx-skill로 HWPX를 새로 만들고 다시 편집합니다.
4. namespace 수정, strict 검사와 layout 검증을 통과시킵니다.
5. 한컴오피스에서 최종 HWPX를 열어 페이지, 표와 글꼴을 눈으로 확인합니다.
6. 위 저장소 검증 게이트를 통과시킵니다.

검증기는 도구를 자동 설치하지 않습니다. Pester 6.1.0 이상과
PSScriptAnalyzer 1.25.0 이상은 사용자가 사용자 범위에 설치하고, PATH의 `uv`와
`uv sync --locked`로 만든 Python 환경도 준비해야 합니다. 실패 원인을 좁힐 때만
저장소 루트에서 다음 개별 명령을 실행합니다.

```powershell
Import-Module Pester -MinimumVersion 6.1.0 -Force
Invoke-Pester -Path .\tests -Output Detailed
Import-Module PSScriptAnalyzer -MinimumVersion 1.25.0 -Force
Invoke-ScriptAnalyzer -Path .\tools -Recurse
Invoke-ScriptAnalyzer -Path .\tests -Recurse
uv lock --check
```

GitHub Actions는 Word 없이 실행 가능한 `tools\verify.ps1 -Tier Static`만
검증합니다. 위 실제 문서 확인과 기본 `-Tier Full`에 필요한 Word나
한컴오피스의 설치 및 사용 가능 상태는 로컬 Windows에서 확인합니다.

검증 자료에는 개인정보가 없는 소형 문서만 사용합니다. 실제 업무 문서와 변환
결과는 `inbox` 또는 `output`에 두고 Git으로 추적하지 않습니다.

## 문제 발생 시

- 새 버전 적용 커밋만 되돌리고 마지막 검증 기준선을 사용합니다.
- 재현 절차와 사용한 버전 또는 커밋을 GitHub Issue에 기록합니다.
- 범용 수정은 개인 fork에만 장기 보관하지 말고 가능하면 upstream에
  Pull Request로 제출합니다.
