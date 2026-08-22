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
  `custom` 브랜치에서 관리합니다.

## 검증된 기준선

2026-08-22에 다음 상태를 확인했습니다.

| 구성 요소 | 기준선 |
|---|---|
| Kordoc | `4.9.0` (`kordoc@4`가 해석한 최신 4.x) |
| hwpx-skill | `v1.17.0` / `0a7709aca5c0e66b9a94f8f335a5b28a5060af19` |
| Python 패키지 | `uv.lock` |

현재 MCP 명령은 `cmd /c npx -y kordoc@4 mcp`이므로 새 4.x가 나오면 자동으로
선택됩니다. 결과 재현성이 더 중요한 작업에서는 검증된 정확한 버전
(`kordoc@4.9.0` 형식)을 사용하고, 새 버전 검증이 끝난 뒤 기준선을 갱신합니다.

## 업데이트 절차

upstream은 한 번에 하나씩 업데이트합니다.

1. `HWPX_wiz`와 대상 upstream의 작업 트리가 깨끗한지 확인합니다.
2. `chore/update-<component>-<date>` 브랜치를 만듭니다.
3. 업데이트 전 버전 또는 커밋을 기록합니다.
4. Kordoc은 후보 버전을 명시해 실행하고, hwpx-skill은 `pull --ff-only`만
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

## 호환성 게이트

업데이트 후 최소한 다음 작업을 실제 문서 표면에서 확인합니다.

1. Kordoc으로 대표 HWPX 또는 PDF를 읽고 구조를 추출합니다.
2. `convert-doc-to-md.bat`으로 DOC → DOCX → Markdown 변환을 수행합니다.
3. hwpx-skill로 HWPX를 새로 만들고 다시 편집합니다.
4. namespace 수정, strict 검사와 layout 검증을 통과시킵니다.
5. 한컴오피스에서 최종 HWPX를 열어 페이지, 표와 글꼴을 눈으로 확인합니다.

검증 자료에는 개인정보가 없는 소형 문서만 사용합니다. 실제 업무 문서와 변환
결과는 `inbox` 또는 `output`에 두고 Git으로 추적하지 않습니다.

## 문제 발생 시

- 새 버전 적용 커밋만 되돌리고 마지막 검증 기준선을 사용합니다.
- 재현 절차와 사용한 버전 또는 커밋을 GitHub Issue에 기록합니다.
- 범용 수정은 개인 fork에만 장기 보관하지 말고 가능하면 upstream에
  Pull Request로 제출합니다.
