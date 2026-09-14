# HWPX_wiz 작업 규칙

- Python 환경과 명령은 반드시 `uv`로 관리한다. 전역 `pip install`을 사용하지 않는다.
- Python 실행은 `uv run python ...`, 환경 동기화는 `uv sync --locked`를 사용한다.
- HWP/HWPX/DOCX/PDF/XLSX 문서의 읽기, 구조 추출, 비교와 변환은 공식 `kordoc` MCP를 우선 사용한다.
- 편집 가능한 HWPX 생성·수정 요청에는 설치된 `hwpx` skill을 사용하고 그 검증 절차를 끝까지 따른다.
- hwpx skill 스크립트는 저장소 루트에서 `uv run python <skill 경로>\scripts\<script>.py`로 실행하며 전역 `pip install`을 하지 않는다.
- 실제 공문 양식을 제공받으면 빌드 전에 해당 skill의 `references/reference-official-letter.md`를 읽는다. Kordoc으로 내용·구조를 파악하고 참고 페이지를 시각적으로 확인하며, 원고 사실과 참고 양식을 분리한다. 자세한 연결·검증은 `docs/HWPX_INTEGRATION.md`를 따른다.
- 구형 Word `.doc`는 Kordoc에 직접 넘기지 않는다. 루트의 `convert-doc-to-docx.bat`과 `convert-doc-to-md.bat`은 모든 인자를 무시하고 프로젝트 `inbox`만 처리한다.
- 특정 파일이나 폴더는 `tools/doc-to-docx/convert-doc-to-docx.ps1` 또는 `tools/doc-to-docx/convert-doc-to-md.ps1`에 `-Path`로 전달한다.
- 사용자가 입력 파일의 경로를 지정하면 `.docx`와 `.md` 결과를 원본 파일과 같은 폴더에 같은 기본 이름으로 저장한다.
- 사용자가 경로를 지정하지 않으면 `inbox`에서 입력을 찾는다. DOC 변환 결과는 이 경우에도 해당 원본 `.doc` 옆에 저장한다.
- `output`은 사용자가 중앙 결과 폴더를 명시했거나, 별도로 생성하는 HWPX·보고서·로그를 모을 때 사용한다.
- 원본 문서는 덮어쓰거나 삭제하지 않는다. 기존 결과 파일도 사용자가 덮어쓰기를 명시하지 않으면 보존한다.
- upstream 저장소를 갱신할 때는 먼저 작업 트리가 깨끗한지 확인하고 `git pull --ff-only`만 사용한다.
- `Kordoc_helper`의 자체 MCP나 애플리케이션 코드는 사용하지 않는다.
