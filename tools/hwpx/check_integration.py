"""Exercise the installed skill from the repository's locked Python environment."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[2]


def check_build_report(report):
    if not (report.get('ok') is True and report.get('published') is True
            and report.get('report_saved') is True):
        raise ValueError('Skill build did not pass and publish with a saved report')
    if report.get('draft') or report.get('status') in ('DRAFT', 'FAIL', 'PARTIAL'):
        raise ValueError('Smoke fixture must be a complete, non-draft build')


def check_render_report(report, document):
    if report.get('ok') is not True or not report.get('page_images'):
        raise ValueError('Hancom page rendering did not succeed')
    digest = hashlib.sha256(document.read_bytes()).hexdigest()
    if report.get('source_sha256') != digest:
        raise ValueError('Rendered source hash differs from final HWPX')
    if not all(Path(p).is_file() for p in report['page_images']):
        raise ValueError('Rendered page evidence is missing')


def run_script(script, *args):
    env = dict(os.environ, PYTHONUTF8='1', PYTHONIOENCODING='utf-8')
    result = subprocess.run(
        [sys.executable, '-X', 'utf8', str(script), *map(str, args)],
        cwd=ROOT, env=env, capture_output=True, text=True, encoding='utf-8',
        timeout=90,
    )
    if result.returncode:
        raise RuntimeError(f'{script.name} failed ({result.returncode}): '
                           f'{result.stdout}\n{result.stderr}')
    return result.stdout


def check_integration(skill_root, hancom='required', render=False):
    skill_root = skill_root.resolve(strict=True)
    required = ['SKILL.md', 'references/reference-official-letter.md',
                'scripts/one_shot.py', 'scripts/render_hwpx.py']
    for name in required:
        if not (skill_root / name).is_file():
            raise ValueError(f'Installed skill lacks {name}; see docs/HWPX_INTEGRATION.md')
    if 'references/reference-official-letter.md' not in (skill_root / 'SKILL.md').read_text(encoding='utf-8'):
        raise ValueError('Installed SKILL.md does not route to reference-letter guidance')
    contract = json.loads((ROOT / 'tools/hwpx/skill-lock.json').read_text(encoding='utf-8'))
    revision = subprocess.run(['git', '-C', str(skill_root), 'rev-parse', 'HEAD'],
                              capture_output=True, text=True, timeout=10, check=True).stdout.strip()
    if revision != contract['commit']:
        raise ValueError('Installed skill commit differs from tools/hwpx/skill-lock.json')
    guide_hash = hashlib.sha256((skill_root / contract['reference_guide']).read_text(
        encoding='utf-8').encode('utf-8')).hexdigest()
    if guide_hash != contract['guide_sha256_lf']:
        raise ValueError('Installed reference-letter guide differs from the verified contract')
    if render and hancom != 'required':
        raise ValueError('Page render requires hancom=required')
    output = ROOT / 'output'
    output.mkdir(exist_ok=True)
    evidence = Path(tempfile.mkdtemp(prefix='hwpx-integration-', dir=output))
    document = evidence / 'letter.hwpx'
    # Synthetic names only. No private documents, seals, contacts or accounts.
    spec = {
        'version': 1, 'kind': 'official-letter', 'output': str(document),
        'metadata_date': '2000-01-01',
        'document': {
            '기관명': '서식검증기관', '수신': '검증수신기관장',
            '제목': '문서 생성 연결 확인', '발신명의': '서식검증기관장',
            'body': ['1. 문서 생성 연결을 확인합니다.',
                     '2. 편집 가능한 본문과 붙임 표시를 검증하여 주시기 바랍니다.'],
            '붙임': ['연결 확인 내역 1부.'], '끝': True,
        },
        'quality': {'hancom': hancom, 'required_text': ['문서 생성 연결 확인', '연결 확인 내역']},
    }
    spec_path = evidence / 'spec.json'
    report_path = evidence / 'quality.json'
    spec_path.write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding='utf-8')
    run_script(skill_root / 'scripts/one_shot.py', spec_path, '--report', report_path)
    report = json.loads(report_path.read_text(encoding='utf-8'))
    check_build_report(report)
    if not document.is_file():
        raise ValueError('Published HWPX is missing')
    if hancom == 'required' and report.get('checks', {}).get('hancom', {}).get('ok') is not True:
        raise ValueError('Required Hancom open check did not pass')
    if render:
        run_script(skill_root / 'scripts/render_hwpx.py', document,
                   '--output-dir', evidence / 'render')
        rendered = json.loads((evidence / 'render/render.json').read_text(encoding='utf-8'))
        check_render_report(rendered, document)
    result = {
        'ok': True, 'skill_root': str(skill_root), 'evidence': str(evidence),
        'skill_commit': revision, 'guide_sha256_lf': guide_hash,
        'document_sha256': hashlib.sha256(document.read_bytes()).hexdigest(),
        'hancom': hancom, 'render': 'passed' if render else 'not_run',
        'visual_review': 'not_run',
    }
    (evidence / 'integration.json').write_text(json.dumps(result, indent=2), encoding='utf-8')
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--skill-root', required=True, type=Path)
    parser.add_argument('--hancom', choices=['required', 'off'], default='required')
    parser.add_argument('--render', action='store_true')
    args = parser.parse_args()
    try:
        result = check_integration(args.skill_root, args.hancom, args.render)
    except (OSError, ValueError, RuntimeError, subprocess.SubprocessError) as exc:
        result = {'ok': False, 'error': str(exc)}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result['ok'] else 1


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
