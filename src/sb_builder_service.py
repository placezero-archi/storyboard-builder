"""스토리보드 빌더 웹 서비스 — 요구사항 → storyboard.json(claude headless) → PPTX 렌더

빌더 본체: /Users/nicesso/ai-demo/storyboard-builder
생성 엔진: `claude -p` (사용자 구독 CLI, headless) — 실패 시 에러 보고(폴백 없음, 품질 우선)
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import threading
import unicodedata
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

SB_ROOT = Path("/Users/nicesso/ai-demo/storyboard-builder")
PROJECTS = SB_ROOT / "projects"
GUIDE = SB_ROOT / "standards" / "generation-guide.md"
STYLE = SB_ROOT / "standards" / "style-spec.md"
FEWSHOT = SB_ROOT / "projects" / "demo-preorder" / "storyboard.json"
RENDER = SB_ROOT / "builder" / "render_pptx.py"
EXTRACT = SB_ROOT / "builder" / "extract_source.py"
VENV_PY = SB_ROOT / ".venv" / "bin" / "python"
SOURCE_TEXT_LIMIT = 200000  # claude 프롬프트에 넣는 원본 텍스트 상한
CLAUDE_BIN = "/Users/nicesso/.local/bin/claude"

_jobs: dict[str, dict[str, Any]] = {}
_lock = threading.Lock()


def _clean_env() -> dict[str, str]:
    """부모 env 상속 + 인증 충돌 변수만 제거 (완전 빈 env는 키체인 접근 실패)"""
    env = dict(os.environ)
    for k in list(env):
        if k.startswith(("ANTHROPIC_", "CLAUDE")):
            env.pop(k)
    env["PATH"] = "/Users/nicesso/.local/bin:/opt/homebrew/bin:/usr/bin:/bin"
    env.setdefault("HOME", os.path.expanduser("~"))
    env.setdefault("LANG", "ko_KR.UTF-8")
    return env


def _run_claude(prompt: str, timeout: int = 900, max_turns: int = 1) -> str:
    result = subprocess.run(
        [CLAUDE_BIN, "-p", prompt, "--max-turns", str(max_turns)],
        capture_output=True, text=True, timeout=timeout, env=_clean_env(),
        cwd=str(SB_ROOT),
    )
    if result.returncode != 0:
        detail = (result.stderr or "").strip()[-200:] + " | " + (result.stdout or "").strip()[-200:]
        raise RuntimeError(f"claude 실행 실패: {detail}")
    return result.stdout


def _extract_json(raw: str) -> dict:
    # 코드펜스/서문 제거 후 가장 바깥 { } 추출
    m = re.search(r"\{.*\}", raw, re.DOTALL)
    if not m:
        raise ValueError("응답에서 JSON을 찾지 못함")
    return json.loads(m.group(0))


def _slugify(title: str) -> str:
    slug = re.sub(r"[^0-9A-Za-z가-힣]+", "-", title).strip("-").lower()
    return f"web-{slug[:40] or uuid.uuid4().hex[:8]}"


def _render(project_dir: Path, data: dict) -> Path:
    json_path = project_dir / "storyboard.json"
    json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    result = subprocess.run(
        [str(VENV_PY), str(RENDER), str(json_path)],
        capture_output=True, text=True, timeout=180, cwd=str(SB_ROOT),
    )
    if result.returncode != 0:
        raise RuntimeError(f"렌더 실패: {(result.stderr or result.stdout)[-400:]}")
    ver = data.get("meta", {}).get("version", "0.1")
    pptx = project_dir / f"{data['meta']['title']}_v{ver}.pptx"
    if not pptx.exists():
        raise RuntimeError("렌더 결과 pptx를 찾지 못함")
    return pptx


def _gen_prompt(form: dict) -> str:
    guide = GUIDE.read_text(encoding="utf-8")
    fewshot = FEWSHOT.read_text(encoding="utf-8")
    today = datetime.now().strftime("%Y. %m. %d")
    return f"""너는 스마일게이트 퍼블리싱기획팀 스타일의 화면설계서(storyboard.json)를 작성하는 전문 기획자다.

## 작성 규칙 (요약 — 반드시 준수)
{guide}

## 참조 예시 (동일 스키마의 완성본 — 구조와 문체를 그대로 따를 것)
```json
{fewshot}
```

## 이번 요청
- 문서 제목: {form.get('title')}
- 카테고리(게임/서비스명): {form.get('category')}
- 플랫폼: {form.get('platform', 'PC/MO')}
- 지원 언어: {form.get('langs', 'ko')}
- 어드민 포함: {form.get('admin', False)}
- 작성자: 장소영 / 퍼블리싱기획팀 / 작성일: {today} / version: 0.1
- 요구사항:
\"\"\"
{form.get('requirements', '')}
\"\"\"

## 출력 지시
1. 위 스키마와 동일한 구조의 storyboard.json **하나만** 출력한다. 설명·코드펜스 금지, 순수 JSON만.
2. 페이지 구성: 정책 1장 + PC 화면 2~4장 + MO 1장{' + 어드민 목록/편집 2장' if form.get('admin') else ''} + 필요 시 플로우차트 1장. 총 5~8장.
3. 요구사항에 없는 수치·기능은 지어내지 말고 ==미정==으로 표기.
4. Description은 개조식 명사형 종결, etype 필수, 값 범위·얼럿 문구 명시.
5. history는 [{{"date": "{datetime.now().strftime('%Y.%m.%d')}", "items": ["최초 작성"], "drafter": "장소영"}}] 하나만.
"""


def _extract_source(path: Path) -> str:
    result = subprocess.run(
        [str(VENV_PY), str(EXTRACT), str(path)],
        capture_output=True, text=True, timeout=180, cwd=str(SB_ROOT),
    )
    if result.returncode != 0:
        raise RuntimeError(f"텍스트 추출 실패: {(result.stderr or result.stdout)[-300:]}")
    text = result.stdout.strip()
    if not text:
        raise RuntimeError("파일에서 텍스트를 추출하지 못했습니다 (이미지 전용 문서일 수 있음)")
    if len(text) > SOURCE_TEXT_LIMIT:
        text = text[:SOURCE_TEXT_LIMIT] + "\n\n[... 이후 내용 생략 (원본이 김) ...]"
    return text


def _import_prompt(form: dict, source_text: str) -> str:
    guide = GUIDE.read_text(encoding="utf-8")
    fewshot = FEWSHOT.read_text(encoding="utf-8")
    today = datetime.now().strftime("%Y. %m. %d")
    return f"""너는 스마일게이트 퍼블리싱기획팀 스타일의 화면설계서(storyboard.json)를 작성하는 전문 기획자다.
아래는 **기존에 작성된 기획서 문서에서 추출한 원문**이다. 이 내용을 storyboard.json 스키마로 변환하라.

## 작성 규칙 (요약 — 반드시 준수)
{guide}

## 참조 예시 (동일 스키마의 완성본 — 구조만 참고)
```json
{fewshot}
```

## 기존 문서 원문 (추출 텍스트)
\"\"\"
{source_text}
\"\"\"

## 변환 지시
1. 위 스키마와 동일한 구조의 storyboard.json **하나만** 출력한다. 설명·코드펜스 금지, 순수 JSON만. 도구 사용·파일 읽기/쓰기 금지 — 최종 응답 본문에 JSON을 직접 출력할 것.
2. **원문 충실 변환이 최우선** — 원문의 정책·수치·문구·화면 구성을 그대로 보존한다. 원문에 없는 내용은 절대 지어내지 말 것.
3. 원문의 각 페이지를 가장 가까운 ptype(policy/screen_pc/screen_mo/admin/flowchart/decision)으로 매핑한다. 애매하면 policy 표로.
4. 원문에 History 표가 있으면 그대로 history 배열로 옮기고, 끝에 {{"date": "{datetime.now().strftime('%Y.%m.%d')}", "items": ["기존 문서 업로드 · 빌더 변환"], "drafter": "장소영"}} 추가.
5. meta.title은 "{form.get('title')}", category는 "{form.get('category', '')}", version은 원문에 있으면 그대로, 없으면 "1.0". 작성일: {today}.
6. 원문이 길어 일부 생략된 경우, 확인 필요한 부분은 ==원본 확인 필요==로 표기.
"""


def _revise_prompt(current: dict, instruction: str) -> str:
    today = datetime.now().strftime("%Y.%m.%d")
    return f"""아래는 화면설계서 storyboard.json이다. 수정 지시에 따라 갱신하라.

## 현재 JSON
```json
{json.dumps(current, ensure_ascii=False)}
```

## 수정 지시
{instruction}

## 규칙
1. meta.version을 0.1 올리고 meta.updated를 "{datetime.now().strftime('%Y. %m. %d')}"로.
2. history 배열 끝에 {{"date": "{today}", "items": ["<영역> > <변경 내용> (p<대략적 페이지>)"], "drafter": "장소영"}} 추가 — 항목은 실제 변경마다 1줄.
3. 폐기되는 스펙은 삭제하지 말고 ~~취소선~~ 처리.
4. 수정된 전체 JSON **하나만** 출력. 설명·코드펜스 금지.
"""


def _job_worker(job_id: str, kind: str, payload: dict):
    job = _jobs[job_id]
    try:
        if kind == "generate":
            job["step"] = "claude 생성 중 (1~3분)"
            raw = _run_claude(_gen_prompt(payload))
            data = _extract_json(raw)
            data.setdefault("meta", {})["title"] = data["meta"].get("title") or payload["title"]
            slug = _slugify(payload["title"])
            project_dir = PROJECTS / slug
            project_dir.mkdir(parents=True, exist_ok=True)
            job["step"] = "PPTX 렌더 중"
            pptx = _render(project_dir, data)
            job.update(status="done", project=slug, pptx=pptx.name)
        elif kind == "import":
            src = Path(payload["source_path"])
            job["step"] = "원본 텍스트 추출 중"
            source_text = _extract_source(src)
            job["step"] = "claude 변환 중 (2~5분)"
            # 대형 문서 변환은 모델이 정리 턴을 쓸 수 있어 여유 턴 허용 (1턴 제한 시 max-turns 에러 이력)
            raw = _run_claude(_import_prompt(payload, source_text), timeout=1800, max_turns=8)
            data = _extract_json(raw)
            data.setdefault("meta", {})["title"] = data["meta"].get("title") or payload["title"]
            project_dir = src.parent.parent  # <project>/source/<file> → <project>
            job["step"] = "PPTX 렌더 중"
            pptx = _render(project_dir, data)
            job.update(status="done", project=project_dir.name, pptx=pptx.name)
        elif kind == "revise":
            slug = payload["project"]
            project_dir = PROJECTS / slug
            current = json.loads((project_dir / "storyboard.json").read_text(encoding="utf-8"))
            job["step"] = "claude 수정 중 (1~3분)"
            raw = _run_claude(_revise_prompt(current, payload["instruction"]))
            data = _extract_json(raw)
            job["step"] = "PPTX 렌더 중"
            pptx = _render(project_dir, data)
            job.update(status="done", project=slug, pptx=pptx.name)
    except Exception as error:
        job.update(status="error", error=str(error)[-500:])


def start_job(kind: str, payload: dict) -> str:
    job_id = uuid.uuid4().hex[:12]
    _jobs[job_id] = {"status": "running", "step": "대기", "kind": kind,
                     "created": datetime.now().isoformat(timespec="seconds")}
    t = threading.Thread(target=_job_worker, args=(job_id, kind, payload), daemon=True)
    t.start()
    return job_id


def save_upload(filename: str, content: bytes, title: str) -> Path:
    """업로드 파일을 프로젝트 디렉터리에 저장하고 경로 반환."""
    slug = _slugify(title)
    if (PROJECTS / slug / "storyboard.json").exists():  # 기존 프로젝트 덮어쓰기 방지
        slug = f"{slug}-{uuid.uuid4().hex[:4]}"
    src_dir = PROJECTS / slug / "source"
    src_dir.mkdir(parents=True, exist_ok=True)
    filename = unicodedata.normalize("NFC", filename)  # macOS/브라우저 NFD 한글 → 정규화 (전부 _ 되는 문제 방지)
    safe_name = re.sub(r"[^0-9A-Za-z가-힣._ -]", "_", filename) or "upload"
    dest = src_dir / safe_name
    dest.write_bytes(content)
    return dest


def job_status(job_id: str) -> dict | None:
    return _jobs.get(job_id)


def list_projects() -> list[dict]:
    out = []
    if not PROJECTS.exists():
        return out
    for d in sorted(PROJECTS.iterdir()):
        sb = d / "storyboard.json"
        if not sb.exists():
            continue
        try:
            data = json.loads(sb.read_text(encoding="utf-8"))
        except Exception:
            continue
        meta = data.get("meta", {})
        pptxs = sorted(d.glob("*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)
        out.append({
            "id": d.name,
            "title": meta.get("title", d.name),
            "category": meta.get("category", ""),
            "version": meta.get("version", ""),
            "updated": meta.get("updated", ""),
            "pages": len(data.get("pages", [])) + 3,
            "history": data.get("history", []),
            "pptx": pptxs[0].name if pptxs else None,
            "mtime": sb.stat().st_mtime,
        })
    out.sort(key=lambda x: -x["mtime"])
    return out


def pptx_path(project_id: str) -> Path | None:
    d = PROJECTS / project_id
    if not d.exists() or not d.is_dir():
        return None
    # 경로 이탈 방지
    if not str(d.resolve()).startswith(str(PROJECTS.resolve())):
        return None
    pptxs = sorted(d.glob("*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return pptxs[0] if pptxs else None
