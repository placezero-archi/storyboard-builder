from __future__ import annotations

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from app.services import sb_builder_service as svc

router = APIRouter(prefix="/api/sb", tags=["sb-builder"])


class GenerateRequest(BaseModel):
    title: str = Field(min_length=2, max_length=80)
    category: str = Field(default="", max_length=40)
    platform: str = Field(default="PC/MO", max_length=30)
    langs: str = Field(default="ko", max_length=60)
    admin: bool = False
    requirements: str = Field(min_length=10, max_length=8000)


class ReviseRequest(BaseModel):
    instruction: str = Field(min_length=5, max_length=4000)


@router.post("/generate")
def generate(req: GenerateRequest):
    job_id = svc.start_job("generate", req.model_dump())
    return {"job_id": job_id}


MAX_UPLOAD = 30 * 1024 * 1024  # 30MB
ALLOWED_EXT = {".pptx", ".pdf"}


@router.post("/upload")
async def upload(
    file: UploadFile = File(...),
    title: str = Form(min_length=2, max_length=80),
    category: str = Form(default="", max_length=40),
):
    ext = ("." + file.filename.rsplit(".", 1)[-1].lower()) if "." in (file.filename or "") else ""
    if ext not in ALLOWED_EXT:
        raise HTTPException(400, "PPTX 또는 PDF 파일만 업로드할 수 있습니다.")
    content = await file.read()
    if len(content) > MAX_UPLOAD:
        raise HTTPException(400, "파일이 너무 큽니다 (최대 30MB).")
    if not content:
        raise HTTPException(400, "빈 파일입니다.")
    src = svc.save_upload(file.filename, content, title)
    job_id = svc.start_job("import", {
        "source_path": str(src), "title": title, "category": category,
    })
    return {"job_id": job_id}


@router.post("/projects/{project_id}/revise")
def revise(project_id: str, req: ReviseRequest):
    if svc.pptx_path(project_id) is None and not (svc.PROJECTS / project_id / "storyboard.json").exists():
        raise HTTPException(404, "프로젝트를 찾을 수 없습니다.")
    job_id = svc.start_job("revise", {"project": project_id, "instruction": req.instruction})
    return {"job_id": job_id}


@router.get("/jobs/{job_id}")
def job(job_id: str):
    st = svc.job_status(job_id)
    if st is None:
        raise HTTPException(404, "작업을 찾을 수 없습니다.")
    return st


@router.get("/projects")
def projects():
    return {"items": svc.list_projects()}


@router.get("/projects/{project_id}/download")
def download(project_id: str):
    p = svc.pptx_path(project_id)
    if p is None:
        raise HTTPException(404, "산출물이 없습니다.")
    return FileResponse(
        str(p),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=p.name,
    )
