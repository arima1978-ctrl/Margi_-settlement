"""FastAPI web UI for margin-settlement.

Design
------
Single-page form that drives ``scripts/run_monthly.py`` as a subprocess and
streams its output back to the browser via periodic polling. Jobs are kept
in-memory; this is fine because only 2 people use the tool on a single
instance. A server restart wipes job history.

Auth is HTTP Basic, configured via env ``WEB_BASIC_AUTH_USERS`` as a
comma-separated ``user:password`` list. Credentials are loaded from ``.env``
at import time so systemd/cron contexts pick them up.

Security note: ``/download`` serves files from under ``MARGIN_BASE_DIR``
only, and resolves the path before checking the prefix to prevent
``..`` traversal.
"""
from __future__ import annotations

import os
import secrets
import subprocess
import sys
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from fastapi import Depends, FastAPI, HTTPException, Request, status
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

# Repo root on sys.path for the src/* imports below
REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from src.notifier import load_dotenv  # noqa: E402

load_dotenv(REPO_ROOT / ".env", override=True)

BASE_DIR = Path(os.environ.get("MARGIN_BASE_DIR")
                or r"Y:\_★20170701作業用\【エデュプラス請求書】")

SERVICE_FOLDERS = {
    "programming": "プログラミング清算書",
    "shogi":       "スマイル将棋清算書",
    "bunri":       "文理ヴィクトリー清算書",
    "sokudoku":    "１００万人の速読　清算書",
}

# eduplus は在来4サービスと違い、源泉 .xlsm を in-place で書き換えるだけで
# ダウンロード可能な新ファイルは生成しない。SERVICE_CHOICES として扱うが
# ファイル一覧 (SERVICE_FOLDERS) からは除外する。
SERVICE_CHOICES: list[str] = list(SERVICE_FOLDERS.keys()) + ["eduplus"]

app = FastAPI(title="margin-settlement web UI")
templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))

security = HTTPBasic()


def _parse_users() -> Dict[str, str]:
    raw = os.environ.get("WEB_BASIC_AUTH_USERS", "")
    users: Dict[str, str] = {}
    for pair in raw.split(","):
        pair = pair.strip()
        if not pair or ":" not in pair:
            continue
        user, _, password = pair.partition(":")
        users[user.strip()] = password.strip()
    return users


USERS = _parse_users()


def authenticate(credentials: HTTPBasicCredentials = Depends(security)) -> str:
    if not USERS:
        # No users configured → fail closed
        raise HTTPException(status_code=500, detail="Basic auth not configured")
    expected = USERS.get(credentials.username)
    if expected is None or not secrets.compare_digest(expected, credentials.password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid credentials",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username


# ---- Job registry --------------------------------------------------------

class Job:
    __slots__ = ("id", "month", "services", "status", "log", "started", "finished", "returncode")

    def __init__(self, job_id: str, month: str, services: List[str]) -> None:
        self.id = job_id
        self.month = month
        self.services = services
        self.status = "queued"  # queued → running → done | failed
        self.log: List[str] = []
        self.started: Optional[datetime] = None
        self.finished: Optional[datetime] = None
        self.returncode: Optional[int] = None

    def as_dict(self) -> dict:
        return {
            "id": self.id,
            "month": self.month,
            "services": self.services,
            "status": self.status,
            "log_count": len(self.log),
            "started": self.started.isoformat() if self.started else None,
            "finished": self.finished.isoformat() if self.finished else None,
            "returncode": self.returncode,
        }


JOBS: Dict[str, Job] = {}
JOBS_LOCK = threading.Lock()


def run_job(job: Job, notify: bool) -> None:
    """Subprocess runner invoked from a daemon thread."""
    job.status = "running"
    job.started = datetime.now()
    # -u: unbuffered stdout so the UI receives log lines as they happen,
    # not only when the subprocess exits.
    cmd = [sys.executable, "-u", str(REPO_ROOT / "scripts" / "run_monthly.py"),
           "--month", job.month, "--yes"]
    if job.services:
        cmd.extend(["--only", *job.services])
    if notify:
        cmd.append("--notify")
    job.log.append(f"$ {' '.join(cmd)}\n")

    # Child inherits PYTHONIOENCODING=utf-8 so print() of Japanese works on
    # Windows cp932 consoles and inside systemd.
    env = os.environ.copy()
    env.setdefault("PYTHONIOENCODING", "utf-8")
    env.setdefault("PYTHONUNBUFFERED", "1")

    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            cwd=str(REPO_ROOT),
            env=env,
        )
        assert proc.stdout is not None
        for line in proc.stdout:
            job.log.append(line)
        proc.wait()
        job.returncode = proc.returncode
        job.status = "done" if proc.returncode == 0 else "failed"
    except Exception as exc:
        job.log.append(f"\n[web-ui error] {type(exc).__name__}: {exc}\n")
        job.status = "failed"
    finally:
        job.finished = datetime.now()


# ---- Routes --------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def index(request: Request, user: str = Depends(authenticate)):
    # Most recent files per service for the file list.
    # Starlette >= 0.47 requires request as first positional arg; passing it
    # via context dict is deprecated.
    files = list_recent_files(limit_per_service=5)
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "user": user,
            "services": SERVICE_CHOICES,
            "base_dir": str(BASE_DIR),
            "files": files,
            "default_month": datetime.now().strftime("%Y-%m"),
        },
    )


@app.post("/run")
async def start_run(
    request: Request,
    user: str = Depends(authenticate),
) -> JSONResponse:
    form = await request.form()
    month = str(form.get("month", "")).strip()
    services = [s for s in form.getlist("services") if s in SERVICE_CHOICES]
    notify = form.get("notify") in ("on", "true", "1")
    if not month or len(month) != 7 or month[4] != "-":
        raise HTTPException(status_code=400, detail="month must be YYYY-MM")

    job_id = uuid.uuid4().hex[:12]
    job = Job(job_id, month, services)
    with JOBS_LOCK:
        JOBS[job_id] = job
    threading.Thread(target=run_job, args=(job, notify), daemon=True).start()
    return JSONResponse({"job_id": job_id})


@app.get("/status/{job_id}")
def job_status(job_id: str, offset: int = 0, user: str = Depends(authenticate)) -> JSONResponse:
    job = JOBS.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="job not found")
    new_lines = job.log[offset:]
    return JSONResponse({
        "job": job.as_dict(),
        "new_lines": new_lines,
        "next_offset": offset + len(new_lines),
    })


@app.get("/files")
def list_files(user: str = Depends(authenticate)) -> JSONResponse:
    return JSONResponse({"files": list_recent_files(limit_per_service=20)})


def list_recent_files(limit_per_service: int = 5) -> List[dict]:
    out: List[dict] = []
    for svc, folder in SERVICE_FOLDERS.items():
        svc_dir = BASE_DIR / folder
        if not svc_dir.exists():
            continue
        files = sorted(svc_dir.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        for f in files[:limit_per_service]:
            out.append({
                "service": svc,
                "name": f.name,
                "path": str(f),
                "size": f.stat().st_size,
                "mtime": datetime.fromtimestamp(f.stat().st_mtime).isoformat(timespec="seconds"),
            })
    return out


@app.get("/download")
def download(path: str, user: str = Depends(authenticate)):
    target = Path(path).resolve()
    base_resolved = BASE_DIR.resolve()
    # Security: only serve files under BASE_DIR
    try:
        target.relative_to(base_resolved)
    except ValueError:
        raise HTTPException(status_code=403, detail="path outside BASE_DIR")
    if not target.is_file():
        raise HTTPException(status_code=404, detail="file not found")
    return FileResponse(target, filename=target.name)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("WEB_PORT", "8081")))
