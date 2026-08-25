---
name: project-local-run
description: Start, restart, and diagnose this gongwen local development project, including the FastAPI backend on port 8010 and Vite/Vue frontend on port 3000. Use when the user asks to start or restart the local frontend/backend, fix localhost access failures, inspect port conflicts, read run logs, or record recurring errors and solutions for this project.
---

# Project Local Run

## Overview

Use this skill to bring up the local `gongwen` stack and keep a reusable issue catalog. The project root is `C:\Users\Auraa\Desktop\gongwen\webword-source`; the backend is FastAPI and the frontend is Vue 3 + Vite.

## Quick Start

For Docker deployment from the project root:

```powershell
powershell -ExecutionPolicy Bypass -File .\skills\project-local-run\scripts\deploy-local-docker.ps1
```

The script runs `docker compose up -d --build webword`, verifies `http://localhost:8000`, `GET /api/health`, the 15-entry AI template library, and tails recent container logs.

Verify:

```powershell
Invoke-WebRequest -UseBasicParsing http://localhost:8000/api/health
Invoke-WebRequest -UseBasicParsing http://localhost:8000
Invoke-RestMethod http://localhost:8000/api/ai/templates
docker logs --tail 80 webword
```

For non-Docker local development from the project root, run:

From the project root, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\skills\project-local-run\scripts\restart-local-project.ps1
```

This script:

- Stops processes listening on ports `8010` and `3000`.
- Ensures frontend dependencies exist.
- Starts backend with `python -m uvicorn app:app --host 0.0.0.0 --port 8010 --reload --reload-dir backend`.
- Starts frontend with `frontend\node_modules\.bin\vite.cmd --host 0.0.0.0 --port 3000 --strictPort`.
- Writes logs to `.codex-runlogs\backend.out.log`, `.codex-runlogs\backend.err.log`, `.codex-runlogs\frontend.out.log`, and `.codex-runlogs\frontend.err.log`.
- Checks `http://localhost:8010/api/health` and `http://localhost:3000`.

## Workflow

1. Inspect current listeners before restart:

```powershell
Get-NetTCPConnection -LocalPort 8010,3000 -ErrorAction SilentlyContinue |
  Select-Object LocalPort,State,OwningProcess
```

2. Run the restart script from the project root.

3. If the script reports a failure, inspect the latest log tail:

```powershell
Get-Content .\.codex-runlogs\backend.err.log -Tail 80
Get-Content .\.codex-runlogs\frontend.err.log -Tail 80
```

4. Compare the error with `references\known-issues.md`. If it is new, fix it and append a concise entry with symptom, cause, fix, and verification.

## Docker Incremental Deployment

Use this flow when the user asks to deploy the current local project to local Docker:

1. Confirm Docker is reachable before debugging the app:

```powershell
docker info
docker ps --filter "name=webword" --format "table {{.Names}}\t{{.Image}}\t{{.Ports}}\t{{.Status}}"
```

2. Ensure Docker build includes everything required at runtime:

- `Dockerfile` copies `backend/`, the frontend `dist`, and the root template source directory `20260623--整理汇总常用公文及范例（环食药侦）`.
- `.dockerignore` excludes `backend/template_library.db`; this SQLite file is a runtime-derived template index and must not carry host absolute paths into the image.
- `backend/ai_writer.py::sync_template_library()` rebuilds `document_templates` from the current runtime template directory.

3. Deploy incrementally:

```powershell
powershell -ExecutionPolicy Bypass -File .\skills\project-local-run\scripts\deploy-local-docker.ps1
```

4. Required proof before reporting success:

```powershell
Invoke-WebRequest -UseBasicParsing http://localhost:8000/api/health
Invoke-WebRequest -UseBasicParsing http://localhost:8000
Invoke-RestMethod http://localhost:8000/api/ai/templates
docker logs --tail 80 webword
```

The template API must return exactly 15 entries ordered from `1决议` through `15纪要`. If it returns 0, the template source directory was not copied into the image. If it returns 30 or duplicate entries, stale `backend/template_library.db` paths were copied or the sync routine failed to clear old rows.

## Project Details

- Backend working directory: `backend`
- Backend port: `8010` (non-Docker local)
- Backend health check: `GET /api/health`
- Frontend working directory: `frontend`
- Frontend port: `3000`
- Vite proxy target for `/api`: `http://127.0.0.1:8010`
- Frontend dependency command: `npm install`
- Frontend build check: `npm run build`
- Docker compose file: `docker-compose.yml`
- Docker container name: `webword`
- Docker app URL: `http://localhost:8000`
- Docker build/redeploy command: `powershell -ExecutionPolicy Bypass -File .\skills\project-local-run\scripts\deploy-local-docker.ps1`

## Rules

- Do not kill unrelated ports unless the user explicitly asks.
- Do not use `git reset --hard` or revert user edits while diagnosing startup problems.
- Prefer direct Vite executable startup on Windows because `npm run dev` may hang without exposing useful output in some shells.
- Keep new troubleshooting knowledge in `references\known-issues.md`.
- Treat Docker port `8000` and non-Docker backend port `8010` as separate deployment modes. Do not reuse `8001` locally; other stacks (e.g. `xianyu_backend`) commonly bind `127.0.0.1:8001`.
- For Docker success, do not stop at container "Up"; verify `/api/health`, homepage, template count, and recent logs.
