# Known Issues

## Vite command is not recognized

Symptom: `npm run build` or `npm run dev` prints `'vite' is not recognized as an internal or external command`.

Cause: `frontend\node_modules` is missing.

Fix: Run `npm install` in `frontend`, then retry. The restart script does this automatically when `node_modules` is absent.

Verification: `npm run build` completes and `http://localhost:3000` returns HTTP 200.

## Port 8000 is already occupied

Symptom: backend cannot bind to port `8000`, or another project is already listening there.

Cause: Docker `webword` uses `8000`; non-Docker local development uses a separate backend port.

Fix: Start this backend on `8010`, not `8000`.

Verification: `http://localhost:8010/api/health` returns `{"status":"ok"}`.

## AI templates stuck on “正在加载模板...”

Symptom: the assistant panel shows `正在加载模板...` forever, or the browser Network tab shows `/api/ai/templates` returning 404/`Not Found`.

Cause: Vite proxies `/api` to `127.0.0.1:8001`, but that address was occupied by another stack (`xianyu_backend` Docker on `127.0.0.1:8001`, plus an unrelated Python `app.py` on `0.0.0.0:8001`). On Windows, the more specific `127.0.0.1` bind wins for localhost, so the frontend never reaches this project's FastAPI. Template data itself was fine (`template_library.db` already had 15 rows).

Fix: Point the Vite proxy and local restart script at backend port `8010` (`frontend/vite.config.js`, `skills/project-local-run/scripts/restart-local-project.ps1`), then restart with the non-Docker script.

Verification: `Invoke-RestMethod http://127.0.0.1:8010/api/ai/templates` returns exactly 15 entries; opening the assistant no longer stays on `正在加载模板...`.

## Frontend import fails for AiWritingAssistant.vue

Symptom: Vite reports that `frontend/src/components/AiWritingAssistant.vue` cannot be resolved.

Cause: `App.vue` imports the component but the component file is missing.

Fix: Restore the component file or remove both the import and the `<AiWritingAssistant />` usage.

Verification: `npm run build` completes.

## PowerShell reports invalid variable reference near a colon

Symptom: PowerShell parser error says `Variable reference is not valid. ':' was not followed by a valid variable name character`.

Cause: a double-quoted string contains a variable immediately followed by `:`, such as `$Port:`.

Fix: wrap the variable name in braces, for example `${Port}:`.

Verification: the script parses and proceeds to port cleanup/startup.

## Start-Process rejects identical output and error log files

Symptom: `Start-Process` fails with `RedirectStandardOutput and RedirectStandardError are same`.

Cause: Windows PowerShell requires stdout and stderr to be redirected to different file paths.

Fix: use separate files such as `backend.out.log` and `backend.err.log`.

Verification: `Start-Process` launches both backend and frontend without that exception.

## Vue build reports missing end tag in AiWritingAssistant

Symptom: `npm run build` fails with `[vite:vue] Element is missing end tag` in `frontend/src/components/AiWritingAssistant.vue`.

Cause: a template edit left an SVG without its closing `</svg>` tag, or corrupted a text badge like `<span>参</span>` into malformed markup.

Fix: inspect the reported line and nearby button markup, restore balanced SVG/span tags, then rebuild.

Verification: `npm run build` completes.

## Restart script cannot write a log file because another process uses it

Symptom: PowerShell reports `The process cannot access the file ... because it is being used by another process` for `.codex-runlogs\*.log`.

Cause: the previous backend or frontend process still has the redirected log file open when the script tries to overwrite it.

Fix: stop port processes first, wait briefly, then initialize log files and start new processes.

Verification: restart script proceeds past log initialization and launches services.

## Uploaded Chinese filenames can become invalid temp paths

Symptom: AI generation returns HTTP 500 with `Invalid argument`, `No such file or directory`, or a temp path containing `??` / `%E2%91`.

Cause: embedding URL-encoded Chinese stems in `backend/temp` paths can exceed Windows path limits or break reads; writing uploads to `backend/temp` while uvicorn `--reload-dir backend` is enabled can also restart the worker mid-request.

Fix: save uploads as short ASCII `{uuid}{suffix}` files with a sidecar `{uuid}{suffix}.upload.json` for the original filename; exclude `temp/*` and `template_library.db` from uvicorn reload in `restart-local-project.ps1`.

Verification: uploading long Chinese filenames returns one generated document per source file with readable output filenames and no temp-path 500.

## Strict similarity control blocks without rewrite

Symptom: with strict reference isolation enabled, generation fails with `严格模式下检测到参考范文事实可能串入成文` or high source-copy overlap, and no usable document is returned.

Cause: the pipeline only raised on similarity risk and retried once with a generic parse-repair prompt that still exposed full reference text.

Fix: on similarity failures, automatically enter semantic rewrite mode up to `AI_MAX_GENERATION_GUARD_ATTEMPTS` (default 4): hide reference body, keep template skeleton + transformation plan, raise drafting temperature slightly, and instruct the model to paraphrase source facts without copying reference phrasing.

Verification: strict-mode generation on long case materials completes or reports `similarityRewrites > 0` in generation metadata instead of failing on the first overlap.

## Missing title aborts generation

Symptom: AI generation fails with `模型仿写缺少目标公文大标题`.

Cause: the drafting model sometimes omitted `[[title]]`, and the pipeline raised immediately instead of filling the missing required field.

Fix: detect missing required roles from the cached template structure JSON, then run permission-isolated targeted fill (rule-based title first, AI only for remaining missing fields) without rewriting already-valid paragraphs; finally render with backend format rules.

Verification: generating with materials that previously failed for missing title returns a document whose first role is `title`, and `generation.targetedFill.filledRoles` may include `title`.

## Docker frontend build reports npm audit vulnerabilities

Symptom: `docker compose up -d --build` completes, but the frontend build stage prints npm audit warnings such as `4 vulnerabilities`.

Cause: npm reports dependency advisory metadata during `npm ci`; it does not necessarily fail the production build.

Fix: Do not treat this as a deployment blocker when the Docker build exits successfully. Track dependency upgrades separately; avoid `npm audit fix --force` during deployment because it may introduce breaking package changes.

Verification: Docker build exits `0`, `webword` container starts, and `http://localhost:8000/api/health` returns `{"status":"ok"}`.

## Docker frontend build reports large chunk warning

Symptom: Vite prints `Some chunks are larger than 500 kB after minification` during Docker build.

Cause: the bundled frontend includes large Element Plus/Vue dependencies in the main chunk.

Fix: Treat as a performance warning, not a failed deployment. If performance becomes a requirement, split vendor chunks in `vite.config.js`.

Verification: Docker build exits `0`, `http://localhost:8000` returns HTTP 200.

## Docker AI template library is empty or duplicated

Symptom: `GET http://localhost:8000/api/ai/templates` returns `0` templates, or returns `30` entries with each of the 15 document types duplicated.

Cause: `0` means the root template source directory `20260623--整理汇总常用公文及范例（环食药侦）` was not copied into the Docker image. Duplicate entries mean a host-generated `backend/template_library.db` with Windows `source_dir` paths was copied into the image, then the container synced Linux paths on startup.

Fix: In `Dockerfile`, copy the root template source directory into `/app`. In `.dockerignore`, exclude `backend/template_library.db`. In `backend/ai_writer.py`, make `sync_template_library()` clear `document_templates` before inserting the current runtime scan.

Verification: rebuild with `docker compose up -d --build webword`, then `Invoke-RestMethod http://localhost:8000/api/ai/templates` returns exactly 15 entries ordered from `1决议` through `15纪要`.

## Local restart script stops Docker Desktop

Symptom: after `restart-local-project.ps1`, `docker info` fails with `open //./pipe/dockerDesktopLinuxEngine: The system cannot find the file specified`, and `com.docker.backend` / `wslrelay` are gone.

Cause: if Docker/WSL is temporarily listening on local backend port `8010`, the restart script's `Stop-PortProcess` force-kills that PID and can take down Docker Desktop.

Fix: restart Docker Desktop (`Start-Process "C:\Program Files\Docker\Docker\Docker Desktop.exe"`), wait until `docker info` succeeds, then redeploy with `deploy-local-docker.ps1`. Prefer verifying port owners before killing; do not stop `com.docker.backend` / `wslrelay` / `Docker Desktop` for local 8010 cleanup.

Verification: `docker ps` works, `webword` is healthy on `http://localhost:8000`, and local `http://127.0.0.1:8010/api/health` still returns ok.
