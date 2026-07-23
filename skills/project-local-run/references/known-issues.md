# Known Issues

## Vite command is not recognized

Symptom: `npm run build` or `npm run dev` prints `'vite' is not recognized as an internal or external command`.

Cause: `frontend\node_modules` is missing.

Fix: Run `npm install` in `frontend`, then retry. The restart script does this automatically when `node_modules` is absent.

Verification: `npm run build` completes and `http://localhost:3000` returns HTTP 200.

## Port 8000 is already occupied

Symptom: backend cannot bind to port `8000`, or another project is already listening there.

Cause: this workspace's Vite proxy targets `localhost:8001`, while another local service may use `8000`.

Fix: Start this backend on `8001`, not `8000`.

Verification: `http://localhost:8001/api/health` returns `{"status":"ok"}`.

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

Symptom: AI generation returns HTTP 500 with `Invalid argument` and a temp path containing `??`.

Cause: the current Windows/Python upload path can lose non-ASCII filename characters when the filename is used directly in the temp file path.

Fix: save upload files using an ASCII-safe encoded temp filename, and decode the original stem only for display/download names.

Verification: uploading multiple files returns one generated document per source file with readable output filenames.

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
