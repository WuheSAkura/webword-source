FROM node:20-alpine AS frontend-build

WORKDIR /app/frontend
COPY frontend/package*.json ./
RUN npm ci
COPY frontend/ ./
RUN npm run build


# 用 Debian 11 (bullseye) 旧版 glibc：创建线程走老的 clone()，
# 避免新 glibc 的 clone3() 在旧 seccomp/旧 Docker 宿主机上被拦 → "can't start new thread"
FROM python:3.12-slim-bullseye

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

COPY backend/requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY backend/ ./backend/
COPY ["20260623--整理汇总常用公文及范例（环食药侦）", "./20260623--整理汇总常用公文及范例（环食药侦）"]
COPY --from=frontend-build /app/frontend/dist ./backend/static

WORKDIR /app/backend

EXPOSE 8000

# 纯 asyncio 事件循环 + h11，避免 uvloop/httptools 在部分 arm64 内核上 SIGTRAP 崩溃
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000", "--loop", "asyncio", "--http", "h11"]
