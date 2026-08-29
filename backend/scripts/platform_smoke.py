"""平台核心接口冒烟测试：健康检查、配置、模板、模型配置、上传与删除。"""
from __future__ import annotations

import io
import sys
import time
from pathlib import Path

from fastapi.testclient import TestClient

BACKEND_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BACKEND_DIR))

from app import app  # noqa: E402
from ai_writer import TEMPLATE_SOURCE_DIR  # noqa: E402

client = TestClient(app)
TEST_CATEGORY = "_cursor_platform_smoke"


def assert_ok(resp, label: str) -> dict:
    assert resp.status_code == 200, f"{label} failed: {resp.status_code} {resp.text}"
    payload = resp.json()
    assert payload.get("success", True), f"{label} payload not successful: {payload}"
    return payload


def test_health() -> None:
    payload = assert_ok(client.get("/api/health"), "health")
    assert payload["status"] == "ok"


def test_config() -> None:
    payload = assert_ok(client.get("/api/config"), "config")
    assert payload.get("templates")
    assert payload.get("labels")


def test_templates() -> None:
    started = time.time()
    payload = assert_ok(client.get("/api/ai/templates"), "templates")
    elapsed = time.time() - started
    assert elapsed < 5, f"templates too slow: {elapsed:.2f}s"
    assert len(payload.get("templates") or []) >= 1


def test_model_config() -> None:
    current = assert_ok(client.get("/api/ai/model-config"), "model-config get")
    saved = assert_ok(
        client.post(
            "/api/ai/model-config",
            json={
                "request_url": current.get("requestUrl") or "",
                "api_key": "",
                "model_name": current.get("modelName") or "",
                "models": current.get("models") or [],
            },
        ),
        "model-config save",
    )
    assert "modelName" in saved


def test_template_upload_delete() -> None:
    test_dir = TEMPLATE_SOURCE_DIR / TEST_CATEGORY
    if test_dir.exists():
        for child in test_dir.iterdir():
            if child.is_file():
                child.unlink()
    else:
        test_dir.mkdir(parents=True, exist_ok=True)

    sample_text = (
        "关于平台冒烟测试函件\n\n"
        "各县（市、区）公安局：\n\n"
        "为进一步规范公文写作，现将有关事项函告如下。\n\n"
        "一、加强组织领导。\n"
        "二、压实工作责任。\n\n"
        "某某市公安局\n"
        "2026年8月29日\n"
    )
    upload = client.post(
        "/api/ai/templates/upload",
        data={
            "category_mode": "new",
            "category_name": TEST_CATEGORY,
            "template_id": "letter",
        },
        files={
            "files": ("冒烟测试函件.txt", io.BytesIO(sample_text.encode("utf-8")), "text/plain"),
        },
    )
    assert upload.status_code == 200, upload.text
    upload_payload = upload.json()
    assert upload_payload["count"] == 1

    templates = upload_payload.get("templates") or []
    matched = next(
        (item for item in templates if item.get("name") == TEST_CATEGORY),
        None,
    )
    assert matched, "upload category missing from template list"
    source_dir = matched["sourceDir"]
    filename = matched["files"][0]

    started = time.time()
    delete_file = client.delete(
        "/api/ai/templates/file",
        params={"source_dir": source_dir, "filename": filename},
    )
    delete_elapsed = time.time() - started
    assert delete_file.status_code == 200, delete_file.text
    assert delete_elapsed < 5, f"delete file too slow: {delete_elapsed:.2f}s"

    upload2 = client.post(
        "/api/ai/templates/upload",
        data={
            "category_mode": "existing",
            "source_dir": source_dir,
        },
        files={
            "files": ("冒烟测试函件2.txt", io.BytesIO(sample_text.encode("utf-8")), "text/plain"),
        },
    )
    assert upload2.status_code == 200, upload2.text

    started = time.time()
    delete_category = client.delete(
        "/api/ai/templates/category",
        params={"source_dir": source_dir},
    )
    delete_category_elapsed = time.time() - started
    assert delete_category.status_code == 200, delete_category.text
    assert delete_category_elapsed < 3, f"delete category too slow: {delete_category_elapsed:.2f}s"


def main() -> None:
    test_health()
    test_config()
    test_templates()
    test_model_config()
    test_template_upload_delete()
    print("platform smoke test passed")


if __name__ == "__main__":
    main()
