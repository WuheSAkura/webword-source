"""模板上传与索引冒烟测试。"""
from __future__ import annotations

import io
import sys
from pathlib import Path

from fastapi.testclient import TestClient

BACKEND_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BACKEND_DIR))

from app import app  # noqa: E402
from ai_writer import scan_template_library, TEMPLATE_SOURCE_DIR  # noqa: E402

client = TestClient(app)


def main() -> None:
    catalog = client.get("/api/ai/templates/catalog")
    assert catalog.status_code == 200, catalog.text
    assert catalog.json()["success"]
    assert any(item["id"] == "letter" for item in catalog.json()["catalog"])

    templates_before = client.get("/api/ai/templates")
    assert templates_before.status_code == 200
    before_count = len(templates_before.json().get("templates") or [])

    category_name = "_cursor_template_upload_test"
    test_dir = TEMPLATE_SOURCE_DIR / category_name
    if test_dir.exists():
        for child in test_dir.iterdir():
            if child.is_file():
                child.unlink()
    else:
        test_dir.mkdir(parents=True, exist_ok=True)

    sample_text = (
        "关于测试函件范例\n\n"
        "各县（市、区）公安局：\n\n"
        "为进一步规范公文写作，现将有关事项函告如下。\n\n"
        "一、加强组织领导。\n"
        "二、压实工作责任。\n\n"
        "某某市公安局\n"
        "2026年8月28日\n"
    )
    files = {
        "files": ("测试函件范例.txt", io.BytesIO(sample_text.encode("utf-8")), "text/plain"),
    }
    data = {
        "category_mode": "new",
        "category_name": category_name,
        "template_id": "letter",
    }
    upload = client.post("/api/ai/templates/upload", data=data, files=files)
    assert upload.status_code == 200, upload.text
    payload = upload.json()
    assert payload["success"]
    assert payload["count"] == 1
    assert payload["assignments"][0]["templateId"] == "letter"

    templates_after = payload.get("templates") or []
    matched = [
        item for item in templates_after
        if item.get("sourceDir") == str(test_dir) or item.get("name") == category_name
    ]
    assert matched, "上传后模板库中未找到新建分类"
    assert matched[0].get("availableReferenceCount", 0) >= 1, "范文未成功解析为可用体例样本"

    # 归入已有分类再传一份
    files2 = {
        "files": ("测试函件范例2.txt", io.BytesIO((sample_text + "\n").encode("utf-8")), "text/plain"),
    }
    data2 = {
        "category_mode": "existing",
        "source_dir": str(test_dir),
    }
    upload2 = client.post("/api/ai/templates/upload", data=data2, files=files2)
    assert upload2.status_code == 200, upload2.text
    assert upload2.json()["count"] == 1

    after_count = len(upload2.json().get("templates") or templates_after)
    assert after_count >= before_count

    print("template upload smoke test passed")
    print(f"category={category_name} references={matched[0].get('availableReferenceCount')}")


if __name__ == "__main__":
    main()
