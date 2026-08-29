"""近成稿材料照录风险验证。"""
from __future__ import annotations

import json
import time
import urllib.request
from pathlib import Path

from docx import Document

BASE = "http://127.0.0.1:8010"
OUT = Path(__file__).resolve().parent.parent / "temp" / "e2e_materials" / "copy_risk_case.docx"


def main() -> int:
    doc = Document()
    for text in [
        "关于商请协助开展专项检查工作的函",
        "市发展改革委：",
        "根据市政府工作安排，拟于2026年3月在全市范围内开展专项检查。现将有关事项函告如下：",
        "一、检查重点。一是核查台账完整性；二是现场抽查落实情况；三是汇总问题并反馈整改。",
        "二、工作要求。请你单位予以配合，组织相关处室做好准备，并于3月底前报送书面情况。",
        "三、联系方式。联系人：张某，电话：010-12345678。",
        "特此函达。",
        "某某市专项检查办公室",
        "2026年2月20日",
    ]:
        doc.add_paragraph(text)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUT)

    boundary = "----WebKitFormBoundaryCopy"
    body = bytearray()
    fields = {
        "prompt": "请据此创作一份正式函，要求公文语体独立成文，不得照录材料原文。",
        "template_id": "letter",
        "speed_mode": "fast",
        "temperature": "0.4",
        "strict_reference_isolation": "false",
        "allow_degradation": "false",
    }
    for key, value in fields.items():
        body.extend(f"--{boundary}\r\n".encode())
        body.extend(f'Content-Disposition: form-data; name="{key}"\r\n\r\n'.encode())
        body.extend(f"{value}\r\n".encode("utf-8"))
    data = OUT.read_bytes()
    body.extend(f"--{boundary}\r\n".encode())
    body.extend(
        f'Content-Disposition: form-data; name="material_files"; filename="{OUT.name}"\r\n'.encode("utf-8")
    )
    body.extend(b"Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n")
    body.extend(data)
    body.extend(b"\r\n")
    body.extend(f"--{boundary}--\r\n".encode())
    req = urllib.request.Request(
        f"{BASE}/api/ai/generate",
        data=bytes(body),
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        method="POST",
    )
    started = time.time()
    with urllib.request.urlopen(req, timeout=600) as resp:
        payload = json.loads(resp.read().decode("utf-8"))
    elapsed = round(time.time() - started, 1)
    docs = payload.get("documents") or ([payload] if payload.get("fileId") else [])
    doc_payload = docs[0] if docs else {}
    gen = doc_payload.get("generation") or {}
    val = gen.get("validation") or {}
    result = {
        "ok": bool(payload.get("success")),
        "elapsedSec": elapsed,
        "plannerFinish": (gen.get("planner") or {}).get("finishReason"),
        "plannerReason": (gen.get("planner") or {}).get("fallbackReason"),
        "copiedParagraphRatio": val.get("copiedParagraphRatio"),
        "sourceCopyCoverage": val.get("sourceCopyCoverage"),
        "similarityRewrites": gen.get("similarityRewrites"),
        "rewriteRetries": gen.get("rewriteRetries"),
        "preview": (doc_payload.get("preview") or "")[:600],
        "warnings": (doc_payload.get("warnings") or [])[:8],
        "error": None if payload.get("success") else (payload.get("error") or payload.get("detail")),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
