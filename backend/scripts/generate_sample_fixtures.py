"""Generate sample fixtures consumed by tests/test_document_parsers.py.

Usage: uv run python scripts/generate_sample_fixtures.py
Output: <repo-root>/sample-data/generated/{docx,csv,json,html,markdown,pptx,xlsx}
"""

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
GENERATED = ROOT / "sample-data" / "generated"

DOCX_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>候选人编号：C001</w:t></w:r></w:p>
    <w:p><w:r><w:t>姓名：李明远</w:t></w:r></w:p>
    <w:p><w:r><w:t>应聘岗位：AI 应用架构师</w:t></w:r></w:p>
    <w:p><w:r><w:t>述职摘要：主导人才评估系统落地，沉淀多 Agent 解析流水线。</w:t></w:r></w:p>
  </w:body>
</w:document>
"""

DOCX_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

DOCX_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head><meta charset="utf-8"><title>内部岗位申请记录</title></head>
<body>
<main>
  <h1>内部岗位申请记录</h1>
  <p>候选人编号 C001 于 2026-07-02 申请 AI 应用架构师岗位。</p>
</main>
</body>
</html>
"""

MARKDOWN = """# 项目复盘

## 背景
人才评估系统需要在两周内完成多格式档案接入。

## 结论
统一解析层上线后，材料入库耗时下降 60%。
"""


def write_docx(path: Path) -> None:
    import zipfile

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", DOCX_CONTENT_TYPES)
        archive.writestr("_rels/.rels", DOCX_RELS)
        archive.writestr("word/document.xml", DOCX_XML)


def write_csv(path: Path) -> None:
    path.write_text(
        "candidate_id,name,department,level\n"
        "C001,李明远,平台研发部,P7\n"
        "C002,王小雨,数据智能部,P6\n",
        encoding="utf-8",
    )


def write_json(path: Path) -> None:
    payload = {
        "candidate_id": "C001",
        "profile": {"name": "李明远", "target_position": "AI 应用架构师"},
        "assessments": [{"round": "技术面", "score": 4.5}],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_html(path: Path) -> None:
    path.write_text(HTML, encoding="utf-8")


def write_markdown(path: Path) -> None:
    path.write_text(MARKDOWN, encoding="utf-8")


def write_pptx(path: Path) -> None:
    from pptx import Presentation
    from pptx.util import Inches

    presentation = Presentation()
    layout = presentation.slide_layouts[5]
    for title, body in [
        ("人才述职与岗位申请", "候选人：李明远 C001"),
        ("岗位匹配度", "AI 应用架构师 · 匹配度 92%"),
        ("关键项目", "人才评估系统多 Agent 解析流水线"),
        ("下一步计划", "完成人才画像与推荐链路联调"),
    ]:
        slide = presentation.slides.add_slide(layout)
        textbox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
        textbox.text_frame.text = f"{title}\n{body}"
    presentation.save(path)


def write_xlsx(path: Path) -> None:
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "员工花名册"
    sheet.append(["candidate_id", "name", "department", "level"])
    sheet.append(["C001", "李明远", "平台研发部", "P7"])
    sheet.append(["C002", "王小雨", "数据智能部", "P6"])
    extra = workbook.create_sheet("绩效记录")
    extra.append(["candidate_id", "period", "score"])
    extra.append(["C001", "2026H1", 4.6])
    workbook.save(path)


def main() -> None:
    writers = [
        ("docx", "C001_述职材料.docx", write_docx),
        ("csv", "C001_hris花名册.csv", write_csv),
        ("json", "C001_人才画像.json", write_json),
        ("html", "C001_内部岗位申请记录.html", write_html),
        ("markdown", "C001_项目复盘.md", write_markdown),
        ("pptx", "C001_人才述职与岗位申请.pptx", write_pptx),
        ("xlsx", "C001_员工花名册.xlsx", write_xlsx),
    ]
    for folder, name, writer in writers:
        target = GENERATED / folder
        target.mkdir(parents=True, exist_ok=True)
        writer(target / name)
        print(f"generated {target / name}")


if __name__ == "__main__":
    main()