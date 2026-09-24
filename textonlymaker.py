import os
import glob
import zipfile
from lxml import etree

# ── 경로 설정 ──────────────────────────────────────────────
# input 폴더에 그림 제거할 docx(예: HWP 변환 결과물)를 넣어두세요.
INPUT_FOLDER = r"C:\KHK\DocuWork\input"
OUTPUT_FOLDER = r"C:\KHK\DocuWork\output"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ── 네임스페이스 ───────────────────────────────────────────
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 문서 본문에서 제거할 "그림/그래픽" 요소들
DRAWING_TAGS = {
    f"{{{W_NS}}}drawing",  # 최신 DrawingML (인라인/떠 있는 그림)
    f"{{{W_NS}}}pict",     # 구형 VML 그림
    f"{{{W_NS}}}object",   # OLE 개체(대개 이미지)
}


def strip_xml_images(xml_bytes: bytes) -> bytes:
    """스토리 XML(document/header/footer 등)에서 그림 요소를 제거."""
    root = etree.fromstring(xml_bytes)
    removed = 0
    for tag in DRAWING_TAGS:
        for el in root.iter(tag):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
                removed += 1
    if removed == 0:
        return xml_bytes
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def strip_rels_images(xml_bytes: bytes) -> bytes:
    """.rels에서 이미지 관계(깨진 참조 방지)를 제거."""
    root = etree.fromstring(xml_bytes)
    removed = 0
    for rel in list(root):
        rtype = rel.get("Type", "")
        target = rel.get("Target", "")
        if rtype.endswith("/image") or "media/" in target.replace("\\", "/"):
            root.remove(rel)
            removed += 1
    if removed == 0:
        return xml_bytes
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def process_docx(src: str, dst: str) -> int:
    """그림 제거 후 dst로 저장. 제거된 이미지 파일 개수를 반환."""
    removed_media = 0
    with zipfile.ZipFile(src, "r") as zin:
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                name = info.filename
                if name.endswith("/"):          # 디렉터리 엔트리 skip
                    continue
                if name.startswith("word/media/"):  # 실제 이미지 파일 통째 제외
                    removed_media += 1
                    continue

                data = zin.read(name)
                if name.endswith(".rels"):
                    data = strip_rels_images(data)
                elif name.endswith(".xml") and (
                    "document.xml" in name or "/header" in name or "/footer" in name
                    or "footnotes.xml" in name or "endnotes.xml" in name or "comments" in name
                ):
                    data = strip_xml_images(data)

                zout.writestr(info, data)
    return removed_media


def human(n: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.1f}{unit}"
        n /= 1024
    return f"{n:.1f}TB"


# ── 실행 ───────────────────────────────────────────────────
docx_files = glob.glob(os.path.join(INPUT_FOLDER, "*.docx"))

print("=" * 62)
print(f"[INPUT] {INPUT_FOLDER}  — 대상 {len(docx_files)}개")
for f in docx_files:
    print(f"   • {os.path.basename(f):40s} {human(os.path.getsize(f)):>10s}")
print("=" * 62)

success, fail = 0, 0
for i, src in enumerate(docx_files, 1):
    filename = os.path.splitext(os.path.basename(src))[0]
    dst = os.path.join(OUTPUT_FOLDER, f"{filename}_textonly.docx")
    try:
        removed = process_docx(src, dst)
        before, after = os.path.getsize(src), os.path.getsize(dst)
        saved = (1 - after / before) * 100 if before else 0
        print(f"[{i}/{len(docx_files)}] OK  {os.path.basename(src)} "
              f"→ 이미지 {removed}개 제거, {human(before)} → {human(after)} ({saved:.0f}% ↓)")
        success += 1
    except Exception as e:
        print(f"[{i}/{len(docx_files)}] FAIL {os.path.basename(src)} : {e}")
        fail += 1

out_files = glob.glob(os.path.join(OUTPUT_FOLDER, "*.docx"))
print("=" * 62)
print(f"[OUTPUT] {OUTPUT_FOLDER}  — 완료 {len(out_files)}개")
for f in out_files:
    print(f"   • {os.path.basename(f):40s} {human(os.path.getsize(f)):>10s}")
print("=" * 62)
print(f"완료! 성공 {success}개 / 실패 {fail}개")