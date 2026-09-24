# webCollect_v4.py  (개선본 4 - '멈춤(hang)' 제거 + URL 직접이동 + 팝업 전체 스크롤)
#
# ▣ v3 실행 로그에서 드러난 두 가지 실제 문제
#   [문제1] 어느 순간 그냥 멈춤(hang)
#     원인: 클릭스루 모드에서 카드를 '리스트 바닥까지' 수집한 뒤, 다시 위쪽 카드를
#           data-id 로 하나씩 찾으려다 가상스크롤(화면 밖 카드)에 걸려 _scroll_find_card 가
#           사실상 무한 탐색 -> 정지.
#   [문제2] 팝업(상세뷰)에서 아래로 스크롤이 안 됨
#     원인: 스크롤 대상을 '본문 요소의 조상' 하나로만 찾았는데, 팝업/모달의 실제
#           스크롤 컨테이너가 그와 달라 안 잡힘.
#
# ▣ v4 해결
#   (1) URL 직접이동 활성화 -> 클릭스루/리스트 재탐색 자체를 제거 -> '멈춤' 소멸.
#       v3 로그: URL 이 ...&memoId=540641 로 바뀜(=URL 로 상세 접근 가능)인데도
#       "URL 템플릿 없음"이 떴던 건, 학습 때 리스트가 바닥이라 ids[0] 카드를 못 찾고
#       '엉뚱한 카드'를 클릭한 뒤 그 memoId 를 원래 ids[0] 와 비교해 불일치 처리했기 때문.
#       v4: 학습 전 리스트를 맨 위로 올리고, '실제로 클릭한 카드의 data-id' 를 확보해
#           URL 의 memoId 와 비교 -> data-id==memoId 면 URL 템플릿을 만들어 전 메모를
#           page.goto(URL) 로 접근(리스트 상호작용 0회 -> 멈춤 없음).
#           폴더마다 folderId 를 다시 읽어 URL 에 반영.
#   (2) 팝업 전체 스크롤 -> 페이지 내 '스크롤 가능한 모든 컨테이너'를 바닥까지, 높이가
#       안정될 때까지 반복 스크롤(모달 스크롤러를 특정하지 않아도 확실히 잡힘) + lazy 이미지 승격.
#   (3) 안전장치 -> _scroll_find_card 최대 반복을 대폭 축소(무한 탐색 방지). 클릭스루로
#       빠지더라도 못 찾으면 '건너뛰기'로 처리하고 절대 멈추지 않음.
#
#   ※ 서식 파서(밑줄/형광펜)는 v3에서 검증 완료되어 그대로 유지.
#
# ⚠️ 만약 실행 후에도 "URL 직접이동: OFF" 로 뜨면(=카드 data-id 가 memoId 와 다른 계정 구조),
#    콘솔 [학습결과]의 clicked_id/memoId 값과 _detail_probe.html 을 공유해 주세요.
#    그 경우 list API 로 memoId 매핑을 잡는 방식으로 바로 보완해 드립니다.


from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_COLOR_INDEX

import requests
import os
import re
import time
import json
import uuid
import shutil


# ─────────────────────────────────────────────
# PATH / CONFIG
# ─────────────────────────────────────────────

SAVE_DIR = r"C:\KHK\DocuWork\webdocu"
IMG_FOLDER = os.path.join(SAVE_DIR, "tmp_imgs")

os.makedirs(SAVE_DIR, exist_ok=True)
os.makedirs(IMG_FOLDER, exist_ok=True)

DEFAULT_TIMEOUT = 15000  # ms
DOWNLOAD_IMAGES = True    # 이미지도 문서에 넣을지 (False면 텍스트만, 훨씬 빠름)
SORT_ASCENDING = True     # True=오름차순(1월→12월), False=내림차순

# 상세뷰 진입 후 본문 로딩 대기 (느린 회선이면 상향) — v2의 900 → 1200 로 상향
DETAIL_WAIT_MS = 1200

# 상세뷰 '끝까지 펼치기(스크롤)' 파라미터
DETAIL_SCROLL_STEPS = 40      # 최대 스크롤 반복
DETAIL_SCROLL_PAUSE_MS = 220  # 한 스텝 후 대기
DETAIL_SCROLL_STABLE = 3      # 높이가 N번 연속 그대로면 '바닥 도달'로 간주

# 상세 본문 컨테이너 후보 selector (자동 탐지에 사용) — 필요 시 맨 앞에 확정 selector 추가
DETAIL_CONTENT_CANDIDATES = [
    ".se-main-container",       # 스마트에디터3 뷰어
    ".se-viewer",
    ".paper_view",
    ".view_content",
    ".memo_view",
    ".content_view",
    ".viewer",
    ".editor_area",
    "[contenteditable='true']",
    "p.paper_contents",         # 최후: 상세에서도 이 클래스를 재사용하는 경우
]

# 상세뷰 '뒤로/닫기' 버튼 후보 (클릭스루 폴백에서 사용)
DETAIL_BACK_CANDIDATES = [
    "button.btn_back", "a.btn_back", ".btn_prev", ".header_back",
    "button[aria-label='뒤로']", "button[aria-label='닫기']", ".btn_close",
]


# ─────────────────────────────────────────────
# IMAGE DOWNLOAD (URL별 캐시)
# ─────────────────────────────────────────────

_img_cache = {}   # url -> local path (or None)

def download_image(url, cookies=None):
    if not url:
        return None
    if url in _img_cache:
        return _img_cache[url]

    candidates = []
    stripped = re.sub(r"\?type=w?\d+$", "", url)
    if stripped != url:
        candidates.append(stripped)
    candidates.append(url)

    headers = {"Referer": "https://memo.naver.com"}
    result = None
    for u in candidates:
        try:
            r = requests.get(u, headers=headers, cookies=cookies, timeout=20)
            if r.status_code != 200:
                continue
            ctype = r.headers.get("Content-Type", "")
            if "image" not in ctype:
                continue
            ext = ".png"
            if "jpeg" in ctype or "jpg" in ctype:
                ext = ".jpg"
            elif "gif" in ctype:
                ext = ".gif"
            elif "webp" in ctype:
                ext = ".webp"
            path = os.path.join(IMG_FOLDER, f"{uuid.uuid4().hex}{ext}")
            with open(path, "wb") as f:
                f.write(r.content)
            result = path
            break
        except Exception:
            continue

    _img_cache[url] = result
    return result


# ─────────────────────────────────────────────
# DATE PARSING  ('일시 :' 라인 → (year, month, day))
# ─────────────────────────────────────────────

def parse_ilsi_date(text):
    if not text:
        return None
    m = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일", text)
    if m:
        return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.search(r"(\d{4})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})", text)
    if m:
        return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def sort_key_of(memo):
    d = memo.get("date")
    if d:
        return (0,) + d
    return (1, 9999, 99, 99)


# ─────────────────────────────────────────────
# 서식 보존 파서 : 상세 HTML(트리) → docx 문단/런
# ─────────────────────────────────────────────

BLOCK_TAGS = {"p", "div", "li", "blockquote", "section", "article",
              "h1", "h2", "h3", "h4", "h5", "h6", "tr", "table", "ul", "ol"}

# 하이라이트 배경색이 아닌 것으로 간주할 값들
_TRANSPARENT = {"transparent", "#fff", "#ffffff", "white", "rgb(255,255,255)",
                "rgba(0,0,0,0)", "none", "inherit", "initial"}


# docx 형광펜 팔레트 (근사색 매핑용). RGB 는 대표값.
_HL_PALETTE = [
    (WD_COLOR_INDEX.YELLOW,       (255, 255, 0)),
    (WD_COLOR_INDEX.BRIGHT_GREEN, (0, 255, 0)),
    (WD_COLOR_INDEX.TURQUOISE,    (0, 255, 255)),
    (WD_COLOR_INDEX.PINK,         (255, 0, 255)),
    (WD_COLOR_INDEX.RED,          (255, 0, 0)),
    (WD_COLOR_INDEX.BLUE,         (0, 0, 255)),
    (WD_COLOR_INDEX.DARK_YELLOW,  (128, 128, 0)),
    (WD_COLOR_INDEX.GREEN,        (0, 128, 0)),
    (WD_COLOR_INDEX.VIOLET,       (128, 0, 128)),
    (WD_COLOR_INDEX.GRAY_25,      (192, 192, 192)),
]

_NAMED_HL = {
    "yellow": WD_COLOR_INDEX.YELLOW, "lime": WD_COLOR_INDEX.BRIGHT_GREEN,
    "cyan": WD_COLOR_INDEX.TURQUOISE, "aqua": WD_COLOR_INDEX.TURQUOISE,
    "magenta": WD_COLOR_INDEX.PINK, "fuchsia": WD_COLOR_INDEX.PINK,
    "pink": WD_COLOR_INDEX.PINK, "red": WD_COLOR_INDEX.RED,
    "blue": WD_COLOR_INDEX.BLUE, "green": WD_COLOR_INDEX.GREEN,
    "gray": WD_COLOR_INDEX.GRAY_25, "grey": WD_COLOR_INDEX.GRAY_25,
}


def _css_color_to_rgb(val):
    """#rrggbb / #rgb / rgb(r,g,b) → RGBColor. 실패 시 None."""
    val = val.strip().lower()
    m = re.match(r"#([0-9a-f]{6})$", val)
    if m:
        h = m.group(1)
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    m = re.match(r"#([0-9a-f]{3})$", val)
    if m:
        h = m.group(1)
        return RGBColor(int(h[0]*2, 16), int(h[1]*2, 16), int(h[2]*2, 16))
    m = re.match(r"rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)", val)
    if m:
        return RGBColor(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def _css_color_to_hl_index(val):
    """형광펜 배경색 → docx 형광펜 인덱스(근사)."""
    rgb = _css_color_to_rgb(val)
    if rgb is None:
        return _NAMED_HL.get(val.strip().lower())
    r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
    best, bestd = None, 1e18
    for idx, (pr, pg, pb) in _HL_PALETTE:
        d = (r - pr) ** 2 + (g - pg) ** 2 + (b - pb) ** 2
        if d < bestd:
            bestd, best = d, idx
    return best


def _class_str(node):
    if not hasattr(node, "get"):
        return ""
    c = node.get("class")
    if not c:
        return ""
    return (" ".join(c) if isinstance(c, list) else str(c)).lower()


def _merge_style(node, inherited):
    """현재 태그의 이름/class/인라인 style 을 보고 서식 플래그를 상속·병합해 반환."""
    f = dict(inherited)
    name = getattr(node, "name", None)

    # 1) 태그명 기반
    if name in ("u", "ins"):
        f["underline"] = True
    if name in ("b", "strong"):
        f["bold"] = True
    if name in ("i", "em"):
        f["italic"] = True
    if name in ("s", "strike", "del"):
        f["strike"] = True
    if name == "mark":
        f["highlight"] = True

    # 2) class 기반 (에디터가 밑줄/형광펜을 class 로 주는 경우 대비)
    cl = _class_str(node)
    if cl:
        if "underline" in cl:
            f["underline"] = True
        if ("strike" in cl) or ("line-through" in cl):
            f["strike"] = True
        if ("highlight" in cl) or ("marker" in cl) or ("hl_" in cl) or ("se-highlight" in cl):
            f["highlight"] = True

    # 3) 인라인 style 기반
    style = (node.get("style") or "").lower() if hasattr(node, "get") else ""
    if style:
        # 밑줄 / 취소선 (text-decoration, text-decoration-line 모두 'underline'/'line-through' 포함)
        if "underline" in style:
            f["underline"] = True
        if "line-through" in style:
            f["strike"] = True
        # 굵게
        mw = re.search(r"font-weight\s*:\s*([a-z0-9]+)", style)
        if mw:
            v = mw.group(1)
            if v in ("bold", "bolder") or (v.isdigit() and int(v) >= 600):
                f["bold"] = True
        # 기울임
        if re.search(r"font-style\s*:\s*italic", style):
            f["italic"] = True
        # 형광펜(배경색) + 실제 색 → docx 형광펜 인덱스 매핑
        mb = re.search(r"background(?:-color)?\s*:\s*([^;]+)", style)
        if mb:
            val = mb.group(1).strip()
            if val and val.replace(" ", "") not in _TRANSPARENT:
                f["highlight"] = True
                idx = _css_color_to_hl_index(val)
                if idx is not None:
                    f["hl_index"] = idx
        # 글자색 (background-color 와 겹치지 않게 경계 지정)
        mc = re.search(r"(?:^|;)\s*color\s*:\s*([^;]+)", style)
        if mc:
            rgb = _css_color_to_rgb(mc.group(1).strip())
            if rgb:
                f["color"] = rgb
    return f


def _tokenize(node, inherited, out):
    """
    HTML 트리를 순회하며 토큰 스트림 생성.
      ("text", 문자열, style)   ("img", src, None)   ("nl", None, None)=줄바꿈
    <br> 과 블록태그 경계를 줄바꿈으로 취급.
    """
    for child in node.children:
        if isinstance(child, NavigableString):
            t = str(child).replace("\xa0", " ")
            if t:
                out.append(("text", t, inherited))
        elif isinstance(child, Tag):
            n = child.name
            if n == "br":
                out.append(("nl", None, None))
            elif n == "img":
                src = child.get("src") or child.get("data-src") or child.get("data-lazy-src") or ""
                if src.startswith("http"):
                    out.append(("img", src, None))
            elif n in BLOCK_TAGS:
                out.append(("nl", None, None))
                _tokenize(child, _merge_style(child, inherited), out)
                out.append(("nl", None, None))
            else:  # span, a, font, u, b, i, mark ... 인라인
                _tokenize(child, _merge_style(child, inherited), out)


def _apply_style(run, style):
    if not style:
        return
    if style.get("underline"):
        run.font.underline = True
    if style.get("bold"):
        run.font.bold = True
    if style.get("italic"):
        run.font.italic = True
    if style.get("strike"):
        run.font.strike = True
    if style.get("highlight"):
        # 실제 형광펜 색이 매핑돼 있으면 그 색, 아니면 노랑으로 근사
        run.font.highlight_color = style.get("hl_index") or WD_COLOR_INDEX.YELLOW
    if style.get("color") is not None:
        try:
            run.font.color.rgb = style["color"]
        except Exception:
            pass


def add_html_to_doc(container_html, doc, cookies=None):
    """상세 본문 HTML → 서식 보존하여 doc 에 문단들로 추가."""
    soup = BeautifulSoup(container_html, "html.parser")

    tokens = []
    _tokenize(soup, {}, tokens)

    # nl 기준으로 문단 분할
    paragraphs = []
    cur = []
    for tok in tokens:
        if tok[0] == "nl":
            paragraphs.append(cur)
            cur = []
        else:
            cur.append(tok)
    paragraphs.append(cur)

    prev_blank = True  # 연속 빈 줄 폭주 방지 (맨 처음 빈 줄도 억제)
    for seg in paragraphs:
        # 텍스트만 추려 공백 여부 판단 (이미지가 있으면 비어있지 않음)
        has_content = any(
            (t[0] == "img") or (t[0] == "text" and t[1].strip())
            for t in seg
        )
        if not has_content:
            if not prev_blank:
                doc.add_paragraph()   # 빈 줄 1개만 유지
                prev_blank = True
            continue

        p = doc.add_paragraph()
        for t in seg:
            if t[0] == "text":
                text = re.sub(r"[ \t\r\n]+", " ", t[1])
                if text == " " and not p.runs:
                    continue  # 문단 맨 앞 홑공백 제거
                run = p.add_run(text)
                _apply_style(run, t[2])
            elif t[0] == "img":
                if DOWNLOAD_IMAGES:
                    path = download_image(t[1], cookies)
                    if path:
                        try:
                            p.add_run().add_picture(path, width=Inches(4.8))
                        except Exception:
                            pass
        prev_blank = False


# ─────────────────────────────────────────────
# 상세 memo → docx (헤더/메타 + 서식 본문)
# ─────────────────────────────────────────────

def write_memo_to_doc(memo, doc, cookies=None):
    doc.add_heading(memo["title"], level=2)

    meta = doc.add_paragraph()
    if memo["date"]:
        y, mth, d = memo["date"]
        run = meta.add_run(f"📅 {y}-{mth:02d}-{d:02d}    🕒 {memo['modified_date']}")
    else:
        run = meta.add_run(f"🕒 {memo['modified_date']}")
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(120, 120, 120)

    doc.add_paragraph()

    # 상세 HTML 이 있으면 서식 보존 파서로, 없으면(폴백) 목록 카드 줄들을 그대로
    if memo.get("detail_html"):
        add_html_to_doc(memo["detail_html"], doc, cookies)
    else:
        for ln in memo.get("lines", []):
            if ln == "":
                doc.add_paragraph()
            else:
                doc.add_paragraph(ln)

    doc.add_paragraph("━" * 35)
    doc.add_paragraph()


# ─────────────────────────────────────────────
# 목록 카드 → 메타(제목/일시/수정일) 파싱  (본문은 상세에서 별도 취득)
# ─────────────────────────────────────────────

def parse_card_meta(item_html, index):
    soup = BeautifulSoup(item_html, "html.parser")
    contents = soup.select_one("p.paper_contents")
    date_el = soup.select_one("span.date")

    modified_date = ""
    if date_el:
        blind = date_el.select_one(".blind")
        if blind:
            blind.extract()
        modified_date = date_el.get_text(strip=True)

    # 카드 미리보기에서 일시/제목만 뽑기(정렬·헤더용)
    lines = []
    if contents:
        buf = ""
        for node in contents.children:
            if isinstance(node, Tag) and node.name == "br":
                lines.append(buf.strip()); buf = ""
            elif isinstance(node, NavigableString):
                buf += str(node).replace("\xa0", " ")
            elif isinstance(node, Tag):
                buf += node.get_text()
        lines.append(buf.strip())

    diary_date, date_line, title_text = None, "", ""
    for ln in lines:
        compact = ln.replace(" ", "")
        if diary_date is None and compact.startswith("일시:"):
            diary_date = parse_ilsi_date(ln)
            date_line = ln
        elif not title_text and compact.startswith("제목:"):
            title_text = ln.split(":", 1)[1].strip()
    if not title_text:
        title_text = modified_date or f"메모 {index}"

    return {
        "date": diary_date,
        "date_line": date_line,
        "title": title_text,
        "lines": lines,             # 폴백용 미리보기
        "modified_date": modified_date,
        "index": index,
        "detail_html": None,        # 상세에서 채움
    }


# ─────────────────────────────────────────────
# FOLDER LIST
# ─────────────────────────────────────────────

def get_folder_tabs(page):
    page.wait_for_selector("div.memo", timeout=DEFAULT_TIMEOUT)
    result = []
    for tab in page.query_selector_all("div.memo"):
        try:
            name_el = tab.query_selector("span.text")
            count_el = tab.query_selector("span.count")
            if not name_el:
                continue
            name = name_el.inner_text().strip()
            count = count_el.inner_text().strip() if count_el else "0"
            if re.search(r"\d{4}", name):
                cnt = re.sub(r"[^\d]", "", count) or "0"
                result.append({"name": name, "count": int(cnt)})
        except Exception:
            continue
    return result


def click_folder(page, folder_name):
    for tab in page.query_selector_all("div.memo"):
        try:
            name_el = tab.query_selector("span.text")
            if not name_el:
                continue
            if name_el.inner_text().strip() == folder_name:
                tab.scroll_into_view_if_needed()
                try:
                    tab.click(timeout=3000)
                except Exception:
                    tab.evaluate("el => el.click()")
                return True
        except Exception:
            continue
    return False


# ─────────────────────────────────────────────
# SCROLL & COLLECT (data-id 만 순서대로 수집)
# ─────────────────────────────────────────────

SCROLL_WAIT_MS = 750
NO_PROGRESS_PATIENCE = 8

SCROLL_TOP_JS = """
() => {
    const ul = document.querySelector('ul.memo_list_view');
    if (!ul) return;
    let el = ul;
    while (el) {
        const s = getComputedStyle(el);
        if ((s.overflowY === 'auto' || s.overflowY === 'scroll')
            && el.scrollHeight > el.clientHeight + 4) { el.scrollTop = 0; return; }
        el = el.parentElement;
    }
    const de = document.scrollingElement || document.documentElement;
    if (de) de.scrollTop = 0;
}
"""


def collect_folder_ids(page, target_total):
    """리스트를 끝까지 스크롤하며 li.paper_item 의 (data-id, 카드 outerHTML)을 순서대로 수집."""
    page.wait_for_selector("li.paper_item", timeout=DEFAULT_TIMEOUT)
    try:
        page.evaluate(SCROLL_TOP_JS)
        page.wait_for_timeout(400)
    except Exception:
        pass
    try:
        page.hover("ul.memo_list_view")
    except Exception:
        pass

    collected = {}   # data-id -> card outerHTML (순서 보존; dict는 삽입순)
    last_marker = None
    stagnant = 0
    max_iters = max(3000, target_total * 4) if target_total > 0 else 3000

    for _ in range(max_iters):
        for it in page.query_selector_all("li.paper_item[data-id]"):
            try:
                did = it.get_attribute("data-id")
                if did and did not in collected:
                    collected[did] = it.evaluate("el => el.outerHTML")
            except Exception:
                continue
        got = len(collected)

        if target_total > 0 and got >= target_total:
            print(f"    ... 수집 {got}개 (목표 도달)          ")
            break

        items = page.query_selector_all("li.paper_item[data-id]")
        if not items:
            break
        try:
            last_id = items[-1].get_attribute("data-id")
        except Exception:
            last_id = None

        marker = (last_id, got)
        if marker == last_marker:
            stagnant += 1
        else:
            stagnant = 0
            print(f"    ... 수집 {got}개", end="\r")
        last_marker = marker
        if stagnant >= NO_PROGRESS_PATIENCE:
            print(f"\n    ... 바닥 도달 (총 {got}개)")
            break

        try:
            items[-1].scroll_into_view_if_needed(timeout=3000)
        except Exception:
            pass
        try:
            page.mouse.wheel(0, 2200)
        except Exception:
            pass
        page.wait_for_timeout(SCROLL_WAIT_MS)

    print()
    return collected


# ─────────────────────────────────────────────
# 상세 본문 컨테이너 selector 자동 탐지 (v3: '헤더만 잡는' 오탐 방지)
# ─────────────────────────────────────────────

def _detect_content_selector(page):
    """
    후보 selector 중 '가장 긴 본문'을 담은 것을 고른다.
    v2는 마커(일시/제목) 가점(+100000)이 지나치게 커서 '헤더만 담은 짧은 요소'가
    본문을 제치고 뽑혔음 → 본문이 통째로 잘림.
    v3은 길이를 1순위로 두고 마커는 소폭(+200) 가점만 준다.
    """
    best_sel, best_score = None, -1
    for sel in DETAIL_CONTENT_CANDIDATES:
        try:
            els = page.query_selector_all(sel)
        except Exception:
            continue
        for el in els:
            try:
                txt = el.inner_text()
            except Exception:
                continue
            L = len((txt or "").strip())
            if L < 1:
                continue
            has_marker = ("일시" in txt or "제목" in txt or "날씨" in txt)
            score = L + (200 if has_marker else 0)   # 길이 우선, 마커는 소폭 가점
            if score > best_score:
                best_score, best_sel = score, sel
    return best_sel


def _grab_fullest_html(page, sel):
    """sel 로 매칭되는 요소들 중 innerText 가 가장 긴 것의 innerHTML 반환."""
    if not sel:
        return None
    try:
        els = page.query_selector_all(sel)
    except Exception:
        els = []
    best_html, best_len = None, -1
    for el in els:
        try:
            L = len(el.inner_text())
            if L > best_len:
                best_len = L
                best_html = el.evaluate("e => e.innerHTML")
        except Exception:
            continue
    return best_html


# ─────────────────────────────────────────────
# 상세뷰 '끝까지 펼치기' 후 innerHTML 캡처 (증상1의 핵심 해결부)
# ─────────────────────────────────────────────

DETAIL_EXPAND_JS = """
async () => {
  const expandWords = ['더보기','펼치기','전체보기','전체 보기','more','See more'];
  document.querySelectorAll('button, a, span, div').forEach(b=>{
    const t=(b.innerText||'').trim();
    if (t && t.length<=8 && expandWords.some(w=>t.includes(w))){ try{ b.click(); }catch(e){} }
  });
  function scrollers(){
    const out=[];
    document.querySelectorAll('*').forEach(el=>{
      const s=getComputedStyle(el);
      if ((s.overflowY==='auto'||s.overflowY==='scroll') && el.scrollHeight>el.clientHeight+8) out.push(el);
    });
    out.push(document.scrollingElement||document.documentElement);
    return out;
  }
  let last=-1, stable=0;
  for (let i=0;i<__STEPS__;i++){
    const scs=scrollers();
    let total=0;
    scs.forEach(sc=>{ try{ sc.scrollTop=sc.scrollHeight; total+=sc.scrollHeight; }catch(e){} });
    window.scrollTo(0, document.body.scrollHeight);
    await new Promise(r=>setTimeout(r,__PAUSE__));
    if (total===last){ stable++; if(stable>=__STABLE__) break; } else { stable=0; last=total; }
  }
  document.querySelectorAll('img[data-src], img[data-lazy-src]').forEach(im=>{
    const s=im.getAttribute('data-src')||im.getAttribute('data-lazy-src');
    if (s && !im.src) im.src=s;
  });
}
""".replace("__STEPS__", str(DETAIL_SCROLL_STEPS)) \
   .replace("__PAUSE__", str(DETAIL_SCROLL_PAUSE_MS)) \
   .replace("__STABLE__", str(DETAIL_SCROLL_STABLE))


def expand_and_capture_detail(page, content_selector):
    """
    상세뷰 진입 직후 호출.
    networkidle 대기 → 마커 대기 → 본문 스크롤러를 바닥까지 반복 스크롤(높이 안정 시까지)
    → '더보기' 펼침 → 그 다음에 '가장 긴 본문 컨테이너'의 innerHTML 캡처.
    반환: (html or None, 사용한 selector)
    """
    # 1) 네트워크/마커 대기 (지연 렌더 대비)
    try:
        page.wait_for_load_state("networkidle", timeout=DEFAULT_TIMEOUT)
    except Exception:
        pass
    try:
        page.wait_for_function(
            "() => /일시|제목|날씨/.test(document.body.innerText)",
            timeout=3000,
        )
    except Exception:
        pass

    sel = content_selector or _detect_content_selector(page)

    # 2) 끝까지 펼치기/스크롤 (가상 DOM/lazy-render 실제화)
    try:
        page.evaluate(DETAIL_EXPAND_JS)   # 모든 스크롤러 바닥까지
    except Exception:
        pass
    page.wait_for_timeout(300)

    # 3) 펼친 뒤 다시 탐지(더 긴 컨테이너가 생겼을 수 있음) 후 캡처
    sel2 = _detect_content_selector(page) or sel
    html = _grab_fullest_html(page, sel2)
    if not html and sel and sel != sel2:
        html = _grab_fullest_html(page, sel)
    return html, (sel2 or sel)


# ─────────────────────────────────────────────
# 상세 HTML 취득 (URL직접이동 우선, 없으면 클릭스루)
# ─────────────────────────────────────────────

def build_detail_url(template, folder_id, memo_id):
    """학습한 URL 템플릿에 현재 folderId 와 memoId 를 반영해 상세 URL 생성."""
    url = template.replace("__ID__", str(memo_id))
    if folder_id:
        url = re.sub(r"folderId=\d+", f"folderId={folder_id}", url)
    return url


def fetch_detail_html_by_url(page, data_id, url_template, content_selector, folder_id=None):
    url = build_detail_url(url_template, folder_id, data_id)
    try:
        page.goto(url, wait_until="domcontentloaded")
    except Exception:
        return None
    page.wait_for_timeout(DETAIL_WAIT_MS)
    html, _sel = expand_and_capture_detail(page, content_selector)
    return html


def fetch_detail_html_by_click(page, data_id, content_selector):
    """클릭스루 폴백: 리스트에서 해당 카드를 찾아 클릭→상세 HTML→복귀."""
    card = page.query_selector(f"li.paper_item[data-id='{data_id}']")
    if not card:
        # 가상 스크롤로 화면 밖이면 리스트에서 스크롤하며 탐색
        card = _scroll_find_card(page, data_id)
    if not card:
        return None
    try:
        card.scroll_into_view_if_needed(timeout=3000)
        card.click(timeout=4000)
    except Exception:
        try:
            card.evaluate("el => el.click()")
        except Exception:
            return None
    page.wait_for_timeout(DETAIL_WAIT_MS)
    html, _sel = expand_and_capture_detail(page, content_selector)
    _return_to_list(page)
    return html


def _scroll_find_card(page, data_id, max_scroll=40):
    try:
        page.evaluate(SCROLL_TOP_JS)
        page.wait_for_timeout(300)
    except Exception:
        pass
    for _ in range(max_scroll):
        card = page.query_selector(f"li.paper_item[data-id='{data_id}']")
        if card:
            return card
        items = page.query_selector_all("li.paper_item[data-id]")
        if not items:
            return None
        try:
            items[-1].scroll_into_view_if_needed(timeout=2000)
            page.mouse.wheel(0, 2200)
        except Exception:
            pass
        page.wait_for_timeout(SCROLL_WAIT_MS)
    return None


# ─────────────────────────────────────────────
# 상세 접근 방법 학습 (URL직접이동 vs 클릭스루) + 본문 컨테이너 selector 자동 탐지
# ─────────────────────────────────────────────

def learn_detail_access(page, first_card_id):
    """
    첫 카드를 '리스트 맨 위로 올린 뒤' 클릭해서:
      - URL 의 memoId 를 추출하고, 실제로 클릭한 카드의 data-id 와 비교
        (data-id == memoId 여야 URL 직접이동으로 전 메모 접근 가능)
      - '끝까지 펼친 뒤' 상세 본문 컨테이너 selector 탐지
      - 첫 상세 HTML 을 _detail_probe.html 로 저장
    반환: (url_template or None, content_selector or None)
    """
    url_before = page.url

    # (핵심) 리스트를 맨 위로 올려 ids[0] 카드를 확실히 보이게 한 뒤 클릭
    try:
        page.evaluate(SCROLL_TOP_JS)
        page.wait_for_timeout(500)
    except Exception:
        pass

    card = page.query_selector(f"li.paper_item[data-id='{first_card_id}']")
    if not card:
        card = _scroll_find_card(page, first_card_id)
    if not card:
        card = page.query_selector("li.paper_item[data-id]")
    if not card:
        print("  ⚠️ [학습] 클릭할 카드를 못 찾음")
        return None, None

    # (핵심) '실제로 클릭하는' 카드의 data-id 를 확보해 URL 의 memoId 와 비교한다
    try:
        clicked_id = card.get_attribute("data-id")
    except Exception:
        clicked_id = first_card_id

    try:
        card.scroll_into_view_if_needed(timeout=3000)
    except Exception:
        pass
    try:
        card.click(timeout=4000)
    except Exception:
        try:
            card.evaluate("el => el.click()")
        except Exception:
            print("  ⚠️ [학습] 카드 클릭 실패")
            return None, None

    page.wait_for_timeout(DETAIL_WAIT_MS + 400)
    url_after = page.url

    # (핵심) memoId 추출 -> 템플릿화 + clicked_id 와 동일한지 확인
    url_template = None
    m = re.search(r"memoId=(\d+)", url_after)
    memo_id = m.group(1) if m else None
    if memo_id:
        tmpl = re.sub(r"memoId=\d+", "memoId=__ID__", url_after)
        if clicked_id and str(clicked_id) == memo_id:
            url_template = tmpl   # data-id == memoId -> URL 직접이동 가능
        # else: data-id != memoId -> 아래 [학습결과]에서 안내, 클릭스루로 진행

    # 끝까지 펼친 뒤 selector 탐지
    try:
        page.wait_for_load_state("networkidle", timeout=DEFAULT_TIMEOUT)
    except Exception:
        pass
    pre_sel = _detect_content_selector(page)
    try:
        page.evaluate(DETAIL_EXPAND_JS)   # 모든 스크롤러 바닥까지
    except Exception:
        pass
    page.wait_for_timeout(300)
    content_selector = _detect_content_selector(page) or pre_sel

    # 확인용 덤프 저장 (가장 긴 컨테이너의 outerHTML, 없으면 전체 페이지)
    try:
        html = None
        if content_selector:
            els = page.query_selector_all(content_selector)
            best_len = -1
            for el in els:
                try:
                    L = len(el.inner_text())
                    if L > best_len:
                        best_len = L
                        html = el.evaluate("el => el.outerHTML")
                except Exception:
                    continue
        if not html:
            html = page.content()
        probe_path = os.path.join(SAVE_DIR, "_detail_probe.html")
        with open(probe_path, "w", encoding="utf-8") as f:
            f.write(html)
    except Exception:
        probe_path = "(저장 실패)"

    nav_mode = "ON (page.goto 로 직접 이동 -> 멈춤 없음/빠름)" if url_template \
               else "OFF (클릭스루 폴백)"
    print("\n  ── [학습결과] ──────────────────────────────")
    print(f"   URL 변화        : {url_before}  ->  {url_after}")
    print(f"   clicked data-id : {clicked_id}")
    print(f"   URL memoId      : {memo_id}")
    print(f"   data-id==memoId : {bool(url_template)}")
    print(f"   URL 직접이동    : {nav_mode}")
    print(f"   URL 템플릿      : {url_template or '없음'}")
    print(f"   본문 selector   : {content_selector or '❗미탐지 (후보 selector 조정 필요)'}")
    print(f"   상세 HTML 덤프  : {probe_path}")
    print("  ───────────────────────────────────────────\n")

    # 리스트로 복귀
    _return_to_list(page)
    return url_template, content_selector


def _return_to_list(page):
    """상세 → 리스트 복귀 (뒤로가기 → 실패 시 닫기버튼 → 최후 메모루트 재진입)."""
    for _ in range(2):
        try:
            page.go_back(wait_until="domcontentloaded", timeout=5000)
        except Exception:
            pass
        page.wait_for_timeout(500)
        if page.query_selector("li.paper_item[data-id]"):
            return True
    for sel in DETAIL_BACK_CANDIDATES:
        try:
            btn = page.query_selector(sel)
            if btn:
                btn.click(timeout=2000)
                page.wait_for_timeout(600)
                if page.query_selector("li.paper_item[data-id]"):
                    return True
        except Exception:
            continue
    try:
        page.goto("https://memo.naver.com", wait_until="domcontentloaded")
        page.wait_for_timeout(1200)
    except Exception:
        pass
    return bool(page.query_selector("li.paper_item[data-id]"))


# ─────────────────────────────────────────────
# GOTO MEMO
# ─────────────────────────────────────────────

def goto_memo(page):
    for _ in range(8):
        try:
            page.goto("https://memo.naver.com", wait_until="domcontentloaded")
        except Exception:
            pass
        time.sleep(2)
        url = page.url
        if "memo.naver.com" in url:
            return True
        if "nid.naver.com" in url:
            print("\n  ⚠️ 네이버 추가 절차(기기 등록/보안 확인) 화면:")
            print(f"     {url}")
            input("  👉 처리 완료 후 Enter... ")
        time.sleep(1)
    return False


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
        )
        context = browser.new_context(
            no_viewport=True,
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/124.0.0.0 Safari/537.36"),
        )
        page = context.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT)

        # ── LOGIN ──
        page.goto("https://nid.naver.com/nidlogin.login")
        print("\n✅ 브라우저에서 네이버 로그인 진행하세요.")
        input("\n👉 로그인이 완전히 끝났으면 Enter... ")

        if not goto_memo(page):
            print("\n  ⚠️ 메모 진입 실패."); input("\n엔터 종료... "); browser.close(); return
        time.sleep(2)

        folders = get_folder_tabs(page)
        print(f"\n📁 연도 폴더 {len(folders)}개 발견")
        if not folders:
            print("  ⚠️ 폴더 없음."); input("\n엔터 종료... "); browser.close(); return

        url_template = None
        content_selector = None
        learned = False

        for folder in folders:
            folder_name = folder["name"]
            target_total = folder["count"]
            print("\n" + "=" * 50)
            print(f"📂 {folder_name} ({target_total}개)")
            print("=" * 50)

            if not click_folder(page, folder_name):
                print("  ⚠️ 폴더 클릭 실패"); continue
            page.wait_for_timeout(1500)

            # (v4) 현재 폴더의 folderId 를 URL 에서 확보 -> 상세 URL 조립에 사용
            mfid = re.search(r"folderId=(\d+)", page.url)
            folder_id = mfid.group(1) if mfid else None

            req_cookies = {c["name"]: c["value"] for c in context.cookies()}

            # 1) data-id 전체 수집
            try:
                cards = collect_folder_ids(page, target_total)
            except PWTimeout:
                print("  ⚠️ 리스트 로딩 실패"); continue

            ids = list(cards.keys())
            got = len(ids)
            if target_total > 0 and got < target_total:
                print(f"  ⚠️ 목표 {target_total} 중 {got}개만 수집됨")
            nav_now = "URL 직접이동" if url_template else ("클릭스루" if learned else "학습 후 결정")
            print(f"  📥 {got}개 카드 → 상세 본문 수집 시작... (방식: {nav_now})")

            # 2) 첫 폴더에서 상세 접근법 학습 (한 번만)
            if not learned and ids:
                url_template, content_selector = learn_detail_access(page, ids[0])
                learned = True
                # 학습 위해 리스트를 벗어났으니, URL직접이동이 아니면 폴더 리스트 재확보
                if not url_template:
                    click_folder(page, folder_name)
                    page.wait_for_timeout(1200)

            # 3) 카드 메타 파싱 + 상세 본문 취득
            memos = []
            for i, did in enumerate(ids, start=1):
                memo = parse_card_meta(cards[did], i)

                if url_template:
                    html = fetch_detail_html_by_url(page, did, url_template, content_selector, folder_id)
                else:
                    html = fetch_detail_html_by_click(page, did, content_selector)

                # v3: 상세 HTML 에 '의미있는 내용'(10자↑)만 있으면 무조건 채택.
                #     (v2의 '미리보기보다 길어야 채택' 조건은 헤더/본문 분리 케이스에서
                #      정상 캡처본을 되레 버려 '...'+무서식 폴백을 유발했음 → 제거)
                if html:
                    dtxt = BeautifulSoup(html, "html.parser").get_text(strip=True)
                    if len(dtxt) >= 10:
                        memo["detail_html"] = html
                memos.append(memo)

                if i % 10 == 0 or i == got:
                    print(f"    상세 {i}/{got}", end="\r")
            print()

            # 4) 정렬
            memos.sort(key=sort_key_of, reverse=not SORT_ASCENDING)
            dated = sum(1 for m in memos if m["date"])
            fulltext = sum(1 for m in memos if m.get("detail_html"))
            fallback = got - fulltext
            direction = "오름차순(1월→12월)" if SORT_ASCENDING else "내림차순"
            print(f"  🔀 정렬 완료: {direction} | 날짜 인식 {dated} | 전체본문 {fulltext}/{got}")
            if fulltext == 0:
                print("  ❗ 상세 본문을 하나도 못 가져왔습니다. "
                      "_detail_probe.html 과 위 [학습결과]를 확인하세요.")
            elif fallback:
                print(f"  ⚠️ 미리보기(요약)로 대체된 메모 {fallback}개 → 이 문서들이 '...'로 "
                      f"잘리고 서식이 없습니다.")
                print(f"     _detail_probe.html 을 열어 본문 컨테이너 class 를 확인 후, "
                      f"DETAIL_CONTENT_CANDIDATES 맨 앞에 추가하세요.")

            # 5) 문서화
            doc = Document()
            doc.add_heading(f"네이버 메모 - {folder_name}", level=1)
            for i, memo in enumerate(memos, start=1):
                try:
                    write_memo_to_doc(memo, doc, req_cookies)
                except Exception as e:
                    print(f"    ⚠️ '{memo.get('title')}' 문서화 실패: {e}")

            safe_name = re.sub(r'[\\/*?:"<>|]', "", folder_name)
            save_file = os.path.join(SAVE_DIR, f"naver_memo_{safe_name}.docx")
            try:
                doc.save(save_file)
                print(f"  💾 저장 완료: {save_file}")
            except Exception as e:
                print(f"  ⚠️ 저장 실패: {e}")

            # URL직접이동이 아니었다면 다음 폴더 위해 리스트 상태 초기화
            if not url_template:
                page.goto("https://memo.naver.com", wait_until="domcontentloaded")
                page.wait_for_timeout(1200)

        browser.close()

    try:
        shutil.rmtree(IMG_FOLDER, ignore_errors=True)
    except Exception:
        pass
    print("\n🎉 전체 완료")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n❌ 치명적 오류 발생"); print(e)
        input("\n엔터 종료...")