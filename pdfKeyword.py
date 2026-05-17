# pdfKeyword.py

import pdfplumber
from konlpy.tag import Okt
from sklearn.feature_extraction.text import TfidfVectorizer
from collections import defaultdict
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import glob, os

# ── 설정 ────────────────────────────────────────────────
FOLDER  = r"C:\KHK\DocuWork\object"
TOP_N   = 20
MIN_LEN = 2

# 불용어 목록 (자유롭게 추가 가능!)
STOPWORDS = {
    "오늘", "어제", "내일", "일단", "지금", "정말", "너무", "그냥", "조금",
    "계속", "다시", "이번", "아직", "때문", "이후", "이전", "이상", "이하",
    "우리", "자신", "그것", "이것", "저것", "무엇", "어떤", "모든", "매우",
    "하나", "이미", "또한", "이제", "결국", "사실", "경우", "정도", "부분",
    "관련", "통해", "위해", "대해", "따라", "대한", "으로", "에서", "에게",
    "날씨", "제목", "일시", "맑음", "흐림", "거의", "그때", "자체", "뭔가",
    "순간", "진짜", "자체", "느낌", "전혀", "바로"
}
# ────────────────────────────────────────────────────────

font_path = "C:/Windows/Fonts/malgun.ttf"
font_prop = fm.FontProperties(fname=font_path)
plt.rcParams['font.family'] = font_prop.get_name()

okt = Okt()
pdf_files = sorted(glob.glob(os.path.join(FOLDER, "*.pdf")))
print(f"총 {len(pdf_files)}개 PDF 발견\n")

# ── 복합명사 포함 명사 추출 함수 ────────────────────────────
def extract_nouns_with_compounds(text, min_len=2, stopwords=set()):
    """
    단일 명사 + 연속 명사 결합(복합명사)을 함께 추출합니다.
    예) ['자연어', '처리'] → '자연어처리'도 후보에 추가
    """
    pos_tags = okt.pos(text, norm=True, stem=True)
    
    result   = []
    compound = []   # 연속 명사 버퍼

    def flush(buf):
        """버퍼에 쌓인 연속 명사를 단일/복합 형태로 모두 결과에 추가"""
        for word in buf:                          # 단일 명사
            if len(word) >= min_len and word not in stopwords:
                result.append(word)
        # 연속 명사가 2개 이상이면 복합명사 생성 (2-gram, 3-gram …)
        for size in range(2, len(buf) + 1):
            for start in range(len(buf) - size + 1):
                combined = "".join(buf[start : start + size])
                if len(combined) >= min_len and combined not in stopwords:
                    result.append(combined)

    for word, tag in pos_tags:
        if tag == "Noun":
            compound.append(word)
        else:
            if compound:
                flush(compound)
                compound = []

    if compound:              # 문장 끝에 명사가 남아 있을 경우 처리
        flush(compound)

    return result
# ─────────────────────────────────────────────────────────────

# ── 1. 각 파일에서 명사 추출 ─────────────────────────────
doc_nouns = {}   # 파일명 → 명사 리스트
doc_texts = []   # TF-IDF용 텍스트 리스트
doc_names = []   # 파일명 리스트

for i, pdf_path in enumerate(pdf_files, 1):
    fname = os.path.basename(pdf_path)
    print(f"[{i}/{len(pdf_files)}] 추출 중: {fname}")

    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""

    if not text.strip():
        print(f"   ⚠️ 텍스트 없음, 스킵")
        continue

    # 명사 추출 + 불용어 제거
    nouns = extract_nouns_with_compounds(text, min_len=MIN_LEN, stopwords=STOPWORDS)

    doc_nouns[fname] = nouns
    doc_texts.append(" ".join(nouns))  # TF-IDF용
    doc_names.append(fname)

# ── 2. TF-IDF 계산 ──────────────────────────────────────
print("\n🔍 TF-IDF 분석 중...")

vectorizer = TfidfVectorizer(max_features=500)
tfidf_matrix = vectorizer.fit_transform(doc_texts)
feature_names = vectorizer.get_feature_names_out()

# ── 3. 파일별 TF-IDF 상위 키워드 출력 + 그래프 ────────────
for i, fname in enumerate(doc_names):
    scores = tfidf_matrix[i].toarray().flatten()
    top_indices = scores.argsort()[::-1][:TOP_N]
    top_keywords = [(feature_names[j], scores[j]) for j in top_indices if scores[j] > 0]

    print(f"\n📄 {fname}")
    print("-" * 40)
    for rank, (word, score) in enumerate(top_keywords[:10], 1):
        print(f"  {rank:>2}. {word:<12} {score:.4f}")

    # 그래프 저장
    words  = [w for w, _ in top_keywords]
    values = [s for _, s in top_keywords]

    plt.figure(figsize=(12, 5))
    bars = plt.bar(words, values, color='mediumseagreen')
    plt.bar_label(bars, fmt='%.3f', padding=3, fontsize=8)
    plt.title(f"{fname} - 의미 키워드 TOP {TOP_N} (TF-IDF)", fontsize=13, fontproperties=font_prop)
    plt.xticks(rotation=30, ha='right', fontproperties=font_prop)
    plt.tight_layout()

    out_path = os.path.join(FOLDER, fname.replace(".pdf", "_tfidf.png"))
    plt.savefig(out_path, dpi=150)
    plt.close()

# ── 4. 전체 평균 TF-IDF 상위 키워드 ─────────────────────
print("\n" + "=" * 45)
print(f"🏆 전체 문서 평균 TF-IDF TOP {TOP_N}")
print("=" * 45)

avg_scores = tfidf_matrix.mean(axis=0).A1
top_avg = avg_scores.argsort()[::-1][:TOP_N]
top_overall = [(feature_names[j], avg_scores[j]) for j in top_avg]

for rank, (word, score) in enumerate(top_overall, 1):
    bar = "█" * int(score * 200)
    print(f"  {rank:>2}. {word:<12} {score:.4f}  {bar}")

# 전체 평균 그래프
words  = [w for w, _ in top_overall]
values = [s for _, s in top_overall]

plt.figure(figsize=(13, 6))
bars = plt.bar(words, values, color='steelblue')
plt.bar_label(bars, fmt='%.3f', padding=3, fontsize=9)
plt.title(f"전체 문서 평균 TF-IDF TOP {TOP_N}", fontsize=15, fontproperties=font_prop)
plt.xticks(rotation=30, ha='right', fontproperties=font_prop)
plt.tight_layout()

total_out = os.path.join(FOLDER, "_전체_tfidf.png")
plt.savefig(total_out, dpi=150)
plt.close()
print(f"\n✅ 전체 그래프 저장: {total_out}")
print("\n🎉 분석 완료!")