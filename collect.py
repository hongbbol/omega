#!/usr/bin/env python3
"""
제품 데이터 자동 수집 파이프라인 (캡처 불필요)
================================================
사용자가 상세페이지를 손으로 캡처하지 않아도, 상세페이지의 '이미지에 잠긴'
스펙(Supplement Facts·IFOS/COA 표 등)을 자동으로 긁어 채점용 토큰 초안을 만든다.

루트(검증됨):
  ① requests(브라우저 UA)로 제품 상세페이지 HTML 수집  (JS 렌더 필요 시 --render: playwright)
  ② 상세 본문의 이미지 URL 자동 추출 → harvest/<번호>/ 에 다운로드
  ③ tesseract(kor+eng) OCR 로 이미지 속 텍스트 추출
  ④ 정규식으로 EPA/DHA mg·순도·어종·형태·인증·COA·포장 등 후보값 추출
  ⑤ build_xlsx.py 의 LOOKUPS 토큰으로 매핑 → products_draft.csv (초안·근거·신뢰도 동봉)
  ⑥ 다운로드 이미지는 보존 → Claude 비전(Read)으로 핵심 수치 정밀 재확인

원칙:
  - 출력은 항상 '초안'. products.csv 를 절대 자동 덮어쓰지 않는다(사람/Claude 검수 후 병합).
  - 추정·가정 금지. 못 읽은 값은 빈칸/'미상'으로 남기고 근거 스니펫을 함께 남긴다.

사용법:
  pip install requests beautifulsoup4 lxml playwright
  apt-get install -y tesseract-ocr tesseract-ocr-kor      # OCR
  python -m playwright install --with-deps chromium        # --render 쓸 때만
  python collect.py                 # sources.csv 의 모든 행 수집
  python collect.py 2 3 16          # 특정 번호만
  python collect.py --render 16     # JS 렌더가 필요한 사이트

입력: sources.csv  (헤더: 번호,url,render)  — render 는 1이면 playwright 사용(기본 0)
출력: harvest/<번호>/{NN.jpg, ocr.txt, proposal.json} 와 products_draft.csv
"""
import csv, json, os, re, sys, subprocess
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

SOURCES = "sources.csv"
DRAFT_OUT = "products_draft.csv"
HARVEST = "harvest"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")

# CSV 토큰 스키마(build_xlsx.py LOOKUPS 와 동일해야 함)
FIELDS = ["밀도_EPADHA_mg","순도_pct","비율표기","형태","어종","항산화제","포장","제형",
          "인증","COA","원료사"]

# ---- 키워드 → 토큰 매핑 사전 (라벨/마케팅 이미지에서 흔한 표기) -----------------
KW = {
 "형태":   [(r"\brTG\b|알티지|초임계|re-?esterified", "rTG"),
            (r"\bTG\b|트리글리세라이드", "TG"),
            (r"\bEE\b|에틸에스터|ethyl ?ester", "EE")],
 "어종":   [(r"멸치|정어리|청어|anchovy|sardine|herring", "소형어"),
            (r"미세조류|조류|algae|schizochytrium|DSM", "조류"),
            (r"연어|salmon|명태|폴락|pollack|참치|tuna|고등어|mackerel|혼합어|일반어", "일반어"),
            (r"아마씨|치아씨|flax|들기름|ALA 단독|식물성 ALA", "식물성")],
 "항산화제":[(r"토코페롤|비타민\s?E|tocopherol|아스타잔틴|astaxanthin|로즈마리|rosemary|항산화", "있음")],
 "포장":   [(r"블리스터|개별포장|PVDC|차광|individual|blister", "우수"),
            (r"대용량|1000\s?ml|500\s?ml|투명\s?용기|투명병|pump|펌프형 대용량", "취약"),
            (r"\b병\b|bottle|밀폐|차광병|pump", "보통")],
 "제형":   [(r"캡슐|capsule|연질|소프트젤|softgel", "캡슐"),
            (r"\b스틱\b|sachet|stick포|짜먹", "스틱"),
            (r"\b츄\b|chew|저작|젤리", "츄"),
            (r"대용량.*액상|1000\s?ml", "대용량액상"),
            (r"펌프|드롭|liquid|oil pump", "펌프액상")],
 "인증":   [(r"IFOS|GOED|ISO\s?\d|HACCP|NSF|cGMP|제3자", "있음")],
 "원료사": [(r"Solutex|KD\s?파마|KD ?Pharma|AlaskaOmega|알래스카오메가|DSM|Croda|BASF|Omegatex|로빈슨", "명시")],
}

def fetch_html(url, render=False):
    if render:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch(args=["--no-sandbox", "--ignore-certificate-errors"])
            # 일부 실행환경은 TLS 가로채기 프록시를 거쳐 CA 불신 → 인증서 오류 무시 필요
            ctx = b.new_context(user_agent=UA, locale="ko-KR", ignore_https_errors=True)
            pg = ctx.new_page()
            pg.goto(url, wait_until="networkidle", timeout=45000)
            pg.wait_for_timeout(2500)
            html = pg.content()
            b.close()
            return html
    r = requests.get(url, headers={"User-Agent": UA, "Accept-Language": "ko-KR,ko;q=0.9",
                                   "Referer": f"{urlparse(url).scheme}://{urlparse(url).netloc}/"},
                     timeout=30)
    r.raise_for_status()
    return r.text

# 상세 본문 컨테이너(플랫폼별) → 없으면 업로드/에디터 경로 이미지로 폴백
DETAIL_SEL = ["#prdDetail", "#detail", ".detail", ".goods_description", ".product-detail",
              "[class*=detail]", "#productDetail", ".se-main-container", ".cont_detail"]
IMG_PAT = re.compile(r"(NNEditor|/web/upload|editor|/detail|/product/|/goods/|cdn).*\.(jpg|jpeg|png)", re.I)

def extract_detail_images(html, base):
    soup = BeautifulSoup(html, "lxml")
    container = None
    for sel in DETAIL_SEL:
        container = soup.select_one(sel)
        if container:
            break
    scope = container if container else soup
    urls = []
    for img in scope.find_all("img"):
        src = img.get("ec-data-src") or img.get("data-src") or img.get("src") or ""
        if not src:
            continue
        full = urljoin(base, src)
        if container or IMG_PAT.search(full):
            if not re.search(r"\.gif$|icon|btn_|sprite|logo|blank", full, re.I):
                urls.append(full)
    # 중복 제거(순서 유지)
    seen = set(); out = []
    for u in urls:
        if u not in seen:
            seen.add(u); out.append(u)
    return out

def download(urls, outdir, base):
    os.makedirs(outdir, exist_ok=True)
    saved = []
    for i, u in enumerate(urls, 1):
        try:
            r = requests.get(u, headers={"User-Agent": UA, "Referer": base}, timeout=30)
            if r.ok and len(r.content) > 3000:  # 아이콘급 소형 제외
                p = os.path.join(outdir, f"{i:02d}.jpg")
                open(p, "wb").write(r.content)
                saved.append(p)
        except Exception as e:
            print(f"   ! 다운로드 실패 {u}: {e}")
    return saved

def ocr_dir(paths):
    texts = {}
    for p in paths:
        try:
            out = subprocess.run(["tesseract", p, "-", "-l", "kor+eng"],
                                 capture_output=True, text=True, timeout=60)
            texts[os.path.basename(p)] = out.stdout
        except Exception as e:
            texts[os.path.basename(p)] = f"(OCR 실패: {e})"
    return texts

def extract_fields(text):
    """OCR 전체 텍스트에서 후보 수치/근거를 정규식으로 추출.
    상세페이지엔 비교표·마케팅 수치가 섞여 모호할 수 있으므로, 모호하면 값을
    비우고 '_flags' 에 '비전확인필요'를 남긴다(추정 금지). 정확값은 Claude 비전이 보완."""
    t = text.replace(" ", "")  # OCR 공백 잡음 완화 (수치 인접 매칭용)
    f = {}; flags = []
    # ① 밀도: 정식 Supplement Facts 형태 "EPA n mg + DHA n mg" 를 최우선(비교표와 구분)
    num = r"(\d{1,4}(?:\.\d+)?)"
    sf = (re.search(rf"EPA{num}mg\+DHA{num}mg", t, re.I)        # EPA 먼저
          or re.search(rf"DHA{num}mg\+EPA{num}mg", t, re.I))    # DHA 먼저(역순 표기)
    if sf:
        f["밀도_EPADHA_mg"] = round(float(sf.group(1)) + float(sf.group(2))); f["비율표기"] = "명시"
        f["_ev_밀도"] = f"{sf.group(0)} (Supplement Facts)"
    else:
        ed = re.search(r"EPA\+?DHA(\d{2,4})mg", t, re.I)
        if ed:
            f["밀도_EPADHA_mg"] = int(ed.group(1)); f["비율표기"] = "일부"
            f["_ev_밀도"] = f"EPA+DHA{ed.group(1)}"
        else:
            flags.append("밀도(EPA/DHA mg) 비전확인필요")
    # ② 순도 %: SF 패널의 'EPA+DHA ÷ 총오일'로 계산하면 비교표 오염을 피해 정확.
    #    예) EPA+DHA 110mg / rTG오일 146.67mg = 75%. 총오일 못 찾으면 라벨 표기 순도로 폴백.
    oil = re.search(r"Omega-?3Oil(\d{2,4}(?:\.\d+)?)mg|오일(\d{2,4}(?:\.\d+)?)mg", t, re.I)
    if sf and oil:
        oilmg = float(oil.group(1) or oil.group(2))
        f["순도_pct"] = round(f["밀도_EPADHA_mg"] / oilmg * 100)
        f["_ev_순도"] = f"{f['밀도_EPADHA_mg']}mg/{oilmg}mg오일"
    else:
        purs = sorted(set(int(x) for x in re.findall(r"순도[^0-9]{0,4}(\d{2,3})%", t)))
        if len(purs) == 1:
            f["순도_pct"] = purs[0]
        elif len(purs) > 1:
            flags.append(f"순도 후보 {purs} 충돌→비전확인필요")
    # ③ 키워드 기반 토큰
    for field, rules in KW.items():
        for pat, token in rules:
            if re.search(pat, text, re.I):
                f.setdefault(field, token)
                break
    # ④ 포장: 블리스터(우수)와 통/병(보통)이 동시 등장하면 비교표 혼입 → 모호
    if re.search(r"블리스터|개별포장|PVDC", text, re.I) and re.search(r"통포장|\b병\b|bottle", text, re.I):
        f.pop("포장", None); flags.append("포장(블리스터/통 혼재)→비전확인필요")
    # ⑤ COA: 산패 실측 수치(과산화물/TOTOX/산가)나 중금속표가 보이면 공개
    if re.search(r"TOTOX|Peroxide|과산화물|산가|Total ?Oxidation|중금속|Mercury|Cadmium", text, re.I):
        f["COA"] = "공개"
    f["_flags"] = flags
    return f

def find_spec_image(per_image):
    """Supplement Facts 패널 이미지 파일명을 특정(Claude 비전 재확인용).
    1순위: EPA·DHA mg 수치가 든 정식 성분표. 2순위: 성분/급여량 약한 신호."""
    for name, txt in per_image.items():
        if re.search(r"EPA\s*\d+[^A-Za-z]{0,6}mg|Active Ingredient|Supplement\s*Fact", txt, re.I):
            return name
    for name, txt in per_image.items():
        if re.search(r"per ?capsule|성분표|함량|급여량", txt, re.I):
            return name
    return ""

def map_row(no, url, fields, spec_img):
    row = {"번호": no, "_source": url}
    conf = []
    for k in FIELDS:
        v = fields.get(k, "")
        row[k] = v
        if v != "":
            conf.append(k)
    row["_evidence"] = fields.get("_ev_밀도", "")
    flags = fields.get("_flags", [])
    row["_flags"] = " / ".join(flags)
    row["_spec_image"] = f"{HARVEST}/{no}/{spec_img}" if spec_img else ""
    row["_confidence"] = f"{len(conf)}/{len(FIELDS)} 자동추출" + (f" · 검토필요 {len(flags)}건" if flags else "")
    return row

def collect_one(no, url, render):
    print(f"\n[{no}] {url}  (render={render})")
    out = os.path.join(HARVEST, str(no))
    base = f"{urlparse(url).scheme}://{urlparse(url).netloc}/"
    cached = sorted([os.path.join(out, x) for x in os.listdir(out)
                     if x.endswith(".jpg")]) if os.path.isdir(out) else []
    if cached:
        saved = cached
        print(f"   캐시 이미지 {len(saved)}개 재사용 → {out}/ (재다운로드 생략)")
    else:
        html = fetch_html(url, render)
        imgs = extract_detail_images(html, url)
        print(f"   상세 이미지 {len(imgs)}개 발견")
        saved = download(imgs, out, base)
        print(f"   {len(saved)}개 다운로드 → {out}/")
    texts = ocr_dir(saved)
    full = "\n".join(texts.values())
    open(os.path.join(out, "ocr.txt"), "w", encoding="utf-8").write(full)
    fields = extract_fields(full)
    spec_img = find_spec_image(texts)
    row = map_row(no, url, fields, spec_img)
    json.dump({"fields": fields, "row": row}, open(os.path.join(out, "proposal.json"), "w",
              encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"   → {row['_confidence']}" + (f" | SF이미지: {row['_spec_image']}" if spec_img else ""))
    if row["_flags"]:
        print(f"     ⚠ {row['_flags']}")
    return row

def main(argv):
    render_all = "--render" in argv
    nums = [a for a in argv if a.isdigit()]
    rows = list(csv.DictReader(open(SOURCES, encoding="utf-8-sig")))
    if nums:
        rows = [r for r in rows if r["번호"] in nums]
    if not rows:
        raise SystemExit(f"{SOURCES} 에 수집할 행이 없습니다. (번호 인자 확인)")
    drafts = []
    for r in rows:
        render = render_all or r.get("render", "0").strip() == "1"
        try:
            drafts.append(collect_one(r["번호"], r["url"], render))
        except Exception as e:
            print(f"   !! 실패: {e}")
    if drafts:
        cols = ["번호"] + FIELDS + ["_evidence", "_flags", "_spec_image", "_confidence", "_source"]
        with open(DRAFT_OUT, "w", encoding="utf-8-sig", newline="") as fp:
            w = csv.DictWriter(fp, fieldnames=cols)
            w.writeheader()
            for d in drafts:
                w.writerow({c: d.get(c, "") for c in cols})
        print(f"\n초안 저장: {DRAFT_OUT} ({len(drafts)}행) — 검수 후 products.csv 에 병합하세요.")
        print("핵심 수치는 harvest/<번호>/*.jpg 를 Claude 비전(Read)으로 재확인 권장.")

if __name__ == "__main__":
    if not os.path.exists(SOURCES):
        raise SystemExit(f"{SOURCES} 가 없습니다. '번호,url,render' 헤더로 만들어 주세요.")
    main(sys.argv[1:])
