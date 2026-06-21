#!/usr/bin/env python3
"""
고양이 오메가3 제품 채점 엔진 (portable)
- 입력: products.csv  (제품 데이터 — 이 파일만 편집하면 됨)
- 출력: 고양이_오메가3_제품_채점표.xlsx  (기준 시트 + 점수표 시트, 수식 자동 채점/순위)
        순위_미리보기.md               (Python 재현 점수·순위 스냅샷 — 엑셀 없이 검증/리뷰용)

사용법:
    pip install openpyxl
    python build_xlsx.py
    # 생성된 xlsx를 Excel 또는 LibreOffice로 열면 수식이 자동 계산됩니다.
    # 콘솔과 순위_미리보기.md 로 점수·순위를 엑셀 없이도 확인할 수 있습니다.

새 제품 추가/수정:
    products.csv 에 행을 추가하거나 값을 채우세요.
    카테고리 컬럼은 아래 RUBRIC의 허용 토큰만 사용해야 점수가 매겨집니다.
    숫자 컬럼(밀도/순도)은 숫자만, 모르면 빈칸으로 두세요(=0점 처리, 가정 입력 금지).

설계 메모(검증 가능성):
    채점은 ❶ Excel 수식과 ❷ Python(score_rows)에서 "같은 상수"로 이중 구현된다.
    밀도/순도 구간(DENS_TIERS·PUR_TIERS)과 LOOKUPS는 단일 출처이며,
    수식 문자열과 Python 점수 둘 다 여기서 생성/참조한다(둘이 어긋날 수 없음).
"""
import csv, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

CSV_IN   = "products.csv"
XLSX_OUT = "고양이_오메가3_제품_채점표.xlsx"
MD_OUT   = "순위_미리보기.md"

# ===== RUBRIC (채점 기준) =====================================================
# 상위 가중치: 성분 35 / 산패 30 / 원료·형태 20 / 투명성 15  (합 100)
# 안전 점수상한: 비해양성(ALA단독)→20, 대구간유·비타민A/D첨가→40, 산패고위험→50
#   ※조류(algae) 오일은 EPA/DHA 직접공급 = 정상 / 폴락 어체유는 간유 아님
# 안전 플래그 자동판정(수식): 비해양성=어종'식물성'(ALA단독) / 산패고위험=포장'취약'+항산화제≠'있음'+COA미공개
#   ※대구간유('대구' 열)·비타민A/D첨가('비타민AD' 열)만 CSV 수동. 비해양성·산패고위험 CSV값은 무시되고 수식으로 계산됨.
# 빈칸/미상 처리: 공개 안 한 항목은 해당 세부점수 0 (제외 아님). 가정 입력 금지.
#   ※예외: 형태·어종의 '미상'은 근거매트릭스 정의상 중간점(형태5·어종3) — 불명은 중간, 비공개 벌은 투명성에서.
LOOKUPS = {
    "비율표기": [("명시",5),("일부",2),("미표기",0),("미상",0)],
    "형태":     [("rTG",10),("TG",10),("EE",5),("미상",5)],
    "어종":     [("소형어",10),("일반어",6),("해양포유류",6),("식물성",0),("조류",8),("미상",3)],
    "항산화제": [("있음",10),("없음",0),("미상",0)],
    "포장":     [("우수",10),("보통",5),("취약",0),("미상",0)],
    "제형":     [("캡슐",10),("소프트젤",10),("츄",6),("스틱",6),("펌프액상",6),("파우더",4),("대용량액상",2),("미상",0)],
    "인증":     [("있음",5),("없음",0),("미상",0)],
    "COA":      [("공개",6),("일부",3),("없음",0),("미상",0)],
    "원료사":   [("명시",4),("일부",2),("없음",0),("미상",0)],
}
# 밀도/순도 구간: (하한 임계값, 배점) 내림차순. 임계값 이상이면 해당 배점. 빈칸/비수치→0.
DENS_TIERS = [(300,20),(200,15),(100,10),(0,5)]   # EPA+DHA mg/단위
PUR_TIERS  = [(60,10),(40,6),(0,3)]               # 순도 %
# =============================================================================

HDR=Font(bold=True,color="FFFFFF",name="Arial"); HFILL=PatternFill("solid",fgColor="06B470")
BOLD=Font(bold=True,name="Arial"); NORM=Font(name="Arial")

def num(v):
    try:
        return float(v) if str(v).strip()!="" else None
    except ValueError:
        return None

def lookup(name, token):
    """LOOKUPS 점수표에서 토큰 점수(없으면 0) — Python 검산용."""
    return dict(LOOKUPS[name]).get((token or "").strip(), 0)

def tier_score(value, tiers):
    """구간 점수(Python). value None/비수치→0."""
    if value is None: return 0
    for thr,pts in tiers:
        if value>=thr: return pts
    return 0

def tier_formula(cell, tiers):
    """구간 점수(Excel 수식 조각) — tier_score와 동일 로직을 수식으로 생성."""
    expr=str(tiers[-1][1])  # 최저 구간 점수(가장 안쪽)
    # 낮은 임계값부터 감싸 올려 가장 높은 임계값이 바깥쪽 IF가 되게 한다(평가 순서 보존).
    for thr,pts in reversed(tiers[:-1]):
        expr=f"IF({cell}>={thr},{pts},{expr})"
    return f"IF(NOT(ISNUMBER({cell})),0,{expr})"

def score_rows(rows):
    """Excel 수식과 동일한 채점을 Python으로 재현 (엑셀 없이 검증/순위 산출)."""
    scored=[]
    for r in rows:
        d=num(r["밀도_EPADHA_mg"]); e=num(r["순도_pct"])
        comp   = tier_score(d,DENS_TIERS)+tier_score(e,PUR_TIERS)+lookup("비율표기",r["비율표기"])
        sanpae = lookup("항산화제",r["항산화제"])+lookup("포장",r["포장"])+lookup("제형",r["제형"])
        wonryo = lookup("형태",r["형태"])+lookup("어종",r["어종"])
        trans  = lookup("인증",r["인증"])+lookup("COA",r["COA"])+lookup("원료사",r["원료사"])
        total  = comp+sanpae+wonryo+trans
        nonmar = (r["어종"].strip()=="식물성")                                  # 비해양성 자동
        risk   = (r["포장"].strip()=="취약" and r["항산화제"].strip()!="있음"
                  and r["COA"].strip() in ("없음","미상"))                       # 산패고위험 자동
        cap=min(20 if nonmar else 100, 40 if r["대구"].strip()=="Y" else 100,
                50 if risk else 100, 40 if r["비타민AD"].strip()=="Y" else 100)
        final=min(total,cap)
        scored.append({**r,"_comp":comp,"_sanpae":sanpae,"_wonryo":wonryo,"_trans":trans,
                       "_total":total,"_cap":cap,"_final":final,"_risk":risk,"_nonmar":nonmar})
    # 순위(내림차순, 동점 동순위 — Excel RANK와 동일)
    finals=sorted((x["_final"] for x in scored),reverse=True)
    for x in scored:
        x["_rank"]=finals.index(x["_final"])+1
    return scored

def write_preview_md(scored):
    """순위 스냅샷 md (엑셀 없이 git/리뷰에서 점수 확인)."""
    ordered=sorted(scored,key=lambda x:(x["_rank"],int(x["번호"])))
    lines=["# 점수·순위 미리보기 (자동 생성 — build_xlsx.py)",
           "",
           "> `python build_xlsx.py` 가 매 생성 시 갱신. **수기 편집 금지**(다음 실행에 덮어씀).",
           "> 점수는 Excel 수식과 동일 로직을 Python으로 재현한 값(공개정보 기반·실측 전).",
           "",
           f"제품 {len(scored)}개 · 안전상한 적용 "
           f"{sum(1 for x in scored if x['_cap']<100)}개 · 산패고위험 자동플래그 "
           f"{sum(1 for x in scored if x['_risk'])}개",
           "",
           "| 순위 | 번호 | 제품명 | 성분/35 | 산패/30 | 원료형태/20 | 투명성/15 | 합계 | 안전상한 | 최종 |",
           "|---|---|---|---|---|---|---|---|---|---|"]
    for x in ordered:
        cap="—" if x["_cap"]>=100 else str(x["_cap"])
        lines.append(f"| {x['_rank']} | {x['번호']} | {x['제품명']} | {x['_comp']} | {x['_sanpae']} | "
                     f"{x['_wonryo']} | {x['_trans']} | {x['_total']} | {cap} | {x['_final']} |")
    open(MD_OUT,"w",encoding="utf-8").write("\n".join(lines)+"\n")

def print_preview(scored):
    ordered=sorted(scored,key=lambda x:x["_rank"])
    print(f"\n{'순위':<4}{'번호':<4}{'제품명':<26}{'성35':>5}{'산30':>5}{'원20':>5}{'투15':>5}{'합':>5}{'상한':>5}{'최종':>5}")
    for x in ordered:
        nm=x["제품명"][:24]
        pad=24-sum(2 if ord(c)>0x1100 else 1 for c in nm)
        cap="" if x["_cap"]>=100 else str(x["_cap"])
        print(f"{x['_rank']:<4}{x['번호']:<4}{nm}{' '*max(0,pad)}"
              f"{x['_comp']:>5}{x['_sanpae']:>5}{x['_wonryo']:>5}{x['_trans']:>5}"
              f"{x['_total']:>5}{cap:>5}{x['_final']:>5}")

def build():
    rows=list(csv.DictReader(open(CSV_IN,encoding="utf-8-sig")))
    wb=Workbook()

    # ---- 기준 시트 (룩업 테이블) ----
    r=wb.active; r.title="기준"
    r["A1"]="고양이 오메가3 제품 채점 기준 (공개정보 기반·실측 전)"; r["A1"].font=Font(bold=True,size=13,name="Arial")
    r["A2"]=("가중치 성분35/산패30/원료·형태20/투명성15 | 안전상한 비해양성(ALA)→20·대구간유/비타민A·D→40·산패고위험→50 "
             "| 조류오일=정상, 폴락=어체유(간유 아님)"); r["A2"].font=Font(italic=True,name="Arial")
    RNG={}; row0=4
    for name,table in LOOKUPS.items():
        r.cell(row0-1,2,name).font=BOLD
        for i,(k,v) in enumerate(table):
            r.cell(row0+i,2,k).font=NORM; r.cell(row0+i,3,v).font=NORM
        RNG[name]=f"기준!$B${row0}:$C${row0+len(table)-1}"
        row0+=len(table)+2
    r.column_dimensions["B"].width=16; r.column_dimensions["C"].width=8

    # ---- 점수표 시트 ----
    s=wb.create_sheet("점수표")
    head=["번호","제품명","분류","밀도(EPA+DHA mg/단위)","순도(%)","비율표기","형태","어종","항산화제","포장","제형",
          "인증","COA","원료사","비해양성","대구","산패고위험","비타민AD",
          "성분/35","산패/30","원료형태/20","투명성/15","합계","안전상한","최종점수","잠정순위",
          "정보검증","적합성","데이터충실도","출처·비고"]
    for j,h in enumerate(head,1):
        c=s.cell(1,j,h); c.font=HDR; c.fill=HFILL
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    thin=Side(style="thin",color="DDDDDD"); bd=Border(thin,thin,thin,thin)
    f0=2; n=len(rows)
    for i,row in enumerate(rows):
        rr=f0+i
        vals=[row["번호"],row["제품명"],row["분류"],num(row["밀도_EPADHA_mg"]),num(row["순도_pct"]),
              row["비율표기"],row["형태"],row["어종"],row["항산화제"],row["포장"],row["제형"],
              row["인증"],row["COA"],row["원료사"],row["비해양성"],row["대구"],row["산패고위험"],row["비타민AD"]]
        for j,v in enumerate(vals,1):
            c=s.cell(rr,j,v); c.font=NORM; c.border=bd
            if j>=4: c.alignment=Alignment(horizontal="center")
        # 안전 플래그 자동판정 (CSV의 비해양성/산패고위험 값 대신 계산; 대구·비타민AD는 CSV 수동값 유지)
        s.cell(rr,15,f'=IF(H{rr}="식물성","Y","N")')  # 비해양성 = 식물성 ALA 단독
        s.cell(rr,17,f'=IF(AND(J{rr}="취약",I{rr}<>"있음",OR(M{rr}="없음",M{rr}="미상")),"Y","N")')  # 산패고위험
        # 점수 수식 (밀도/순도 구간은 DENS_TIERS·PUR_TIERS에서 생성 = Python 검산과 동일 출처)
        comp = (f'{tier_formula(f"D{rr}",DENS_TIERS)}'
                f'+{tier_formula(f"E{rr}",PUR_TIERS)}'
                f'+IFERROR(VLOOKUP(F{rr},{RNG["비율표기"]},2,FALSE),0)')
        sanpae=(f'IFERROR(VLOOKUP(I{rr},{RNG["항산화제"]},2,FALSE),0)+IFERROR(VLOOKUP(J{rr},{RNG["포장"]},2,FALSE),0)'
                f'+IFERROR(VLOOKUP(K{rr},{RNG["제형"]},2,FALSE),0)')
        wonryo=(f'IFERROR(VLOOKUP(G{rr},{RNG["형태"]},2,FALSE),0)+IFERROR(VLOOKUP(H{rr},{RNG["어종"]},2,FALSE),0)')
        trans =(f'IFERROR(VLOOKUP(L{rr},{RNG["인증"]},2,FALSE),0)+IFERROR(VLOOKUP(M{rr},{RNG["COA"]},2,FALSE),0)'
                f'+IFERROR(VLOOKUP(N{rr},{RNG["원료사"]},2,FALSE),0)')
        # 비타민AD 는 col R(18) CSV 수동값. 점수 컬럼은 S(19)부터.
        s.cell(rr,19,"="+comp); s.cell(rr,20,"="+sanpae); s.cell(rr,21,"="+wonryo); s.cell(rr,22,"="+trans)
        s.cell(rr,23,f"=SUM(S{rr}:V{rr})")
        s.cell(rr,24,f'=MIN(IF(O{rr}="Y",20,100),IF(P{rr}="Y",40,100),IF(Q{rr}="Y",50,100),IF(R{rr}="Y",40,100))')
        s.cell(rr,25,f"=MIN(W{rr},X{rr})")
        s.cell(rr,26,f"=RANK(Y{rr},$Y${f0}:$Y${f0+n-1},0)")
        s.cell(rr,27,row["정보검증"]); s.cell(rr,28,row["적합성"])
        # 데이터충실도: 채워진 항목 수/11. 빈칸·'미상'은 미입력으로 간주.
        comp_cnt=("+".join([f'IF(ISNUMBER(D{rr}),1,0)',f'IF(ISNUMBER(E{rr}),1,0)']
                  +[f'IF(AND({col}{rr}<>"미상",{col}{rr}<>""),1,0)' for col in "FGHIJKLMN"]))
        s.cell(rr,29,f'=({comp_cnt})&"/11"')
        s.cell(rr,30,row["출처비고"])
        for j in range(19,31):
            c=s.cell(rr,j); c.font=NORM; c.border=bd
            if j<29: c.alignment=Alignment(horizontal="center")
    widths={"A":5,"B":26,"C":13,"D":13,"E":8,"F":9,"G":7,"H":8,"I":9,"J":8,"K":11,"L":7,"M":7,"N":8,
            "O":8,"P":6,"Q":10,"R":9,"S":8,"T":8,"U":11,"V":10,"W":7,"X":9,"Y":9,"Z":9,"AA":11,"AB":9,"AC":11,"AD":44}
    for k,v in widths.items(): s.column_dimensions[k].width=v
    s.freeze_panes="D2"; s.row_dimensions[1].height=42
    fill=PatternFill("solid",fgColor="EAF7F0")
    for rr in range(f0,f0+n):
        for j in range(19,27): s.cell(rr,j).fill=fill
    wb.save(XLSX_OUT)

    # ---- Python 검산 + 미리보기 산출 ----
    scored=score_rows(rows)
    write_preview_md(scored)
    print(f"생성 완료: {XLSX_OUT}  (제품 {n}개)")
    print(f"순위 스냅샷: {MD_OUT}  (엑셀 없이 점수 확인 가능)")
    print_preview(scored)
    print("\n→ xlsx를 Excel/LibreOffice로 열면 동일 수식이 자동 계산됩니다.")

if __name__=="__main__":
    if not os.path.exists(CSV_IN):
        raise SystemExit(f"{CSV_IN} 가 없습니다. 같은 폴더에 두고 실행하세요.")
    build()
