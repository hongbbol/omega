#!/usr/bin/env python3
"""
핏펫몰(fitpetmall.com) 상품 검색 헬퍼 — collect.py 용 URL 발굴기
=================================================================
핏펫몰은 Next.js SPA라 상품 URL이 정적 HTML/검색엔진에 안 나온다. 그러나 백엔드
GraphQL(`api.fitpetmall.com/mall/graphql`)이 익명 쿼리를 허용하므로, 키워드로
상품 ID·이름을 받아 표준 상품 URL(`/mall/goods/<번호>`)을 만들 수 있다.

사용:
    python fitpetmall_search.py 오메가3
    python fitpetmall_search.py "연어오일" --pet CAT
출력된 URL을 sources.csv 에 (render=1 로) 넣고 `python collect.py <번호>` 실행.

주의: 상세 스펙 이미지는 lazy-load라 collect.py 의 --render(스크롤 포함)로 받는다.
      일부 상품은 상세 이미지가 빈약할 수 있다(공개정보 자체가 적은 경우).
"""
import sys, json, base64, urllib.request

API = "https://api.fitpetmall.com/mall/graphql"
H = {"Content-Type": "application/json", "Origin": "https://www.fitpetmall.com",
     "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0 Safari/537.36"}

def gq(query, variables):
    data = json.dumps({"query": query, "variables": variables}).encode()
    req = urllib.request.Request(API, data=data, headers=H)
    return json.load(urllib.request.urlopen(req, timeout=25))

def search_ids(text, pet):
    q = "query($t:String!,$p:String!){searchProducts(text:$t,petType:$p){edges{node{id}}}}"
    return [e["node"]["id"] for e in gq(q, {"t": text, "p": pet})["data"]["searchProducts"]["edges"]]

def product(gid):
    q = "query($id:ID!){product(id:$id){name brand{name} isSoldOut}}"
    p = gq(q, {"id": gid})["data"]["product"]
    num = base64.b64decode(gid).decode().split(":")[1]  # ProductType:<번호>
    return num, p

def main(argv):
    if not argv:
        raise SystemExit("사용: python fitpetmall_search.py <키워드> [--pet DOG|CAT]")
    pet = "DOG"
    if "--pet" in argv:
        i = argv.index("--pet"); pet = argv[i + 1]; argv = argv[:i] + argv[i + 2:]
    keyword = " ".join(argv)
    seen = {}
    for p in ([pet] if pet != "ALL" else ["DOG", "CAT"]):
        for gid in search_ids(keyword, p):
            num, info = product(gid)
            if num not in seen and info:
                seen[num] = info
    print(f"'{keyword}' 검색 결과 {len(seen)}건\n")
    for num, info in seen.items():
        brand = info["brand"]["name"] if info.get("brand") else ""
        sold = " [품절]" if info.get("isSoldOut") else ""
        print(f"https://www.fitpetmall.com/mall/goods/{num}  [{brand}] {info['name']}{sold}")

if __name__ == "__main__":
    main(sys.argv[1:])
