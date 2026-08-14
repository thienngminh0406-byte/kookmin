# -*- coding: utf-8 -*-
"""
매일 아침 GitHub Actions가 이 스크립트를 실행해서 data/listings.xlsx를
처음부터 새로 만듭니다 (고방 API 크롤링 + 카카오 편의시설 조회).

로컬에서 테스트하려면:
    pip install requests pandas openpyxl
    KAKAO_API_KEY=발급받은키 python scripts/collect_data.py
"""

import json
import math
import os
import re
import sys
import time
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parent.parent
OUTPUT_XLSX = ROOT / "data" / "listings.xlsx"

KOOKMIN_LAT, KOOKMIN_LNG = 37.6099, 127.0165

# 국민대 기준으로 훑을 지역들 (법정동코드 앞 5자리). 필요하면 추가하세요.
DONGLI_CODES = {
    "성북구": "11290",
    "강북구": "11305",
}

# 고방이 쓰는 매물 유형 코드. GOSIWON/ONE_ROOM_TEL/MOTEL은 확인된 값이고,
# 나머지는 실제 존재할 가능성이 있어 시도해보는 값입니다.
# 브라우저 개발자도구 Network 탭에서 실제 값을 다시 확인하면 더 정확해질 수 있어요.
HOUSE_TYPES_WIDE = ["GOSIWON", "ONE_ROOM_TEL", "MOTEL", "SHARE_HOUSE", "ONE_ROOM", "TWO_ROOM"]
HOUSE_TYPES_SAFE = ["GOSIWON", "ONE_ROOM_TEL"]

PAGE_SIZE = 40
RANDOM_SEED = 9410

GOBANG_HOUSES_URL = "https://api.gobang.kr/v2/houses"
GOBANG_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://gobang.kr/",
    "Origin": "https://gobang.kr",
}

TYPE_MAP = {
    "GOSIWON": "고시원", "ONE_ROOM_TEL": "원룸텔", "MOTEL": "모텔",
    "SHARE_HOUSE": "쉐어하우스", "ONE_ROOM": "원룸", "TWO_ROOM": "투룸",
}

KAKAO_CATEGORY_SEARCH_URL = "https://dapi.kakao.com/v2/local/search/category.json"
KAKAO_KEYWORD_SEARCH_URL = "https://dapi.kakao.com/v2/local/search/keyword.json"
SEARCH_RADIUS_M = 500

CATEGORY_TARGETS = {"편의점": "CS2", "카페": "CE7", "마트": "MT1", "병원": "HP8"}
KEYWORD_TARGETS = {"세탁소": "세탁소", "우체국": "우체국"}


def haversine_km(lat1, lng1, lat2, lng2):
    r = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlng = math.radians(lng2 - lng1)
    a = (math.sin(dlat / 2) ** 2
         + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlng / 2) ** 2)
    return r * 2 * math.asin(math.sqrt(a))


def fetch_gobang_page(dongli_code, house_types, page_no, headers):
    params = {
        "dongliCode": dongli_code,
        "houseTypes": house_types,
        "pageNo": page_no,
        "pageSize": PAGE_SIZE,
        "randomSeed": RANDOM_SEED,
    }
    resp = requests.get(GOBANG_HOUSES_URL, headers=headers, params=params, timeout=15)
    return resp


def crawl_gobang():
    all_items = {}
    for region, dongli_code in DONGLI_CODES.items():
        # 넓은 유형 목록으로 먼저 시도하고, 실패하면 확인된 값으로만 재시도
        house_types = HOUSE_TYPES_WIDE
        test = fetch_gobang_page(dongli_code, house_types, 1, GOBANG_HEADERS)
        if test.status_code != 200:
            print(f"[{region}] 넓은 유형 목록 실패({test.status_code}), 기본 유형으로 재시도")
            house_types = HOUSE_TYPES_SAFE

        page_no = 1
        while page_no <= 100:
            resp = fetch_gobang_page(dongli_code, house_types, page_no, GOBANG_HEADERS)
            if resp.status_code != 200:
                print(f"[{region}] {page_no}페이지 실패({resp.status_code})")
                break
            try:
                data = resp.json()
            except json.JSONDecodeError:
                print(f"[{region}] {page_no}페이지 JSON 파싱 실패")
                break

            items = None
            for key in ["content", "data", "items", "list", "houses", "result", "results"]:
                if isinstance(data, dict) and key in data and isinstance(data[key], list):
                    items = data[key]
                    break
            if items is None and isinstance(data, list):
                items = data
            if not items:
                break

            for it in items:
                it["_region"] = region
                key = it.get("id") or it.get("no")
                if key is not None:
                    all_items[key] = it

            print(f"[{region}] {page_no}페이지: {len(items)}건")
            if len(items) < PAGE_SIZE:
                break
            page_no += 1
            time.sleep(0.4)

    return list(all_items.values())


def parse_house_types(raw):
    try:
        types = json.loads(raw) if isinstance(raw, str) else (raw or [])
    except Exception:
        types = []
    labels = [TYPE_MAP.get(t, t) for t in types]
    return "/".join(dict.fromkeys(labels)) if labels else "정보없음"


def parse_tags(raw):
    try:
        tags = json.loads(raw) if isinstance(raw, str) else (raw or [])
    except Exception:
        tags = []
    return [t.get("name", "") for t in tags if isinstance(t, dict)]


def has_tag(names, keyword):
    return any(keyword in n for n in names)


def nearest_subway(raw):
    try:
        subs = json.loads(raw) if isinstance(raw, str) else (raw or [])
    except Exception:
        subs = []
    if not subs:
        return None, None
    nearest = min(subs, key=lambda d: d.get("distance", 999))
    return nearest.get("name"), round(nearest.get("distance", 0) * 1000)


def nearest_by_category(x, y, code, headers):
    params = {"category_group_code": code, "x": x, "y": y, "radius": SEARCH_RADIUS_M, "sort": "distance", "size": 1}
    resp = requests.get(KAKAO_CATEGORY_SEARCH_URL, headers=headers, params=params, timeout=10)
    if resp.status_code != 200:
        return None, None
    docs = resp.json().get("documents", [])
    return (docs[0]["place_name"], docs[0]["distance"]) if docs else (None, None)


def nearest_by_keyword(x, y, keyword, headers):
    params = {"query": keyword, "x": x, "y": y, "radius": SEARCH_RADIUS_M, "sort": "distance", "size": 1}
    resp = requests.get(KAKAO_KEYWORD_SEARCH_URL, headers=headers, params=params, timeout=10)
    if resp.status_code != 200:
        return None, None
    docs = resp.json().get("documents", [])
    return (docs[0]["place_name"], docs[0]["distance"]) if docs else (None, None)


def main():
    api_key = os.environ.get("KAKAO_API_KEY")
    if not api_key:
        print("[오류] 환경변수 KAKAO_API_KEY가 없습니다. GitHub Secrets에 등록되어 있는지 확인하세요.")
        sys.exit(1)
    kakao_headers = {"Authorization": f"KakaoAK {api_key}"}

    print("고방 매물 크롤링 시작...")
    raw_items = crawl_gobang()
    print(f"총 {len(raw_items)}건 수집 (중복 제거 후)")

    if not raw_items:
        print("[오류] 수집된 매물이 없습니다. 스크립트를 중단합니다 (기존 파일을 덮어쓰지 않음).")
        sys.exit(1)

    rows = []
    for idx, r in enumerate(raw_items):
        lat = r.get("latitude")
        lng = r.get("longitude")
        if lat is None or lng is None:
            continue
        dist_km = haversine_km(KOOKMIN_LAT, KOOKMIN_LNG, lat, lng)
        tags = parse_tags(r.get("tags"))
        subway_name, subway_dist_m = nearest_subway(r.get("nearSubways"))

        min_price, max_price = r.get("minPrice"), r.get("maxPrice")
        min_deposit, max_deposit = r.get("minDeposit"), r.get("maxDeposit")

        print(f"[{idx+1}/{len(raw_items)}] {r.get('name')} - 편의시설 조회 중...")
        row = {
            "시설명": r.get("name"),
            "유형": parse_house_types(r.get("houseTypes")),
            "지역": r.get("_region"),
            "거리(km)": round(dist_km, 2),
            "거리추정여부": "N",
            "주소정확여부": "Y",
            "월세최소": min_price,
            "월세최대": max_price,
            "월세원본": None,
            "보증금": (f"{min_deposit}~{max_deposit}" if min_deposit != max_deposit else (min_deposit or "없음")) if min_deposit is not None else "없음",
            "관리비": "정보없음",
            "개인화장실": "O(개인)" if has_tag(tags, "개인화장실") else "정보없음",
            "에어컨": "O(개인)" if has_tag(tags, "개인에어컨") else "정보없음",
            "지하철역": subway_name or "",
            "지하철역거리m": subway_dist_m or "",
            "전화번호": "",
            "주소": r.get("eupmyeondongFullName") or "",
            "비고": (f"{subway_name} 도보권({subway_dist_m}m) · " if subway_name else "") + ", ".join(
                [tg for tg in tags if "화장실" not in tg and "에어컨" not in tg][:6]
            ),
            "고방링크": f"https://gobang.kr/place/{r.get('no')}" if r.get("no") else "",
            "카카오맵링크": "",
            "위도": lat,
            "경도": lng,
        }

        for label, code in CATEGORY_TARGETS.items():
            name, dist = nearest_by_category(lng, lat, code, kakao_headers)
            row[label] = name or ""
            row[f"{label}거리m"] = dist or ""
            time.sleep(0.12)
        for label, kw in KEYWORD_TARGETS.items():
            name, dist = nearest_by_keyword(lng, lat, kw, kakao_headers)
            row[label] = name or ""
            row[f"{label}거리m"] = dist or ""
            time.sleep(0.12)

        rows.append(row)

    df = pd.DataFrame(rows)
    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(OUTPUT_XLSX, index=False)
    print(f"완료: {len(df)}개 매물 -> {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
