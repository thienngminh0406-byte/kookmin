# -*- coding: utf-8 -*-
"""
data/listings.xlsx 를 읽어서 docs/data.json 을 만드는 스크립트.

GitHub Actions가 이 스크립트를 자동으로 실행합니다.
로컬에서 테스트하고 싶으면:
    pip install pandas openpyxl
    python scripts/build_data.py
"""

import json
import math
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
INPUT_XLSX = ROOT / "data" / "listings.xlsx"
OUTPUT_JSON = ROOT / "docs" / "data.json"

AMENITY_COLUMNS = [
    ("편의점", "🏪"),
    ("카페", "☕"),
    ("마트", "🛒"),
    ("병원", "🏥"),
    ("세탁소", "🧺"),
    ("우체국", "📮"),
]


def clean(v):
    """빈 값/NaN을 None으로, 그 외는 그대로."""
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, str) and v.strip() == "":
        return None
    return v


def to_num(v):
    v = clean(v)
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def to_bool(v, default=False):
    v = clean(v)
    if v is None:
        return default
    return str(v).strip().upper() in ("Y", "YES", "TRUE", "1")


def build_amenities(row):
    amenities = []
    for label, icon in AMENITY_COLUMNS:
        name = clean(row.get(label))
        if not name:
            continue
        dist = clean(row.get(f"{label}거리m"))
        if dist is not None:
            try:
                dist_str = f"{int(float(dist))}m"
            except (TypeError, ValueError):
                dist_str = str(dist)
            amenities.append(f"{icon} {name} {dist_str}")
        else:
            amenities.append(f"{icon} {name}")
    return amenities


def main():
    if not INPUT_XLSX.exists():
        print(f"[오류] {INPUT_XLSX} 파일이 없습니다.")
        sys.exit(1)

    df = pd.read_excel(INPUT_XLSX)
    required = ["시설명", "유형", "지역", "거리(km)", "주소"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"[오류] 필수 컬럼이 없습니다: {missing}")
        sys.exit(1)

    records = []
    for _, row in df.iterrows():
        rent_min = to_num(row.get("월세최소"))
        rent_max = to_num(row.get("월세최대"))
        rec = {
            "name": clean(row.get("시설명")),
            "type": clean(row.get("유형")),
            "region": clean(row.get("지역")),
            "distanceKm": to_num(row.get("거리(km)")),
            "distanceEstimated": to_bool(row.get("거리추정여부")),
            "addressPrecise": to_bool(row.get("주소정확여부"), default=True),
            "rentMin": rent_min,
            "rentMax": rent_max,
            "rentRaw": clean(row.get("월세원본")) or (
                str(int(rent_min)) if rent_min is not None and rent_min == rent_max
                else (f"{rent_min:g}~{rent_max:g}" if rent_min is not None and rent_max is not None else None)
            ),
            "deposit": clean(row.get("보증금")),
            "maintenance": clean(row.get("관리비")),
            "privateBathroom": clean(row.get("개인화장실")),
            "ac": clean(row.get("에어컨")),
            "subwayName": clean(row.get("지하철역")),
            "subwayDistanceM": to_num(row.get("지하철역거리m")),
            "phone": clean(row.get("전화번호")),
            "address": clean(row.get("주소")),
            "notes": clean(row.get("비고")),
            "gobangUrl": clean(row.get("고방링크")),
            "kakaoUrl": clean(row.get("카카오맵링크")),
            "lat": to_num(row.get("위도")),
            "lng": to_num(row.get("경도")),
            "amenities": build_amenities(row),
        }
        if not rec["name"]:
            continue  # 빈 행 스킵
        records.append(rec)

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=0)

    print(f"완료: {len(records)}개 매물 -> {OUTPUT_JSON}")


if __name__ == "__main__":
    main()
