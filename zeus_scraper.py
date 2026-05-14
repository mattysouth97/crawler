#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ZEUS 장비 통계 자동 수집 (반도체 + 자동차, 전체 지역)
=============================================================

[필수 - 실행 전 1회 확인 (DevTools, 5분)]
  1. Chrome 에서 https://www.zeus.go.kr/stat/equip 접속
  2. F12 → Network 탭 → Fetch/XHR 필터 ON
  3. 업종 필터에서 "반도체" 선택 → 호출되는 요청 1건 클릭
  4. Headers 탭에서:
       - Request URL          → ENDPOINT 에 반영
       - Request Method       → METHOD 에 반영 (POST/GET)
  5. Payload 탭에서:
       - 업종 파라미터 이름     → INDUSTRY_KEY 에 반영
       - 반도체/자동차 실제 코드 → INDUSTRY_CODES 값에 반영 (숫자코드일 수 있음)
       - 페이지 파라미터 이름   → PAGE_KEY, SIZE_KEY 에 반영
  6. Response 탭에서:
       - 리스트가 들어있는 키    → LIST_PATH 에 반영 (예: data.list)
       - 전체건수 키            → TOTAL_PATH 에 반영

[진단 모드 (필드명/응답 구조 먼저 확인)]
  $ python zeus_scraper.py --inspect
    → 반도체 1페이지 raw JSON 을 콘솔에 덤프해서 KEY_MAP 매핑에 활용

[실행]
  $ pip install requests pandas openpyxl
  $ python zeus_scraper.py
"""

import json
import sys
import time

import pandas as pd
import requests

# =============================================================
# ▼▼▼  DevTools 에서 확인 후 이 블록만 수정  ▼▼▼
# =============================================================

BASE_URL = "https://www.zeus.go.kr"
REFERER  = f"{BASE_URL}/stat/equip"

# 1) 실제 AJAX URL / METHOD
ENDPOINT = f"{BASE_URL}/stat/equip/list"   # TODO: DevTools Request URL
METHOD   = "POST"                          # TODO: "GET" 또는 "POST"

# 2) 업종 파라미터
INDUSTRY_KEY = "useIndst"                  # TODO: payload 의 업종 키
INDUSTRY_CODES = {                         # TODO: 실제 코드값으로 교체
    "반도체": "반도체",
    "자동차": "자동차",
}

# 3) 지역 = 전체 (보통 빈 문자열 또는 "00")
REGION_KEY  = "sido"                       # TODO: 지역 키 (없으면 None)
REGION_CODE = ""                           # TODO: 전체에 해당하는 값

# 4) 페이지네이션
PAGE_KEY  = "pageIndex"                    # TODO
SIZE_KEY  = "pageSize"                     # TODO
PAGE_SIZE = 100

# 5) 응답 JSON 경로
LIST_PATH  = ("list",)                     # TODO: 예) ("result","items")
TOTAL_PATH = ("totalCount",)               # TODO: 예) ("paginationInfo","totalRecordCount")

# 6) 응답 필드명 → 한글 컬럼 매핑 (--inspect 로 본 뒤 채워넣기)
KEY_MAP = {
    "equipNm":     "장비명",
    "equipNmKor":  "장비명",
    "mdlNm":       "모델명",
    "mfrNm":       "제조사",
    "useIndstNm":  "사용업종",
    "indstNm":     "사용업종",
    "instlAddr":   "구축위치",
    "instlPlc":    "구축위치",
    "addr":        "구축위치",
    "sidoNm":      "시도",
    "orgNm":       "기관명",
    "instNm":      "기관명",
    "acqsYr":      "구축연도",
    "acqsAmt":     "구축비용",
}

OUTPUT_FILE = "zeus_장비_반도체_자동차.xlsx"

# =============================================================
# ▲▲▲  설정 끝  ▲▲▲
# =============================================================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
    "Referer": REFERER,
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE_URL,
}


def get_nested(d, path):
    cur = d
    for k in path:
        if not isinstance(cur, dict):
            return None
        cur = cur.get(k)
    return cur


def make_session():
    s = requests.Session()
    s.headers.update(HEADERS)
    # 일부 한국 정부 사이트는 JSESSIONID 선발급이 필요
    r = s.get(REFERER, timeout=15)
    r.raise_for_status()
    return s


def fetch_page(session, industry_code, page=1):
    payload = {
        INDUSTRY_KEY: industry_code,
        PAGE_KEY: page,
        SIZE_KEY: PAGE_SIZE,
    }
    if REGION_KEY:
        payload[REGION_KEY] = REGION_CODE

    if METHOD.upper() == "POST":
        r = session.post(ENDPOINT, data=payload, timeout=20)
    else:
        r = session.get(ENDPOINT, params=payload, timeout=20)
    r.raise_for_status()
    try:
        return r.json()
    except json.JSONDecodeError:
        print(f"[!] JSON 파싱 실패. 응답 일부:\n{r.text[:500]}", file=sys.stderr)
        raise


def fetch_all(session, label, code):
    records, page = [], 1
    while True:
        print(f"  - {label}  page {page} ...", end=" ")
        data = fetch_page(session, code, page)
        items = get_nested(data, LIST_PATH) or []
        print(f"{len(items)}건")
        if not items:
            break
        records.extend(items)
        total = get_nested(data, TOTAL_PATH) or 0
        if total and len(records) >= total:
            break
        page += 1
        time.sleep(0.4)
    return records


def normalize(records, label):
    out = []
    for r in records:
        row = {"조회업종": label}
        for k, v in r.items():
            row[KEY_MAP.get(k, k)] = v
        out.append(row)
    return out


def write_excel(rows, path):
    df = pd.DataFrame(rows)

    priority = ["조회업종", "장비명", "사용업종", "구축위치", "시도",
                "기관명", "모델명", "제조사", "구축연도", "구축비용"]
    cols = [c for c in priority if c in df.columns]
    cols += [c for c in df.columns if c not in cols]
    df = df[cols]

    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="ZEUS_장비", index=False)
        ws = w.sheets["ZEUS_장비"]
        for i, col in enumerate(df.columns, 1):
            max_len = max(
                len(str(col)),
                df[col].astype(str).map(len).max() if len(df) else 0,
            )
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(
                max(max_len + 2, 12), 50
            )
    print(f"\n✓ 저장: {path}  ({len(df):,} 건)")


def inspect():
    """첫 페이지 raw JSON 덤프 — 응답 구조 확인용."""
    s = make_session()
    code = next(iter(INDUSTRY_CODES.values()))
    print(f"[inspect] 업종코드='{code}' 1페이지 호출\n")
    data = fetch_page(s, code, 1)
    print(json.dumps(data, ensure_ascii=False, indent=2)[:3000])
    print("\n... (3000자 잘림) ...")


def main():
    if "--inspect" in sys.argv:
        inspect()
        return

    print(f"[ZEUS 수집] 업종={list(INDUSTRY_CODES)}, 지역=전체\n")
    s = make_session()
    rows = []
    for label, code in INDUSTRY_CODES.items():
        recs = fetch_all(s, label, code)
        print(f"  → {label}: 총 {len(recs)} 건")
        rows.extend(normalize(recs, label))
    if not rows:
        print("\n[!] 수집 0건. ENDPOINT/파라미터/응답경로 재확인 필요.")
        sys.exit(1)
    write_excel(rows, OUTPUT_FILE)


if __name__ == "__main__":
    main()
