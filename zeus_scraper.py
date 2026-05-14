#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ZEUS 장비 통계 수집 — 인증 세션 기반 (polite scraper)
=====================================================

본인 로그인 계정의 세션으로만 동작합니다. 우회/위장 기능 없음:
  - 단일 정직한 User-Agent (로테이션 없음)
  - 고정 요청 간격 (서버 부하 최소화)
  - 403 / 429 / 로그인 리다이렉트 = 즉시 중단 (재시도로 뚫지 않음)
  - 중단 시 JSONL 체크포인트로 정확히 재개

[준비 — 로그인 상태에서 DevTools, 1회]
  1. Chrome 으로 https://www.zeus.go.kr/stat/equip 에 로그인
  2. F12 → Network 탭 → Fetch/XHR 필터
  3. 업종 "반도체" 필터 클릭 → 목록 XHR 요청 1건 선택
       - Request URL / Method        → ENDPOINT / METHOD
       - Payload 의 업종·페이지 키    → INDUSTRY_KEY / PAGE_KEY / SIZE_KEY
       - Response 의 목록·총건수 키   → LIST_PATH / TOTAL_PATH
  4. 같은 요청의 Request Headers 에서 Cookie 값 전체 복사
       → 파일  .zeus_session  에 한 줄로 붙여넣기
       (이 파일은 .gitignore 처리됨 — 절대 커밋 금지)

[실행]
  $ pip install requests pandas openpyxl
  $ python zeus_scraper.py --inspect     # 첫 페이지 raw JSON 확인
  $ python zeus_scraper.py               # 수집
  $ python zeus_scraper.py --fresh       # 체크포인트 무시하고 처음부터
"""

import json
import sys
import time
from pathlib import Path

import pandas as pd
import requests

# =============================================================
# ▼▼▼  DevTools(로그인 상태)에서 확인 후 이 블록만 수정  ▼▼▼
# =============================================================

BASE_URL = "https://www.zeus.go.kr"
REFERER  = f"{BASE_URL}/stat/equip"

ENDPOINT = f"{BASE_URL}/stat/equip/list"   # TODO: 실제 목록 XHR URL
METHOD   = "POST"                          # TODO: "GET" 또는 "POST"

INDUSTRY_KEY = "useIndst"                  # TODO: payload 의 업종 키
INDUSTRY_CODES = {                         # TODO: 실제 코드값으로 교체
    "반도체": "반도체",
    "자동차": "자동차",
}

REGION_KEY  = "sido"                       # TODO: 지역 키 (없으면 None)
REGION_CODE = ""                           # TODO: 전체에 해당하는 값

PAGE_KEY  = "pageIndex"                    # TODO
SIZE_KEY  = "pageSize"                     # TODO
PAGE_SIZE = 100

LIST_PATH  = ("list",)                     # TODO: 예) ("result", "items")
TOTAL_PATH = ("totalCount",)               # TODO: 예) ("paginationInfo", "totalRecordCount")

# 본인 브라우저의 실제 User-Agent 로 교체 권장 (navigator.userAgent)
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)

REQUEST_DELAY = 1.0   # 페이지 간 고정 간격(초). 서버가 버거워하면 늘리세요.

# =============================================================
# ▲▲▲  설정 끝  ▲▲▲
# =============================================================

SESSION_FILE    = Path(".zeus_session")
CHECKPOINT_FILE = Path(".zeus_checkpoint.jsonl")
OUTPUT_FILE     = "zeus_장비_반도체_자동차.xlsx"

KEY_HINTS = {
    "장비명":   ["equipnm", "equipname", "facnm", "equpnm"],
    "사용업종": ["useindst", "indstr", "indst", "industry"],
    "구축위치": ["instlplc", "instladdr", "addr", "location"],
    "시도":     ["sido", "ctprv"],
    "기관명":   ["orgnm", "instnm", "organm", "ownnm"],
    "모델명":   ["mdlnm", "model"],
    "제조사":   ["mfrnm", "manuf", "maker", "makernm"],
    "구축연도": ["acqsyr", "buyyr", "year"],
    "구축비용": ["acqsamt", "amt", "price", "buyamt"],
}


def get_nested(d, path):
    cur = d
    for k in path:
        if not isinstance(cur, dict):
            return None
        cur = cur.get(k)
    return cur


def load_session_cookie() -> str:
    if not SESSION_FILE.exists():
        print(f"[!] {SESSION_FILE} 없음. 로그인 후 DevTools 에서 Cookie 값을 "
              f"이 파일에 붙여넣으세요. (docstring 의 '준비' 참고)", file=sys.stderr)
        sys.exit(1)
    raw = SESSION_FILE.read_text(encoding="utf-8").strip()
    if not raw:
        print(f"[!] {SESSION_FILE} 가 비어 있습니다.", file=sys.stderr)
        sys.exit(1)
    return raw


def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": USER_AGENT,
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
        "Referer": REFERER,
        "X-Requested-With": "XMLHttpRequest",
        "Origin": BASE_URL,
        "Cookie": load_session_cookie(),
    })
    return s


def _looks_like_login(resp: requests.Response) -> bool:
    if "nsso" in resp.url or "login" in resp.url.lower():
        return True
    ctype = resp.headers.get("Content-Type", "")
    if "text/html" in ctype and ("로그인" in resp.text or "login" in resp.text.lower()):
        return True
    return False


def fetch_page(session: requests.Session, industry_code, page: int = 1):
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

    # 차단/만료 신호는 뚫지 않고 즉시 중단한다.
    if r.status_code in (401, 403, 429):
        print(f"\n[중단] HTTP {r.status_code} — 서버가 요청을 거부했습니다.", file=sys.stderr)
        if r.status_code == 429:
            print("       요청 간격(REQUEST_DELAY)을 늘린 뒤 다시 시도하세요.", file=sys.stderr)
        else:
            print("       세션이 만료됐을 수 있습니다. 다시 로그인해 .zeus_session 을 갱신하세요.",
                  file=sys.stderr)
        sys.exit(1)
    if r.status_code != 200:
        print(f"\n[중단] 예상치 못한 HTTP {r.status_code}", file=sys.stderr)
        sys.exit(1)
    if _looks_like_login(r):
        print("\n[중단] 로그인 페이지가 반환됐습니다. .zeus_session 쿠키를 갱신하세요.",
              file=sys.stderr)
        sys.exit(1)

    try:
        return r.json()
    except json.JSONDecodeError:
        print(f"\n[중단] JSON 파싱 실패. 응답 일부:\n{r.text[:500]}", file=sys.stderr)
        sys.exit(1)


def normalize_record(rec: dict, label: str) -> dict:
    row = {"조회업종": label}
    for k, v in rec.items():
        kl = k.lower().replace("_", "")
        mapped = None
        for kor, hints in KEY_HINTS.items():
            if any(h in kl for h in hints):
                mapped = kor
                break
        row[mapped or k] = v
    return row


def load_checkpoint():
    if not CHECKPOINT_FILE.exists():
        return [], set()
    rows, done = [], set()
    for line in CHECKPOINT_FILE.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        e = json.loads(line)
        rows.append(e["row"])
        done.add((e["label"], e["page"]))
    return rows, done


def append_checkpoint(label: str, page: int, rows: list):
    with CHECKPOINT_FILE.open("a", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps({"label": label, "page": page, "row": r},
                               ensure_ascii=False) + "\n")


def fetch_all(session, label, code, rows, done):
    page = 1
    while True:
        if (label, page) in done:
            page += 1
            continue

        print(f"  - {label}  page {page} ...", end=" ", flush=True)
        data = fetch_page(session, code, page)
        items = get_nested(data, LIST_PATH) or []
        print(f"{len(items)}건")

        if not items:
            break

        page_rows = [normalize_record(it, label) for it in items]
        rows.extend(page_rows)
        append_checkpoint(label, page, page_rows)

        total = get_nested(data, TOTAL_PATH) or 0
        collected = sum(1 for r in rows if r["조회업종"] == label)
        if total and collected >= total:
            break
        if len(items) < PAGE_SIZE:
            break

        page += 1
        time.sleep(REQUEST_DELAY)


def write_excel(rows: list, path: str):
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
            data_max = df[col].astype(str).map(len).max() if len(df) else 0
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = (
                min(max(len(str(col)), data_max) + 2, 50)
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
    args = set(sys.argv[1:])

    if "--inspect" in args:
        inspect()
        return

    if "--fresh" in args:
        CHECKPOINT_FILE.unlink(missing_ok=True)

    print(f"[ZEUS 수집] 업종={list(INDUSTRY_CODES)}, 지역=전체\n")
    s = make_session()

    rows, done = load_checkpoint()
    if rows:
        print(f"  ↺ 체크포인트 복원: {len(rows)}건\n")

    for label, code in INDUSTRY_CODES.items():
        fetch_all(s, label, code, rows, done)
        print(f"  → {label}: 누적 {sum(1 for r in rows if r['조회업종'] == label)} 건")

    if not rows:
        print("\n[!] 수집 0건. ENDPOINT/파라미터/응답경로(LIST_PATH) 재확인 필요.")
        sys.exit(1)

    write_excel(rows, OUTPUT_FILE)


if __name__ == "__main__":
    main()
