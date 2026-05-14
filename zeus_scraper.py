#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ZEUS 등록장비 Excel 내보내기 — 사이트 공식 '엑셀자료 다운로드' 기능 사용
=====================================================================
HTML 스크래핑 없음. 사이트가 제공하는 내보내기 기능을 그대로 호출합니다:
  1) GET  /search?...&category=PRODUCT&subCategory=EQUIP   → 총 건수 확인
  2) POST /search/equip/excel/create?<검색조건>            → searchFileHistSeq
  3) GET  /search/equip/excel/down?searchFileHistSeq=<seq> → .xlsx 다운로드

본인 로그인 세션으로만 동작 — 우회/위장 없음:
  - 단일 User-Agent, 요청 간 고정 간격
  - 403 / 로그인 리다이렉트 = 즉시 중단 (재시도로 뚫지 않음)
  - 1회 내보내기 상한 5,000건(사이트 제한) → 초과 시 자동 분할 후 병합

[중요 — ZEUS 세션은 수 분 내 만료됩니다]
  쿠키를 복사한 '직후' 바로 실행하세요. 'fn_login' / 로그인 페이지 메시지가
  나오면 .zeus_session 을 다시 복사해 재실행하면 됩니다.

[준비]
  1. 브라우저에서 https://www.zeus.go.kr 로그인 → 등록장비 검색이 보이는 상태로 둠
  2. F12 → Network → 아무 요청 클릭 → Request Headers 의 Cookie 값 전체 복사
  3. 이 폴더에 .zeus_session 파일로 저장 (한 줄). 이미 .gitignore 처리됨 — 커밋 금지.

[실행]
  pip install requests pandas openpyxl
  python zeus_scraper.py
"""

import re
import sys
import time
from pathlib import Path
from urllib.parse import urlencode

import pandas as pd
import requests

# ============================ 설정 ============================

BASE_URL = "https://www.zeus.go.kr"

# 수집 대상: {엑셀 시트명: 검색어}
SEARCHES = {
    "반도체": "반도체 장비",
    "자동차": "자동차 장비",
}

# 검색 조건 (등록장비 / 전체 지역) — DevTools 캡처와 동일
SEARCH_PARAMS = {
    "category": "PRODUCT",
    "subCategory": "EQUIP",
    "nfecNo": "", "stTakeDt": "", "edTakeDt": "",
    "includeKwd": "", "exclusiveKwd": "",
    "ntisEqStPrice": "0", "ntisEqEdPrice": "0",
    "sort": "d", "srchFd": "all",
    "orKwd": "false", "alwaysKwd": "false", "optionOnOff": "false",
    "startDate": "", "endDate": "", "date": "all",
    "eqAreaCd": "",            # 빈 값 = 전체 지역
    "pageNum": "1", "imageRowNo": "0",
    "reSrchFlag": "false", "kolasYn": "",
    "locMapX": "", "locMapY": "",
    "customWhere": "",
}

# 내보낼 컬럼 (사이트 엑셀 다운로드 폼의 selectColumn 값 — 전체 선택)
EXPORT_COLUMNS = [
    "takeDtYyyy", "korNm", "engNm", "equipNo", "manufacture", "madeNm",
    "modelNm", "organNm", "equipNm", "takeDt", "takePrc", "useScopeNm",
    "idleDisuseNm", "setupNm", "detail", "subjectNm", "busiNm",
]

# 요청 목적 구분 코드 — 브라우저에서 본인이 선택한 값. 실제 목적에 맞게 설정하세요.
# (이 스크립트는 실행할 때마다 이 목적과 동의를 본인 명의로 제출합니다.)
REQ_REASON_CD = "5"

CHUNK_SIZE    = 5000      # 사이트 1회 내보내기 상한
REQUEST_DELAY = 1.5       # 요청 간 고정 간격(초). 서버가 버거워하면 늘리세요.

SESSION_FILE = Path(".zeus_session")
OUTPUT_FILE  = "zeus_등록장비_반도체_자동차.xlsx"
RAW_DIR      = Path("zeus_raw")     # 사이트가 내려준 원본 xlsx 보관

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")

# ===============================================================


def _stop(msg: str):
    print(f"\n[중단] {msg}", file=sys.stderr)
    sys.exit(1)


def load_session_cookie() -> str:
    if not SESSION_FILE.exists():
        _stop(f"{SESSION_FILE} 없음. 로그인 후 브라우저의 Cookie 값을 "
              f"이 파일에 붙여넣으세요. (docstring 의 '준비' 참고)")
    raw = SESSION_FILE.read_text(encoding="utf-8").strip()
    if not raw:
        _stop(f"{SESSION_FILE} 가 비어 있습니다.")
    return raw


def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": UA,
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
        "Cookie": load_session_cookie(),
    })
    return s


def _is_login(resp: requests.Response) -> bool:
    if "nsso" in resp.url or "login_check" in resp.url:
        return True
    ctype = resp.headers.get("Content-Type", "")
    return "text/html" in ctype and "연동시스템 로그인" in resp.text


def _guard(resp: requests.Response, what: str):
    """차단/만료 신호는 뚫지 않고 즉시 중단한다."""
    if resp.status_code in (401, 403, 429):
        _stop(f"{what}: HTTP {resp.status_code} — 서버가 요청을 거부했습니다.\n"
              f"       세션 만료 시 .zeus_session 갱신 후 재실행하세요.")
    if resp.status_code != 200:
        _stop(f"{what}: 예상치 못한 HTTP {resp.status_code}")
    if _is_login(resp):
        _stop(f"{what}: 로그인 페이지가 반환됐습니다.\n"
              f"       ZEUS 세션은 수 분 내 만료됩니다 — .zeus_session 을 다시 "
              f"복사해 바로 재실행하세요.")


def _request(session: requests.Session, method: str, url: str, what: str, **kw):
    """네트워크 오류 / 5xx 에만 지수 백오프 재시도. 4xx/로그인은 _guard 가 처리."""
    last = None
    for attempt in range(3):
        try:
            r = session.request(method, url, timeout=60, **kw)
            if r.status_code >= 500:
                last = f"HTTP {r.status_code}"
                time.sleep(2 ** attempt)
                continue
            return r
        except requests.RequestException as e:
            last = repr(e)
            time.sleep(2 ** attempt)
    _stop(f"{what}: 네트워크 오류 — {last}")


def _search_query(keyword: str) -> str:
    p = dict(SEARCH_PARAMS)
    p["keyword"] = keyword
    p["preKwd"] = keyword
    return urlencode(p)


def get_total(session: requests.Session, keyword: str) -> int:
    url = f"{BASE_URL}/search?{_search_query(keyword)}"
    r = _request(session, "GET", url, "총 건수 조회")
    _guard(r, "총 건수 조회")
    m = re.search(r'var\s+searchTotal\s*=\s*"(\d+)"', r.text)
    if not m:
        _stop("총 건수 조회: 페이지에서 searchTotal 을 찾지 못했습니다. "
              "검색 조건이나 페이지 구조가 바뀌었을 수 있습니다.")
    return int(m.group(1))


def create_export(session: requests.Session, keyword: str, start: int, end: int) -> str:
    q = _search_query(keyword)
    url = f"{BASE_URL}/search/equip/excel/create?{q}"
    body = (
        [("equipIds", ""), ("targetTypeCd", "B"), ("selColAllCheck", "Y")]
        + [("selectColumn", c) for c in EXPORT_COLUMNS]
        + [("startNo", str(start)), ("endNo", str(end)),
           ("reqReasonCd", REQ_REASON_CD), ("agree", "on"), ("agree", "on")]
    )
    headers = {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": BASE_URL,
        "Referer": f"{BASE_URL}/search?{q}",
        "Accept": "text/plain, */*; q=0.01",
    }
    r = _request(session, "POST", url, "엑셀 생성 요청", data=body, headers=headers)
    _guard(r, "엑셀 생성 요청")
    seq = r.text.strip()
    if seq == "0":
        _stop("엑셀 생성 요청: 서버가 0(실패)을 반환했습니다. "
              "검색 조건 또는 다운로드 권한을 확인하세요.")
    if not seq.isdigit():
        _stop(f"엑셀 생성 요청: 예상치 못한 응답 — {seq[:200]!r}")
    return seq


def download_export(session: requests.Session, seq: str, dest: Path) -> Path:
    url = f"{BASE_URL}/search/equip/excel/down?searchFileHistSeq={seq}"
    r = _request(session, "GET", url, "엑셀 다운로드")
    _guard(r, "엑셀 다운로드")
    if r.content[:4] != b"PK\x03\x04":
        _stop(f"엑셀 다운로드: xlsx 형식이 아닙니다 (응답 {len(r.content)} bytes). "
              f"세션이 만료됐을 수 있습니다 — .zeus_session 갱신 후 재실행.")
    dest.write_bytes(r.content)
    return dest


def main():
    RAW_DIR.mkdir(exist_ok=True)
    session = make_session()
    sheets: dict[str, pd.DataFrame] = {}

    for label, keyword in SEARCHES.items():
        print(f"\n[{label}] 검색어='{keyword}'")
        total = get_total(session, keyword)
        print(f"  총 {total:,}건")
        if total == 0:
            print("  → 0건, 건너뜀")
            continue

        parts, start = [], 1
        while start <= total:
            end = min(start + CHUNK_SIZE - 1, total)
            print(f"  - {start:,}~{end:,} 생성 요청 ...", end=" ", flush=True)
            seq = create_export(session, keyword, start, end)
            print(f"seq={seq}", end=" ", flush=True)
            time.sleep(REQUEST_DELAY)

            raw = RAW_DIR / f"{label}_{start}-{end}.xlsx"
            download_export(session, seq, raw)
            df = pd.read_excel(raw)
            parts.append(df)
            print(f"→ {len(df):,}행  ({raw})")

            start = end + 1
            time.sleep(REQUEST_DELAY)

        sheets[label] = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()

    if not sheets:
        _stop("수집 결과가 없습니다.")

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        for label, df in sheets.items():
            sheet = label[:31]
            df.to_excel(w, sheet_name=sheet, index=False)
            ws = w.sheets[sheet]
            for i, col in enumerate(df.columns, 1):
                dmax = df[col].astype(str).map(len).max() if len(df) else 0
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = (
                    min(max(len(str(col)), dmax) + 2, 60)
                )

    total_rows = sum(len(d) for d in sheets.values())
    print(f"\n✓ 저장: {OUTPUT_FILE}  (시트 {len(sheets)}개, 총 {total_rows:,}행)")
    print(f"  원본 파일: {RAW_DIR}/")


if __name__ == "__main__":
    main()
