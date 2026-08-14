# -*- coding: utf-8 -*-
"""
오피넷 유가 일일 수집기 (GitHub Actions 용)
------------------------------------------------
매일 실행되어 data/oil_history.csv 에 하루치 유가를 누적한다.
저장 항목: 날짜, 지역(전국/경북), 유종, 가격
- 전국 평균: avgAllPrice.do
- 시도 평균: avgSidoPrice.do (경북 코드 47 만 추출)

환경변수 OPINET_KEY 로 API 키를 받는다 (Actions secrets).
누적 CSV 는 마지막 한 달치만 유지(파일 비대화 방지, 필요시 조정).
"""
import os
import sys
import datetime
import requests
import pandas as pd

OPINET_KEY = os.environ.get("OPINET_KEY", "")
BASE = "https://www.opinet.co.kr/api"
HISTORY_PATH = "data/oil_history.csv"

# 오피넷 유종코드 → 표시명
PROD = {
    "B027": "휘발유",
    "D047": "경유",
    "K015": "자동차용LPG",
}
GYEONGBUK_SIDO = "47"      # 경북 시도코드
KEEP_DAYS = 62             # 약 2개월치 유지


def fetch_nationwide():
    """전국 평균 유가 → [(유종명, 가격), ...]"""
    r = requests.get(f"{BASE}/avgAllPrice.do",
                     params={"code": OPINET_KEY, "out": "json"}, timeout=15)
    r.raise_for_status()
    rows = r.json()["RESULT"]["OIL"]
    out = []
    for row in rows:
        code = row.get("PRODCD")
        if code in PROD:
            out.append((PROD[code], float(row["PRICE"])))
    return out


def fetch_gyeongbuk():
    """경북 시도 평균 유가 → [(유종명, 가격), ...]"""
    r = requests.get(f"{BASE}/avgSidoPrice.do",
                     params={"code": OPINET_KEY, "out": "json"}, timeout=15)
    r.raise_for_status()
    rows = r.json()["RESULT"]["OIL"]
    out = []
    for row in rows:
        if row.get("SIDOCD") == GYEONGBUK_SIDO and row.get("PRODCD") in PROD:
            out.append((PROD[row["PRODCD"]], float(row["PRICE"])))
    return out


def main():
    if not OPINET_KEY:
        print("OPINET_KEY 없음. 종료.")
        sys.exit(1)

    today = datetime.date.today().isoformat()
    records = []
    try:
        for name, price in fetch_nationwide():
            records.append({"날짜": today, "지역": "전국", "유종": name, "가격": price})
        for name, price in fetch_gyeongbuk():
            records.append({"날짜": today, "지역": "경북", "유종": name, "가격": price})
    except Exception as e:
        print(f"수집 실패: {e}")
        sys.exit(1)

    if not records:
        print("수집된 데이터 없음.")
        sys.exit(1)

    new_df = pd.DataFrame(records)

    # 기존 누적본과 병합
    if os.path.exists(HISTORY_PATH):
        old = pd.read_csv(HISTORY_PATH)
        # 같은 날짜 중복 제거(재실행 대비)
        old = old[old["날짜"] != today]
        merged = pd.concat([old, new_df], ignore_index=True)
    else:
        merged = new_df

    # 최근 KEEP_DAYS 일만 유지
    merged["날짜"] = merged["날짜"].astype(str)
    cutoff = (datetime.date.today() - datetime.timedelta(days=KEEP_DAYS)).isoformat()
    merged = merged[merged["날짜"] >= cutoff]
    merged = merged.sort_values(["날짜", "지역", "유종"]).reset_index(drop=True)

    os.makedirs("data", exist_ok=True)
    merged.to_csv(HISTORY_PATH, index=False, encoding="utf-8-sig")
    print(f"{today} 수집 완료: {len(new_df)}건 추가, 누적 {len(merged)}건")


if __name__ == "__main__":
    main()
