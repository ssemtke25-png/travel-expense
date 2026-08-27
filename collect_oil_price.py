# -*- coding: utf-8 -*-
"""
오피넷 유가 일일 수집기 (GitHub Actions 용)
------------------------------------------------
매일 새벽(KST 06:00) 실행되어 data/oil_history.csv 에 '전날' 확정 유가를 누적한다.

왜 전날인가:
  오피넷 avgSidoPrice.do 는 '호출 시점의 잠정 평균'을 반환한다. 낮에 호출하면
  그날 장중 값이라 마감 후 확정값과 미세하게 다르다(예: 장중 1854.47 → 확정 1854.52).
  새벽 06:00(KST)에 호출하면 그 값은 사실상 '전날 하루가 마감된 확정 평균'에 가깝다.
  따라서 이 값을 '전날(어제)' 날짜로 라벨링해 저장한다. 그래야 앱에서
  '8/26 출장 → 8/26 유가' 조회가 오피넷 국내유가통계 확정값과 일치한다.

저장 항목: 날짜, 지역(전국/경북), 유종, 가격
  - 전국 평균: avgAllPrice.do
  - 시도 평균: avgSidoPrice.do (경북만 SIDONM 으로 추출)
환경변수 OPINET_KEY 로 API 키를 받는다 (Actions secrets).
누적 CSV 는 최근 KEEP_DAYS 일치만 유지(파일 비대화 방지).
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

# avgSidoPrice 는 자체 순번 코드(전국00, 서울01...)를 쓴다.
# 행정표준코드(경북47)와 다르므로, 코드 대신 시도명(SIDONM)으로 매칭한다.
GYEONGBUK_NAMES = {"경북", "경상북도"}
KEEP_DAYS = 62             # 약 2개월치 유지

# KST 기준으로 '수집 대상 날짜'(= 어제)를 정한다.
# GitHub Actions 러너는 UTC 이므로 KST(UTC+9)로 변환 후 하루를 뺀다.
KST = datetime.timezone(datetime.timedelta(hours=9))


def target_date() -> str:
    """수집 대상 날짜 = KST 기준 어제 (ISO 문자열)."""
    now_kst = datetime.datetime.now(KST)
    yesterday = (now_kst - datetime.timedelta(days=1)).date()
    return yesterday.isoformat()


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
    """경북 시도 평균 유가 → [(유종명, 가격), ...]. 시도명으로 매칭."""
    r = requests.get(f"{BASE}/avgSidoPrice.do",
                     params={"code": OPINET_KEY, "out": "json"}, timeout=15)
    r.raise_for_status()
    rows = r.json()["RESULT"]["OIL"]
    out = []
    for row in rows:
        name = str(row.get("SIDONM", "")).strip()
        if name in GYEONGBUK_NAMES and row.get("PRODCD") in PROD:
            out.append((PROD[row["PRODCD"]], float(row["PRICE"])))
    return out


def main():
    if not OPINET_KEY:
        print("OPINET_KEY 없음. 종료.")
        sys.exit(1)

    day = target_date()   # 어제(KST)
    records = []
    try:
        for name, price in fetch_nationwide():
            records.append({"날짜": day, "지역": "전국", "유종": name, "가격": price})
        for name, price in fetch_gyeongbuk():
            records.append({"날짜": day, "지역": "경북", "유종": name, "가격": price})
    except Exception as e:
        print(f"수집 실패: {e}")
        sys.exit(1)

    if not records:
        print("수집된 데이터 없음.")
        sys.exit(1)

    new_df = pd.DataFrame(records)

    # 기존 누적본과 병합 (같은 날짜 중복 제거 → 재실행/수동실행 대비)
    if os.path.exists(HISTORY_PATH):
        old = pd.read_csv(HISTORY_PATH)
        old["날짜"] = old["날짜"].astype(str)
        old = old[old["날짜"] != day]
        merged = pd.concat([old, new_df], ignore_index=True)
    else:
        merged = new_df

    # 최근 KEEP_DAYS 일만 유지 (KST 어제 기준으로 컷오프)
    merged["날짜"] = merged["날짜"].astype(str)
    cutoff = (datetime.datetime.now(KST).date()
              - datetime.timedelta(days=KEEP_DAYS)).isoformat()
    merged = merged[merged["날짜"] >= cutoff]
    merged = merged.sort_values(["날짜", "지역", "유종"]).reset_index(drop=True)

    os.makedirs("data", exist_ok=True)
    merged.to_csv(HISTORY_PATH, index=False, encoding="utf-8-sig")
    print(f"{day}(KST 어제) 수집 완료: {len(new_df)}건 추가, 누적 {len(merged)}건")


if __name__ == "__main__":
    main()
