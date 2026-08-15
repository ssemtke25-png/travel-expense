# -*- coding: utf-8 -*-
"""
공무원 출장 여비 산정 시스템 (골격 버전 + 지오코딩 진단 + 관내 자동판정)
--------------------------------------------------
구성
  1) 유가 조회 탭      : 오피넷 무료 API (전국/시도/시군)
  2) 여비 계산 탭      : 주소→좌표(카카오 지오코딩) → 도로거리(카카오모빌리티)
                         → 규정 표준 연비/전비 → 유류비(왕복),
                         그 위에 일비·식비·숙박비·통행료 구조
                         + 관내 자동판정(호명읍 특례 + 12km) — 제안 후 사용자 전환
  3) 일괄 처리 탭      : 출장자 명단 엑셀 업로드 → 전원 계산 → 여비 지급명세서 엑셀 출력

설계 원칙
  - 저장 기능 없음 (세션 동안만 유지). 개인정보 서버 미보관.
  - 자동 계산값을 기본값으로 채우되, 모든 항목 수기 수정 가능 (3층 구조).
  - 여비 조례 단가/출력 양식은 사무실 실물 자료 확보 후 채울 자리만 확보.
  - 연비는 「공무원 여비업무 처리기준(2023.1.18.)」 유종별 표준값을 기본값으로 사용.

API 키는 .streamlit/secrets.toml (또는 Streamlit Cloud Secrets)에서 로드.
"""

import io
import time
from datetime import date
from modules.travel_expense_excel import (
    Person, Group,
    build_workbook, workbook_to_bytes,
    get_road_distance_waypoints,
)
from modules.tab3_integration import render_tab3
import requests
import pandas as pd
import streamlit as st

# =============================================================
# 설정 / 상수
# =============================================================
st.set_page_config(page_title="출장 여비 산정 시스템", page_icon="🚗", layout="wide")

# ---- API 키 로드 (없어도 UI는 뜨도록 방어) ----
def _get_secret(key: str) -> str:
    try:
        return st.secrets[key]
    except Exception:
        return ""

OPINET_KEY = _get_secret("OPINET_KEY")          # 오피넷 무료 API 키
KAKAO_REST_KEY = _get_secret("KAKAO_REST_KEY")  # 카카오 REST API 키 (지오코딩+길찾기 공용)

# ---- 여비 단가 ----
ALLOWANCE_RATES = {
    "일비_1일": 25000,
    "식비_1일": 25000,
    "숙박비_상한_서울": 100000,
    "숙박비_상한_광역시": 80000,
    "숙박비_상한_그밖": 70000,
    "관내_4시간이상": 20000,
    "관내_4시간미만": 10000,
    "관내_공무용차량감액": 10000,
    "관내_거리기준_km": 12.0,
}

# ---- 유종별 표준 연비/전비 ----
# (modules.travel_expense_excel 에도 동일 정의가 있으나, app.py 단건 계산용으로 여기 유지)
FUEL_STANDARDS = {
    "휘발유":            {"eff": 11.97, "unit": "km/L"},
    "경유":              {"eff": 12.52, "unit": "km/L"},
    "LPG":               {"eff": 8.83,  "unit": "km/L"},
    "하이브리드":         {"eff": 15.37, "unit": "km/L"},
    "플러그인하이브리드":  {"eff": 10.61, "unit": "km/L"},
    "전기":              {"eff": 5.22,  "unit": "km/kWh"},
}
PHEV_ELEC_EFF = 2.84

DEFAULT_ELEC_PRICE = 300.0

OIL_PRICE_FUEL_MAP = {
    "휘발유": "휘발유",
    "경유": "경유",
    "LPG": "자동차용LPG",
    "하이브리드": "휘발유",
    "플러그인하이브리드": "휘발유",
    "전기": None,
}

SUKBAK_REGION = {
    "서울특별시": "숙박비_상한_서울",
    "광역시": "숙박비_상한_광역시",
    "그 밖의 지역": "숙박비_상한_그밖",
}

OPINET_PRODUCTS = {
    "휘발유": "B027",
    "고급휘발유": "B034",
    "경유": "D047",
    "실내등유": "C004",
    "자동차용LPG": "K015",
}

OPINET_SIDO = {
    "서울": "01", "경기": "31", "강원": "42", "충북": "43", "충남": "44",
    "전북": "45", "전남": "46", "경북": "47", "경남": "48", "부산": "26",
    "대구": "27", "인천": "28", "광주": "29", "대전": "30", "울산": "31",
    "제주": "50", "세종": "51",
}

OPINET_BASE = "https://www.opinet.co.kr/api"
KAKAO_GEOCODE_URL = "https://dapi.kakao.com/v2/local/search/address.json"
KAKAO_KEYWORD_URL = "https://dapi.kakao.com/v2/local/search/keyword.json"
KAKAO_DIRECTIONS_URL = "https://apis-navi.kakaomobility.com/v1/directions"


# =============================================================
# 관내 자동판정 (호명읍 특례 + 12km) — 제안용
# =============================================================
DOCHEONG_ADDR = "경상북도 안동시 풍천면 도청대로 455"


def is_from_docheong(origin: str) -> bool:
    if not origin:
        return False
    o = origin.replace(" ", "")
    return ("도청대로455" in o) or ("경상북도청" in o) or ("풍천면도청대로" in o)


def check_gwannae_candidate(origin: str, dest: str, dist_km_oneway):
    reasons = []
    dest_c = (dest or "").replace(" ", "")
    if is_from_docheong(origin):
        if "안동시" in dest_c:
            reasons.append("도청 출발 + 안동시(도청 소재지, 관내)")
        if "예천군호명읍" in dest_c or "호명읍" in dest_c:
            reasons.append("도청 출발 + 예천군 호명읍(관내 특례)")
    if dist_km_oneway is not None and dist_km_oneway < ALLOWANCE_RATES["관내_거리기준_km"]:
        reasons.append(f"편도 {dist_km_oneway:,.1f}km < 12km")
    return (len(reasons) > 0, " · ".join(reasons))


def _switch_to_gwannae():
    st.session_state["force_gwannae"] = True
    st.session_state.pop("gwannae_candidate", None)
    st.session_state.pop("gwannae_candidate_reason", None)


# =============================================================
# 데이터 로드
# =============================================================
@st.cache_data(ttl=1800)
def load_oil_history() -> pd.DataFrame:
    try:
        df = pd.read_csv("data/oil_history.csv")
        df["날짜"] = df["날짜"].astype(str)
        return df
    except Exception:
        return pd.DataFrame(columns=["날짜", "지역", "유종", "가격"])


def lookup_oil_price(hist: pd.DataFrame, date_str: str, region: str, fuel: str):
    if hist.empty:
        return None
    m = hist[(hist["날짜"] == date_str) & (hist["지역"] == region) & (hist["유종"] == fuel)]
    if m.empty:
        return None
    return float(m.iloc[0]["가격"])


# =============================================================
# 오피넷 유가 API
# =============================================================
@st.cache_data(ttl=3600)
def fetch_oil_price_nationwide() -> pd.DataFrame:
    if not OPINET_KEY:
        return pd.DataFrame()
    url = f"{OPINET_BASE}/avgAllPrice.do"
    params = {"code": OPINET_KEY, "out": "json"}
    try:
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        rows = r.json()["RESULT"]["OIL"]
        return pd.DataFrame(rows)
    except Exception as e:
        st.warning(f"전국 유가 조회 실패: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=3600)
def fetch_oil_price_sido() -> pd.DataFrame:
    if not OPINET_KEY:
        return pd.DataFrame()
    url = f"{OPINET_BASE}/avgSidoPrice.do"
    params = {"code": OPINET_KEY, "out": "json"}
    try:
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        rows = r.json()["RESULT"]["OIL"]
        return pd.DataFrame(rows)
    except Exception as e:
        st.warning(f"시도별 유가 조회 실패: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=3600)
def fetch_oil_price_sigun(sido_code: str, product_code: str) -> pd.DataFrame:
    if not OPINET_KEY:
        return pd.DataFrame()
    url = f"{OPINET_BASE}/avgSigunPrice.do"
    params = {"code": OPINET_KEY, "out": "json", "sido": sido_code, "prodcd": product_code}
    try:
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        rows = r.json()["RESULT"]["OIL"]
        return pd.DataFrame(rows)
    except Exception as e:
        st.warning(f"시군별 유가 조회 실패: {e}")
        return pd.DataFrame()


# =============================================================
# 카카오 지오코딩 + 길찾기
# =============================================================
def geocode_debug(address: str) -> dict:
    if not KAKAO_REST_KEY:
        return {"ok": False, "xy": None, "method": "-",
                "reason": "카카오 REST 키가 secrets에 설정되지 않았습니다."}
    if not address or not address.strip():
        return {"ok": False, "xy": None, "method": "-", "reason": "주소가 비어 있습니다."}

    addr = address.strip()
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_KEY}"}

    try:
        r = requests.get(KAKAO_GEOCODE_URL, headers=headers,
                         params={"query": addr}, timeout=10)
        if r.status_code == 401:
            return {"ok": False, "xy": None, "method": "address",
                    "reason": "401 인증 실패 — REST 키가 틀렸거나 카카오 앱에 등록된 키가 아닙니다."}
        if r.status_code == 403:
            return {"ok": False, "xy": None, "method": "address",
                    "reason": "403 권한 거부 — 카카오 앱에서 '카카오맵/로컬(주소검색)' 사용이 꺼져 있거나 "
                              "플랫폼(도메인/IP) 제한에 걸렸습니다."}
        r.raise_for_status()
        docs = r.json().get("documents", [])
        if docs:
            return {"ok": True,
                    "xy": (float(docs[0]["x"]), float(docs[0]["y"])),
                    "method": "address", "reason": "성공(주소검색)"}
    except Exception as e:
        return {"ok": False, "xy": None, "method": "address",
                "reason": f"주소검색 호출 예외: {type(e).__name__}: {e}"}

    try:
        r2 = requests.get(KAKAO_KEYWORD_URL, headers=headers,
                          params={"query": addr}, timeout=10)
        if r2.status_code in (401, 403):
            return {"ok": False, "xy": None, "method": "keyword",
                    "reason": f"{r2.status_code} — 키워드 검색 권한 문제. 카카오 앱 로컬 API 설정 확인 필요."}
        r2.raise_for_status()
        docs2 = r2.json().get("documents", [])
        if docs2:
            return {"ok": True,
                    "xy": (float(docs2[0]["x"]), float(docs2[0]["y"])),
                    "method": "keyword", "reason": "성공(키워드검색 폴백)"}
    except Exception as e:
        return {"ok": False, "xy": None, "method": "keyword",
                "reason": f"키워드검색 호출 예외: {type(e).__name__}: {e}"}

    return {"ok": False, "xy": None, "method": "both",
            "reason": f"주소·키워드 모두 검색결과 0건: '{addr}'. 시·군 포함 전체 주소로 다시 입력해 보세요."}


@st.cache_data(ttl=86400)
def geocode(address: str):
    result = geocode_debug(address)
    return result["xy"] if result["ok"] else None


@st.cache_data(ttl=86400)
def geocode_full(address: str) -> dict:
    return geocode_debug(address)


def get_road_distance(origin_xy, dest_xy):
    if not KAKAO_REST_KEY:
        return {"ok": False, "reason": "카카오 키가 설정되지 않았습니다."}
    if not origin_xy or not dest_xy:
        return {"ok": False, "reason": "좌표가 없습니다."}
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_KEY}"}
    params = {
        "origin": f"{origin_xy[0]},{origin_xy[1]}",
        "destination": f"{dest_xy[0]},{dest_xy[1]}",
        "priority": "RECOMMEND",
        "car_fuel": "GASOLINE",
        "car_hipass": "true",       # 하이패스 기준(카카오맵 웹 기본과 일치)
        "alternatives": "true",     # 대안경로 여러 개 받기 → 통행료 최저 선택
    }
    try:
        r = requests.get(KAKAO_DIRECTIONS_URL, headers=headers, params=params, timeout=10)
        if r.status_code in (401, 403):
            return {"ok": False, "reason": "길찾기 API 사용 권한이 아직 승인되지 않았습니다. "
                                           "카카오 데브톡 승인 후 자동으로 작동합니다."}
        r.raise_for_status()
        routes = r.json().get("routes", [])
        # result_code != 0 인 경로는 제외(길찾기 실패 경로)
        valid = [rt for rt in routes if rt.get("result_code", 0) == 0 and "summary" in rt]
        if not valid:
            valid = [rt for rt in routes if "summary" in rt]
        if not valid:
            return {"ok": False, "reason": "경로를 찾지 못했습니다."}
        # 통행료 최저 경로 선택 (동일하면 거리 짧은 쪽)
        def _toll(rt):
            return rt["summary"].get("fare", {}).get("toll", 0)
        best = min(valid, key=lambda rt: (_toll(rt), rt["summary"].get("distance", 0)))
        summary = best["summary"]
        return {
            "ok": True,
            "distance_m": summary["distance"],
            "duration_s": summary["duration"],
            "toll": summary.get("fare", {}).get("toll", 0),
            "n_routes": len(valid),
        }
    except Exception as e:
        return {"ok": False, "reason": f"길찾기 호출 오류: {e}"}


# =============================================================
# 계산 로직
# =============================================================
def floor10(value) -> int:
    try:
        return (int(value) // 10) * 10
    except (TypeError, ValueError):
        return 0


def calc_fuel_cost(distance_km: float, efficiency: float, unit_price: float) -> float:
    if efficiency <= 0:
        return 0.0
    return (distance_km / efficiency) * unit_price


def calculate_travel_allowance(
    *,
    is_gwannae: bool,
    hours_over_4: bool,
    use_official_car: bool,
    days: int,
    nights: int,
    sukbak_region_key: str,
    fuel_cost: float,
    manual_transport: float,
    toll: float = 0.0,
    num_passengers: int = 0,
) -> dict:
    result = {}

    if is_gwannae:
        base = (ALLOWANCE_RATES["관내_4시간이상"] if hours_over_4
                else ALLOWANCE_RATES["관내_4시간미만"])
        if use_official_car:
            base = max(0, base - ALLOWANCE_RATES["관내_공무용차량감액"])
        base = floor10(base)
        pax_n = max(0, int(num_passengers))
        result.update({
            "구분": "근무지 내(관내)",
            "관내정액": base,
            "운전자정액": base,
            "동승자정액": base,
            "동승자수": pax_n,
            "동승자합계": base * pax_n,
            "합계": base,
            "총합계": base + base * pax_n,
        })
        return result

    ilbi_unit = ALLOWANCE_RATES["일비_1일"]
    ilbi = ilbi_unit * days
    if use_official_car:
        ilbi = ilbi // 2

    sikbi = ALLOWANCE_RATES["식비_1일"] * days

    sukbak_cap = ALLOWANCE_RATES.get(sukbak_region_key, ALLOWANCE_RATES["숙박비_상한_그밖"])
    sukbak = sukbak_cap * max(0, nights)

    if use_official_car:
        transport = 0
        toll_val = 0
    else:
        transport = manual_transport if manual_transport > 0 else fuel_cost
        transport = int(round(transport))
        toll_val = max(0, int(round(toll)))

    ilbi = floor10(ilbi)
    sikbi = floor10(sikbi)
    sukbak = floor10(sukbak)
    transport = floor10(transport)
    toll_val = floor10(toll_val)

    driver = {
        "운임": transport,
        "일비": ilbi,
        "식비": sikbi,
        "숙박비": sukbak,
        "통행료": toll_val,
    }
    driver["합계"] = sum(driver.values())

    pax_unit = {
        "운임": 0,
        "일비": ilbi,
        "식비": sikbi,
        "숙박비": sukbak,
        "통행료": 0,
    }
    pax_unit["합계"] = sum(pax_unit.values())

    pax_n = max(0, int(num_passengers))
    pax_total = pax_unit["합계"] * pax_n

    result.update({
        "구분": "근무지 외(관외)",
        "운전자": driver,
        "동승자단가": pax_unit,
        "동승자수": pax_n,
        "동승자합계": pax_total,
        "총합계": driver["합계"] + pax_total,
    })
    return result


# =============================================================
# UI
# =============================================================
st.title("🚗 공무원 출장 여비 산정 시스템")
st.caption("골격 버전 · 저장 기능 없음(세션 유지) · 여비 조례/출력양식은 실물 자료 확보 후 반영")

with st.expander("🔑 API 연결 상태 확인 / 지오코딩 진단", expanded=False):
    c1, c2, c3 = st.columns(3)
    c1.metric("오피넷 유가", "연결됨" if OPINET_KEY else "미설정")
    c2.metric("카카오 지오코딩", "연결됨" if KAKAO_REST_KEY else "미설정")
    c3.metric("카카오 길찾기", "키 있음(승인 확인 필요)" if KAKAO_REST_KEY else "미설정")
    st.info("길찾기 API는 카카오 데브톡 사용 권한 승인 후 정상 작동합니다.")

    st.markdown("---")
    st.markdown("**🔍 지오코딩 단독 테스트** — 주소 하나만 넣어 실패 사유를 바로 확인")
    test_addr = st.text_input("테스트 주소", value="경상북도 안동시 풍천면 도청대로 455",
                              key="geocode_test_addr")
    if st.button("주소 → 좌표 변환 테스트", key="geocode_test_btn"):
        res = geocode_debug(test_addr)
        if res["ok"]:
            st.success(f"성공 [{res['method']}] → x={res['xy'][0]}, y={res['xy'][1]}")
        else:
            st.error(f"실패 [{res['method']}] → {res['reason']}")
        st.caption("여기서 뜨는 메시지가 실제 원인입니다. "
                   "401/403이면 카카오 앱 설정, 0건이면 주소 형식, 미설정이면 Secrets 문제입니다.")

tab2, tab3 = st.tabs(["🧮 여비 계산", "📋 명단 일괄 처리"])


# -------------------------------------------------------------
# TAB 2 : 여비 계산 (단건)
# -------------------------------------------------------------
SIGUN_PREFIX = {
    "포항시청": "경상북도 포항시 남구 시청로 1",
    "포항남": "경상북도 포항시 남구",
    "포항북": "경상북도 포항시 북구",
    "경주": "경상북도 경주시",
    "김천": "경상북도 김천시",
    "구미": "경상북도 구미시",
    "영주": "경상북도 영주시",
    "영천": "경상북도 영천시",
    "상주": "경상북도 상주시",
    "문경": "경상북도 문경시",
    "경산": "경상북도 경산시",
    "의성": "경상북도 의성군",
    "청송": "경상북도 청송군",
    "영양": "경상북도 영양군",
    "영덕": "경상북도 영덕군",
    "청도": "경상북도 청도군",
    "고령": "경상북도 고령군",
    "성주": "경상북도 성주군",
    "칠곡": "경상북도 칠곡군",
    "예천": "경상북도 예천군",
    "봉화": "경상북도 봉화군",
    "울진": "경상북도 울진군",
    "울릉": "경상북도 울릉군",
}
SIGUN_ORDER = list(SIGUN_PREFIX.keys())
DEST_OPTIONS = SIGUN_ORDER + ["기타(직접입력)"]

DEFAULT_ORIGIN = "경상북도 안동시 풍천면 도청대로 455"


def build_dest_address(sigun_choice: str, detail: str) -> str:
    detail = (detail or "").strip()
    if sigun_choice == "기타(직접입력)":
        return detail
    prefix = SIGUN_PREFIX.get(sigun_choice, "")
    if sigun_choice == "포항시청":
        return prefix
    if not detail:
        return prefix
    core = prefix.replace("경상북도 ", "")
    if detail.startswith("경상북도") or core in detail:
        return detail
    return f"{prefix} {detail}"


with tab2:
    st.subheader("단건 여비 계산")

    _default_idx = 1 if st.session_state.pop("force_gwannae", False) else 0
    g1, g2 = st.columns(2)
    trip_type = g1.radio(
        "출장 구분", ["근무지 외(관외)", "근무지 내(관내)"],
        index=_default_idx,
        horizontal=True,
        help="같은 시·군 안 또는 여행거리 12km 미만이면 관내입니다. "
             "관내는 정액이라 경로·유종·유가 입력이 필요 없습니다.",
    )
    use_car = g2.checkbox(
        "공무용 차량 이용",
        help="관외: 운임·통행료 0(제15조) + 일비 1/2(제16조③) / 관내: 1만원 감액(제18조①)",
    )
    is_gwannae = (trip_type == "근무지 내(관내)")
    if use_car:
        st.caption("🚙 공무용 차량 이용 체크됨 → 관외: **운임·통행료 미지급**, 일비 1/2 자동 적용 · "
                   "관내: 정액 1만원 감액. (기름값·통행료는 기관 부담)")

    st.markdown("---")

    if is_gwannae:
        st.markdown("#### 근무지 내(관내) 출장 — 정액 여비")
        st.caption("「공무원 여비 규정」 제18조: 4시간 이상 20,000원 / 4시간 미만 10,000원. "
                   "공무용 차량 이용 시 10,000원 감액. 운전자·동승자 동일 정액. (경로·유종·유가 입력 불필요)")

        gi0, gi1, gi2 = st.columns([1, 1, 1])
        trip_date_g = gi0.date_input("출장일자", value=date.today(),
                                     help="엑셀 추출용. 관내는 당일 기준입니다.")
        driver_name = gi1.text_input("운전자 성명", placeholder="예) 홍길동",
                                     help="엑셀 추출용")
        hours_over_4 = gi2.radio("출장 시간", ["4시간 이상", "4시간 미만"],
                                 horizontal=True) == "4시간 이상"

        num_pax = st.number_input("동승자 수 (최대 4)", min_value=0, max_value=4, value=0, step=1)
        passenger_names = []
        if num_pax > 0:
            pcols = st.columns(int(num_pax))
            for i in range(int(num_pax)):
                pname = pcols[i].text_input(f"동승자 {i+1}", key=f"gpax_{i}", placeholder="성명")
                passenger_names.append(pname)
        st.caption("ⓘ 관내는 정액이라 운전자·동승자 모두 같은 금액을 각자 받습니다.")

        if st.button("여비 계산", type="primary", key="calc_gwannae"):
            res = calculate_travel_allowance(
                is_gwannae=True,
                hours_over_4=hours_over_4,
                use_official_car=use_car,
                days=0, nights=0,
                sukbak_region_key="숙박비_상한_그밖",
                fuel_cost=0.0, manual_transport=0.0, toll=0.0,
                num_passengers=int(num_pax),
            )
            st.markdown(f"**구분: {res['구분']}** · 출장일자: {trip_date_g:%Y-%m-%d}")
            if use_car:
                st.markdown(
                    """<div style="background:#fff4e5;border:2px solid #f39c12;border-radius:8px;
                    padding:10px 14px;margin:6px 0;">
                    🚙 <b>공무용 차량 이용</b> → 관내 정액에서 <b style="color:#c0392b;">1만원 감액</b> 적용됨
                    (기름값·통행료는 기관 부담)</div>""",
                    unsafe_allow_html=True,
                )
            st.markdown(f"**🚗 운전자: {driver_name or '(성명 미입력)'}** → {res['운전자정액']:,} 원")
            if res["동승자수"] > 0:
                st.markdown(f"**🧑‍🤝‍🧑 동승자 {res['동승자수']}명** (1인당 {res['동승자정액']:,} 원)")
                for i, pname in enumerate(passenger_names):
                    st.write(f"　- 동승자 {i+1}: {pname or '(성명 미입력)'} → {res['동승자정액']:,} 원")
                st.caption(f"동승자 소계: {res['동승자정액']:,} 원 × {res['동승자수']}명 = {res['동승자합계']:,} 원")
            st.success(f"총 여비 합계 (운전자 + 동승자 {res['동승자수']}명): {res['총합계']:,} 원")
            st.caption("※ 십원 미만 버림 적용. 관내는 일비·식비·숙박비·운임 별도 지급 없음(정액).")

        st.info("정식 여비 정산서(엑셀/HWPX) 출력 양식은 실물 양식 확보 후 셀 매핑으로 연결합니다. "
                "출장일자·성명은 그때 엑셀 추출에 사용됩니다.")

    else:
        st.markdown("##### 1) 출장 경로")
        origin = st.text_input(
            "출발지 주소",
            value=st.session_state.get("origin", DEFAULT_ORIGIN),
            help="기본값은 경상북도청입니다. 필요 시 수정하세요.",
        )
        st.session_state["origin"] = origin

        st.markdown("**도착지 시군 선택** (직재순)")
        sigun_choice = st.radio(
            "도착 시군", DEST_OPTIONS, horizontal=True, label_visibility="collapsed",
        )
        if sigun_choice == "기타(직접입력)":
            dest_detail = st.text_input(
                "도착지 주소(전체 입력)",
                placeholder="예) 대구광역시 북구 ○○로 123",
            )
        elif sigun_choice == "포항시청":
            dest_detail = ""
            st.caption("ⓘ 포항시청(본청)은 고정 지점입니다 — 세부주소 입력 없이 시청으로 계산됩니다.")
        else:
            dest_detail = st.text_input(
                f"세부 주소 — '{SIGUN_PREFIX[sigun_choice]}' 뒤에 이어질 부분만 입력",
                placeholder="예) 풍천면 도청대로 455  (시·도·시군명은 생략 가능)",
            )
        dest = build_dest_address(sigun_choice, dest_detail)
        if dest:
            st.caption(f"🎯 최종 도착지: **{dest}**")

        st.markdown("##### 2) 출장자")
        st.caption("소속 / 직급 / 성명 순으로 입력하세요. 명세서 출장자 칸에 '소속 직급 성명'으로 출력됩니다.")
        nm0, nm1, nm2 = st.columns([1.4, 1, 1])
        driver_dept = nm0.text_input("운전자 소속", value="건설도시국 토지정보과",
                                     placeholder="예) 건설도시국 토지정보과", key="drv_dept")
        driver_rank = nm1.text_input("운전자 직급", placeholder="예) 지방시설주사", key="drv_rank")
        driver_name = nm2.text_input("운전자 성명", placeholder="예) 홍길동", key="drv_name")

        num_pax = st.number_input("동승자 수 (최대 4)", min_value=0, max_value=4, value=0, step=1)
        passengers_info = []   # [(소속, 직급, 성명), ...]
        passenger_names = []   # 하위호환(기존 표시 코드용)
        if num_pax > 0:
            for i in range(int(num_pax)):
                pc0, pc1, pc2 = st.columns([1.4, 1, 1])
                p_dept = pc0.text_input(f"동승자 {i+1} 소속", value=driver_dept,
                                        key=f"pax_dept_{i}", placeholder="소속")
                p_rank = pc1.text_input(f"동승자 {i+1} 직급", key=f"pax_rank_{i}", placeholder="직급")
                p_name = pc2.text_input(f"동승자 {i+1} 성명", key=f"pax_{i}", placeholder="성명")
                passengers_info.append((p_dept, p_rank, p_name))
                passenger_names.append(p_name)
        st.caption("ⓘ 동승자는 같은 차량 탑승 기준입니다. 운임·통행료는 운전자만, "
                   "일비·식비·숙박비는 전원 각자 지급됩니다.")

        st.markdown("##### 3) 출장 목적")
        trip_purpose = st.text_input(
            "출장 목적", placeholder="예) ○○ 지적재조사 지구 현장 확인",
            help="운전자·동승자 모두 동일하게 출력됩니다.",
        )

        st.markdown("##### 4) 출장 기간")
        hours_over_4 = True

        def _sync_end_date():
            s = st.session_state.get("trip_start_key")
            if s is not None:
                st.session_state["trip_end_key"] = s

        if "trip_start_key" not in st.session_state:
            st.session_state["trip_start_key"] = date.today()
        if "trip_end_key" not in st.session_state:
            st.session_state["trip_end_key"] = date.today()

        d0, d1 = st.columns(2)
        trip_start = d0.date_input(
            "출장 시작일", key="trip_start_key", on_change=_sync_end_date,
            help="달력에서 출장 시작일을 고르세요. 종료일이 자동으로 같은 날로 초기화됩니다(당일치기 기본). "
                 "이 날짜가 유가 기준일로도 자동 적용됩니다.")
        trip_end = d1.date_input(
            "출장 종료일", key="trip_end_key",
            help="당일치기면 그대로 두세요. 1박 이상이면 시작일을 먼저 고른 뒤 종료일만 뒤로 조정하세요.")

        # 출장 시간 (기본 09:00~18:00, 수정 가능) — 명세서 '출장월일(출장시간)' 칸에 사용
        tc0, tc1 = st.columns(2)
        trip_time_start = tc0.text_input("출장 시작시간", value="09:00", key="trip_time_start",
                                         help="명세서 출장월일 칸에 함께 표기됩니다. 기본 09:00, 수정 가능.")
        trip_time_end = tc1.text_input("출장 종료시간", value="18:00", key="trip_time_end",
                                       help="기본 18:00, 수정 가능.")

        if trip_end < trip_start:
            st.error("종료일이 시작일보다 빠릅니다. 날짜를 다시 확인하세요.")
            trip_end = trip_start

        span = (trip_end - trip_start).days
        days = span + 1
        nights = span
        st.markdown(
            f"""<div style="background:#eaf4ff;border:2px solid #2b7cff;border-radius:8px;
            padding:12px 16px;margin:8px 0;font-size:1.05rem;">
            📅 <b>출장 기간: {trip_start:%Y-%m-%d} ~ {trip_end:%Y-%m-%d}</b>
            &nbsp;→&nbsp; <b style="color:#c0392b;font-size:1.15rem;">{nights}박 {days}일</b>
            &nbsp;·&nbsp; <b>{trip_time_start} ~ {trip_time_end}</b>
            <br><span style="color:#333;">일비·식비 <b>{days}일</b>分 · 숙박 <b>{nights}박</b>分 자동 계산</span>
            </div>""",
            unsafe_allow_html=True,
        )

        st.markdown("##### 5) 차량 유종 / 연비 (규정 표준값)")
        st.caption("「공무원 여비업무 처리기준(2023.1.18.)」 유종별 표준 연비를 기본값으로 적용합니다. "
                   "실측값이 있으면 수정하세요.")
        c3, c4 = st.columns(2)
        fuel_type = c3.selectbox("유종", list(FUEL_STANDARDS.keys()),
                                 help="규정 표준 연비/전비가 자동 적용됩니다.")
        std = FUEL_STANDARDS[fuel_type]
        is_electric = (std["unit"] == "km/kWh")
        eff_label = "표준 전비(km/kWh)" if is_electric else "표준 연비(km/L)"
        eff = c4.number_input(eff_label, min_value=0.0, value=float(std["eff"]), step=0.01,
                              help="규정 표준값. 수정 가능.")
        if fuel_type == "플러그인하이브리드":
            st.caption(f"ⓘ 플러그인하이브리드는 규정상 연비 10.61 km/L(휘발유 운행 기준)와 "
                       f"전비 {PHEV_ELEC_EFF} km/kWh가 모두 제시됩니다. 기본은 연비 기준이며, "
                       f"전기 운행분으로 계산하려면 유종을 '전기'로 바꾸거나 위 값을 전비로 수정하세요.")

        st.markdown("##### 6) 단가 (유가 또는 전기요금)")
        hist = load_oil_history()

        if is_electric:
            c_e1, c_e2 = st.columns([1, 3])
            oil_price = c_e1.number_input("전기요금(원/kWh)", min_value=0.0,
                                          value=DEFAULT_ELEC_PRICE, step=10.0,
                                          help="오피넷 미제공 항목이라 수기 입력입니다. "
                                               "기본값은 공용 급속 대략 단가(300원/kWh).")
            c_e2.caption("ⓘ 전기차는 유가가 아니라 전기요금(원/kWh)으로 계산합니다. "
                         "충전 방식(급속/완속·공용/자가)에 따라 단가가 크게 다르니 실제 단가로 수정하세요.")
            region = "전국"
        else:
            c6, c7, c8, c9 = st.columns([1.2, 1, 1.2, 1])
            oil_fuel = OIL_PRICE_FUEL_MAP.get(fuel_type, "휘발유")
            c6.text_input("유가 기준 유종", value=oil_fuel, disabled=True,
                          help="하이브리드·PHEV는 휘발유 유가를 기준으로 합니다.")
            region = c7.selectbox("지역 기준", ["경북", "전국"],
                                  help="여비 규정에 명시된 유가 기준에 맞춰 선택하세요. 기본값은 경북입니다.")
            sel_date = trip_start.isoformat()
            c8.text_input("유가 기준일 (= 출장 시작일)", value=sel_date, disabled=True,
                          help="규정상 출장 시작일의 유가를 적용합니다. 위 '출장 시작일'을 바꾸면 자동으로 바뀝니다.")
            auto_price = None
            if not hist.empty:
                avail_dates = sorted(hist["날짜"].unique().tolist())
                auto_price = lookup_oil_price(hist, sel_date, region, oil_fuel)
            else:
                avail_dates = []
            default_price = auto_price if auto_price is not None else 0.0
            oil_price = c9.number_input("적용 유가(원/L)", min_value=0.0,
                                        value=float(default_price), step=1.0,
                                        help="출장 시작일·지역·유종의 오피넷 값이 자동 입력됩니다. 수정 가능.")
            if auto_price is not None:
                st.caption(f"✅ {sel_date}(출장 시작일) · {region} · {oil_fuel} 유가 자동 적용: {auto_price:,.2f} 원/L")
            elif avail_dates:
                st.caption(f"⚠️ {sel_date}(출장 시작일) · {region} · {oil_fuel} 유가 데이터가 없어 수기 입력이 필요합니다. "
                           f"(유가 보유 범위: {avail_dates[0]} ~ {avail_dates[-1]} · 주말·공휴일 등은 수집 누락일 수 있음)")
            else:
                st.caption("ⓘ 유가 누적 데이터가 아직 쌓이지 않았습니다. GitHub Actions가 매일 수집하며, "
                           "며칠 후부터 자동 적용됩니다. 그전까지는 유가를 수기 입력하세요.")

        st.markdown("##### 7) 여비 산정 입력")
        st.caption("「공무원 여비 규정」 준용. 단가는 상단 ALLOWANCE_RATES 기본값(국가 규정)입니다. "
                   "경상북도 별도 조례가 없어 국가 규정을 그대로 적용합니다. 운임은 왕복(편도×2)으로 계산합니다.")

        sukbak_region_label = st.selectbox(
            "숙박지역(상한)", list(SUKBAK_REGION.keys()),
            index=2, help="서울 10만 / 광역시 8만 / 그 밖 7만 (실비, 상한 이내)")
        sukbak_region_key = SUKBAK_REGION[sukbak_region_label]

        daily_commute = (int(nights) == 0 and int(days) >= 2)
        round_multiplier = 2 * (int(days) if daily_commute else 1)
        if daily_commute:
            st.caption(f"🚗 숙박 없이 {int(days)}일 출장 → **매일 출퇴근**으로 보아 운임·통행료를 "
                       f"편도 × 2(왕복) × {int(days)}일 = **×{round_multiplier}**로 계산합니다.")
        else:
            st.caption("🚗 운임·통행료는 **편도 × 2(왕복 1회)**로 계산합니다.")

        manual_transport = st.number_input(
            "운임(대중교통 등, 원) — 비우면 자가차 유류/전기비 사용", min_value=0, value=0, step=1000,
            disabled=use_car,
            help="자가차 출장이면 비워두세요(왕복 유류/전기비가 운임으로 들어갑니다). "
                 "KTX·버스 등 대중교통이면 실비를 입력하세요. "
                 "공무용 차량이면 운임은 0으로 처리되어 입력이 비활성화됩니다.")
        toll_input = st.number_input(
            "통행료(원, 왕복)", min_value=0, value=0, step=100,
            disabled=use_car,
            help="운전자가 실제 낸 통행료(왕복)를 직접 입력하세요. 하이패스 할인·경로 선택에 따라 "
                 "사람마다 다르므로 직접 입력합니다. 입력한 값이 명세서 통행료행에 반영되어 "
                 "계·청구액이 자동 재계산됩니다. 공무용 차량이면 통행료는 0으로 처리됩니다.")
        if use_car:
            st.caption("ⓘ 공무용 차량 이용이므로 운임·통행료는 계산에서 0 처리됩니다.")

        st.markdown("---")
        if st.button("💰 여비 합계 계산", type="primary", key="calc_all"):
            if not origin or not dest:
                st.error("출발지와 도착지 주소를 모두 입력하세요. (도착지: 시군 선택 후 세부주소 또는 기타 입력)")
            else:
                o = geocode_debug(origin)
                d = geocode_debug(dest)
                if not o["ok"] or not d["ok"]:
                    if not o["ok"]:
                        st.error(f"출발지 실패 [{o['method']}] → {o['reason']}")
                    if not d["ok"]:
                        st.error(f"도착지 실패 [{d['method']}] → {d['reason']}")
                    st.caption("↑ 상단 '🔑 API 연결 상태 확인 / 지오코딩 진단'에서 단독 테스트로 원인을 좁힐 수 있습니다.")
                else:
                    o_xy, d_xy = o["xy"], d["xy"]
                    route = get_road_distance(o_xy, d_xy)
                    if not route["ok"]:
                        st.warning(f"도로거리 계산 대기: {route['reason']}")
                        st.info("주소 인식은 정상입니다. 길찾기 승인 후 이 버튼만 다시 누르면 계산됩니다.")
                    else:
                        dist_km_oneway = route["distance_m"] / 1000
                        st.session_state["last_dist_km_oneway"] = dist_km_oneway

                        cand, why = check_gwannae_candidate(origin, dest, dist_km_oneway)
                        if cand:
                            st.warning(f"🏠 **관내 출장 후보**입니다 — {why}\n\n"
                                       "같은 시·군 안이거나 여행거리 12km 미만이면 관내(정액)로 "
                                       "처리해야 할 수 있습니다. 아래 버튼으로 전환하거나, 관외 결과를 그대로 사용하세요.")
                            st.button("→ 관내(정액)로 전환", key="switch_to_gwannae",
                                      type="primary", on_click=_switch_to_gwannae)

                        applied_dist = dist_km_oneway * round_multiplier
                        # 통행료는 사용자 입력칸 값을 그대로 사용(운전자마다 실제 금액 상이).
                        applied_toll = int(toll_input)

                        fuel_cost_val = calc_fuel_cost(applied_dist, eff, oil_price)

                        m1, m2, m3 = st.columns(3)
                        m1.metric("도로거리(편도)", f"{dist_km_oneway:,.1f} km")
                        m2.metric(f"적용거리(×{round_multiplier})", f"{applied_dist:,.1f} km")
                        m3.metric("통행료(입력값, 왕복)", f"{applied_toll:,} 원")
                        st.caption(f"경로 좌표 변환 ✓ (출발지:{o['method']} / 도착지:{d['method']}) · "
                                   f"예상 소요(편도) {route['duration_s']//60}분 · "
                                   f"통행료는 위 입력칸에 직접 입력한 {applied_toll:,}원이 적용됩니다.")

                        res = calculate_travel_allowance(
                            is_gwannae=False,
                            hours_over_4=hours_over_4,
                            use_official_car=use_car,
                            days=int(days),
                            nights=int(nights),
                            sukbak_region_key=sukbak_region_key,
                            fuel_cost=fuel_cost_val,
                            manual_transport=float(manual_transport),
                            toll=float(applied_toll),
                            num_passengers=int(num_pax),
                        )

                        st.markdown(f"**구분: {res['구분']}** · 목적: {trip_purpose or '(미입력)'}")
                        if use_car:
                            st.markdown(
                                """<div style="background:#fff4e5;border:2px solid #f39c12;border-radius:8px;
                                padding:10px 14px;margin:6px 0;">
                                🚙 <b>공무용 차량 이용</b> → <b style="color:#c0392b;">운임·통행료 0원</b>,
                                <b style="color:#c0392b;">일비 1/2 감액</b> 적용됨 (기름값·통행료는 기관 부담)</div>""",
                                unsafe_allow_html=True,
                            )
                        drv = res["운전자"]
                        st.markdown(f"**🚗 운전자: {driver_name or '(성명 미입력)'}**")
                        cc = st.columns(6)
                        cc[0].metric("운임", f"{drv['운임']:,} 원")
                        cc[1].metric("일비", f"{drv['일비']:,} 원")
                        cc[2].metric("식비", f"{drv['식비']:,} 원")
                        cc[3].metric("숙박비", f"{drv['숙박비']:,} 원")
                        cc[4].metric("통행료", f"{drv['통행료']:,} 원")
                        cc[5].metric("운전자 합계", f"{drv['합계']:,} 원")

                        if res["동승자수"] > 0:
                            pu = res["동승자단가"]
                            st.markdown(f"**🧑‍🤝‍🧑 동승자 {res['동승자수']}명** "
                                        f"(1인당 일비+식비+숙박비 = {pu['합계']:,} 원 · 목적 동일)")
                            for i, pname in enumerate(passenger_names):
                                st.write(f"　- 동승자 {i+1}: {pname or '(성명 미입력)'} → {pu['합계']:,} 원")
                            st.caption(f"동승자 소계: {pu['합계']:,} 원 × {res['동승자수']}명 = {res['동승자합계']:,} 원")

                        st.success(f"총 여비 합계 (운전자 + 동승자 {res['동승자수']}명): {res['총합계']:,} 원")
                        st.caption("※ 십원 미만 버림 적용. 운임은 왕복(편도×2, 매일 출퇴근이면 ×2×일수). "
                                   "운임·통행료는 운전자만, 일비·식비·숙박비는 전원 각자. "
                                   "숙박비는 실비 정산(상한 이내), 통행료는 실비가 원칙입니다.")

                        # --- 명세서 생성용 데이터 세션 저장 (다운로드 버튼은 아래 블록 밖에서 렌더) ---
                        st.session_state["mx_ready"] = True
                        st.session_state["mx_data"] = {
                            "driver": {
                                "소속": driver_dept, "직급": driver_rank, "성명": driver_name,
                                "목적": trip_purpose,
                                "시작일": trip_start.isoformat(), "종료일": trip_end.isoformat(),
                                "시작시간": trip_time_start, "종료시간": trip_time_end,
                                "출발지명": origin, "도착지명": dest,
                                "유종": fuel_type, "공무차": bool(use_car),
                                "갈때거리": dist_km_oneway, "올때거리": dist_km_oneway,
                                "유가": float(oil_price), "eff": float(eff),
                                "통행료왕복": int(applied_toll),
                                "숙박지역키": sukbak_region_key,
                                "round_mult": int(round_multiplier),
                            },
                            "passengers": [
                                {"소속": pd_, "직급": pr_, "성명": pn_}
                                for (pd_, pr_, pn_) in passengers_info
                            ],
                        }

        # =========================================================
        # 여비 지급명세서 엑셀 다운로드 (계산 결과가 세션에 있을 때만)
        # =========================================================
        if st.session_state.get("mx_ready") and st.session_state.get("mx_data"):
            _mx = st.session_state.get("mx_data")
            if _mx:
                st.markdown("---")
                st.markdown("##### 📄 여비 지급명세서 엑셀")
                st.caption("위 계산 결과를 HWP 「여비 지급명세서」 양식과 동일한 엑셀로 내려받습니다. "
                           "소속·직급·성명, 출장월일·시간, 통행료(입력값)가 그대로 반영됩니다.")
                if st.button("📄 명세서 엑셀 생성", key="mx_make"):
                    try:
                        d = _mx["driver"]
                        # 숙박지역키(app.py) → 엔진 키(서울/광역시/그밖)
                        _region_map = {
                            "숙박비_상한_서울": "서울",
                            "숙박비_상한_광역시": "광역시",
                            "숙박비_상한_그밖": "그밖",
                        }
                        region_e = _region_map.get(d["숙박지역키"], "그밖")
                        s_date = date.fromisoformat(d["시작일"])
                        e_date = date.fromisoformat(d["종료일"])
                        # 통행료 왕복 → 갈때/올때 반반 분배(엔진은 갈때+올때 합산으로 통행료행 표기)
                        toll_half = int(d["통행료왕복"]) // 2
                        toll_go = toll_half
                        toll_back = int(d["통행료왕복"]) - toll_half

                        driver_p = Person(
                            소속=d["소속"], 직급=d["직급"], 성명=d["성명"],
                            목적=d["목적"], 시작일=s_date, 종료일=e_date,
                            출발지명=d["출발지명"], 도착지명=d["도착지명"],
                            유종=d["유종"], is_driver=True, use_official_car=d["공무차"],
                            갈때거리_km=d["갈때거리"] * d["round_mult"] / 2,
                            올때거리_km=d["올때거리"] * d["round_mult"] / 2,
                            유가=d["유가"],
                            통행료_갈때=toll_go, 통행료_올때=toll_back,
                            숙박지역=region_e,
                        )
                        passengers_p = [
                            Person(소속=p["소속"], 직급=p["직급"], 성명=p["성명"],
                                   목적=d["목적"], 시작일=s_date, 종료일=e_date,
                                   도착지명=d["도착지명"], 유종=d["유종"],
                                   is_driver=False, use_official_car=d["공무차"],
                                   숙박지역=region_e)
                            for p in _mx["passengers"]
                        ]
                        grp = Group(driver=driver_p, passengers=passengers_p)
                        # 시간 표기를 엔진에 반영하기 위해 목적 아래 출장월일 칸에 시간 병기
                        # (엔진은 09:00~18:00 고정이므로, 사용자 시간으로 바꾸려면 전역 치환)
                        wb = build_workbook(
                            [grp],
                            dept_footer=f"경상북도 {d['소속']}",
                            signer=(d["성명"] or "성명"),
                        )
                        # 시간 커스텀: 엔진 기본 '(09:00 ~ 18:00)'을 사용자 입력으로 치환
                        _custom_time = f"({d['시작시간']} ~ {d['종료시간']})"
                        if _custom_time != "(09:00 ~ 18:00)":
                            ws = wb.active
                            for row in ws.iter_rows():
                                for cell in row:
                                    if isinstance(cell.value, str) and "(09:00 ~ 18:00)" in cell.value:
                                        cell.value = cell.value.replace("(09:00 ~ 18:00)", _custom_time)
                        xbytes = workbook_to_bytes(wb)
                        st.download_button(
                            "📥 여비 지급명세서 다운로드", xbytes,
                            file_name=f"여비지급명세서_{d['성명'] or '단건'}_{d['시작일']}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary", key="mx_dl")
                        st.success("명세서 생성 완료 — 위 버튼으로 내려받으세요.")
                    except Exception as e:
                        st.error(f"명세서 생성 오류: {e}")

        st.info("계산 후 위 '📄 명세서 엑셀 생성'을 누르면 지급명세서를 내려받을 수 있습니다.")


# -------------------------------------------------------------
# TAB 3 : 명단 일괄 처리 → 여비 지급명세서 엑셀 출력
# -------------------------------------------------------------
with tab3:
    render_tab3(
        geocode_full=geocode_full,
        KAKAO_REST_KEY=KAKAO_REST_KEY,
        fuel_price_map=None,
    )
