# -*- coding: utf-8 -*-
"""
공무원 출장 여비 산정 시스템 (골격 버전 + 지오코딩 진단)
--------------------------------------------------
구성
  1) 유가 조회 탭      : 오피넷 무료 API (전국/시도/시군)
  2) 여비 계산 탭      : 주소→좌표(카카오 지오코딩) → 도로거리(카카오모빌리티)
                         → 규정 표준 연비/전비 → 유류비(왕복),
                         그 위에 일비·식비·숙박비·통행료 구조
  3) 일괄 처리 탭      : 출장자 명단 엑셀 업로드 → 전원 계산 → 엑셀 다운로드

설계 원칙
  - 저장 기능 없음 (세션 동안만 유지). 개인정보 서버 미보관.
  - 자동 계산값을 기본값으로 채우되, 모든 항목 수기 수정 가능 (3층 구조).
  - 여비 조례 단가/출력 양식은 사무실 실물 자료 확보 후 채울 자리만 확보.
  - 연비는 「공무원 여비업무 처리기준(2023.1.18.)」 유종별 표준값을 기본값으로 사용.

API 키는 .streamlit/secrets.toml (또는 Streamlit Cloud Secrets)에서 로드.
"""

import io
import time
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
# 근거: 「공무원 여비 규정」(대통령령) 별표2 및 제18조.
# 지방공무원은 제29조의2에 따라 "해당 지방자치단체 여비 조례"가 우선 적용됨.
#   → 대부분 지자체가 국가 별표2 금액을 그대로 준용하나,
#     아래 값은 반드시 경상북도 여비 조례 별표로 최종 확인해 교체할 것.
# 아래는 국가 규정 기준 기본값(교체 전 임시값).
ALLOWANCE_RATES = {
    "일비_1일": 25000,          # 별표2: 1일 25,000원
    "식비_1일": 25000,          # 별표2: 1일 25,000원 (국내는 정액)
    # 숙박비 상한은 지역별 상이(실비 정산, 상한 이내)
    "숙박비_상한_서울": 100000,
    "숙박비_상한_광역시": 80000,
    "숙박비_상한_그밖": 70000,
    # 근무지 내(관내) 국내출장 정액 — 제18조
    "관내_4시간이상": 20000,     # 4시간 이상 2만원
    "관내_4시간미만": 10000,     # 4시간 미만 1만원
    "관내_공무용차량감액": 10000,  # 공무용 차량 이용 시 1만원 감액
    # 관내 출장 판정 거리 기준(km) — 제18조 제2항 제2호
    "관내_거리기준_km": 12.0,
}

# ---- 유종별 표준 연비/전비 ----
# 근거: 「공무원 여비업무 처리기준 변경사항 안내」(2023.1.18.)
#       "가. 국내자동차운임 지급 기준 현행화 — 승용차 유종별 연비(전비)기준 세분화"
# unit:
#   "km/L"  → 유류비 = 왕복거리 ÷ 연비 × 유가(원/L)
#   "km/kWh"→ 전기료 = 왕복거리 ÷ 전비 × 전기요금(원/kWh)
FUEL_STANDARDS = {
    "휘발유":            {"eff": 11.97, "unit": "km/L"},
    "경유":              {"eff": 12.52, "unit": "km/L"},
    "LPG":               {"eff": 8.83,  "unit": "km/L"},
    "하이브리드":         {"eff": 15.37, "unit": "km/L"},
    "플러그인하이브리드":  {"eff": 10.61, "unit": "km/L"},
    "전기":              {"eff": 5.22,  "unit": "km/kWh"},
}
# 플러그인하이브리드는 규정상 연비(10.61 km/L)와 전비(2.84 km/kWh)가 모두 제시됨.
# 여기서는 유류(휘발유) 운행 기준의 연비(10.61 km/L)를 기본값으로 사용.
# 전기 운행분으로 계산하려면 유종을 '전기'로 바꾸거나 아래 값을 참고해 수기 수정.
PHEV_ELEC_EFF = 2.84  # 플러그인하이브리드 전비(참고용)

# 전기차 전기요금 기본 단가(원/kWh). 공용 급속 대략값. 오피넷 미제공 → 수기 수정 가능.
DEFAULT_ELEC_PRICE = 300.0

# 유가 조회(오피넷)에서 쓰는 유종 ↔ 표준 유종 매핑
#   오피넷 유가는 휘발유/경유/자동차용LPG 3종만 제공.
OIL_PRICE_FUEL_MAP = {
    "휘발유": "휘발유",
    "경유": "경유",
    "LPG": "자동차용LPG",
    "하이브리드": "휘발유",         # 하이브리드는 휘발유가 기준
    "플러그인하이브리드": "휘발유",  # PHEV도 휘발유가 기준
    "전기": None,                  # 전기는 유가 아님(전기요금)
}

# 숙박비 상한 지역 구분 라벨 → 상수 키 매핑
SUKBAK_REGION = {
    "서울특별시": "숙박비_상한_서울",
    "광역시": "숙박비_상한_광역시",
    "그 밖의 지역": "숙박비_상한_그밖",
}

# ---- 오피넷 유종 코드 ----
OPINET_PRODUCTS = {
    "휘발유": "B027",
    "고급휘발유": "B034",
    "경유": "D047",
    "실내등유": "C004",
    "자동차용LPG": "K015",
}

# ---- 오피넷 시도 코드 ----
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
# 데이터 로드
# =============================================================
@st.cache_data(ttl=1800)
def load_oil_history() -> pd.DataFrame:
    """
    유가 누적 CSV 로드 (GitHub Actions가 매일 수집).
    컬럼: 날짜, 지역(전국/경북), 유종, 가격
    아직 파일이 없으면 빈 DataFrame 반환 (앱은 정상 작동).
    """
    try:
        df = pd.read_csv("data/oil_history.csv")
        df["날짜"] = df["날짜"].astype(str)
        return df
    except Exception:
        return pd.DataFrame(columns=["날짜", "지역", "유종", "가격"])


def lookup_oil_price(hist: pd.DataFrame, date_str: str, region: str, fuel: str):
    """누적 CSV에서 특정 날짜·지역·유종의 유가를 찾는다. 없으면 None."""
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
    """전국 평균 유가 (avgAllPrice)."""
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
    """시도별 평균 유가 (avgSidoPrice)."""
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
    """시군구별 평균 유가 (avgSigunPrice)."""
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
    """
    주소 → 좌표 + 실패 사유.
    반환: {'ok': bool, 'xy': (x,y)|None, 'reason': str, 'method': str}

    1차: 주소 검색 API(address.json)
    2차: 1차 실패 시 키워드 검색 API(keyword.json)로 폴백
         (도로명 상세가 정확히 안 맞거나 건물명/약식 주소일 때 대응)
    """
    if not KAKAO_REST_KEY:
        return {"ok": False, "xy": None, "method": "-",
                "reason": "카카오 REST 키가 secrets에 설정되지 않았습니다."}
    if not address or not address.strip():
        return {"ok": False, "xy": None, "method": "-", "reason": "주소가 비어 있습니다."}

    addr = address.strip()
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_KEY}"}

    # --- 1차: 주소 검색 ---
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

    # --- 2차: 키워드 검색 폴백 ---
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
    """주소 → (경도 x, 위도 y). 실패 시 None. (일괄 처리용 간단 래퍼)"""
    result = geocode_debug(address)
    return result["xy"] if result["ok"] else None


@st.cache_data(ttl=86400)
def get_road_distance(origin_xy, dest_xy):
    """
    좌표 → 도로 주행거리(m, 편도), 소요시간(s), 통행료(원, 편도).
    카카오모빌리티 자동차 길찾기 API. ★사용 권한 승인 필요★
    반환: {"ok": True, ...} 또는 {"ok": False, "reason": "..."}

    ★ 여기서 나오는 distance_m·toll은 모두 '편도' 기준.
      출장은 왕복이므로 유류비·통행료 계산 시 ×2 처리는 호출부에서 수행.
    """
    if not KAKAO_REST_KEY:
        return {"ok": False, "reason": "카카오 키가 설정되지 않았습니다."}
    if not origin_xy or not dest_xy:
        return {"ok": False, "reason": "좌표가 없습니다."}
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_KEY}"}
    params = {
        "origin": f"{origin_xy[0]},{origin_xy[1]}",
        "destination": f"{dest_xy[0]},{dest_xy[1]}",
        "priority": "RECOMMEND",
    }
    try:
        r = requests.get(KAKAO_DIRECTIONS_URL, headers=headers, params=params, timeout=10)
        # 권한 미승인 시 401/403 계열
        if r.status_code in (401, 403):
            return {"ok": False, "reason": "길찾기 API 사용 권한이 아직 승인되지 않았습니다. "
                                           "카카오 데브톡 승인 후 자동으로 작동합니다."}
        r.raise_for_status()
        summary = r.json()["routes"][0]["summary"]
        return {
            "ok": True,
            "distance_m": summary["distance"],                 # 편도(m)
            "duration_s": summary["duration"],
            "toll": summary.get("fare", {}).get("toll", 0),    # 편도 통행료
        }
    except Exception as e:
        return {"ok": False, "reason": f"길찾기 호출 오류: {e}"}


# =============================================================
# 계산 로직
# =============================================================
def calc_fuel_cost(distance_km: float, efficiency: float, unit_price: float) -> float:
    """
    운행비 = 거리 ÷ 연비(또는 전비) × 단가.
      - 내연차: 연비(km/L) × 유가(원/L)
      - 전기차: 전비(km/kWh) × 전기요금(원/kWh)
    distance_km 에는 '왕복' 거리를 넣어야 왕복 운임이 나온다.
    """
    if efficiency <= 0:
        return 0.0
    return (distance_km / efficiency) * unit_price


def calculate_travel_allowance(
    *,
    is_gwannae: bool,          # 관내(근무지 내) 출장 여부
    hours_over_4: bool,        # (관내) 4시간 이상 여부
    use_official_car: bool,    # 공무용 차량 이용 여부
    days: int,                 # 출장일수(관외 일비·식비 계산용)
    nights: int,               # 숙박 밤 수
    sukbak_region_key: str,    # 숙박비 상한 지역 키
    fuel_cost: float,          # 운임(자가차 유류/전기비, 왕복). 운임 대체로 사용 가능
    manual_transport: float,   # 운임(대중교통 등) 수기 입력값
    toll: float = 0.0,         # 통행료(왕복, 수기 수정 반영된 최종값) — 별도 항목
) -> dict:
    """
    「공무원 여비 규정」 준용 여비 산출.

    - 관내(근무지 내, 제18조): 정액 지급.
        4시간 이상 20,000 / 미만 10,000, 공무용 차량 이용 시 10,000 감액.
        (일비·식비·숙박비 별도 지급 없음)
    - 관외(근무지 외, 별표2): 운임 + 일비 + 식비 + 숙박비 + 통행료 합산.
        일비 = 일수 × 단가 (공무용 차량 이용 시 1/2)
        식비 = 일수 × 단가
        숙박비 = 밤 수 × 지역별 상한(실비, 상한 이내)
        운임 = 대중교통 수기입력, 자가차면 왕복 유류/전기비를 운임 자리에 사용
        통행료 = 왕복 기본값(편도×2)을 사용자가 수기 수정 (국도 이용 시 0)

    반환: 각 항목 금액과 합계 딕셔너리
    """
    result = {}

    if is_gwannae:
        base = (ALLOWANCE_RATES["관내_4시간이상"] if hours_over_4
                else ALLOWANCE_RATES["관내_4시간미만"])
        if use_official_car:
            base = max(0, base - ALLOWANCE_RATES["관내_공무용차량감액"])
        result["구분"] = "근무지 내(관내)"
        result["관내정액"] = base
        result["합계"] = base
        return result

    # --- 관외 ---
    ilbi_unit = ALLOWANCE_RATES["일비_1일"]
    ilbi = ilbi_unit * days
    if use_official_car:
        ilbi = ilbi // 2  # 공무용 차량 이용 시 일비 1/2 (제16조 제3항)

    sikbi = ALLOWANCE_RATES["식비_1일"] * days

    sukbak_cap = ALLOWANCE_RATES.get(sukbak_region_key, ALLOWANCE_RATES["숙박비_상한_그밖"])
    sukbak = sukbak_cap * max(0, nights)

    # 운임: 자가차면 왕복 유류/전기비, 대중교통이면 수기입력. 둘 다 있으면 수기입력 우선.
    transport = manual_transport if manual_transport > 0 else fuel_cost

    # 통행료: 별도 항목으로 합산 (운임에 섞지 않음)
    toll_val = max(0, int(round(toll)))

    total = int(round(ilbi + sikbi + sukbak + transport + toll_val))
    result.update({
        "구분": "근무지 외(관외)",
        "운임": int(round(transport)),
        "일비": int(ilbi),
        "식비": int(sikbi),
        "숙박비": int(sukbak),
        "통행료": toll_val,
        "합계": total,
    })
    return result


# =============================================================
# UI
# =============================================================
st.title("🚗 공무원 출장 여비 산정 시스템")
st.caption("골격 버전 · 저장 기능 없음(세션 유지) · 여비 조례/출력양식은 실물 자료 확보 후 반영")

# API 키 상태 표시
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

tab1, tab2, tab3 = st.tabs(["⛽ 유가 조회", "🧮 여비 계산", "📋 명단 일괄 처리"])

# -------------------------------------------------------------
# TAB 1 : 유가 조회
# -------------------------------------------------------------
with tab1:
    st.subheader("오피넷 유가 조회")
    view = st.radio("조회 범위", ["전국 평균", "시도별", "시군별"], horizontal=True)

    if view == "전국 평균":
        df = fetch_oil_price_nationwide()
        if not df.empty:
            st.dataframe(df, width="stretch")
        else:
            st.info("오피넷 API 키를 secrets에 설정하면 조회됩니다.")

    elif view == "시도별":
        df = fetch_oil_price_sido()
        if not df.empty:
            st.dataframe(df, width="stretch")
        else:
            st.info("오피넷 API 키를 secrets에 설정하면 조회됩니다.")

    else:  # 시군별
        cc1, cc2 = st.columns(2)
        sido_name = cc1.selectbox("시도", list(OPINET_SIDO.keys()),
                                  index=list(OPINET_SIDO.keys()).index("경북"))
        prod_name = cc2.selectbox("유종", list(OPINET_PRODUCTS.keys()))
        if st.button("시군별 조회"):
            df = fetch_oil_price_sigun(OPINET_SIDO[sido_name], OPINET_PRODUCTS[prod_name])
            if not df.empty:
                st.dataframe(df, width="stretch")
            else:
                st.info("결과가 없거나 API 키가 설정되지 않았습니다.")

    st.caption("※ 유가는 일 단위 갱신. 2개월치 누적은 별도 수집(GitHub Actions 등)으로 확장 예정.")

# -------------------------------------------------------------
# TAB 2 : 여비 계산 (단건)
# -------------------------------------------------------------
with tab2:
    st.subheader("단건 여비 계산")

    # --- 세션 유지: 공통값(출발지 등) 재사용 ---
    st.markdown("##### 1) 출장 경로")
    c1, c2 = st.columns(2)
    origin = c1.text_input("출발지 주소", value=st.session_state.get("origin", ""),
                           placeholder="예) 경상북도 안동시 풍천면 도청대로 455")
    dest = c2.text_input("도착지 주소", placeholder="예) 경상북도 경주시 ○○면 ○○리 123")
    st.session_state["origin"] = origin  # 다음 건에서 출발지 재사용

    st.markdown("##### 2) 차량 유종 / 연비 (규정 표준값)")
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

    st.markdown("##### 3) 단가 (유가 또는 전기요금)")
    hist = load_oil_history()

    if is_electric:
        # 전기: 전기요금(원/kWh) 수기 입력, 기본값 300원/kWh
        c_e1, c_e2 = st.columns([1, 3])
        oil_price = c_e1.number_input("전기요금(원/kWh)", min_value=0.0,
                                      value=DEFAULT_ELEC_PRICE, step=10.0,
                                      help="오피넷 미제공 항목이라 수기 입력입니다. "
                                           "기본값은 공용 급속 대략 단가(300원/kWh).")
        c_e2.caption("ⓘ 전기차는 유가가 아니라 전기요금(원/kWh)으로 계산합니다. "
                     "충전 방식(급속/완속·공용/자가)에 따라 단가가 크게 다르니 실제 단가로 수정하세요.")
        region = "전국"  # 전기는 지역 유가 불필요(자리만 유지)
        sel_date = None
        auto_price = None
    else:
        c6, c7, c8, c9 = st.columns([1.2, 1, 1.2, 1])

        # 오피넷 유가 조회용 유종으로 변환 (하이브리드→휘발유 등)
        oil_fuel = OIL_PRICE_FUEL_MAP.get(fuel_type, "휘발유")
        c6.text_input("유가 기준 유종", value=oil_fuel, disabled=True,
                      help="하이브리드·PHEV는 휘발유 유가를 기준으로 합니다.")
        region = c7.selectbox("지역 기준", ["전국", "경북"],
                              help="여비 규정에 명시된 유가 기준에 맞춰 선택하세요.")

        # 누적 CSV에 있는 날짜 목록 (최신순). 없으면 안내.
        if not hist.empty:
            avail_dates = sorted(hist["날짜"].unique().tolist(), reverse=True)
        else:
            avail_dates = []

        auto_price = None
        if avail_dates:
            sel_date = c8.selectbox("유가 기준일", avail_dates,
                                    help="누적 수집된 날짜 중 선택. 최대 약 2개월치.")
            auto_price = lookup_oil_price(hist, sel_date, region, oil_fuel)
        else:
            c8.selectbox("유가 기준일", ["(누적 데이터 없음)"], disabled=True)
            sel_date = None

        default_price = auto_price if auto_price is not None else 0.0
        oil_price = c9.number_input("적용 유가(원/L)", min_value=0.0,
                                    value=float(default_price), step=1.0,
                                    help="선택한 날짜·지역·유종의 오피넷 값이 자동 입력됩니다. 수정 가능.")

        if avail_dates and auto_price is not None:
            st.caption(f"✅ {sel_date} · {region} · {oil_fuel} 유가 자동 적용: {auto_price:,.2f} 원/L")
        elif avail_dates and auto_price is None:
            st.caption(f"⚠️ {sel_date} · {region} · {oil_fuel} 데이터가 없어 수기 입력이 필요합니다.")
        else:
            st.caption("ⓘ 유가 누적 데이터가 아직 쌓이지 않았습니다. GitHub Actions가 매일 수집하며, "
                       "며칠 후부터 날짜 선택이 가능해집니다. 그전까지는 유가를 수기 입력하세요.")

    # --- 계산 (단계별 에러 구분) ---
    if st.button("거리 · 운임 계산", type="primary"):
        if not origin or not dest:
            st.error("출발지와 도착지 주소를 모두 입력하세요.")
        else:
            o = geocode_debug(origin)
            d = geocode_debug(dest)
            # 1단계: 지오코딩 실패 시 구체적 사유 표시
            if not o["ok"] or not d["ok"]:
                if not o["ok"]:
                    st.error(f"출발지 실패 [{o['method']}] → {o['reason']}")
                if not d["ok"]:
                    st.error(f"도착지 실패 [{d['method']}] → {d['reason']}")
                st.caption("↑ 위 상단 '🔑 API 연결 상태 확인 / 지오코딩 진단' 을 펼쳐 "
                           "단독 테스트로 원인을 좁힐 수 있습니다.")
            else:
                o_xy, d_xy = o["xy"], d["xy"]
                st.success(f"주소 좌표 변환 성공 ✓ (출발지:{o['method']} / 도착지:{d['method']})")
                # 2단계: 길찾기
                route = get_road_distance(o_xy, d_xy)
                if not route["ok"]:
                    st.warning(f"도로거리 계산 대기: {route['reason']}")
                    st.info("주소 인식은 정상입니다. 길찾기 승인 후 이 버튼만 다시 누르면 거리가 나옵니다.")
                else:
                    dist_km_oneway = route["distance_m"] / 1000
                    dist_km_round = dist_km_oneway * 2  # 왕복
                    st.session_state["last_dist_km_oneway"] = dist_km_oneway
                    st.session_state["last_dist_km_round"] = dist_km_round
                    # 편도 통행료 → 왕복 기본값 저장
                    st.session_state["last_toll_oneway"] = int(route["toll"])
                    st.session_state["last_toll_roundtrip"] = int(route["toll"]) * 2

                    fuel_cost = calc_fuel_cost(dist_km_round, eff, oil_price)  # 왕복 운임
                    m1, m2, m3 = st.columns(3)
                    m1.metric("도로거리(편도)", f"{dist_km_oneway:,.1f} km")
                    m2.metric("왕복거리", f"{dist_km_round:,.1f} km")
                    unit_txt = "전기비" if is_electric else "유류비"
                    m3.metric(f"{unit_txt}(왕복)", f"{fuel_cost:,.0f} 원")
                    st.caption(f"예상 소요시간(편도): {route['duration_s']//60} 분")
                    if route["toll"]:
                        st.caption(f"통행료(편도): {route['toll']:,} 원 · 왕복 기본값 {route['toll']*2:,} 원 "
                                   "(아래 여비 산정 통행료 칸에 자동 반영, 수정 가능)")

    # --- 여비 항목 (관내/관외 분기) ---
    st.markdown("##### 4) 여비 산정")
    st.caption("「공무원 여비 규정」 준용. 단가는 상단 ALLOWANCE_RATES 기본값(국가 규정)이며, "
               "경상북도 여비 조례 값으로 최종 확인·교체하세요.")

    # 관내/관외 자동 추천 (편도 거리 기준). 계산된 거리가 있으면 참고로 안내.
    last_km_oneway = st.session_state.get("last_dist_km_oneway")
    last_km_round = st.session_state.get("last_dist_km_round")
    auto_gwannae = None
    if last_km_oneway is not None:
        auto_gwannae = last_km_oneway < ALLOWANCE_RATES["관내_거리기준_km"]
        hint = (f"이번 경로 편도 {last_km_oneway:,.1f}km → "
                f"{'관내(12km 미만)' if auto_gwannae else '관외(12km 이상)'} 추정")
        st.caption(f"ⓘ {hint}. 같은 시·군이면 거리와 무관하게 관내입니다. 아래에서 직접 선택하세요.")

    g1, g2 = st.columns(2)
    trip_type = g1.radio(
        "출장 구분",
        ["근무지 외(관외)", "근무지 내(관내)"],
        index=(1 if auto_gwannae else 0),
        horizontal=True,
        help="같은 시·군 안 또는 여행거리 12km 미만이면 관내입니다.",
    )
    use_car = g2.checkbox("공무용 차량 이용",
                          help="관내: 1만원 감액 / 관외: 일비 1/2 지급")

    is_gwannae = (trip_type == "근무지 내(관내)")

    if is_gwannae:
        hours_over_4 = st.radio("출장 시간", ["4시간 이상", "4시간 미만"],
                                horizontal=True) == "4시간 이상"
        days = 0
        nights = 0
        sukbak_region_key = "숙박비_상한_그밖"
        manual_transport = 0.0
        toll_input = 0.0
    else:
        hours_over_4 = True
        e1, e2, e3 = st.columns(3)
        days = e1.number_input("출장일수(일비·식비)", min_value=0, value=1, step=1)
        nights = e2.number_input("숙박 밤 수", min_value=0, value=0, step=1)
        sukbak_region_label = e3.selectbox(
            "숙박지역(상한)", list(SUKBAK_REGION.keys()),
            index=2, help="서울 10만 / 광역시 8만 / 그 밖 7만 (실비, 상한 이내)")
        sukbak_region_key = SUKBAK_REGION[sukbak_region_label]

        f1, f2 = st.columns(2)
        manual_transport = f1.number_input(
            "운임(대중교통 등, 원) — 비우면 위 왕복 운임 사용", min_value=0, value=0, step=1000,
            help="자가차 출장이면 비워두세요(위에서 계산된 왕복 유류/전기비가 운임 자리에 들어갑니다). "
                 "KTX·버스 등 대중교통이면 실비를 입력하세요.")

        # 통행료: 카카오 왕복 기본값(편도×2)을 채우되 수기 수정 가능. 국도 이용이면 0으로.
        default_toll = int(st.session_state.get("last_toll_roundtrip", 0))
        toll_input = f2.number_input(
            "통행료(왕복, 원)", min_value=0, value=default_toll, step=100,
            help="카카오 경로 통행료의 왕복 기본값(편도×2)이 자동 입력됩니다. "
                 "국도 이용·하이패스 할인 등으로 다르면 직접 수정(국도면 0).")
        if default_toll > 0:
            st.caption(f"ⓘ 통행료 왕복 기본값 {default_toll:,} 원 자동 반영됨 "
                       f"(편도 {st.session_state.get('last_toll_oneway', 0):,} 원 × 2). 필요 시 수정.")

    # 운임(왕복 유류/전기비): 직전 계산의 왕복거리로 재계산해서 사용
    fuel_cost_val = 0.0
    if last_km_round is not None:
        fuel_cost_val = calc_fuel_cost(last_km_round, eff, oil_price)

    if st.button("여비 합계 계산", type="primary", key="calc_allowance"):
        res = calculate_travel_allowance(
            is_gwannae=is_gwannae,
            hours_over_4=hours_over_4,
            use_official_car=use_car,
            days=int(days),
            nights=int(nights),
            sukbak_region_key=sukbak_region_key,
            fuel_cost=fuel_cost_val,
            manual_transport=float(manual_transport),
            toll=float(toll_input),
        )
        st.markdown(f"**구분: {res['구분']}**")
        if is_gwannae:
            st.metric("관내 정액 여비", f"{res['관내정액']:,} 원")
        else:
            cc = st.columns(5)
            cc[0].metric("운임(왕복)", f"{res['운임']:,} 원")
            cc[1].metric("일비", f"{res['일비']:,} 원")
            cc[2].metric("식비", f"{res['식비']:,} 원")
            cc[3].metric("숙박비", f"{res['숙박비']:,} 원")
            cc[4].metric("통행료", f"{res['통행료']:,} 원")
        st.success(f"여비 합계: {res['합계']:,} 원")
        st.caption("※ 단가·관내 판정·차량 감액은 경상북도 여비 조례로 최종 확인하세요. "
                   "숙박비는 실비 정산(상한 이내), 통행료는 실비가 원칙입니다.")

    st.info("정식 여비 정산서(엑셀/HWPX) 출력 양식은 실물 양식 확보 후 셀 매핑으로 연결합니다.")

# -------------------------------------------------------------
# TAB 3 : 명단 일괄 처리
# -------------------------------------------------------------
with tab3:
    st.subheader("출장자 명단 일괄 처리")
    st.markdown(
        "명단 엑셀을 업로드하면 전원 거리·운임을 한 번에 계산합니다. "
        "**업로드 파일은 처리만 하며 서버에 저장하지 않습니다.**"
    )

    # 템플릿 다운로드
    template = pd.DataFrame({
        "성명": ["홍길동"],
        "직급": ["주무관"],
        "출발지": ["경상북도 안동시 풍천면 도청대로 455"],
        "도착지": ["경상북도 경주시 ○○로 123"],
        "유종": ["휘발유"],
        "출장일수": [1],
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        template.to_excel(w, index=False, sheet_name="명단")
    st.download_button("📥 명단 양식 다운로드", buf.getvalue(),
                       file_name="출장자_명단_양식.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("※ 유종은 휘발유/경유/LPG/하이브리드/플러그인하이브리드/전기 중 입력. "
               "연비는 규정 표준값이 자동 적용됩니다.")

    up = st.file_uploader("작성한 명단 업로드 (.xlsx)", type=["xlsx"])
    if up is not None:
        try:
            df_in = pd.read_excel(up)
            st.write("업로드된 명단:")
            st.dataframe(df_in, width="stretch")

            if st.button("전원 계산 실행", type="primary"):
                results = []
                prog = st.progress(0.0)
                for i, r in df_in.iterrows():
                    o_xy = geocode(str(r.get("출발지", "")))
                    d_xy = geocode(str(r.get("도착지", "")))
                    route = get_road_distance(o_xy, d_xy) if (o_xy and d_xy) else None
                    ok = bool(route and route.get("ok"))
                    dist_km_oneway = route["distance_m"] / 1000 if ok else None
                    dist_km_round = dist_km_oneway * 2 if dist_km_oneway is not None else None
                    toll_oneway = int(route["toll"]) if ok else None

                    # 유종 → 규정 표준 연비
                    ft = str(r.get("유종", "휘발유")).strip()
                    std = FUEL_STANDARDS.get(ft)
                    eff_val = std["eff"] if std else None
                    eff_unit = std["unit"] if std else "-"

                    results.append({
                        "성명": r.get("성명"),
                        "직급": r.get("직급"),
                        "출발지": r.get("출발지"),
                        "도착지": r.get("도착지"),
                        "유종": ft,
                        "표준연비/전비": f"{eff_val} {eff_unit}" if eff_val else "유종 확인",
                        "편도(km)": round(dist_km_oneway, 1) if dist_km_oneway else "조회실패",
                        "왕복(km)": round(dist_km_round, 1) if dist_km_round else "조회실패",
                        "통행료(왕복,원)": toll_oneway * 2 if toll_oneway is not None else "조회실패",
                        "출장일수": r.get("출장일수"),
                    })
                    prog.progress((i + 1) / len(df_in))
                    time.sleep(0.05)  # API 과호출 방지

                res_df = pd.DataFrame(results)
                st.success("계산 완료")
                st.dataframe(res_df, width="stretch")

                # 결과 엑셀 다운로드
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine="openpyxl") as w:
                    res_df.to_excel(w, index=False, sheet_name="여비산정결과")
                st.download_button("📤 결과 엑셀 다운로드", out.getvalue(),
                                   file_name="여비산정_결과.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"파일 처리 오류: {e}")

    st.caption("※ 여비 최종 합산·조례 반영·정식 출력양식은 실물 자료 확보 후 반영 예정.")
