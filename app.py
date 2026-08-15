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
from datetime import date
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
PHEV_ELEC_EFF = 2.84  # 플러그인하이브리드 전비(참고용)

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
DOCHEONG_ADDR = "경상북도 안동시 풍천면 도청대로 455"  # 경상북도청


def is_from_docheong(origin: str) -> bool:
    """출발지가 경북도청인지 대략 판정(핵심어 포함 여부)."""
    if not origin:
        return False
    o = origin.replace(" ", "")
    return ("도청대로455" in o) or ("경상북도청" in o) or ("풍천면도청대로" in o)


def check_gwannae_candidate(origin: str, dest: str, dist_km_oneway):
    """
    관내 후보 판정 (자동전환 X, 제안용).
    기준:
      1) 출발지=도청 AND 도착지에 '예천군 호명읍' 포함  → 호명읍 특례
      2) 편도거리 < 12km                                 → 거리 기준
    반환: (is_candidate: bool, reason: str)
    """
    reasons = []
    dest_c = (dest or "").replace(" ", "")
    if is_from_docheong(origin):
        # 도청 소재지(안동시)는 항상 관내
        if "안동시" in dest_c:
            reasons.append("도청 출발 + 안동시(도청 소재지, 관내)")
        # 예천군 호명읍 관내 특례
        if "예천군호명읍" in dest_c or "호명읍" in dest_c:
            reasons.append("도청 출발 + 예천군 호명읍(관내 특례)")
    if dist_km_oneway is not None and dist_km_oneway < ALLOWANCE_RATES["관내_거리기준_km"]:
        reasons.append(f"편도 {dist_km_oneway:,.1f}km < 12km")
    return (len(reasons) > 0, " · ".join(reasons))


def _switch_to_gwannae():
    """관내 전환 버튼 콜백. 콜백은 재실행 사이클 앞단에서 실행돼 DOM 충돌이 없다."""
    st.session_state["force_gwannae"] = True
    st.session_state.pop("gwannae_candidate", None)
    st.session_state.pop("gwannae_candidate_reason", None)


# =============================================================
# 데이터 로드
# =============================================================
@st.cache_data(ttl=1800)
def load_oil_history() -> pd.DataFrame:
    """유가 누적 CSV 로드. 없으면 빈 DataFrame."""
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
    """일괄처리용: 좌표와 실패사유를 함께 반환(캐시)."""
    return geocode_debug(address)


@st.cache_data(ttl=86400)
def get_road_distance(origin_xy, dest_xy):
    """좌표 → 도로 주행거리(m, 편도), 소요시간(s), 통행료(원, 편도)."""
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
        if r.status_code in (401, 403):
            return {"ok": False, "reason": "길찾기 API 사용 권한이 아직 승인되지 않았습니다. "
                                           "카카오 데브톡 승인 후 자동으로 작동합니다."}
        r.raise_for_status()
        summary = r.json()["routes"][0]["summary"]
        return {
            "ok": True,
            "distance_m": summary["distance"],
            "duration_s": summary["duration"],
            "toll": summary.get("fare", {}).get("toll", 0),
        }
    except Exception as e:
        return {"ok": False, "reason": f"길찾기 호출 오류: {e}"}


# =============================================================
# 계산 로직
# =============================================================
def floor10(value) -> int:
    """십원 미만 버림(일의 자리 절사). 예: 1,863 → 1,860, 25,000 → 25,000."""
    try:
        return (int(value) // 10) * 10
    except (TypeError, ValueError):
        return 0


def calc_fuel_cost(distance_km: float, efficiency: float, unit_price: float) -> float:
    """운행비 = 거리 ÷ 연비(전비) × 단가. distance_km 에는 최종 주행거리를 넣는다."""
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
    num_passengers: int = 0,   # 동승자 수 (운전자 제외)
) -> dict:
    """
    「공무원 여비 규정」 준용 여비 산출 (운전자 + 동승자 분리).

    운전자: 운임 + 일비 + 식비 + 숙박비 + 통행료
    동승자: 일비 + 식비 + 숙박비 (운임·통행료 없음 — 같은 차 이용)
      * 관내 정액 출장은 동승자 개념을 적용하지 않음(정액 1인 기준).

    반환:
      - 관내: {구분, 관내정액, 합계}
      - 관외: {구분, 운전자:{...}, 동승자단가:{...}, 동승자수, 동승자합계, 총합계}
    """
    result = {}

    if is_gwannae:
        base = (ALLOWANCE_RATES["관내_4시간이상"] if hours_over_4
                else ALLOWANCE_RATES["관내_4시간미만"])
        if use_official_car:
            base = max(0, base - ALLOWANCE_RATES["관내_공무용차량감액"])
        base = floor10(base)
        # 관내는 정액. 운전자·동승자 모두 동일 금액(감액도 동일 적용).
        pax_n = max(0, int(num_passengers))
        result.update({
            "구분": "근무지 내(관내)",
            "관내정액": base,          # 1인 정액(하위호환)
            "운전자정액": base,
            "동승자정액": base,
            "동승자수": pax_n,
            "동승자합계": base * pax_n,
            "합계": base,              # 운전자 1인 기준(하위호환)
            "총합계": base + base * pax_n,
        })
        return result

    # --- 관외 ---
    ilbi_unit = ALLOWANCE_RATES["일비_1일"]
    ilbi = ilbi_unit * days
    if use_official_car:
        ilbi = ilbi // 2  # 공무용 차량 이용 시 일비 1/2 (제16조 제3항)

    sikbi = ALLOWANCE_RATES["식비_1일"] * days

    sukbak_cap = ALLOWANCE_RATES.get(sukbak_region_key, ALLOWANCE_RATES["숙박비_상한_그밖"])
    sukbak = sukbak_cap * max(0, nights)

    # 운임: 공무용 차량이면 운임 미지급(제15조) → 0.
    #       자가차면 유류/전기비(배수 반영), 대중교통이면 수기입력.
    if use_official_car:
        transport = 0
        toll_val = 0            # 관용차는 통행료도 기관 부담 → 개인 여비 0
    else:
        transport = manual_transport if manual_transport > 0 else fuel_cost
        transport = int(round(transport))
        toll_val = max(0, int(round(toll)))

    # 각 항목 십원 미만 버림
    ilbi = floor10(ilbi)
    sikbi = floor10(sikbi)
    sukbak = floor10(sukbak)
    transport = floor10(transport)
    toll_val = floor10(toll_val)

    # 운전자 1인
    driver = {
        "운임": transport,
        "일비": ilbi,
        "식비": sikbi,
        "숙박비": sukbak,
        "통행료": toll_val,
    }
    driver["합계"] = sum(driver.values())

    # 동승자 1인 단가 (운임·통행료 없음 — 같은 차 탑승)
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
# 도착지 시군 버튼 (직재순, 군위 제외) → 지오코딩 접두 주소 매핑
#   포항: 시청(본청) + 남구/북구. 세부주소 없어도 각 대표지점으로 지오코딩되도록
#   접두를 시청 실주소·남구·북구로 구분한다.
SIGUN_PREFIX = {
    "포항시청": "경상북도 포항시 남구 시청로 1",
    "포항남": "경상북도 포항시 남구",
    "포항북": "경상북도 포항시 북구",
    "경주": "경상북도 경주시",
    "김천": "경상북도 김천시",
    # 안동은 도청 소재지 → 항상 관내(정액 2만원)이므로 도착지 버튼에서 제외
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
SIGUN_ORDER = list(SIGUN_PREFIX.keys())  # 직재순
DEST_OPTIONS = SIGUN_ORDER + ["기타(직접입력)"]

DEFAULT_ORIGIN = "경상북도 안동시 풍천면 도청대로 455"  # 경상북도청


def build_dest_address(sigun_choice: str, detail: str) -> str:
    """
    시군 버튼 선택 + 세부주소 → 최종 도착지 주소.
    - '기타' 선택: 세부주소를 그대로 사용.
    - 시군 선택: 접두(예:'경상북도 안동시') + 세부주소.
      단 세부주소가 이미 '경상북도'로 시작하거나 시군명을 포함하면
      접두를 중복으로 붙이지 않고 세부주소를 그대로 쓴다.
    """
    detail = (detail or "").strip()
    if sigun_choice == "기타(직접입력)":
        return detail
    prefix = SIGUN_PREFIX.get(sigun_choice, "")
    # 포항시청은 고정 지점 → 세부주소 무시하고 시청 주소 그대로 사용
    if sigun_choice == "포항시청":
        return prefix
    if not detail:
        return prefix  # 세부주소 없으면 시군 대표지점(시/군청 인근)만
    # 접두 중복 방지: 이미 광역/시군명 들어간 전체주소를 넣은 경우
    core = prefix.replace("경상북도 ", "")  # 예: '안동시'
    if detail.startswith("경상북도") or core in detail:
        return detail
    return f"{prefix} {detail}"


with tab2:
    st.subheader("단건 여비 계산")

    # =========================================================
    # 0) 출장 구분 (최상단) — 관내면 경로·유종·유가 불필요
    # =========================================================
    # 관내 자동판정에서 '관내로 전환' 버튼을 누르면 force_gwannae 신호가 들어옴 → 기본 '관내' 선택
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

    # =========================================================
    # 관내(정액) : 경로/유종/유가 전부 생략, 최소 입력만
    # =========================================================
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

    # =========================================================
    # 관외 : 경로 → 출장자 → 출장목적 → 유종/연비 → 단가 → 여비 합계(한 번에)
    # =========================================================
    else:
        # --- 1) 출장 경로 ---
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

        # --- 2) 출장자 (운전자 + 동승자) ---
        st.markdown("##### 2) 출장자")
        nm1, nm2 = st.columns([1, 1])
        driver_name = nm1.text_input("운전자 성명", placeholder="예) 홍길동",
                                     help="엑셀 추출용")
        num_pax = nm2.number_input("동승자 수 (최대 4)", min_value=0, max_value=4, value=0, step=1)
        passenger_names = []
        if num_pax > 0:
            pcols = st.columns(int(num_pax))
            for i in range(int(num_pax)):
                pname = pcols[i].text_input(f"동승자 {i+1}", key=f"pax_{i}", placeholder="성명")
                passenger_names.append(pname)
        st.caption("ⓘ 동승자는 같은 차량 탑승 기준입니다. 운임·통행료는 운전자만, "
                   "일비·식비·숙박비는 전원 각자 지급됩니다.")

        # --- 3) 출장 목적 (운전자·동승자 공통 출력) ---
        # --- 3) 출장 목적 (운전자·동승자 공통 출력) ---
        st.markdown("##### 3) 출장 목적")
        trip_purpose = st.text_input(
            "출장 목적", placeholder="예) ○○ 지적재조사 지구 현장 확인",
            help="운전자·동승자 모두 동일하게 출력됩니다. (실물 양식 확보 후 위치·문구 조정 예정)",
        )

        # --- 4) 출장 기간 (시작일·종료일) ---
        #   유가 기준일 = 출장 시작일 (규정/실무: 출발일 유가 적용).
        #   시작일을 먼저 받아 아래 단가의 유가 기준일로 자동 연결한다.
        st.markdown("##### 4) 출장 기간")
        hours_over_4 = True

        # 종료일 자동 연동: 시작일을 바꾸면, 종료일이 시작일보다 앞서 있을 때만 시작일로 맞춘다.
        #   → 당일치기는 시작일만 누르면 종료일이 자동으로 같은 날이 되고,
        #     1박 이상은 종료일을 뒤로 조정하면 그대로 유지된다.
        def _sync_end_date():
            s = st.session_state.get("trip_start_key")
            e = st.session_state.get("trip_end_key")
            if s is not None and (e is None or e < s):
                st.session_state["trip_end_key"] = s

        if "trip_start_key" not in st.session_state:
            st.session_state["trip_start_key"] = date.today()
        if "trip_end_key" not in st.session_state:
            st.session_state["trip_end_key"] = date.today()

        d0, d1 = st.columns(2)
        trip_start = d0.date_input(
            "출장 시작일", key="trip_start_key", on_change=_sync_end_date,
            help="달력에서 출장 시작일을 고르세요. 종료일이 자동으로 같은 날로 맞춰집니다(당일치기 기본). "
                 "이 날짜가 유가 기준일로도 자동 적용됩니다.")
        trip_end = d1.date_input(
            "출장 종료일", key="trip_end_key",
            help="당일치기면 그대로 두세요. 1박 이상이면 종료일만 뒤로 조정하세요.")

        # 종료일 < 시작일 방지 (콜백 이후에도 안전장치)
        if trip_end < trip_start:
            st.error("종료일이 시작일보다 빠릅니다. 날짜를 다시 확인하세요.")
            trip_end = trip_start

        # 여행일수·밤수 자동 계산
        #   일비·식비 일수 = (종료일 - 시작일) + 1  (당일치기 = 1일)
        #   숙박 밤 수     = (종료일 - 시작일)       (당일치기 = 0박)
        span = (trip_end - trip_start).days
        days = span + 1
        nights = span
        st.markdown(
            f"""<div style="background:#eaf4ff;border:2px solid #2b7cff;border-radius:8px;
            padding:12px 16px;margin:8px 0;font-size:1.05rem;">
            📅 <b>출장 기간: {trip_start:%Y-%m-%d} ~ {trip_end:%Y-%m-%d}</b>
            &nbsp;→&nbsp; <b style="color:#c0392b;font-size:1.15rem;">{nights}박 {days}일</b>
            <br><span style="color:#333;">일비·식비 <b>{days}일</b>分 · 숙박 <b>{nights}박</b>分 자동 계산</span>
            </div>""",
            unsafe_allow_html=True,
        )

        # --- 5) 차량 유종 / 연비 ---
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

        # --- 6) 단가 (유가 기준일 = 출장 시작일 자동) ---
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
            # 유가 기준일 = 출장 시작일 (자동, 수정 불가 표시)
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

        # --- 7) 여비 산정 입력 (숙박지역·운임·통행료) ---
        st.markdown("##### 7) 여비 산정 입력")
        st.caption("「공무원 여비 규정」 준용. 단가는 상단 ALLOWANCE_RATES 기본값(국가 규정)입니다. "
                   "경상북도 별도 조례가 없어 국가 규정을 그대로 적용합니다. 운임은 왕복(편도×2)으로 계산합니다.")

        sukbak_region_label = st.selectbox(
            "숙박지역(상한)", list(SUKBAK_REGION.keys()),
            index=2, help="서울 10만 / 광역시 8만 / 그 밖 7만 (실비, 상한 이내)")
        sukbak_region_key = SUKBAK_REGION[sukbak_region_label]

        # 왕복 배수 결정: 기본 왕복 1회(×2). 단, 숙박 0박 + 일수 2일 이상이면 매일 출퇴근 → ×2×일수
        daily_commute = (int(nights) == 0 and int(days) >= 2)
        round_multiplier = 2 * (int(days) if daily_commute else 1)
        if daily_commute:
            st.caption(f"🚗 숙박 없이 {int(days)}일 출장 → **매일 출퇴근**으로 보아 운임·통행료를 "
                       f"편도 × 2(왕복) × {int(days)}일 = **×{round_multiplier}**로 계산합니다.")
        else:
            st.caption("🚗 운임·통행료는 **편도 × 2(왕복 1회)**로 계산합니다.")

        # 통행료: 카카오 편도값이 있으면 왕복배수 적용을 기본값으로, 수정 가능
        last_toll_oneway = int(st.session_state.get("last_toll_oneway", 0))
        applied_toll_default = last_toll_oneway * round_multiplier

        manual_transport = st.number_input(
            "운임(대중교통 등, 원) — 비우면 자가차 유류/전기비 사용", min_value=0, value=0, step=1000,
            disabled=use_car,
            help="자가차 출장이면 비워두세요(왕복 유류/전기비가 운임으로 들어갑니다). "
                 "KTX·버스 등 대중교통이면 실비를 입력하세요. "
                 "공무용 차량이면 운임은 0으로 처리되어 입력이 비활성화됩니다.")
        toll_input = st.number_input(
            "통행료(원, 왕복)", min_value=0, value=int(applied_toll_default), step=100,
            disabled=use_car,
            help="카카오 편도 통행료 × 왕복배수 값이 자동 입력됩니다. 국도 이용·하이패스 할인 등으로 다르면 수정(국도면 0). "
                 "공무용 차량이면 통행료도 0으로 처리됩니다.")
        if use_car:
            st.caption("ⓘ 공무용 차량 이용이므로 운임·통행료는 계산에서 0 처리됩니다.")
        elif last_toll_oneway > 0:
            st.caption(f"ⓘ 통행료 기본값 {applied_toll_default:,} 원 "
                       f"(직전 계산 편도 {last_toll_oneway:,} × {round_multiplier})")

        # --- 7) 여비 합계 계산 (거리→운임→여비 한 번에) ---
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
                        toll_oneway = int(route["toll"])
                        st.session_state["last_dist_km_oneway"] = dist_km_oneway
                        st.session_state["last_toll_oneway"] = toll_oneway

                        # 관내 자동판정 (버튼은 최상단에서 처리 → 여기선 세션에 신호만 저장)
                        cand, why = check_gwannae_candidate(origin, dest, dist_km_oneway)
                        if cand:
                            st.warning(f"🏠 **관내 출장 후보**입니다 — {why}\n\n"
                                       "같은 시·군 안이거나 여행거리 12km 미만이면 관내(정액)로 "
                                       "처리해야 할 수 있습니다. 아래 버튼으로 전환하거나, 관외 결과를 그대로 사용하세요.")
                            st.button("→ 관내(정액)로 전환", key="switch_to_gwannae",
                                      type="primary", on_click=_switch_to_gwannae)

                        # 왕복 주행거리
                        applied_dist = dist_km_oneway * round_multiplier
                        # 통행료: 사용자가 입력한 값 사용. (기본값은 편도×배수였음)
                        applied_toll = int(toll_input)

                        # 자가차 유류/전기비
                        fuel_cost_val = calc_fuel_cost(applied_dist, eff, oil_price)

                        # 경로·거리 요약
                        m1, m2, m3 = st.columns(3)
                        m1.metric("도로거리(편도)", f"{dist_km_oneway:,.1f} km")
                        m2.metric(f"적용거리(×{round_multiplier})", f"{applied_dist:,.1f} km")
                        m3.metric("통행료(왕복)", f"{applied_toll:,} 원")
                        st.caption(f"경로 좌표 변환 ✓ (출발지:{o['method']} / 도착지:{d['method']}) · "
                                   f"예상 소요(편도) {route['duration_s']//60}분")

                        # 여비 산출
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

        st.info("정식 여비 정산서(엑셀/HWPX) 출력 양식은 실물 양식 확보 후 셀 매핑으로 연결합니다. "
                "운전자·동승자 성명, 출장 목적은 그때 엑셀 추출에 사용됩니다.")


# -------------------------------------------------------------
# TAB 3 : 명단 일괄 처리
# -------------------------------------------------------------
with tab3:
    st.subheader("출장자 명단 일괄 처리")
    st.markdown(
        "명단 엑셀을 업로드하면 전원 거리·운임을 한 번에 계산합니다. "
        "**업로드 파일은 처리만 하며 서버에 저장하지 않습니다.**"
    )

    template = pd.DataFrame({
        "성명": ["홍길동"],
        "직급": ["주무관"],
        "출발지": ["경상북도 안동시 풍천면 도청대로 455"],
        "도착지": ["경상북도 경주시 ○○로 123"],
        "유종": ["휘발유"],
        "출장일수": [1],
        "운임배수": [2],
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        template.to_excel(w, index=False, sheet_name="명단")
    st.download_button("📥 명단 양식 다운로드", buf.getvalue(),
                       file_name="출장자_명단_양식.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("※ 유종은 휘발유/경유/LPG/하이브리드/플러그인하이브리드/전기 중 입력. "
               "연비는 규정 표준값이 자동 적용됩니다. 운임배수는 편도에 곱할 값(표준 왕복=2).")

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
                    origin_addr = str(r.get("출발지", ""))
                    dest_addr = str(r.get("도착지", ""))
                    o_res = geocode_full(origin_addr)
                    d_res = geocode_full(dest_addr)
                    o_xy = o_res["xy"] if o_res["ok"] else None
                    d_xy = d_res["xy"] if d_res["ok"] else None

                    # 실패사유 정리 (지오코딩 단계)
                    fail_reasons = []
                    if not o_res["ok"]:
                        fail_reasons.append(f"출발지[{o_res['method']}]:{o_res['reason']}")
                    if not d_res["ok"]:
                        fail_reasons.append(f"도착지[{d_res['method']}]:{d_res['reason']}")

                    route = get_road_distance(o_xy, d_xy) if (o_xy and d_xy) else None
                    ok = bool(route and route.get("ok"))
                    if (o_xy and d_xy) and not ok and route is not None:
                        fail_reasons.append(f"길찾기:{route.get('reason', '알 수 없음')}")

                    dist_km_oneway = route["distance_m"] / 1000 if ok else None
                    toll_oneway = int(route["toll"]) if ok else None

                    try:
                        mult = int(r.get("운임배수", 2))
                    except Exception:
                        mult = 2
                    if mult < 1:
                        mult = 1

                    dist_applied = dist_km_oneway * mult if dist_km_oneway is not None else None
                    toll_applied = toll_oneway * mult if toll_oneway is not None else None

                    ft = str(r.get("유종", "휘발유")).strip()
                    std = FUEL_STANDARDS.get(ft)
                    eff_val = std["eff"] if std else None
                    eff_unit = std["unit"] if std else "-"

                    # 관내 자동판정 (일괄에서도 사유 표기)
                    cand, why = check_gwannae_candidate(origin_addr, dest_addr, dist_km_oneway)

                    results.append({
                        "성명": r.get("성명"),
                        "직급": r.get("직급"),
                        "출발지": r.get("출발지"),
                        "도착지": r.get("도착지"),
                        "유종": ft,
                        "표준연비/전비": f"{eff_val} {eff_unit}" if eff_val else "유종 확인",
                        "편도(km)": round(dist_km_oneway, 1) if dist_km_oneway else "조회실패",
                        "운임배수": mult,
                        "적용거리(km)": round(dist_applied, 1) if dist_applied else "조회실패",
                        "통행료(원)": toll_applied if toll_applied is not None else "조회실패",
                        "출장일수": r.get("출장일수"),
                        "관내후보": "예" if cand else "아니오",
                        "관내판정사유": why if cand else "",
                        "실패사유": " / ".join(fail_reasons) if fail_reasons else "",
                    })
                    prog.progress((i + 1) / len(df_in))
                    time.sleep(0.05)

                res_df = pd.DataFrame(results)
                st.success("계산 완료")
                st.dataframe(res_df, width="stretch")

                # 실패 행 요약
                fail_cnt = int((res_df["실패사유"] != "").sum())
                if fail_cnt > 0:
                    st.warning(f"⚠️ 조회 실패 {fail_cnt}건 — '실패사유' 열에서 원인을 확인하세요. "
                               "(주소 형식 문제면 시·군 포함 전체주소로 수정, 401/403이면 카카오 앱 설정)")
                cand_cnt = int((res_df["관내후보"] == "예").sum())
                if cand_cnt > 0:
                    st.info(f"🏠 관내 출장 후보 {cand_cnt}건 — '관내판정사유' 열을 확인하세요. "
                            "(호명읍 특례 또는 편도 12km 미만)")

                out = io.BytesIO()
                with pd.ExcelWriter(out, engine="openpyxl") as w:
                    res_df.to_excel(w, index=False, sheet_name="여비산정결과")
                st.download_button("📤 결과 엑셀 다운로드", out.getvalue(),
                                   file_name="여비산정_결과.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"파일 처리 오류: {e}")

    st.caption("※ 여비 최종 합산·정식 출력양식은 실물 자료 확보 후 반영 예정.")
