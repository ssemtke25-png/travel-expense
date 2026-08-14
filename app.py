# -*- coding: utf-8 -*-
"""
공무원 출장 여비 산정 시스템 (골격 버전)
--------------------------------------------------
구성
  1) 유가 조회 탭      : 오피넷 무료 API (전국/시도/시군)
  2) 여비 계산 탭      : 주소→좌표(카카오 지오코딩) → 도로거리(카카오모빌리티)
                         → 연비 → 유류비, 그 위에 일비·식비·숙박비 3층 구조
  3) 일괄 처리 탭      : 출장자 명단 엑셀 업로드 → 전원 계산 → 엑셀 다운로드

설계 원칙
  - 저장 기능 없음 (세션 동안만 유지). 개인정보 서버 미보관.
  - 자동 계산값을 기본값으로 채우되, 모든 항목 수기 수정 가능 (3층 구조).
  - 여비 조례 단가/출력 양식은 사무실 실물 자료 확보 후 채울 자리만 확보.
  - 관용차 계산법은 추후 추가 (fuel_type 분기 자리 확보).

API 키는 .streamlit/secrets.toml 에서 로드.
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

# ---- 여비 조례 단가 (자리만 확보. 사무실 조례 별표로 교체) ----
# TODO: 경상북도 공무원 여비 조례·규칙 별표 값으로 채울 것
ALLOWANCE_RATES = {
    "일비_1일": 0,      # 예) 25000
    "식비_1일": 0,      # 예) 25000
    "숙박비_상한_1박": 0,  # 실비 정산, 상한만. 예) 지역별 상이
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
KAKAO_DIRECTIONS_URL = "https://apis-navi.kakaomobility.com/v1/directions"


# =============================================================
# 데이터 로드
# =============================================================
@st.cache_data
def load_fuel_efficiency() -> pd.DataFrame:
    """모델별 대표 연비 CSV 로드."""
    try:
        return pd.read_csv("data/fuel_efficiency.csv")
    except Exception:
        return pd.DataFrame(columns=["제조사", "모델명", "연료", "복합연비"])


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
@st.cache_data(ttl=86400)
def geocode(address: str):
    """주소 → (경도 x, 위도 y). 지번/도로명 모두 지원. 권한 신청 불필요."""
    if not KAKAO_REST_KEY:
        return None
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_KEY}"}
    params = {"query": address}
    try:
        r = requests.get(KAKAO_GEOCODE_URL, headers=headers, params=params, timeout=10)
        r.raise_for_status()
        docs = r.json().get("documents", [])
        if not docs:
            return None
        return float(docs[0]["x"]), float(docs[0]["y"])
    except Exception:
        return None


@st.cache_data(ttl=86400)
def get_road_distance(origin_xy, dest_xy):
    """
    좌표 → 도로 주행거리(m), 소요시간(s).
    카카오모빌리티 자동차 길찾기 API. ★사용 권한 승인 필요★
    캐싱으로 동일 경로 재호출 최소화.
    """
    if not KAKAO_REST_KEY or not origin_xy or not dest_xy:
        return None
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_KEY}"}
    params = {
        "origin": f"{origin_xy[0]},{origin_xy[1]}",
        "destination": f"{dest_xy[0]},{dest_xy[1]}",
        "priority": "RECOMMEND",
    }
    try:
        r = requests.get(KAKAO_DIRECTIONS_URL, headers=headers, params=params, timeout=10)
        r.raise_for_status()
        summary = r.json()["routes"][0]["summary"]
        return {
            "distance_m": summary["distance"],
            "duration_s": summary["duration"],
            "toll": summary.get("fare", {}).get("toll", 0),
        }
    except Exception as e:
        st.warning(f"길찾기 실패 (권한 승인 여부 확인): {e}")
        return None


# =============================================================
# 계산 로직
# =============================================================
def calc_fuel_cost(distance_km: float, efficiency: float, oil_price: float) -> float:
    """유류비 = 거리 ÷ 연비 × 유가."""
    if efficiency <= 0:
        return 0.0
    return (distance_km / efficiency) * oil_price


def calculate_travel_allowance(**kwargs) -> dict:
    """
    여비 최종 산출 (자리만 확보).
    TODO: 경북 여비 조례 로직 반영.
      - 근무지 내/외 구분(4시간 기준)
      - 일비·식비 = 일수 × 단가
      - 숙박비 = 실비(상한 이내)
      - 관용차 이용 시 유류비 처리 분기
    현재는 유류비만 반환하는 골격.
    """
    return {}


# =============================================================
# UI
# =============================================================
st.title("🚗 공무원 출장 여비 산정 시스템")
st.caption("골격 버전 · 저장 기능 없음(세션 유지) · 여비 조례/출력양식은 실물 자료 확보 후 반영")

# API 키 상태 표시
with st.expander("🔑 API 연결 상태 확인", expanded=False):
    c1, c2, c3 = st.columns(3)
    c1.metric("오피넷 유가", "연결됨" if OPINET_KEY else "미설정")
    c2.metric("카카오 지오코딩", "연결됨" if KAKAO_REST_KEY else "미설정")
    c3.metric("카카오 길찾기", "키 있음(승인 확인 필요)" if KAKAO_REST_KEY else "미설정")
    st.info("길찾기 API는 카카오 데브톡 사용 권한 승인 후 정상 작동합니다.")

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
            st.dataframe(df, use_container_width=True)
        else:
            st.info("오피넷 API 키를 secrets에 설정하면 조회됩니다.")

    elif view == "시도별":
        df = fetch_oil_price_sido()
        if not df.empty:
            st.dataframe(df, use_container_width=True)
        else:
            st.info("오피넷 API 키를 secrets에 설정하면 조회됩니다.")

    else:  # 시군별
        cc1, cc2 = st.columns(2)
        sido_name = cc1.selectbox("시도", list(OPINET_SIDO.keys()), index=list(OPINET_SIDO.keys()).index("경북"))
        prod_name = cc2.selectbox("유종", list(OPINET_PRODUCTS.keys()))
        if st.button("시군별 조회"):
            df = fetch_oil_price_sigun(OPINET_SIDO[sido_name], OPINET_PRODUCTS[prod_name])
            if not df.empty:
                st.dataframe(df, use_container_width=True)
            else:
                st.info("결과가 없거나 API 키가 설정되지 않았습니다.")

    st.caption("※ 유가는 일 단위 갱신. 2개월치 누적은 별도 수집(GitHub Actions 등)으로 확장 예정.")

# -------------------------------------------------------------
# TAB 2 : 여비 계산 (단건)
# -------------------------------------------------------------
with tab2:
    st.subheader("단건 여비 계산")
    fuel_df = load_fuel_efficiency()

    # --- 세션 유지: 공통값(출발지 등) 재사용 ---
    st.markdown("##### 1) 출장 경로")
    c1, c2 = st.columns(2)
    origin = c1.text_input("출발지 주소", value=st.session_state.get("origin", ""),
                           placeholder="예) 경상북도 안동시 풍천면 도청대로 455")
    dest = c2.text_input("도착지 주소", placeholder="예) 경상북도 경주시 ○○면 ○○리 123")
    st.session_state["origin"] = origin  # 다음 건에서 출발지 재사용

    st.markdown("##### 2) 차량 / 연비")
    c3, c4, c5 = st.columns(3)
    makers = ["(선택)"] + sorted(fuel_df["제조사"].unique().tolist())
    maker = c3.selectbox("제조사", makers)
    if maker != "(선택)":
        models = fuel_df[fuel_df["제조사"] == maker]["모델명"].tolist()
    else:
        models = []
    model = c4.selectbox("모델", ["(선택)"] + models)

    default_eff = 0.0
    default_fuel = "휘발유"
    if model != "(선택)":
        row = fuel_df[(fuel_df["제조사"] == maker) & (fuel_df["모델명"] == model)].iloc[0]
        default_eff = float(row["복합연비"])
        default_fuel = row["연료"]
    eff = c5.number_input("복합연비(km/L)", min_value=0.0, value=default_eff, step=0.1,
                          help="모델 선택 시 자동 입력. 수정 가능.")

    st.markdown("##### 3) 유가")
    c6, c7 = st.columns(2)
    fuel_type = c6.selectbox("유종", ["휘발유", "경유", "자동차용LPG"],
                             index=["휘발유", "경유", "자동차용LPG"].index(
                                 default_fuel if default_fuel in ["휘발유", "경유"] else
                                 ("자동차용LPG" if default_fuel == "LPG" else "휘발유")))
    oil_price = c7.number_input("적용 유가(원/L)", min_value=0.0, value=0.0, step=1.0,
                                help="오피넷 조회값을 참고해 입력. 추후 자동 연동 예정.")

    # --- 계산 ---
    if st.button("거리 · 유류비 계산", type="primary"):
        o_xy = geocode(origin) if origin else None
        d_xy = geocode(dest) if dest else None
        if not o_xy or not d_xy:
            st.error("주소를 좌표로 변환하지 못했습니다. 주소를 확인하거나 카카오 키 설정을 확인하세요.")
        else:
            route = get_road_distance(o_xy, d_xy)
            if not route:
                st.error("도로거리 조회 실패. 길찾기 API 권한 승인 여부를 확인하세요.")
            else:
                dist_km = route["distance_m"] / 1000
                st.session_state["last_dist_km"] = dist_km
                fuel_cost = calc_fuel_cost(dist_km, eff, oil_price)
                m1, m2, m3 = st.columns(3)
                m1.metric("도로 주행거리", f"{dist_km:,.1f} km")
                m2.metric("예상 소요시간", f"{route['duration_s']//60} 분")
                m3.metric("유류비(참고)", f"{fuel_cost:,.0f} 원")
                if route["toll"]:
                    st.caption(f"통행료 참고: {route['toll']:,} 원")

    # --- 3층 구조: 여비 항목 (자동값 기본, 수기 수정 가능) ---
    st.markdown("##### 4) 여비 항목 (자동값 기본 · 모두 수정 가능)")
    st.caption("조례 단가는 사무실 자료 확보 후 자동 채움. 현재는 수기 입력 상태.")
    d1, d2, d3, d4 = st.columns(4)
    days = d1.number_input("출장일수", min_value=0, value=1, step=1)
    ilbi = d2.number_input("일비(원)", min_value=0, value=ALLOWANCE_RATES["일비_1일"] * 1, step=1000)
    sikbi = d3.number_input("식비(원)", min_value=0, value=ALLOWANCE_RATES["식비_1일"] * 1, step=1000)
    sukbak = d4.number_input("숙박비(실비, 원)", min_value=0, value=0, step=1000)

    st.info("여비 최종 합산·조례 반영·엑셀 출력양식은 실물 자료 확보 후 연결됩니다. "
            "현재는 각 항목이 독립적으로 수정 가능한 상태입니다.")

# -------------------------------------------------------------
# TAB 3 : 명단 일괄 처리
# -------------------------------------------------------------
with tab3:
    st.subheader("출장자 명단 일괄 처리")
    st.markdown(
        "명단 엑셀을 업로드하면 전원 거리·유류비를 한 번에 계산합니다. "
        "**업로드 파일은 처리만 하며 서버에 저장하지 않습니다.**"
    )

    # 템플릿 다운로드
    template = pd.DataFrame({
        "성명": ["홍길동"],
        "직급": ["주무관"],
        "출발지": ["경상북도 안동시 풍천면 도청대로 455"],
        "도착지": ["경상북도 경주시 ○○로 123"],
        "제조사": ["현대"],
        "모델명": ["아반떼"],
        "출장일수": [1],
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        template.to_excel(w, index=False, sheet_name="명단")
    st.download_button("📥 명단 양식 다운로드", buf.getvalue(),
                       file_name="출장자_명단_양식.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    up = st.file_uploader("작성한 명단 업로드 (.xlsx)", type=["xlsx"])
    if up is not None:
        try:
            df_in = pd.read_excel(up)
            st.write("업로드된 명단:")
            st.dataframe(df_in, use_container_width=True)

            if st.button("전원 계산 실행", type="primary"):
                fuel_df = load_fuel_efficiency()
                results = []
                prog = st.progress(0.0)
                for i, r in df_in.iterrows():
                    o_xy = geocode(str(r.get("출발지", "")))
                    d_xy = geocode(str(r.get("도착지", "")))
                    route = get_road_distance(o_xy, d_xy) if (o_xy and d_xy) else None
                    dist_km = route["distance_m"] / 1000 if route else None

                    eff = None
                    match = fuel_df[(fuel_df["제조사"] == r.get("제조사")) &
                                    (fuel_df["모델명"] == r.get("모델명"))]
                    if not match.empty:
                        eff = float(match.iloc[0]["복합연비"])

                    results.append({
                        "성명": r.get("성명"),
                        "직급": r.get("직급"),
                        "출발지": r.get("출발지"),
                        "도착지": r.get("도착지"),
                        "도로거리(km)": round(dist_km, 1) if dist_km else "조회실패",
                        "연비(km/L)": eff if eff else "미등록",
                        "출장일수": r.get("출장일수"),
                    })
                    prog.progress((i + 1) / len(df_in))
                    time.sleep(0.05)  # API 과호출 방지

                res_df = pd.DataFrame(results)
                st.success("계산 완료")
                st.dataframe(res_df, use_container_width=True)

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
