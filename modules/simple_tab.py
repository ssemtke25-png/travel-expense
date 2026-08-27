# -*- coding: utf-8 -*-
"""
간소화 여비 계산 탭 (render_simple_tab)
------------------------------------------------
서무용 빠른 계산기. 경로·기간·유종·유가·공무용차량만 정하면 여비 총액 즉시 산출.
고정값: 숙박=그 밖의 지역(7만), 동승자 0명, 통행료=카카오 자동값, 관외 전용.

app.py 의 기존 함수/상수를 인자로 받아 재사용 → 로직 중복 없음.
"""
from datetime import date
import streamlit as st


def render_simple_tab(*, ctx):
    """
    ctx: app.py에서 넘겨주는 함수/상수 묶음(dict)
      geocode_debug, get_road_distance, calculate_travel_allowance,
      calc_fuel_cost, load_oil_history, lookup_oil_price,
      FUEL_STANDARDS, OIL_PRICE_FUEL_MAP, DEFAULT_ELEC_PRICE,
      SIGUN_PREFIX, SIGUN_ORDER, DEST_OPTIONS, DEFAULT_ORIGIN,
      build_dest_address
    """
    geocode_debug = ctx["geocode_debug"]
    get_road_distance = ctx["get_road_distance"]
    calculate_travel_allowance = ctx["calculate_travel_allowance"]
    calc_fuel_cost = ctx["calc_fuel_cost"]
    load_oil_history = ctx["load_oil_history"]
    lookup_oil_price = ctx["lookup_oil_price"]
    FUEL_STANDARDS = ctx["FUEL_STANDARDS"]
    OIL_PRICE_FUEL_MAP = ctx["OIL_PRICE_FUEL_MAP"]
    DEFAULT_ELEC_PRICE = ctx["DEFAULT_ELEC_PRICE"]
    SIGUN_PREFIX = ctx["SIGUN_PREFIX"]
    DEST_OPTIONS = ctx["DEST_OPTIONS"]
    DEFAULT_ORIGIN = ctx["DEFAULT_ORIGIN"]
    build_dest_address = ctx["build_dest_address"]

    st.markdown("#### ⚡ 빠른 여비 계산")
    st.caption("경로 · 기간 · 유종 · 유가 · 공무용차량만 정하면 여비 총액이 바로 나옵니다. "
               "숙박=그 밖 지역, 통행료=카카오 자동값, 동승자 없음으로 계산합니다. "
               "정확한 정산(숙박지역·통행료 실비·동승자)은 '여비 계산' 탭에서.")

    # ── ① 출발지 + 공무용차량 ──
    top = st.columns([3, 1])
    origin = top[0].text_input("출발지", value=DEFAULT_ORIGIN, key="sx_origin")
    use_car = top[1].checkbox("🚙 공무용 차량", key="sx_car",
                              help="체크 시 운임·통행료 0, 일비 1/2 자동 적용")

    # ── ② 도착 시군 + 세부주소 ──
    dc = st.columns([2.4, 2])
    sigun = dc[0].selectbox("도착 시군", DEST_OPTIONS, key="sx_sigun")
    if sigun == "기타(직접입력)":
        detail = dc[1].text_input("도착지 전체주소", key="sx_detail",
                                  placeholder="예) 대구광역시 북구 ○○로 123")
    elif sigun == "포항시청":
        detail = ""
        dc[1].text_input("세부주소", value="(본청 고정)", disabled=True, key="sx_detail_fix")
    else:
        detail = dc[1].text_input("세부주소", key="sx_detail",
                                  placeholder="읍·면·동 이하만")
    dest = build_dest_address(sigun, detail)
    if dest:
        st.caption(f"🎯 도착지: **{dest}**")

    # ── ③ 시작일 · 종료일 · 유종 · 유가지역 · 적용유가 ──
    dt = st.columns([1.2, 1.2, 1.3, 0.9, 1.3])
    if "sx_start" not in st.session_state:
        st.session_state["sx_start"] = date.today()
    if "sx_end" not in st.session_state:
        st.session_state["sx_end"] = date.today()

    def _sync_end():
        s = st.session_state.get("sx_start")
        if s is not None and st.session_state.get("sx_end", s) < s:
            st.session_state["sx_end"] = s

    start = dt[0].date_input("출장 시작일", key="sx_start", on_change=_sync_end)
    end = dt[1].date_input("출장 종료일", key="sx_end")
    fuel = dt[2].selectbox("유종", list(FUEL_STANDARDS.keys()), key="sx_fuel")
    std = FUEL_STANDARDS[fuel]
    is_elec = (std["unit"] == "km/kWh")
    region = dt[3].selectbox("유가지역", ["경북", "전국"], key="sx_region")

    # 유가 자동 조회
    hist = load_oil_history()
    if is_elec:
        oil_price = float(DEFAULT_ELEC_PRICE)
        dt[4].metric("적용 전기요금", f"{oil_price:,.0f} 원/kWh")
    else:
        oil_fuel = OIL_PRICE_FUEL_MAP.get(fuel, "휘발유")
        auto_p = None
        if not hist.empty:
            auto_p = lookup_oil_price(hist, start.isoformat(), region, oil_fuel)
        oil_price = float(auto_p) if auto_p is not None else 0.0
        dt[4].metric("적용 유가", f"{oil_price:,.1f} 원/L" if oil_price else "데이터 없음")

    if end < start:
        st.error("종료일이 시작일보다 빠릅니다.")
        end = start
    span = (end - start).days
    days, nights = span + 1, max(0, span)

    # 매일 출퇴근(숙박0 + 2일 이상) 배수 — 기존 앱과 동일
    daily_commute = (nights == 0 and days >= 2)
    round_mult = 2 * (days if daily_commute else 1)

    st.markdown(
        f"""<div style="background:#eaf4ff;border:1.5px solid #2b7cff;border-radius:8px;
        padding:8px 14px;margin:4px 0;">
        📅 <b>{start:%m-%d} ~ {end:%m-%d}</b> →
        <b style="color:#c0392b;">{nights}박 {days}일</b>
        · 일비·식비 {days}일 · 숙박 {nights}박{' · 매일 출퇴근 ×'+str(round_mult) if daily_commute else ''}</div>""",
        unsafe_allow_html=True)

    if oil_price <= 0 and not is_elec:
        st.warning(f"⚠️ {start.isoformat()} · {region} · {oil_fuel} 유가 데이터가 없어 계산이 부정확할 수 있습니다. "
                   "정확한 계산은 '여비 계산' 탭에서 유가를 수기 입력하세요.")

    calc = st.button("⚡ 여비 계산", type="primary", key="sx_calc", use_container_width=True)

    if calc:
        if not origin or not dest:
            st.error("출발지와 도착지를 모두 입력하세요.")
            return
        o = geocode_debug(origin)
        d = geocode_debug(dest)
        if not o["ok"] or not d["ok"]:
            if not o["ok"]:
                st.error(f"출발지 실패 [{o['method']}] → {o['reason']}")
            if not d["ok"]:
                st.error(f"도착지 실패 [{d['method']}] → {d['reason']}")
            return
        route = get_road_distance(o["xy"], d["xy"])
        if not route["ok"]:
            st.warning(f"도로거리 계산 대기: {route['reason']}")
            return

        dist_oneway = route["distance_m"] / 1000
        applied_dist = dist_oneway * round_mult
        # 통행료: 카카오 자동값 사용(간소화). 공무차면 계산 함수가 0 처리.
        auto_toll = int(route.get("toll", 0)) * (round_mult // 2 if round_mult >= 2 else 1)
        fuel_cost_val = calc_fuel_cost(applied_dist, std["eff"], oil_price)

        res = calculate_travel_allowance(
            is_gwannae=False, hours_over_4=True, use_official_car=use_car,
            days=days, nights=nights, sukbak_region_key="숙박비_상한_그밖",
            fuel_cost=fuel_cost_val, manual_transport=0.0,
            toll=float(auto_toll), num_passengers=0,
        )

        st.markdown("---")
        if use_car:
            st.markdown(
                """<div style="background:#fff4e5;border:1.5px solid #f39c12;border-radius:8px;
                padding:8px 14px;margin:4px 0;">🚙 <b>공무용 차량</b> →
                <b style="color:#c0392b;">운임·통행료 0 · 일비 1/2</b> 적용</div>""",
                unsafe_allow_html=True)
        drv = res["운전자"]
        r = st.columns(6)
        r[0].metric("운임", f"{drv['운임']:,} 원")
        r[1].metric("일비", f"{drv['일비']:,} 원")
        r[2].metric("식비", f"{drv['식비']:,} 원")
        r[3].metric("숙박비", f"{drv['숙박비']:,} 원")
        r[4].metric("통행료", f"{drv['통행료']:,} 원")
        r[5].metric("총 여비", f"{drv['합계']:,} 원")
        st.success(f"총 여비 합계: {drv['합계']:,} 원  ·  편도 {dist_oneway:,.1f}km · "
                   f"적용거리 {applied_dist:,.1f}km")
        st.caption("※ 십원 미만 버림. 숙박=그 밖 지역, 통행료=카카오 자동값, 동승자 없음. "
                   "정확한 정산은 '여비 계산' 탭에서.")
