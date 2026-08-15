# -*- coding: utf-8 -*-
"""
TAB 3 (명단 일괄 처리) → 여비 지급명세서 엑셀 출력 통합 스니펫
=====================================================================
기존 main 스트림릿 파일의 `with tab3:` 블록을 이 내용으로 교체.
travel_expense_excel.py 를 modules/ 에 두고 import 해서 사용.

엑셀 입력 컬럼 (양식과 100% 매핑):
  그룹ID    : 같은 차 = 같은 값. 그룹 내 첫 행(또는 동승여부='운전자')이 운전자.
  성명, 직급, 소속
  목적
  출장시작일, 출장종료일   (YYYY-MM-DD)
  출발지, 경유지(선택,쉼표), 도착지   (지오코딩용 전체주소)
  출발지명, 도착지명(선택)  (표기용. 비우면 출발지/도착지 앞부분 사용)
  유종
  동승여부  : '운전자' 또는 '동승' (비우면 그룹 내 순서로 첫 행=운전자)
  공무용차량 : 'Y'/'N' (비우면 N)
  숙박지역  : 서울/광역시/그밖 (비우면 그밖)

카카오/유가 규칙:
  · 갈때 = 출발→경유지들→최종도착 (waypoints 1회 호출) → 거리·통행료
  · 올때 = 최종도착→출발 (직행) → 거리·통행료
  · 유가는 UI에서 유종별 입력받거나 오피넷 자동값 사용
"""

import io
import time
from datetime import datetime
import pandas as pd
import streamlit as st

# ⬇️ 실제 배포 시: from modules.travel_expense_excel import ...
from modules.travel_expense_excel import (
    Person, Group, FUEL_STANDARDS,
    build_workbook, workbook_to_bytes,
    get_road_distance_waypoints,
)


def _parse_date(v):
    if pd.isna(v) or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    try:
        return pd.to_datetime(str(v)).date()
    except Exception:
        return None


def _short_place(full_addr: str) -> str:
    """전체주소에서 표기용 짧은 이름 추출 (마지막 시군구 or 앞 2어절)."""
    if not full_addr:
        return ""
    parts = str(full_addr).split()
    for p in parts:
        if p.endswith(("시", "군", "구")) and p not in ("경상북도",):
            return p
    return " ".join(parts[:2]) if len(parts) >= 2 else full_addr


def render_tab3(*, geocode_full, KAKAO_REST_KEY, fuel_price_map=None):
    """
    geocode_full: 기존 코드의 geocode_full(address)->{'ok','xy','reason'}
    KAKAO_REST_KEY: 카카오 REST 키
    fuel_price_map: {유종: 유가} dict (없으면 UI에서 입력)
    """
    st.subheader("출장자 명단 → 여비 지급명세서 엑셀 출력")
    st.markdown(
        "명단 엑셀을 업로드하면 전원 거리·유류비·통행료·여비를 계산하고, "
        "**HWP 「여비 지급명세서」 양식과 동일한 엑셀**을 생성합니다. "
        "업로드 파일은 처리만 하며 서버에 저장하지 않습니다."
    )

    # ---- 양식 다운로드 ----
    template = pd.DataFrame({
        "그룹ID": [1, 1],
        "성명": ["박00", "서00"],
        "직급": ["지방시설주사", "지방시설사무관"],
        "소속": ["건설도시국 토지정보과", "건설도시국 토지정보과"],
        "목적": ["2026년 제2회 지적재조사위원회 서면 심의", "2026년 제2회 지적재조사위원회 서면 심의"],
        "출장시작일": ["2026-07-23", "2026-07-23"],
        "출장종료일": ["2026-07-23", "2026-07-23"],
        "출발지": ["경상북도 안동시 풍천면 도청대로 455", ""],
        "경유지": ["", ""],
        "도착지": ["대구광역시 북구 연암로 40", ""],
        "출발지명": ["경북도청", ""],
        "도착지명": ["대구", "대구, 경산"],
        "유종": ["휘발유", ""],
        "동승여부": ["운전자", "동승"],
        "공무용차량": ["N", "N"],
        "숙박지역": ["그밖", "그밖"],
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        template.to_excel(w, index=False, sheet_name="명단")
    st.download_button("📥 명단 양식 다운로드", buf.getvalue(),
                       file_name="여비명세서_명단양식.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("※ 그룹ID: 같은 차=같은 값. 동승여부='운전자'인 행이 운임·통행료 산정 기준. "
               "동승자는 출발지·경유지·도착지 비워도 됨(운전자 경로 사용). "
               "경유지는 쉼표로 여러 개 입력 가능.")

    # ---- 유가 입력 ----
    st.markdown("##### 유종별 적용 유가 (원/L, 전기는 원/kWh)")
    price_cols = st.columns(len(FUEL_STANDARDS))
    prices = {}
    defaults = {"휘발유": 1862.0, "경유": 1700.0, "LPG": 1100.0,
                "하이브리드": 1862.0, "플러그인하이브리드": 1862.0, "전기": 300.0}
    for i, ft in enumerate(FUEL_STANDARDS.keys()):
        prices[ft] = price_cols[i].number_input(
            ft, min_value=0.0, value=float((fuel_price_map or defaults).get(ft, 1862.0)),
            step=10.0, key=f"price_{ft}")

    up = st.file_uploader("작성한 명단 업로드 (.xlsx)", type=["xlsx"], key="tab3_up")
    if up is None:
        return

    try:
        df_in = pd.read_excel(up, dtype=str)
    except Exception as e:
        st.error(f"파일 읽기 오류: {e}")
        return

    st.write("업로드된 명단:")
    st.dataframe(df_in, width="stretch")

    if not st.button("전원 계산 → 명세서 생성", type="primary", key="tab3_run"):
        return

    # ---- 그룹핑 (그룹ID 기준) ----
    if "그룹ID" not in df_in.columns:
        df_in["그룹ID"] = range(1, len(df_in) + 1)  # 없으면 전원 개별 운전자

    groups = []
    calc_log = []
    prog = st.progress(0.0)
    group_ids = list(dict.fromkeys(df_in["그룹ID"].tolist()))  # 순서 유지 유니크

    for gi, gid in enumerate(group_ids):
        rows = df_in[df_in["그룹ID"] == gid].reset_index(drop=True)

        # 운전자 판정: 동승여부='운전자' 우선, 없으면 첫 행
        driver_idx = 0
        for idx, rr in rows.iterrows():
            if str(rr.get("동승여부", "")).strip() == "운전자":
                driver_idx = idx
                break
        drow = rows.iloc[driver_idx]

        # --- 운전자 경로 계산 (갈때: 경유지 포함, 올때: 직행) ---
        origin_addr = str(drow.get("출발지", "") or "")
        dest_addr = str(drow.get("도착지", "") or "")
        via_raw = str(drow.get("경유지", "") or "").strip()
        via_addrs = [v.strip() for v in via_raw.split(",") if v.strip()] if via_raw else []

        o_res = geocode_full(origin_addr)
        d_res = geocode_full(dest_addr)
        o_xy = o_res["xy"] if o_res.get("ok") else None
        d_xy = d_res["xy"] if d_res.get("ok") else None
        via_xy = []
        for va in via_addrs:
            vr = geocode_full(va)
            if vr.get("ok"):
                via_xy.append(vr["xy"])

        갈때거리 = 올때거리 = 0.0
        통행_갈 = 통행_올 = 0
        fail = []
        if o_xy and d_xy:
            go = get_road_distance_waypoints(o_xy, d_xy, KAKAO_REST_KEY,
                                             waypoints_xy=via_xy or None)
            back = get_road_distance_waypoints(d_xy, o_xy, KAKAO_REST_KEY)
            if go.get("ok"):
                갈때거리 = go["distance_m"] / 1000
                통행_갈 = int(go.get("toll", 0))
            else:
                fail.append(f"갈때:{go.get('reason')}")
            if back.get("ok"):
                올때거리 = back["distance_m"] / 1000
                통행_올 = int(back.get("toll", 0))
            else:
                fail.append(f"올때:{back.get('reason')}")
        else:
            if not o_xy:
                fail.append(f"출발지 지오코딩 실패:{o_res.get('reason')}")
            if not d_xy:
                fail.append(f"도착지 지오코딩 실패:{d_res.get('reason')}")

        # --- Person 생성 ---
        def _mk(rr, is_driver):
            유종 = str(rr.get("유종", "") or drow.get("유종", "") or "휘발유").strip() or "휘발유"
            출발명 = str(rr.get("출발지명", "") or _short_place(origin_addr))
            도착명 = str(rr.get("도착지명", "") or _short_place(dest_addr))
            경유명 = ", ".join(_short_place(v) for v in via_addrs) if via_addrs else ""
            공무 = str(rr.get("공무용차량", "N")).strip().upper() in ("Y", "YES", "예", "1", "TRUE")
            return Person(
                소속=str(rr.get("소속", "") or ""),
                직급=str(rr.get("직급", "") or ""),
                성명=str(rr.get("성명", "") or ""),
                목적=str(rr.get("목적", "") or drow.get("목적", "") or ""),
                시작일=_parse_date(rr.get("출장시작일") or drow.get("출장시작일")),
                종료일=_parse_date(rr.get("출장종료일") or drow.get("출장종료일")),
                출발지명=출발명, 도착지명=도착명, 경유지명=경유명,
                유종=유종, is_driver=is_driver, use_official_car=공무,
                갈때거리_km=갈때거리 if is_driver else 0.0,
                올때거리_km=올때거리 if is_driver else 0.0,
                유가=prices.get(유종, 1862.0) if is_driver else 0.0,
                통행료_갈때=통행_갈 if is_driver else 0,
                통행료_올때=통행_올 if is_driver else 0,
                숙박지역=str(rr.get("숙박지역", "그밖") or "그밖").strip(),
            )

        driver = _mk(drow, True)
        passengers = [_mk(rows.iloc[i], False)
                      for i in range(len(rows)) if i != driver_idx]
        groups.append(Group(driver=driver, passengers=passengers))

        driver.calc()
        calc_log.append({
            "그룹": gid, "운전자": driver.성명,
            "갈때(km)": round(갈때거리, 1), "올때(km)": round(올때거리, 1),
            "유류비": driver._r["운임"], "통행료": driver._r["통행료"],
            "운전자 계": driver._r["계"], "동승자수": len(passengers),
            "실패사유": " / ".join(fail),
        })
        prog.progress((gi + 1) / len(group_ids))
        time.sleep(0.03)

    # ---- 결과 요약 ----
    log_df = pd.DataFrame(calc_log)
    st.success(f"계산 완료 — 그룹 {len(groups)}개")
    st.dataframe(log_df, width="stretch")
    fail_cnt = int((log_df["실패사유"] != "").sum())
    if fail_cnt:
        st.warning(f"⚠️ 경로 조회 실패 {fail_cnt}건 — '실패사유'를 확인하세요. "
                   "(주소는 시·군 포함 전체주소, 401/403이면 카카오 앱 설정)")

    # ---- 명세서 엑셀 생성 ----
    dept = str(df_in.iloc[0].get("소속", "경상북도 건설도시국 토지정보과"))
    signer = str(groups[0].driver.성명 or "박00")
    wb = build_workbook(groups, dept_footer=f"경상북도 {dept}", signer=signer)
    data = workbook_to_bytes(wb)
    st.download_button(
        "📤 여비 지급명세서 엑셀 다운로드", data,
        file_name=f"여비지급명세서_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary")
    st.caption("※ HWP 양식과 셀·병합 동일. 유류비=편도÷표준연비×유가(갈때·올때 각각), "
               "십원 미만 버림. 관용차는 운임·통행료 0. 동승자는 일비·식비·숙박비만.")
