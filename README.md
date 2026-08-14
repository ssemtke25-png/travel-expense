# 공무원 출장 여비 산정 시스템 (골격)

## 구성
- **⛽ 유가 조회** : 오피넷 무료 API (전국/시도/시군, 휘발유·경유·LPG)
- **🧮 여비 계산** : 주소→좌표(카카오 지오코딩) → 도로거리(카카오모빌리티) → 연비 → 유류비
  + 일비·식비·숙박비 3층 구조 (자동값 기본, 모두 수기 수정 가능)
- **📋 명단 일괄 처리** : 출장자 명단 엑셀 업로드 → 전원 계산 → 엑셀 다운로드

## 설계 원칙
- 저장 기능 없음 (세션 동안만 유지). 개인정보 서버 미보관.
- 업로드 파일은 처리만, 저장 안 함.

## 로컬 실행
```bash
pip install -r requirements.txt
# .streamlit/secrets.toml.example 를 secrets.toml 로 복사 후 키 입력
streamlit run app.py
```

## 키 준비 상태
| API | 신청 | 비고 |
|---|---|---|
| 오피넷 유가 | ✅ 완료 | secrets 에 OPINET_KEY |
| 카카오 지오코딩(주소→좌표) | ✅ 즉시 사용 | 권한 신청 불필요 |
| 카카오 길찾기(도로거리) | ⏳ 데브톡 승인 대기 | 승인 후 바로 작동 |

## 승인 대기 중 지금 테스트 가능한 것
- 유가 조회 전체
- 주소→좌표 변환 (지오코딩)
- 연비 CSV 선택 → 유류비 계산 (거리만 수동 입력하면)

## 다음 단계 (사무실 자료 확보 후)
1. `ALLOWANCE_RATES` : 경북 여비 조례 별표 단가 입력
2. `calculate_travel_allowance()` : 근무지 내/외, 일비·식비·숙박비 합산 로직
3. 정식 여비 출력 양식(엑셀/HWPX) 셀 좌표에 값 매핑
4. 관용차 계산법 추가 (연비 CSV 연료열 + 계산 분기)
5. 유가 2개월 누적 (GitHub Actions 일일 수집 → CSV 커밋)
6. 오피넷 조회값 → 여비 계산 탭 유가 자동 연동

## 파일 구조
```
travel_app/
├── app.py                      # 메인
├── data/fuel_efficiency.csv    # 모델별 대표 연비 (샘플)
├── requirements.txt
├── .gitignore                  # secrets.toml 제외
└── .streamlit/
    └── secrets.toml.example    # 키 템플릿
```
