# -*- coding: utf-8 -*-
#결제단가표의 서비스유형/종류/시간대 코드 체계, 단가 계산 규칙 정리
from app.gosi_reader import show_price
from app.param_store import load_params

# 사업 고정값: 결제단가표 전체 행과 동일 ──
BUSINESS = {
    "사업구분": "HW01",
    "사업구분명": "장애인활동지원",
    "사업유형ID": "HWG001",
    "사업유형명": "장애인활동지원",
}

# 시간 코드: 결제단가표에 쓰이는 시간대 이름
TIME_NAMES = {
    0: "0분부터 30분미만",
    30: "30분부터 60분미만",
    40: "40분부터 59분까지",
    60: "60분이상",
}

# 단가_파라미터를 이용해 서비스유형/종류 코드와 단가 계산 방식
# 단가 계산 방식 두 가지:
#   1) 시간비율   : 기준 금액 하나에 시간대별 비율을 곱해 파생 (활동보조/방문목욕/지시서)
#   2) 시간대별금액: 시간대마다 고시에 금액이 따로 명시됨 (방문간호)
SERVICE_MASTER = [
    {
        "서비스유형ID": "ST0001",
        "서비스유형명": "활동보조",
        "종류": [
            ("SK0001", "사회활동지원"),
            ("SK0002", "신체활동지원"),
            ("SK0003", "가사활동지원"),
            ("SK0004", "기타서비스"),
        ],
        "단가": {
            "방식": "시간비율",
            "원천": {"공통": ("활동보조", "일반")},  # 모든 종류가 같은 시간당 금액
            # (시간코드, 비율, 근거) - 30분 50%는 고시 제3장 1. 표의 비고
            "시간대": [
                (60, 1.0, "고시 시간당 금액 그대로"),
                (30, 0.5, "시간당 금액의 50% (제3장 1. 표 비고), 10원 미만 절사(제1장 6.)"),
            ],
            # 활동보조만 심야 단가가 따로 있음: 주간단가 x 1.5 절사
            # 계산값이 고시에 명시된 심야 금액과 일치하는지 검증한다
            "심야배율": 1.5,
            "심야확인": ("활동보조", "심야"),
        },
    },
    {
        "서비스유형ID": "ST0002",
        "서비스유형명": "방문간호",
        "종류": [
            ("SK0005", "기본간호"),
            ("SK0006", "치료간호"),
            ("SK0007", "교육상담"),
        ],
        "단가": {
            "방식": "시간대별금액",
            # 시간대마다 고시 금액이 따로 있음 (제3장 3.)
            "시간대": [
                (0, ("방문간호", "30분미만"), "고시 금액 그대로 (제3장 3.)"),
                (30, ("방문간호", "30분이상60분미만"), "고시 금액 그대로 (제3장 3.)"),
                (60, ("방문간호", "60분이상"), "고시 금액 그대로 (제3장 3.)"),
            ],
        },
    },
    {
        "서비스유형ID": "ST0003",
        "서비스유형명": "방문목욕",
        "종류": [
            ("SK0008", "차량내입욕"),
            ("SK0009", "가정내입욕"),
        ],
        "단가": {
            "방식": "시간비율",
            "원천": {  # 종류마다 고시 금액이 다름
                "SK0008": ("방문목욕", "차량내입욕"),
                "SK0009": ("방문목욕", "가정내입욕"),
            },
            # 40~59분 80%는 고시 제3장 2. 나.
            "시간대": [
                (60, 1.0, "고시 금액 그대로"),
                (40, 0.8, "급여비용의 80% (제3장 2. 나.), 10원 미만 절사(제1장 6.)"),
            ],
        },
    },
    {
        "서비스유형ID": "ST0004",
        "서비스유형명": "방문간호지시서",
        "종류": [
            ("SK0010", "의료기관 방문"),
            ("SK0011", "의료기관 의사내방"),
            ("SK0012", "보건기관 방문"),
            ("SK0013", "보건기관 의사내방"),
        ],
        "단가": {
            "방식": "시간비율",
            "원천": {
                "SK0010": ("방문간호지시서", "의료기관_방문"),
                "SK0011": ("방문간호지시서", "의료기관_의사내방"),
                "SK0012": ("방문간호지시서", "보건기관_방문"),
                "SK0013": ("방문간호지시서", "보건기관_의사내방"),
            },
            "시간대": [(60, 1.0, "고시 1회당 발급비용 그대로 (제4장)")],
        },
    },
]

# 10원 미만 절사 (고시 제1장 6.)
def floor10(x):
    return int(x) // 10 * 10

# 단가_파라미터 X 서비스유형/종류 코드 -> (유형 x 종류 x 시간대) 단가표
#     반환: [{서비스유형ID, 서비스유형명, 서비스종류, 서비스종유명,
#         서비스시간, 서비스시간명, 단위기준단가, 단위기준심야단가, 계산, 출처}]
def build_service_prices(prices):
    rows = []
    for svc in SERVICE_MASTER:
        rule = svc["단가"]

        for sk_code, sk_name in svc["종류"]:
            if rule["방식"] == "시간비율":
                # 종류별(또는 공통) 기준 금액에 시간대 비율을 곱해 파생
                src_key = rule["원천"].get(sk_code) or rule["원천"]["공통"]
                base = prices[src_key[0]][src_key[1]]

                for time_code, ratio, why in rule["시간대"]:
                    day = floor10(base["금액"] * ratio)

                    if "심야배율" in rule:
                        night = floor10(day * rule["심야배율"])
                        # 계산한 심야단가가 고시에 명시된 심야 금액과 맞는지 확인
                        # (기준 시간대인 비율 1.0에서만 직접 비교 가능)
                        chk = rule["심야확인"]
                        stated = prices[chk[0]][chk[1]]
                        if ratio == 1.0 and night != stated["금액"]:
                            raise ValueError(
                                f"{svc['서비스유형명']} 심야단가 계산값이 고시 명시값과 다릅니다.\n"
                                f"  계산값     : {night:,}원 = 주간 {day:,}원 x {rule['심야배율']} (10원 미만 절사)\n"
                                f"  고시 명시값: {show_price(stated)}\n"
                                "조치: 고시 원문에서 위 출처의 심야 금액을 확인하세요. "
                                "금액 오류면 params/단가_파라미터.json의 '금액'만 수정하면 되고, "
                                "배율 자체가 개정된 것이라면 SERVICE_MASTER의 '심야배율'을 근거와 함께 갱신해야 합니다."
                            )
                    else:
                        night = day  # 심야 구분이 없는 서비스는 주간과 동일

                    rows.append({
                        "서비스유형ID": svc["서비스유형ID"],
                        "서비스유형명": svc["서비스유형명"],
                        "서비스종류": sk_code,
                        "서비스종유명": sk_name,
                        "서비스시간": time_code,
                        "서비스시간명": TIME_NAMES[time_code],
                        "단위기준단가": day,
                        "단위기준심야단가": night,
                        "계산": why,
                        "출처": base["출처"],
                    })

            elif rule["방식"] == "시간대별금액":
                # 시간대마다 고시 금액이 따로 명시된 경우
                for time_code, src_key, why in rule["시간대"]:
                    base = prices[src_key[0]][src_key[1]]
                    rows.append({
                        "서비스유형ID": svc["서비스유형ID"],
                        "서비스유형명": svc["서비스유형명"],
                        "서비스종류": sk_code,
                        "서비스종유명": sk_name,
                        "서비스시간": time_code,
                        "서비스시간명": TIME_NAMES[time_code],
                        "단위기준단가": base["금액"],
                        "단위기준심야단가": base["금액"],  # 심야 구분 없음
                        "계산": why,
                        "출처": base["출처"],
                    })

            else:
                raise ValueError(f"알 수 없는 단가 계산 방식: {rule['방식']}")

    _validate_rows(prices, rows)
    return rows

# 유형별 조합 수, 금액 상식, 심야=공휴일 전제를 확인해 어긋나면 에러
def _validate_rows(prices, rows):
    # 유형별 조합 수: 활동보조 4x2=8, 방문간호 3x3=9, 방문목욕 2x2=4, 지시서 4x1=4
    expected = {"ST0001": 8, "ST0002": 9, "ST0003": 4, "ST0004": 4}
    from collections import Counter
    counts = Counter(r["서비스유형ID"] for r in rows)
    if dict(counts) != expected:
        raise ValueError(f"유형별 조합 수가 기대와 다릅니다: {dict(counts)} != {expected}")

    for r in rows:
        if r["단위기준단가"] <= 0 or r["단위기준심야단가"] < r["단위기준단가"]:
            raise ValueError(f"단가가 상식에 어긋납니다: {r}")

    # 결제단가표에는 심야단가 열만 있고 공휴일 열이 없다
    # 고시에서 '심야 금액=공휴일 금액' 인 사실을 이용해 공휴일, 심야 금액 점검
    h = prices["활동보조"]
    if h["심야"]["금액"] != h["공휴일"]["금액"]:
        raise ValueError(
            "고시의 심야 금액과 공휴일 금액이 다릅니다.\n"
            f"  심야  : {show_price(h['심야'])}\n"
            f"  공휴일: {show_price(h['공휴일'])}\n"
            "결제단가표 형식(심야단가 열 하나)으로는 두 값을 담을 수 없으므로, "
            "값 수정이 아니라 표 형식 확장(공휴일단가 열 추가 등)을 검토해야 합니다."
        )

# 생성한 결제단가표를 실제단가표와 전부 대조(근데 금액까지? 대조?)
def verify_against_sample(rows, sample_xlsx):
    import pandas as pd

    cols = ["서비스유형ID", "서비스유형명", "서비스종류", "서비스종유명",
            "서비스시간", "서비스시간명", "단위기준단가", "단위기준심야단가"]
    sample = pd.read_excel(sample_xlsx, engine="openpyxl")
    expected = {tuple(t) for t in sample[cols].drop_duplicates().itertuples(index=False)}
    actual = {tuple(r[c] for c in cols) for r in rows}

    if actual != expected:
        missing = expected - actual
        extra = actual - expected
        raise ValueError(
            f"샘플과 불일치!\n샘플에만 있는 조합 {len(missing)}개: {sorted(missing)}\n"
            f"생성본에만 있는 조합 {len(extra)}개: {sorted(extra)}"
        )
    return len(actual)


if __name__ == "__main__":
    import sys

    sys.stdout.reconfigure(encoding="utf-8")

    prices = load_params()  # params/단가_파라미터.json (담당자 수정분 포함)
    rows = build_service_prices(prices)

    print(f"서비스 단가표 {len(rows)}개 조합 생성")
    for r in rows:
        print(
            f"  {r['서비스유형ID']} {r['서비스유형명']:<7} | {r['서비스종류']} "
            f"{r['서비스종유명']:<9} | {r['서비스시간명']:<10} | "
            f"주간 {r['단위기준단가']:>7,} / 심야 {r['단위기준심야단가']:>7,} | {r['계산']}"
        )

    if len(sys.argv) >= 2:
        n = verify_against_sample(rows, sys.argv[1])
        print(f"\n샘플 대조 통과: {n}개 조합 전부 일치 ({sys.argv[1]})")
