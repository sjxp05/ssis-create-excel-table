# -*- coding: utf-8 -*-
"""결제단가표 생성기

이미 만들어진 두 재료를 결합한다 (새로운 금액 계산 없음):

    기본급여 단가표 (등급구분/등급명/지원량/정부지원금/본인부담금, 950행)
            x
    서비스 단가 25개 조합 (payment_master.build_service_prices)
            |
            v
    결제단가표 949등급 x 25조합 = 23,725행 (27열)

데이터로 확인된 생성 규칙:
- 부적합(등급구분 9999) 행은 결제단가로 전개되지 않는다 -> 950 - 1 = 949등급
- 긴급활동지원(D599)은 일반 등급과 똑같이 모든 서비스유형에 전개된다
- 임시공휴일 추가지원(D799, 2025년) 같은 연중 추가 등급은 기본 생성 대상이 아니다
  (연중 공문에 따라 별도 차수로 추가되는 행)
- 행 순서: 서비스유형(ST0001~4) -> 등급구분 알파벳순
  (등급 안의 종류x시간 순서는 공식 파일에서도 일정한 규칙이 없어 고정 순서로 생성)
- 정부지원금율 = 정부지원금액/지원량, 본인부담금율 = 본인부담금액/지원량

단독 실행:
    python -m app.payment_writer <기본급여단가표.xlsx> <출력.xlsx> <연도> <차수> [검증샘플.xlsx]
"""

import pandas as pd

from app.payment_master import BUSINESS, build_service_prices
from app.param_store import load_params

# 결제단가표의 열 순서 (실제 결제단가 파일과 동일)
COLUMNS = [
    "순번", "사업년도", "차수", "사업구분", "사업구분명", "사업유형ID", "사업유형명",
    "서비스유형ID", "서비스유형명", "서비스종류", "서비스종유명", "등급구분", "등급명",
    "서비스시간", "서비스시간명", "단위기준단가", "단위기준심야단가",
    "지원량", "정부지원금액", "본인부담금액", "정부지원금율", "본인부담금율",
    "사용여부", "서비스시작일자", "서비스종료일자", "단가관리레벨", "비고",
]

# 결제로 전개하지 않는 등급 (결제가 발생할 수 없는 코드)
EXCLUDED_GRADES = {"9999"}  # 부적합


def read_basic_table(basic_xlsx):
    """기본급여 단가표에서 결제단가에 필요한 열만 읽는다.

    우리 파이프라인의 생성본(write_basic_table 결과)과
    공식 배포본 모두 같은 열 이름을 쓰므로 둘 다 입력이 될 수 있다.
    """
    df = pd.read_excel(basic_xlsx, engine="openpyxl")
    need = ["등급구분", "등급명", "지원량", "정부지원금", "본인부담금"]
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise ValueError(
            f"기본급여 단가표에 필요한 열이 없습니다: {missing}\n"
            f"파일: {basic_xlsx}\n열 목록: {list(df.columns)}"
        )
    df = df[need].copy()

    # 부적합 등 결제 불가 등급 제외
    before = len(df)
    df = df[~df["등급구분"].astype(str).isin(EXCLUDED_GRADES)]
    print(f"기본급여 단가표 {before}행 중 결제 제외 등급 {before - len(df)}행 제외 -> {len(df)}등급")

    # 등급구분은 유일해야 한다 (중복되면 어떤 금액이 맞는지 알 수 없음)
    dup = df[df["등급구분"].duplicated(keep=False)]
    if not dup.empty:
        raise ValueError(f"등급구분이 중복된 행이 있습니다:\n{dup.to_string()}")

    # 항등식: 지원량 = 정부지원금 + 본인부담금 (기본 단가표 자체의 무결성)
    bad = df[df["지원량"] != df["정부지원금"] + df["본인부담금"]]
    if not bad.empty:
        raise ValueError(
            "지원량 = 정부지원금 + 본인부담금 항등식이 깨진 행이 있습니다. "
            f"기본급여 단가표를 확인하세요:\n{bad.head(10).to_string()}"
        )

    # 율 계산에 지원량이 0이면 안 된다 (부적합 제외 후에는 없어야 정상)
    zero = df[df["지원량"] <= 0]
    if not zero.empty:
        raise ValueError(
            "지원량이 0 이하인 등급이 있어 부담율을 계산할 수 없습니다. "
            f"결제 제외 대상인지 확인하세요:\n{zero.to_string()}"
        )
    return df


def build_payment_table(basic_df, service_rows, year, order_no):
    """등급(기본 단가표) x 서비스 조합(25개) -> 결제단가 DataFrame"""
    # 행 순서: 서비스유형 블록 -> 등급구분 알파벳순 (실제 파일과 동일한 배치)
    grades = basic_df.sort_values("등급구분").to_dict("records")

    rows = []
    # 유형 블록 단위로 묶기 위해 유형ID별로 그룹화
    by_type = {}
    for svc in service_rows:
        by_type.setdefault(svc["서비스유형ID"], []).append(svc)

    for type_id in sorted(by_type):
        for grade in grades:
            for svc in by_type[type_id]:
                rows.append({
                    "사업년도": year,
                    "차수": order_no,
                    **BUSINESS,
                    "서비스유형ID": svc["서비스유형ID"],
                    "서비스유형명": svc["서비스유형명"],
                    "서비스종류": svc["서비스종류"],
                    "서비스종유명": svc["서비스종유명"],
                    "등급구분": grade["등급구분"],
                    "등급명": grade["등급명"],
                    "서비스시간": svc["서비스시간"],
                    "서비스시간명": svc["서비스시간명"],
                    "단위기준단가": svc["단위기준단가"],
                    "단위기준심야단가": svc["단위기준심야단가"],
                    "지원량": grade["지원량"],
                    "정부지원금액": grade["정부지원금"],
                    "본인부담금액": grade["본인부담금"],
                    "정부지원금율": grade["정부지원금"] / grade["지원량"],
                    "본인부담금율": grade["본인부담금"] / grade["지원량"],
                    "사용여부": "Y",
                    "서비스시작일자": int(f"{year}0101"),
                    "서비스종료일자": int(f"{year}1231"),
                    "단가관리레벨": 0,
                    "비고": " ",
                })

    df = pd.DataFrame(rows)
    df.insert(0, "순번", range(1, len(df) + 1))
    df = df[COLUMNS]

    # 자가 검증: 행수 = 등급수 x 조합수
    expected = len(grades) * len(service_rows)
    if len(df) != expected:
        raise ValueError(f"행수가 기대와 다릅니다: {len(df)} != 등급 {len(grades)} x 조합 {len(service_rows)}")
    return df


def write_payment_table(basic_xlsx, out_xlsx, year, order_no):
    """결제단가표를 생성해 엑셀로 저장한다.

    '단가근거' 시트에 25개 조합의 계산 근거(고시 조항, 출처)를 함께 남겨
    담당자가 결과 파일만 열어도 단가의 출처를 확인할 수 있게 한다.
    """
    prices = load_params()  # 고시에서 추출된 단가 (담당자 수정분 포함, 로드 시 검증)
    service_rows = build_service_prices(prices)
    basic_df = read_basic_table(basic_xlsx)
    df = build_payment_table(basic_df, service_rows, year, order_no)

    # 단가근거 시트: 각 조합의 금액이 고시 어디에서 어떻게 계산됐는지
    evidence = pd.DataFrame([
        {
            "서비스유형명": s["서비스유형명"],
            "서비스종유명": s["서비스종유명"],
            "서비스시간명": s["서비스시간명"],
            "단위기준단가": s["단위기준단가"],
            "단위기준심야단가": s["단위기준심야단가"],
            "계산": s["계산"],
            "고시출처": " > ".join(
                x for x in [s["출처"]["문맥"].get("장"), s["출처"]["문맥"].get("항목")] if x
            ) + f' > 표{s["출처"]["표번호"]} 행{s["출처"]["행"]} "{s["출처"]["라벨"]}"',
        }
        for s in service_rows
    ])

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as xl:
        df.to_excel(xl, sheet_name="결제단가", index=False)
        evidence.to_excel(xl, sheet_name="단가근거", index=False)

    print(f"결제단가표 {len(df):,}행 생성 완료: {out_xlsx} (근거 시트 포함)")
    return df


def verify_against_sample(df, sample_xlsx):
    """생성본을 실제 결제단가 샘플과 내용 기준으로 전수 대조한다.

    순번과 행 배치는 공식 파일도 일정한 규칙이 없으므로(수작업 흔적),
    순번을 제외한 26개 열의 내용을 집합으로 비교한다.
    """
    sample = pd.read_excel(sample_xlsx, engine="openpyxl")

    def key_set(d):
        d = d.drop(columns=["순번"]).copy()
        # 부담율은 실수라 미세오차 방지를 위해 10자리에서 반올림해 비교
        for c in ["정부지원금율", "본인부담금율"]:
            d[c] = d[c].round(10)
        return {tuple(t) for t in d.itertuples(index=False)}

    ours, theirs = key_set(df), key_set(sample)
    if ours != theirs:
        missing = theirs - ours
        extra = ours - theirs
        raise ValueError(
            f"샘플과 불일치! 샘플에만 있는 행 {len(missing)}개, 생성본에만 있는 행 {len(extra)}개\n"
            f"샘플에만: {list(missing)[:3]}\n생성본에만: {list(extra)[:3]}"
        )
    return len(ours)


if __name__ == "__main__":
    import sys

    sys.stdout.reconfigure(encoding="utf-8")

    if len(sys.argv) < 5:
        print("사용법: python -m app.payment_writer <기본급여단가표.xlsx> <출력.xlsx> <연도> <차수> [검증샘플.xlsx]")
        sys.exit(1)

    basic_xlsx, out_xlsx, year, order_no = sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4])
    df = write_payment_table(basic_xlsx, out_xlsx, year, order_no)

    if len(sys.argv) >= 6:
        n = verify_against_sample(df, sys.argv[5])
        print(f"샘플 대조 통과: {n:,}개 행 내용 전부 일치 ({sys.argv[5]})")
