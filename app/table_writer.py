from copy import deepcopy

import pandas as pd

# 헤더
DF_HEADER = [
    "안",
    "사업유형ID",
    "사업년도",
    "차수",
    "등급구분",
    "등급명",
    "바우처구분",
    "지원량",
    "정부지원금",
    "본인부담금",
    "",
    "재판정여부",
    "재판정기간",
    "소득기준ID",
    "소득기준년도",
    "소득구분",
    "정렬순서",
    "물품코드",
]


# 등급 수, 구간 수
IJ_GRADE_CNT = 4  # 인정조사 등급 수
JH_RANGE_CNT = 15  # 종합조사 구간 수
SJ_CNT = 60  # 산정특례 종류 수
INCOME_RANGE_CNT = 6  # 소득구간 수
EXTENDED_CNT = 2  # 주간확장 여부

# 소득구간 이름, 설명
INCOME_RANGE_NAME = ["가", "나", "다", "라", "마", "바"]

IJ_INCOME_RANGE_DESCRIPTION = [
    "기초수급자",
    "차상위",
    "전국가구평균소득50%이하",
    "전국가구평균소득50%초과~100%이하",
    "전국가구평균소득100%초과~150%이하",
    "전국가구평균소득150%초과~200%이하",
]

JH_SJ_INCOME_RANGE_DESCRIPTION = [
    "기초수급자",
    "차상위",
    "기준중위소득70%이하",
    "기준중위소득70%초과~120%이하",
    "기준중위소득120%초과~180%이하",
    "기준중위소득180%초과",
]


def write_basic_table(filename, ij, jh, sj=None):
    # df = pd.read_excel(filename, engine="openpyxl", header=None)
    df = pd.DataFrame(data=[DF_HEADER])
    print(df)

    # 안 및 정렬순서로 들어갈 번호
    order_num = 1

    # 인정조사 df, 각 시작 위치
    ij_df: pd.DataFrame = ij[0]
    ij_grade_name = ij[1]
    ij_monthly_limit = ij[2]
    ij_copayment = ij[3]

    # 종합조사 df, 각 시작 위치
    jh_df: pd.DataFrame = jh[0]
    jh_monthly_limit = jh[1]
    jh_copayment = jh[2]

    # 산정특례 df, 각 시작 위치
    # sj_df: pd.DataFrame = sj[0]
    # sj_grade_name = sj[1]
    # sj_monthly_limit = sj[2]
    # sj_copayment = sj[3]

    for ex in range(EXTENDED_CNT):  # 주간확장 여부: 0~1
        # 인정조사
        for g in range(IJ_GRADE_CNT):  # 등급: 시작 행 번호 + 0~3
            for ir in range(INCOME_RANGE_CNT):  # 소득: 시작 열 번호 + 0~5
                # 행 추가
                df.loc[order_num] = [None] * len(df.columns)

                # 정렬 번호 쓰기
                df.iat[order_num, DF_HEADER.index("안")] = order_num
                df.iat[order_num, DF_HEADER.index("정렬순서")] = order_num

                # 등급명 쓰기 ex) 1등급(가형)
                grade_name = (
                    ij_df.iat[ij_grade_name[0] + g, ij_grade_name[1]]
                    + "("
                    + INCOME_RANGE_NAME[ir]
                    + "형)"
                    + ("_주간확장" if ex == 1 else "")
                )
                df.iat[order_num, DF_HEADER.index("등급명")] = grade_name

                # 지원량(월한도액) 쓰기
                monthly_limit = int(
                    ij_df.iat[ij_monthly_limit[ex][0] + g, ij_monthly_limit[ex][1]]
                )
                df.iat[order_num, DF_HEADER.index("지원량")] = monthly_limit

                # 본인부담금 쓰기
                copayment = ij_df.iat[ij_copayment[ex][0] + g, ij_copayment[ex][1] + ir]
                if copayment == "면제":
                    copayment = 0
                df.iat[order_num, DF_HEADER.index("본인부담금")] = copayment

                # 정부지원금 (지원량 - 본인부담금) 쓰기
                df.iat[order_num, DF_HEADER.index("정부지원금")] = (
                    monthly_limit - copayment
                )

                # 소득구분 쓰기
                df.iat[order_num, DF_HEADER.index("소득구분")] = (
                    IJ_INCOME_RANGE_DESCRIPTION[ir]
                )

                order_num += 1

        # 종합조사
        for g in range(JH_RANGE_CNT):
            for ir in range(INCOME_RANGE_CNT):  # 기초, 차상위, 0~3번째 행
                # 행 추가
                df.loc[order_num] = [None] * len(df.columns)

                # 정렬 번호 쓰기
                df.iat[order_num, DF_HEADER.index("안")] = order_num
                df.iat[order_num, DF_HEADER.index("정렬순서")] = order_num

                # 등급명 쓰기 ex) 1구간(가형)
                grade_name = (
                    str(g + 1)
                    + "구간("
                    + INCOME_RANGE_NAME[ir]
                    + "형)"
                    + ("_주간확장" if ex == 1 else "")
                )
                df.iat[order_num, DF_HEADER.index("등급명")] = grade_name

                # 지원량(월한도액) 쓰기 (*종합은 역순으로 되어있음)
                monthly_limit = int(
                    jh_df.iat[jh_monthly_limit[ex][0], jh_monthly_limit[ex][1] - g]
                )
                df.iat[order_num, DF_HEADER.index("지원량")] = monthly_limit

                # 본인부담금 쓰기
                if ir == 0:  # 기초수급자(가형): 0원
                    copayment = 0
                elif ir == 1:  # 차상위(나형): 20000원
                    copayment = 20000
                else:  # 다형~바형
                    copayment = int(
                        jh_df.iat[
                            jh_copayment[ex][0] + (ir - 2), jh_copayment[ex][1] - g
                        ]
                    )
                df.iat[order_num, DF_HEADER.index("본인부담금")] = copayment

                # 정부지원금 (지원량 - 본인부담금) 쓰기
                df.iat[order_num, DF_HEADER.index("정부지원금")] = (
                    monthly_limit - copayment
                )

                # 소득구분 쓰기
                df.iat[order_num, DF_HEADER.index("소득구분")] = (
                    JH_SJ_INCOME_RANGE_DESCRIPTION[ir]
                )

                order_num += 1

        # TODO: 산정특례까지 포함된 테이블 만들기

    df.to_excel(filename, engine="openpyxl", header=None, index=None)


def write_extra_table(filename, ij):
    pass
