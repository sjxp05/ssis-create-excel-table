from copy import deepcopy

import pandas as pd
from itertools import product
from decimal import Decimal

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

JH_INCOME_RANGE_DESCRIPTION = [
    "기초수급자",
    "차상위",
    "기준중위소득70%이하",
    "기준중위소득70%초과~120%이하",
    "기준중위소득120%초과~180%이하",
    "기준중위소득180%초과",
]


def write_basic_table(filename, ij, sj=None, jh=None):
    # df = pd.read_excel(filename, engine="openpyxl", header=None)
    df = pd.DataFrame(data=[DF_HEADER])
    print(df)

    # 안 및 정렬순서로 들어갈 번호
    order_num = 1

    # 인정조사 테이블 먼저 만들기
    ij_df: pd.DataFrame = ij[0]
    ij_grade = ij[1]#등급명 시작좌표
    ij_monthly_limit = ij[2]
    ij_copayment = ij[3]

    for ex in range(EXTENDED_CNT):  # 주간확장 여부: 0~1
        for g in range(IJ_GRADE_CNT):  # 등급: 시작 행 번호 + 0~3
            for ir in range(INCOME_RANGE_CNT):  # 소득: 시작 열 번호 + 0~5
                # 행 추가
                df.loc[order_num] = [None] * len(df.columns)

                # 정렬 번호 쓰기
                df.iat[order_num, DF_HEADER.index("안")] = order_num
                df.iat[order_num, DF_HEADER.index("정렬순서")] = order_num

                # 등급명 쓰기
                grade_name = (
                    ij_df.iat[ij_grade[0] + g, ij_grade[1]]
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
                #ij_monthly_limit          # ((8, 5), (8, 6))   ← 좌표 튜플 2개가 든 튜플
                #ij_monthly_limit[ex]      # (8, 5)             ← ex번째 좌표 하나 선택
                #ij_monthly_limit[ex][0]   # 8                  ← 그 좌표의 행
                #ij_monthly_limit[ex][1]   # 5                  ← 그 좌표의 열
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

        # TODO: 종합조사까지 포함된 테이블 만들기

    df.to_excel(filename, engine="openpyxl", header=None, index=None)


# TK_GROUP = [
#     ("1등급", "최중증취약가구"), ("1등급", "최중증1인가구"),
#     ("1등급", "1등급취약가구"), ("1등급", "1등급1인가구"),
#     ("1등급", "나머지가구구성원의직장생활등"), ("1등급", None),
#     ("2등급", "2등급이하취약가구"), ("2등급", "2등급이하1인가구"), ("2등급", None),
#     ("3등급", "2등급이하취약가구"), ("3등급", "2등급이하1인가구"), ("3등급", None),
#     ("4등급", "2등급이하취약가구"), ("4등급", "2등급이하1인가구"), ("4등급", None),
# ]
# TWO    = [(True, True), (True, False), (False, True), (False, False)]
# TYPES  = ["가형", "나형", "다형", "라형", "마형", "바형"]
# INCOME = ["기초수급자", "차상위", "기준중위소득70%이하", "기준중위소득70%초과~120%이하",
#           "기준중위소득120%초과~180%이하", "기준중위소득180%초과"]

# def write_tk_talbe(table):
#     rows_normal,rows_ext=[],[]
#     #Product를 이용해 모든 조합생성
#     #("1등급","최중증취약가구") 조합 생성 후 start=1, (1,("1등급","최중증취약가구")(True,True)) n이 start에서 특례 올라가는것
#     for n,((grade,gagu_type),(school,work))in enumerate(product(TK_GROUP,TWO),start=1):
#         info = table[(grade, gagu_type, school, work)] 
#         for k in range(6):
#             #일반
#             limit, copay=info["주간기본금액"], info["주간기본부담"][k]
#             rows_normal.append({
#                 "순번": "",
#                 "등급구분": "",
#                 "등급명": f"특례{n}({TYPES[k]})",
#                 "바우처구분": "포인트",
#                 "지원량": limit,
#                 "정부지원금": limit - copay,
#                 "본인부담금": copay,
#                 "재판정여부": "불가능",
#                 "재판정기간":"",
#                 "소득기준ID":"",
#                 "소득구분": INCOME[k],
#                 "정렬순서": "",
#                 "물품코드":""
#             })

#             #주간확장
#             limit, copay = info["주간확장금액"], info["주간확장부담"][k]
#             rows_ext.append({
#                 "순번": "",
#                 "등급구분": "",
#                 "등급명": f"특례{n}({TYPES[k]})_주간확장",
#                 "바우처구분": "포인트",
#                 "지원량": limit,
#                 "정부지원금": limit - copay,
#                 "본인부담금": copay,
#                 "재판정여부": "불가능",
#                 "재판정기간":"",
#                 "소득기준ID":"",
#                 "소득구분": INCOME[k],
#                 "정렬순서": "",
#                 "물품코드":""
#             })
#     return pd.DataFrame(rows_normal + rows_ext) 


#추가 급여 table
ADD_HEADER=[
    "순번",
    "등급구분",
    "등급명",
    "지원량",
    "정부지원금",
    "본인부담금",
    "소득구분",
    "추가급여구분",
]

ADD_BENEFIT_CLASSIFICATION=[
    ("출산가구","출산","출산가구여부","A025"),
    ("학교생활","학교생활","학교생활여부","A031"),
    ("직장생활","직장생활","직장생활여부","A037"),
    ("자립준비","자립준비", "자립준비여부","A043"),
    ("보호자일시부재","보호자일시부재","보호자일시부재","A061"),
    ("가족의직장생활","나머지가구구성원의직장생활등","가족의직장생활","A067"),
    ("최중증1인가구","최중증1인가구","최중증1인가구여부","A073"),
    ("1등급1인가구","1등급1인가구","1등급1인가구","A079"),
    ("2등급1인가구","2등급이하1인가구","2등급이하1인가구","A085"),
    ("최중증취약가구","최중증취약가구","최중증취약가구","A091"),
    ("1등급취약가구","1등급취약가구","1등급취약가구","A097"),
    ("2등급취약가구","2등급이하취약가구","2등급이하취약가구","A103"),
]

def write_add_table(filename,gagu_amount,add_paypercent):
    rates=[0,0]+list(add_paypercent.values())
    table=[]
    seq=1
    for name,amount_key,category,code in ADD_BENEFIT_CLASSIFICATION:
        limit=gagu_amount[amount_key]#지원량
        code=int(code[1:])#A뗀 수

        for k in range(6):
            if rates[k]==0:
                copay=0
            else:
                raw = Decimal(limit) * Decimal(str(rates[k]))
                copay = int(raw // 100) * 100 #100원단위절사

            table.append({
                "순번":seq,
                "등급구분":f"A{code+k:03d}",
                "등급명":f"{name}_{INCOME_RANGE_NAME[k]}형",
                "지원량":limit,
                "정부지원금":limit-copay,
                "본인부담금":copay,
                "소득구분":IJ_INCOME_RANGE_DESCRIPTION[k],
                "추가급여구분":category,
            })
            seq+=1
    df=pd.DataFrame(table)
    df.to_excel(filename, engine="openpyxl", header=True, index=None)