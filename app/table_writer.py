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


def write_basic_table(filename, ij, jh, sj):
    # df = pd.read_excel(filename, engine="openpyxl", header=None)
    df = pd.DataFrame(data=[DF_HEADER])
    print(df)

    # 안 및 정렬순서로 들어갈 번호
    order_num = 1

    # 인정조사 df, 각 시작 위치
    ij_df: pd.DataFrame = ij[0]
    ij_grade_name = ij[1]#등급명 시작좌표
    ij_monthly_limit = ij[2]
    ij_copayment = ij[3]

    # 종합조사 df, 각 시작 위치
    jh_df: pd.DataFrame = jh[0]
    jh_monthly_limit = jh[1]
    jh_copayment = jh[2]

    # 산정특례 df, 각 시작 위치
    sj_df: pd.DataFrame = sj[0]
    sj_grade_name = sj[1]
    sj_monthly_limit = sj[2]
    sj_copayment = sj[3]

    # 조견표 특례 시트 상 기준 위치
    GRADE_R = sj_grade_name[0]      # 데이터 시작 행
    GRADE_C = sj_grade_name[1]      # 등급 열
    MAIN_C = GRADE_C + 1            # main 추가급여 열
    SCHOOL_C = GRADE_C + 11         # 학교생활 지원시간 열
    WORK_C = GRADE_C + 12           # 직장생활 지원시간 열

    # 추가 특성
    main_1_keys = ['최중증취약가구', '최중증1인가구', '1등급취약가구', '1등급1인가구', '나머지가구구성원의직장생활등', '-']
    main_other_keys = ['2등급이하취약가구', '2등급이하1인가구', '-']
                
    # (학교, 직장 유무) 조합
    cond_types = [(True, True), (True, False), (False, True), (False, False)]

    # 특례 1~60 순서 찾는 함수
    def get_row_offset(t_grade, t_main, is_school, is_work):
        for i in range(SJ_CNT):
            row = GRADE_R + i
            g_val = sj_df.iat[row, GRADE_C]
            m_val = sj_df.iat[row, MAIN_C]
            s_val = sj_df.iat[row, SCHOOL_C]
            w_val = sj_df.iat[row, WORK_C]
                        
            has_school = pd.notna(s_val)
            has_work = pd.notna(w_val)
                        
            if g_val != t_grade:
                continue
                            
            # 엑셀 특이점 처리: main이 '-'여야 하는데 학교나 직장이 있으면 main에 그 이름이 들어감
            if t_main == '-':
                if is_school and is_work:
                    if m_val == '학교생활' and has_work: return i
                elif is_school:
                    if m_val == '학교생활' and not has_work: return i
                elif is_work:
                    if m_val == '직장생활' and not has_school: return i
                else:
                    if m_val == '-' and not has_school and not has_work: return i
            else:
                if m_val == t_main and has_school == is_school and has_work == is_work:
                    return i
        raise ValueError(f"해당 조건의 특례를 엑셀에서 찾을 수 없습니다: {t_grade}, {t_main}, 학교={is_school}, 직장={is_work}")

    # 특례 순서 저장할 리스트
    SJ_ORDER_MAP = []
                
    # 1등급
    for m in main_1_keys:
        for s_flag, w_flag in cond_types:
            SJ_ORDER_MAP.append(get_row_offset('1등급', m, s_flag, w_flag))
                        
    # 2~4등급
    for g in ['2등급', '3등급', '4등급']:
        for m in main_other_keys:
            for s_flag, w_flag in cond_types:
                    SJ_ORDER_MAP.append(get_row_offset(g, m, s_flag, w_flag))

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
                
                # 등급구분 쓰기
                prefix = "D" if ex == 0 else "C"
                code_num = (g * INCOME_RANGE_CNT) + ir + 1
                df.iat[order_num, DF_HEADER.index("등급구분")] = f"{prefix}{code_num:03d}"

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

                # 등급구분 쓰기
                prefix = "D" if ex == 0 else "C"
                code_num = 500 + (g * INCOME_RANGE_CNT) + ir + 1
                df.iat[order_num, DF_HEADER.index("등급구분")] = f"{prefix}{code_num:03d}"


                # 지원량(월한도액) 쓰기 (*종합은 역순으로 되어있음)
                monthly_limit = int(
                    jh_df.iat[jh_monthly_limit[ex][0], jh_monthly_limit[ex][1] - g]
                )
                df.iat[order_num, DF_HEADER.index("지원량")] = monthly_limit

                # 본인부담금 쓰기
                if ir == 0:  # 기초수급자(가형): 0원
                    copayment = 0
                elif ir == 1:  # 차상위(나형): 20000원
                    # 단, 주간확장(주간활동 이용으로 월 한도액을 조정한 경우)은 면제
                    # (고시 제2장 4. 단서 / 2025·2026 공식 단가표에서 확인)
                    copayment = 20000 if ex == 0 else 0
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

    for ex in range(EXTENDED_CNT): 
        for s in range(SJ_CNT):  # 0~59 (특례1 ~ 특례60)
            actual_row = SJ_ORDER_MAP[s]

            for ir in range(INCOME_RANGE_CNT):  # 0~5 (가형 ~ 바형)
                # 행 추가
                df.loc[order_num] = [None] * len(df.columns)

                # 정렬 번호 쓰기
                df.iat[order_num, DF_HEADER.index("안")] = order_num
                df.iat[order_num, DF_HEADER.index("정렬순서")] = order_num

                # 등급명 쓰기 ex) 특례1(가형), 특례1(가형)_주간확장
                grade_name = (
                    f"특례{s + 1}"
                    + "("
                    + INCOME_RANGE_NAME[ir]
                    + "형)"
                    + ("_주간확장" if ex == 1 else "")
                )
                df.iat[order_num, DF_HEADER.index("등급명")] = grade_name

                # 등급구분 쓰기
                prefix = "D" if ex == 0 else "C"
                group_char = ["A", "B", "C", "D"][s // 15] # 15개 단위로 A, B, C, D 할당
                code_num = ((s % 15) * INCOME_RANGE_CNT) + ir + 1
                df.iat[order_num, DF_HEADER.index("등급구분")] = f"{prefix}{group_char}{code_num:02d}"

                # 지원량(월한도액) 쓰기 - actual_row 위치의 데이터를 가져옴
                monthly_limit_val = sj_df.iat[sj_monthly_limit[ex][0] + actual_row, sj_monthly_limit[ex][1]]
                
                # 이상한 텍스트가 들어오면 0으로 처리하도록 예외처리
                try:
                    monthly_limit = int(monthly_limit_val) if pd.notna(monthly_limit_val) else 0
                except ValueError:
                    monthly_limit = 0
                    
                df.iat[order_num, DF_HEADER.index("지원량")] = monthly_limit

                # 본인부담금 쓰기 - actual_row 위치의 데이터를 가져옴
                copayment = sj_df.iat[sj_copayment[ex][0] + actual_row, sj_copayment[ex][1] + ir]
                if copayment == "면제" or pd.isna(copayment):
                    copayment = 0
                else:
                    try:
                        copayment = int(copayment)
                    except ValueError:
                        copayment = 0
                        
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

    # 긴급활동지원 (949번째 행)
    # 지원량 = 종합조사 13구간 월 한도액 (2025: 1,997,000 / 2026: 2,076,000 결제단가에서 확인됨)
    # 본인부담금 없음, 소득구분은 시스템상 ':::선택:::'
    df.loc[order_num] = [None] * len(df.columns)
    df.iat[order_num, DF_HEADER.index("안")] = order_num
    df.iat[order_num, DF_HEADER.index("정렬순서")] = order_num
    df.iat[order_num, DF_HEADER.index("등급구분")] = "D599"
    df.iat[order_num, DF_HEADER.index("등급명")] = "긴급활동지원"
    urgent_limit = int(
        jh_df.iat[jh_monthly_limit[0][0], jh_monthly_limit[0][1] - 12]  # 13구간 (기본형)
    )
    df.iat[order_num, DF_HEADER.index("지원량")] = urgent_limit
    df.iat[order_num, DF_HEADER.index("정부지원금")] = urgent_limit
    df.iat[order_num, DF_HEADER.index("본인부담금")] = 0
    df.iat[order_num, DF_HEADER.index("소득구분")] = ":::선택:::"
    order_num += 1

    # 부적합 (950번째 행): 결제가 발생하지 않는 코드, 전부 0원
    df.loc[order_num] = [None] * len(df.columns)
    df.iat[order_num, DF_HEADER.index("안")] = order_num
    df.iat[order_num, DF_HEADER.index("정렬순서")] = order_num
    df.iat[order_num, DF_HEADER.index("등급구분")] = "9999"
    df.iat[order_num, DF_HEADER.index("등급명")] = "부적합"
    df.iat[order_num, DF_HEADER.index("지원량")] = 0
    df.iat[order_num, DF_HEADER.index("정부지원금")] = 0
    df.iat[order_num, DF_HEADER.index("본인부담금")] = 0
    df.iat[order_num, DF_HEADER.index("소득구분")] = ":::선택:::"
    order_num += 1

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