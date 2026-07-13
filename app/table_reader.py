import pandas as pd
from itertools import product

# 인정조사 시트 읽고 월 한도액, 등급, 본인부담금 등 위치 찾아서 반환
def read_ij_table(filename, sheet_name):
    # 파일 열기
    df = pd.read_excel(
        io=filename,
        sheet_name=sheet_name,
        engine="openpyxl",
        header=None,
    )
    print(df)

    # 활동지원등급 시작위치 찾기
    #.to_numpy().nonzero() true인 위치들의 행,열 배열 반환 / zip(*...)->행,열 좌표 쌍으로 묶음
    for row, col in zip(*((df == "활동지원등급").to_numpy().nonzero())):
        GRADE_HEADER = (row + 2, col)

    # 월 한도액(기본/확장) 시작위치 찾기
    BASIC_MONTHLY_LIMIT = []#주간활동기본형
    EXTENDED_MONTHLY_LIMIT = []#주간활동 확장형
    for row, col in zip(*((df == "월 한도액").to_numpy().nonzero())):
        BASIC_MONTHLY_LIMIT.append((row + 2, col + 1))
        EXTENDED_MONTHLY_LIMIT.append((row + 2, col + 2))

    # 구간별 본인부담금(기본형) 시작위치 찾기
    BASIC_COPAYMENT = []
    for row, col in zip(*((df == "주간활동 본인부담금(기본형)").to_numpy().nonzero())):
        BASIC_COPAYMENT.append((row + 2, col))

    # 구간별 본인부담금(확장형) 시작위치 찾기
    EXTENDED_COPAYMENT = []
    for row, col in zip(*((df == "주간활동 본인부담금(확장형)").to_numpy().nonzero())):
        EXTENDED_COPAYMENT.append((row + 2, col))

    print(
        GRADE_HEADER,
        BASIC_MONTHLY_LIMIT[0],
        EXTENDED_MONTHLY_LIMIT[0],
        BASIC_COPAYMENT[0],
        EXTENDED_COPAYMENT[0],
    )

    return (
        df,
        GRADE_HEADER,
        BASIC_MONTHLY_LIMIT,
        EXTENDED_MONTHLY_LIMIT,
        BASIC_COPAYMENT,
        EXTENDED_COPAYMENT,
    )


# 산정조사 시트

# def read_sj_table(filename, sheet_name):
#     # 파일 열기
#     df = pd.read_excel(
#         io=filename,
#         sheet_name=sheet_name,
#         engine="openpyxl",
#         header=None,
#     )
#     print(df)

#     # 표 시작 위치 찾기, main 추가급여가 바로 엽에 있는 활동지원등급 찾기
#     pos = (df == "활동지원등급").to_numpy().nonzero()
#     for r,c in zip(*pos):
#         if str(df.iat[r,c+1]).strip()=="main 추가급여":
#             r0,c0=r,c
#             break
    
#     header0 = df.iloc[r0].astype(str).str.replace(" ", "")#본인부담금 블록명
#     header1 = df.iloc[r0 + 1].astype(str).str.replace(" ", "")# 학교/직장, 기본형/확장형

#     school_cols = header1[header1 == "학교생활"].index.tolist()
#     work_cols   = header1[header1 == "직장생활"].index.tolist()
#     school_time_col = school_cols[1]                
#     work_time_col   = work_cols[1]
#     basic_col   = header1[header1 == "기본형(시간/금액)"].index[0] 
#     extend_col  = header1[header1 == "확장형(시간/금액)"].index[0]

#     #본인부담금 기본형, 확장형
#     def block_cols(name):
#         start=header0[header0 == name].index[0]
#         return list(range(start,start+6))
    
#     day_basic_cols    = block_cols("주간활동본인부담금(기본형)")
#     day_ext_cols      = block_cols("주간활동본인부담금(확장형)")

#     GAGU = ["최중증취약가구", "최중증1인가구", "1등급취약가구", "1등급1인가구",
#         "2등급이하취약가구", "2등급이하1인가구", "나머지가구구성원의직장생활등"]

#         #학교생활, 직장생활 o
#     def flag(i, time_col, main, name):
#         if main == name:
#             return True
#         return pd.notna(df.iat[i, time_col])
    
#     def to_won(v):#면제와 콤마 처리
#         s = str(v).strip()
#         return 0 if s == "면제" else int(float(s.replace(",", "")))
    
#     table={}
#     for i in range(r0+2,len(df)):
#         grade=df.iat[i,c0]
#         if pd.isna(grade) or str(grade).strip() not in ["1등급", "2등급", "3등급", "4등급"]:
#             continue
#         grade = str(grade).strip()
#         main = str(df.iat[i, c0 + 1]).strip()
#         gagu_type= main if main in GAGU else None
#         key = (grade, gagu_type,
#                flag(i, school_time_col, main, "학교생활"),
#                flag(i, work_time_col,   main, "직장생활"))
#         table[key]={
#             "주간기본금액":to_won(df.iat[i,basic_col+1]),
#             "주간확장금액": to_won(df.iat[i, extend_col+1]),
#             "주간기본부담": [to_won(df.iat[i, c]) for c in day_basic_cols],
#             "주간확장부담": [to_won(df.iat[i, c]) for c in day_ext_cols],
#         }

#     return table


# 종합조사 시트
def read_jh_table(filename, sheet_name):
    pass


# 기본급여 단가표 작성에 필요한 Dataframe과 헤더 위치 정보 가져오기
#BASIC_MONTHLY_LIMIT.append((8, 5))      # 위쪽 표의 기본형 열
#EXTENDED_MONTHLY_LIMIT.append((8, 6))   # 위쪽 표의 확장형 열
#BASIC_MONTHLY_LIMIT.append((30, 5))     # 아래쪽 표의 기본형 열
#EXTENDED_MONTHLY_LIMIT.append((30, 6))  # 아래쪽 표의 확장형 열
#BASIC_MONTHLY_LIMIT    = [(8, 5), (30, 5)]   # 기본형만 모임 ← 맞아요!
#EXTENDED_MONTHLY_LIMIT = [(8, 6), (30, 6)]   # 확장형만 모임 ← 맞아요!
#[0]=위쪽 표      [1]=아래쪽 표
#                     (기본급여)       (추가급여)
#BASIC (기본형)   →    (8, 5)          (30, 5)
#EXTENDED (확장형) →   (8, 6)          (30, 6)
def get_basic_df_headers(filename, sheet_names):

    ij = read_ij_table(filename, sheet_names[0])
    # sj = read_sj_table(filename, sheet_names[1])
    # jh = read_jh_table(filename, sheet_names[2])

    return (ij[0], ij[1], (ij[2][0], ij[3][0]), (ij[4][0], ij[5][0]))
    # , (sj[0], sj[1], ...)
    # , (jh[0], jh[1], ...)


# 추가급여 단가표 작성에 필요한 dataframe과 헤더 위치 정보 가져오기
# 인정조사 등급만 해당됨
#(30, 5)->ij[2][1]
def get_extended_headers(filename, sheet_name):

    ij = read_ij_table(filename, sheet_name)

    return (ij[0], ij[1], (ij[2][1], ij[3][1]), (ij[4][1], ij[5][1]))




def read_ij_sj_table_extra(filename, sheet_name1, sheet_name2):
    #파일 열기
    dfs = pd.read_excel(filename, sheet_name=[sheet_name1, sheet_name2],
                    engine="openpyxl", header=None)
    df1, df2 = dfs[sheet_name1], dfs[sheet_name2]

    BASE_ITEMS = [
        "최중증1인가구", "1등급1인가구", "2등급이하1인가구",
        "최중증취약가구", "1등급취약가구", "2등급이하취약가구",
        "출산", "자립준비", "학교생활", "직장생활",
        "보호자일시부재", "나머지가구구성원의직장생활등",
    ]

    #최중층1인가구 시작위치 찾기
    rows,cols=(df1=="최중증1인가구").to_numpy().nonzero()
    assert len(rows) >= 1, "'최중증1인가구' 없음"
    r,c=rows[0],cols[0]
    GAGU=[]
    for i in range(len(BASE_ITEMS)):
        cell=df1.iat[r,c+i]
        text=str(cell)#문자열로 변환
        text=text.strip()#앞 뒤 공백 제거
        GAGU.append(text)
    assert GAGU==BASE_ITEMS, f"해더 불일치:{GAGU}" #조건에 안 맞으면 해더 불일치

    GAGU_AMOUNTS={}
    for i in range(len(BASE_ITEMS)):
        name=BASE_ITEMS[i]
        cell=df1.iat[r+1,c+i]
        text=str(cell)
        text=text.replace(",","")
        amount=int(text)#숫자로 변환
        GAGU_AMOUNTS[name]=amount
    
    #추가 부담률 시작 위치 찾기
    ADD_PERCENT=[
        "50% 이하","100% 이하","150% 이하","150% 초과",
    ]
    rows,cols=(df2=="추가 부담률").to_numpy().nonzero()
    assert len(rows) == 1, f"'추가 부담률' {len(rows)}개"#중복채크
    r,c=rows[0],cols[0]
    ADD_PAYPERCENT={}
    for i in range(len(ADD_PERCENT)):
        name=ADD_PERCENT[i]
        cell=df2.iat[r,c+i+1]
        text=str(cell)#문자열로 변환
        amount=float(text)#숫자로 변환
        ADD_PAYPERCENT[name]=amount

    return (
        GAGU_AMOUNTS,
        ADD_PAYPERCENT,
    )