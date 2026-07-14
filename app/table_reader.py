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

    return (
        df,
        GRADE_HEADER,
        BASIC_MONTHLY_LIMIT,
        EXTENDED_MONTHLY_LIMIT,
        BASIC_COPAYMENT,
        EXTENDED_COPAYMENT,
    )


# 산정조사 시트
def read_sj_table(filename, sheet_name):
    pass


# 종합조사 시트
def read_jh_table(filename, sheet_name):
    # 파일 열기
    df = pd.read_excel(
        io=filename,
        sheet_name=sheet_name,
        engine="openpyxl",
        header=None,
    )
    print(df)
    # 주간활동 기본형 월 한도액, 본인부담률 시작 위치 찾기
    for row, col in zip(*((df == "주간활동 기본형").to_numpy().nonzero())):
        BASIC_MONTHLY_LIMIT = (row + 1, col + 17)
        BASIC_COPAYMENT = (row + 2, col + 17)

    # 주간활동 확장형 표 찾기
    for row, col in zip(*((df == "주간활동 확장형").to_numpy().nonzero())):
        EXTENDED_MONTHLY_LIMIT = (row + 1, col + 17)
        EXTENDED_COPAYMENT = (row + 2, col + 17)

    return (
        df,
        BASIC_MONTHLY_LIMIT,
        EXTENDED_MONTHLY_LIMIT,
        BASIC_COPAYMENT,
        EXTENDED_COPAYMENT,
    )


# 산정특례 시트
def read_sj_table(filename, sheet_name):
    # 파일 열기
    df = pd.read_excel(
        io=filename,
        sheet_name=sheet_name,
        engine="openpyxl",
        header=None,
    )
    print(df)

    # 활동지원등급 시작위치 찾기
    TARGET_ROW = -1
    for row, col in zip(*((df == "활동지원등급").to_numpy().nonzero())):
        if row + 1 < len(df) and "기초" in df.iloc[row + 1].values:
            TARGET_ROW = row
            break

    for row, col in zip(*((df == "활동지원등급").to_numpy().nonzero())):
        if row == TARGET_ROW:  # 위쪽 예시 표 무시
            GRADE_HEADER = (row + 2, col)
            break

    # 월 한도액(기본/확장) 시작위치 찾기
    BASIC_MONTHLY_LIMIT = []
    EXTENDED_MONTHLY_LIMIT = []
    for row, col in zip(*((df == "월 한도액").to_numpy().nonzero())):
        if row == TARGET_ROW:
            BASIC_MONTHLY_LIMIT.append((row + 2, col+2))
            EXTENDED_MONTHLY_LIMIT.append((row + 2, col+4))

    # 구간별 본인부담금(기본형) 시작위치 찾기
    BASIC_COPAYMENT = []
    for row, col in zip(*((df == "주간활동 본인부담금(기본형)").to_numpy().nonzero())):
        BASIC_COPAYMENT.append((row + 2, col))

    # 구간별 본인부담금(확장형) 시작위치 찾기
    EXTENDED_COPAYMENT = []
    for row, col in zip(*((df == "주간활동 본인부담금(확장형)").to_numpy().nonzero())):
        EXTENDED_COPAYMENT.append((row + 2, col))

    return (
        df,
        GRADE_HEADER,
        BASIC_MONTHLY_LIMIT,
        EXTENDED_MONTHLY_LIMIT,
        BASIC_COPAYMENT,
        EXTENDED_COPAYMENT,
    )


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
    jh = read_jh_table(filename, sheet_names[2])
    sj = read_sj_table(filename, sheet_names[1])

    return (
        ij[0],
        ij[1],
        (ij[2][0], ij[3][0]),
        (ij[4][0], ij[5][0]),
    ), (
        jh[0],
        (jh[1], jh[2]),
        (jh[3], jh[4]),
    ), (
        sj[0],
        sj[1],
        (sj[2][0], sj[3][0]),
        (sj[4][0], sj[5][0]),
    )


# 추가급여 단가표 작성에 필요한 dataframe과 헤더 위치 정보 가져오기
# 인정조사 등급만 해당됨
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