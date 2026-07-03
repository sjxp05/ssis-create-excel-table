import pandas as pd


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
    for row, col in zip(*((df == "활동지원등급").to_numpy().nonzero())):
        GRADE_HEADER = (row + 2, col)

    # 월 한도액(기본/확장) 시작위치 찾기
    BASIC_MONTHLY_LIMIT = []
    EXTENDED_MONTHLY_LIMIT = []
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
def read_sj_table(filename, sheet_name):
    pass


# 종합조사 시트
def read_jh_table(filename, sheet_name):
    pass


# 기본급여 단가표 작성에 필요한 Dataframe과 헤더 위치 정보 가져오기
def get_basic_df_headers(filename, sheet_names):

    ij = read_ij_table(filename, sheet_names[0])
    # sj = read_sj_table(filename, sheet_names[1])
    # jh = read_jh_table(filename, sheet_names[2])

    return (ij[0], ij[1], (ij[2][0], ij[3][0]), (ij[4][0], ij[5][0]))
    # , (sj[0], sj[1], ...)
    # , (jh[0], jh[1], ...)


# 추가급여 단가표 작성에 필요한 dataframe과 헤더 위치 정보 가져오기
# 인정조사 등급만 해당됨
def get_extended_headers(filename, sheet_name):

    ij = read_ij_table(filename, sheet_name)

    return (ij[0], ij[1], (ij[2][1], ij[3][1]), (ij[4][1], ij[5][1]))
