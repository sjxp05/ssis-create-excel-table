# -*- coding: utf-8 -*-
"""실제(내부등록용) 조견표 대응 리더

팀의 table_reader.read_jh_table 은 조견표_샘플 기준으로 작성되어 있는데,
실제 조견표의 종합조사 시트에는 같은 앵커 문구("주간활동 확장형")가
불완전한 블록(15등급만 있고 나머지는 빈 표)에 한 번 더 나타나서
마지막 발견 위치를 쓰는 기존 로직이 NaN을 읽고 실패한다.

이 모듈은 read_jh_table 만 보완하고(기존 팀 코드는 그대로 둠),
인정조사/산정특례는 팀의 리더를 그대로 재사용한다.

보완 방식 (gosi_reader와 같은 원칙):
1. 앵커와 같은 행에 완전한 등급 헤더('1등급'까지)가 있는 블록만 진짜 표로 인정
2. 열 위치를 고정 오프셋(+17)이 아니라 '1등급' 라벨의 실제 위치에서 찾음
3. 읽을 자리의 값이 숫자가 아니면 조용히 넘어가지 않고 즉시 에러

사용: main.py 에서 table_reader 대신 이 모듈의 get_basic_df_headers 를 import
"""

import pandas as pd

from app.table_reader import read_ij_table, read_sj_table


import re


def _grade_columns(df, row):
    """헤더 행에서 'N등급' 라벨 -> 열 위치 매핑을 만든다."""
    cols = {}
    for col in range(df.shape[1]):
        v = df.iat[row, col]
        if isinstance(v, str) and re.fullmatch(r"\d+등급", v.strip()):
            cols[v.strip()] = col
    return cols


def _find_block(df, anchor):
    """앵커 문구의 블록들을 찾아 (완전한 본표 1개, 부분 보정표들)로 나눈다.

    - 본표: 같은 행에 '1등급'까지 15개 등급 헤더를 모두 갖춘 표
    - 보정표: 일부 등급만 있는 표. 실제 조견표에는 확장형 15구간처럼
      본표와 값이 다른 구간을 하단에 별도 표로 두는 경우가 있다
      (2026 조견표에서 확인: 하단 표의 661,000이 고시 제2장 표와 일치)

    반환: ((본표 앵커 행, '1등급' 열), [(보정표 앵커 행, {등급라벨: 열}), ...])
    """
    complete, partial = [], []
    for row, col in zip(*((df == anchor).to_numpy().nonzero())):
        grades = _grade_columns(df, int(row))
        if "1등급" in grades and len(grades) >= 15:
            complete.append((int(row), grades["1등급"]))
        elif grades:  # 일부 등급만 있는 블록 = 보정표
            partial.append((int(row), grades))
    if len(complete) != 1:
        raise ValueError(
            f"종합조사 시트에서 완전한 '{anchor}' 표를 {len(complete)}개 찾았습니다 (1개여야 함). "
            f"조견표 형식이 바뀌었는지 확인하세요."
        )
    return complete[0], partial


def _apply_corrections(df, anchor, main_row, partials):
    """보정표의 값(월 한도액 1행 + 소득구간 4행)을 본표의 해당 등급 열에 덮어쓴다."""
    main_grades = _grade_columns(df, main_row)
    for p_row, p_grades in partials:
        for label, p_col in p_grades.items():
            if label not in main_grades:
                raise ValueError(
                    f"'{anchor}' 보정표의 '{label}' 열이 본표에 없습니다. 조견표 형식을 확인하세요."
                )
            m_col = main_grades[label]
            for dr in range(1, 6):  # +1 월한도액, +2~+5 소득구간 4행
                v = df.iat[p_row + dr, p_col]
                if pd.isna(v):
                    raise ValueError(
                        f"'{anchor}' 보정표({label})의 값이 비어 있습니다: 행 {p_row + dr}"
                    )
                df.iat[main_row + dr, m_col] = v
            print(f"[조견표] '{anchor}' {label} 값을 하단 보정표 값으로 대체했습니다.")


def _check_numeric(df, name, row, col, count):
    """(row, col)에서 왼쪽으로 count개 칸이 전부 숫자인지 확인 (구간 값 검증)"""
    for g in range(count):
        v = df.iat[row, col - g]
        if pd.isna(v) or not isinstance(v, (int, float)):
            raise ValueError(
                f"종합조사 '{name}'의 {g + 1}구간 값이 숫자가 아닙니다: "
                f"행 {row}, 열 {col - g} = {v!r}. 조견표 형식을 확인하세요."
            )


def read_jh_table(filename, sheet_name):
    """종합조사 시트 읽기 - 실제 조견표 대응판 (반환 형식은 팀 코드와 동일)

    반환: (df, 기본형 월한도액 위치, 확장형 월한도액 위치,
           기본형 본인부담금 위치, 확장형 본인부담금 위치)
    각 위치는 (행, 1구간 열) 튜플이고, write_basic_table 이
    열 - g (g=0~14) 로 1~15구간을 읽는 방식과 호환된다.
    """
    df = pd.read_excel(io=filename, sheet_name=sheet_name, engine="openpyxl", header=None)

    (r, c), partials = _find_block(df, "주간활동 기본형")
    _apply_corrections(df, "주간활동 기본형", r, partials)
    BASIC_MONTHLY_LIMIT = (r + 1, c)   # 앵커 행 바로 아래 = 월 한도액
    BASIC_COPAYMENT = (r + 2, c)       # 그 아래부터 = 70%이하~180%초과 부담금

    (r, c), partials = _find_block(df, "주간활동 확장형")
    _apply_corrections(df, "주간활동 확장형", r, partials)
    EXTENDED_MONTHLY_LIMIT = (r + 1, c)
    EXTENDED_COPAYMENT = (r + 2, c)

    # 읽을 자리가 전부 실제 숫자인지 검증: 월한도액 15구간, 부담금 4개 소득행 x 15구간
    for name, (row, col) in [
        ("기본형 월 한도액", BASIC_MONTHLY_LIMIT),
        ("확장형 월 한도액", EXTENDED_MONTHLY_LIMIT),
    ]:
        _check_numeric(df, name, row, col, 15)
    for name, (row, col) in [
        ("기본형 본인부담금", BASIC_COPAYMENT),
        ("확장형 본인부담금", EXTENDED_COPAYMENT),
    ]:
        for ir in range(4):  # 70%이하 / 120%이하 / 180%이하 / 180%초과
            _check_numeric(df, f"{name}({ir + 1}번째 소득행)", row + ir, col, 15)

    return (
        df,
        BASIC_MONTHLY_LIMIT,
        EXTENDED_MONTHLY_LIMIT,
        BASIC_COPAYMENT,
        EXTENDED_COPAYMENT,
    )


def get_basic_df_headers(filename, sheet_names):
    """팀의 get_basic_df_headers 와 동일한 반환 형식.

    인정조사/산정특례는 팀 리더를 그대로 쓰고,
    종합조사만 이 모듈의 보완 리더를 사용한다.
    """
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
