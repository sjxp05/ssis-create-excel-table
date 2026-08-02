from __future__ import annotations
import pandas as pd
import itertools
import re

# 조견표에서 찾는 문구 상수로 정리 - 서식 변경시 여기만 변경
BASE_PRICE = "기본단가"
A_VALUE = "A값"
BASIC_RATE = "기본 부담률"
GRADE_HEADER = "활동지원등급"
INCOME_HEADER = "기준중위소득"
JH_BASIC = "주간활동 기본형"
JH_EXTENDED = "주간활동 확장형"

IJ_GRADES = ("1등급", "2등급", "3등급", "4등급")
JH_ZONES = tuple(f"{i}등급" for i in range(1, 16))
ADD_ITEMS = (
    "최중증1인가구",
    "1등급1인가구",
    "2등급이하1인가구",
    "최중증취약가구",
    "1등급취약가구",
    "2등급이하취약가구",
    "출산",
    "자립준비",
    "학교생활",
    "직장생활",
    "보호자일시부재",
    "나머지가구구성원의직장생활등",
)

class ExtractError(Exception):
    """ 에러용 클래스 """

# 공백·줄바꿈·비단절공백
_WS = re.compile(r"[\s ]+")

# _WS 처리
def _squeeze(text):
    return _WS.sub("", str(text))

# 파일 읽고 원본(df)과 공백 제거한 검색용 사본(norm)을 함께 반환
def _load(filename, sheet_name):
    df = pd.read_excel(filename, sheet_name, engine="openpyxl", header=None)
    norm = df.astype(str).apply(lambda s: s.str.replace(_WS, "", regex=True))
    return df, norm

# 키워드가 나오는 모든 칸을 읽는 순서(위→아래, 왼→오른쪽)로 반환
def _find_all(norm, keyword):
    key = _squeeze(keyword)
    # 모든 칸을 True, False 로 표시
    hit = norm.apply(lambda s: s.str.contains(key, regex=False, na=False))
    found = sorted(
        (int(r), int(c)) for r, c in zip(*hit.to_numpy().nonzero(), strict=True)
    )
    if not found:
        raise ExtractError(f"'{keyword}'를 찾지 못했습니다.")
    return found

# 한 곳에만 있어야 하는 문구, 여러 번 나오면 예외
def _find_one(norm, keyword):
    found = _find_all(norm, keyword)
    if len({r for r, _ in found}) > 1:
        raise ExtractError(f"'{keyword}'가 여러 표에 있습니다")
    return found[0]

# 여러 곳에 반복되는 게 정상인 문구. 첫 번째만 쓴다
def _find_first(norm, keyword):
    return _find_all(norm, keyword)[0]

# 엑셀에서 온 value 받아서 계산가능한 숫자로 반환
def _num(value):
    item = getattr(value, "item", None)
    value = item() if callable(item) else value
    
    numeric = isinstance(value, (int, float)) and not isinstance(value, bool)
    if numeric and not pd.isna(value):
        return int(value) if float(value).is_integer() else float(value)
    
    text = _squeeze(value).replace(",", "").replace("%", "")
    if text.lower() in ("", "nan", "none"):
        raise ExtractError("값이 비어 있습니다")
    try:
        number = float(text)
    except ValueError:
        raise ExtractError(f"수치로 읽을 수 없는 값: {value!r}") from None
    return int(number) if number.is_integer() else number
    
    
# 인정조사 시트 읽고 A값, 본인부담금 상한액, 본인부담률 등 값 반환
def read_ij_value(filename, sheet_name):
    # 파일 열기
    df, norm = _load(filename, sheet_name)

    # A값, 본인부담금 상한액
    r, c = _find_one(norm, A_VALUE)
    a_value = _num(df.iat[r, c + 1])
    copay_cap = _num(df.iat[r, c + 2])

    # 기본단가
    r, c = _find_one(norm, BASE_PRICE)
    base_price = _num(df.iat[r, c + 1])

    # 본인 부담률
    r, c = _find_one(norm, BASIC_RATE)
    basic_rates = [_num(df.iat[r, c + i]) for i in range(1, 5)]
    add_rates = [_num(df.iat[r + 1, c + i]) for i in range(1,5)]

    # 월 한도액
    r, c = _find_one(norm, GRADE_HEADER)
    base_r, base_c = r + 2, c
    row_map: dict[str, int] = {}
    for i in range(5):
        if base_r + i >= len(norm.index):
            break
        name = norm.iat[base_r + i, base_c]
        if name in IJ_GRADES and name not in row_map:
            row_map[name] = i
    
    missing = [g for g in IJ_GRADES if g not in row_map]
    if missing:
        raise ExtractError(f"등급을 찾지 못했습니다: {missing}")

    basic_limits, extended_limits = [], []
    for grade in IJ_GRADES:
        i = row_map[grade]
        basic_limits.append(
            _num(df.iat[base_r + i, base_c + 3])
        )
        extended_limits.append(
            _num(df.iat[base_r + i, base_c + 4])
        )

    return {
        "기본단가": base_price,
        "A값": a_value,
        "본인부담금 상한액": copay_cap,
        "인정조사 본인부담률 (기본급여)": basic_rates,
        "인정조사 본인부담률 (추가급여)": add_rates,
        "인정조사 월한도액 (기본형)": basic_limits,
        "인정조사 월한도액 (확장형)": extended_limits,
    }

# 산정 특례
def read_sj_value(filename, sheet_name):
    df, norm = _load(filename, sheet_name)

    # 본인부담률
    r, c = _find_one(norm, INCOME_HEADER)
    rates = [ _num(df.iat[r + 1, c + i]) for i in range(4)]

    # 추가급여 월 한도액 
    base_r, base_c = _find_first(norm, ADD_ITEMS[0])
    col_map: dict[str, int] = {}
    for i in range(len(norm.columns) - base_c):
        name = norm.iat[base_r, base_c + i]
        if name in ADD_ITEMS and name not in col_map:
            col_map[name] = i

    missing = [k for k in ADD_ITEMS if k not in col_map]
    if missing:
        raise ExtractError(
            f"추가급여 항목을 찾지 못했습니다: {missing}"
        )

    limits = {
        k: _num(df.iat[base_r + 1, base_c + col_map[k]])
        for k in ADD_ITEMS
    }
    return {"종합조사/산정특례 본인부담률": rates, "추가급여 월한도액": limits}

def _read_jh_row(df, norm, keyword):
    r, _ = _find_first(norm, keyword)
    base_r = r + 1

    col_map: dict[str, int] = {}
    for offset in (-2, -1, 0):
        row = base_r + offset
        if not 0 <= row < len(norm.index):
            continue
        for col in range(len(norm.columns)):
            name = norm.iat[row, col]
            if name in JH_ZONES and name not in col_map:
                col_map[name] = col

    missing = [z for z in JH_ZONES if z not in col_map]
    if missing:
        raise ExtractError(f"'{keyword}' 표에서 구간을 찾지 못했습니다: {missing}")
    return [
        _num(df.iat[base_r, col_map[z]]) for z in JH_ZONES
    ]


def read_jh_value(filename, sheet_name):
    df, norm = _load(filename, sheet_name)
    where = f"종합조사 시트('{sheet_name}')"
    return {
        "종합조사 월한도액 (기본형)": _read_jh_row(df, norm, JH_BASIC),
        "종합조사 월한도액 (확장형)": _read_jh_row(df, norm, JH_EXTENDED),
    }


_EXPECTED_LEN = {
    "인정조사 본인부담률 (기본급여)": 4,
    "인정조사 본인부담률 (추가급여)": 4,
    "종합조사/산정특례 본인부담률": 4,
    "인정조사 월한도액 (기본형)": 4,
    "인정조사 월한도액 (확장형)": 4,
    "종합조사 월한도액 (기본형)": 15,
    "종합조사 월한도액 (확장형)": 15,
}

_AMOUNT_KEYS = ("기본단가", "A값", "본인부담금 상한액")
_LIMIT_KEYS = tuple(k for k in _EXPECTED_LEN if "월한도액" in k)

def _validate(data):
    for key, size in _EXPECTED_LEN.items():
        actual = len(data[key])
        if actual != size:
            raise ExtractError(f"'{key}' 항목이 {size}개여야 하는데 {actual}개입니다.")
    if len(data["추가급여 월한도액"]) != len(ADD_ITEMS):
        raise ExtractError(f"'추가급여 월한도액'이 {len(ADD_ITEMS)}개가 아닙니다.")

    # 금액은 모두 양수여야 한다
    amounts = {k: data[k] for k in _AMOUNT_KEYS}
    for key in _LIMIT_KEYS:
        amounts.update({f"{key}[{i + 1}]": v for i, v in enumerate(data[key])})
    amounts.update(
        {f"추가급여 월한도액[{k}]": v for k, v in data["추가급여 월한도액"].items()}
    )
    bad = {k: v for k, v in amounts.items() if v <= 0}
    if bad:
        raise ExtractError(f"금액이 0 이하인 항목이 있습니다: {bad}")

    # 종합조사 월 한도액은 구간이 올라갈수록 감소
    for key in ("종합조사 월한도액 (기본형)", "종합조사 월한도액 (확장형)"):
        series = data[key]
        if any(a <= b for a, b in itertools.pairwise(series)):
            raise ExtractError(
                f"'{key}'가 구간 순으로 감소하지 않습니다: {series}"
                "\n조치: 조견표의 구간 배치가 바뀌었는지 확인하세요."
            )



def get_data_for_app(filename, sheet_names):
    readers = (read_ij_value, read_sj_value, read_jh_value)
    data = {}
    for reader, sheet in zip(readers, sheet_names, strict=True):
        try:
            data.update(reader(filename, sheet))
        except ExtractError as error:
            raise ExtractError(f"'{sheet}' 시트 — {error}") from None
    _validate(data)
    return data