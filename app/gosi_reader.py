# -*- coding: utf-8 -*-
"""고시(hwpx)에서 급여비용 단가를 추출하는 모듈

설계 원칙
1. 표 위치(몇 번째 표)가 아니라 서비스의 공식 명칭(핵심 키워드)이
   제목(장/항목)에 나타나는 표를 찾는다 -> TABLE_ANCHORS 참고
   -> 고시가 개정되어 조항/표 순서나 세부 문구가 바뀌어도 동작
2. 추출한 값은 불변식으로 자가 검증하고, 어긋나면 조용히 틀리지 않고 에러를 낸다
3. 모든 금액에 출처(섹션/표번호/행/열/원문/문맥)를 저장해
   나중에 고시 원문에서 근거 문구와 전후 맥락을 찾아갈 수 있게 한다

사용 예:
    from app.gosi_reader import read_gosi_prices
    prices = read_gosi_prices("고시.hwpx")
    prices["활동보조"]["일반"]["금액"]   # 17270
    prices["활동보조"]["일반"]["출처"]   # 어느 표 몇 행에서 나왔는지

단독 실행(추출 결과 눈으로 확인용):
    python -m app.gosi_reader "고시파일.hwpx"
"""

import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET

# hwpx 본문 XML의 네임스페이스 (한글 hwpml 표준)
HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"

# ─────────────────────────────────────────────────────────────────────
# 앵커 키워드: 각 단가표는 [제목] 키워드가 고시의 장/항목 제목에
# 나타나는 표로 찾는다. 예: "제3장 ... > 2. 방문목욕" 아래의 표 = 방문목욕 단가표
#
# 실제 설정은 params/앵커_키워드.json 파일에 있다 (없으면 아래 기본값으로 자동 생성).
# 고시에서 서비스 명칭이 바뀌면 담당자는 그 JSON 파일만 수정하면 되고,
# 코드는 열 필요 없다. 아래 기본값은 파일이 없을 때의 원본일 뿐이다.
# ─────────────────────────────────────────────────────────────────────
ANCHORS_PATH = "params/앵커_키워드.json"

DEFAULT_ANCHORS = {
    "활동보조": {"제목": "활동보조", "표안": []},
    "방문목욕": {"제목": "방문목욕", "표안": []},
    "방문간호": {"제목": "방문간호", "표안": []},
    "방문간호지시서": {"제목": "방문간호지시서", "표안": []},
}

_ANCHORS_GUIDE = (
    "고시에서 각 단가표를 찾는 앵커 키워드입니다. "
    "고시 개정으로 서비스 명칭이 바뀌면 이 파일에서 해당 '제목'만 새 명칭으로 수정하세요. "
    "같은 제목 아래 표가 여러 개 생기면 '표안'에 그 표에만 있는 문구를 추가해 좁히세요. "
    "코드는 열 필요 없습니다. 이 파일을 삭제하면 다음 실행 때 기본값으로 다시 생성됩니다."
)


def load_anchors(path=ANCHORS_PATH):
    """앵커 키워드 설정을 파일에서 읽는다. 파일이 없으면 기본값으로 만들어 준다.

    담당자는 이 파일의 '제목'/'표안' 문구만 고치면 되고 코드는 열 필요 없다.
    """
    if not os.path.exists(path):
        folder = os.path.dirname(path)
        if folder:
            os.makedirs(folder, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"_안내": _ANCHORS_GUIDE, **DEFAULT_ANCHORS},
                      f, ensure_ascii=False, indent=2)
        print(f"앵커 키워드 설정 파일을 새로 생성했습니다: {path}")

    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    anchors = {k: v for k, v in data.items() if not k.startswith("_")}

    # 필요한 서비스가 모두 있고 형식이 맞는지 확인 (수동 수정 실수 방지)
    for name in DEFAULT_ANCHORS:
        if name not in anchors:
            raise ValueError(
                f"앵커 설정에 '{name}' 항목이 없습니다: {path}\n"
                "조치: 항목을 다시 추가하거나, 파일을 삭제하면 기본값으로 재생성됩니다."
            )
        a = anchors[name]
        if not (isinstance(a.get("제목"), str) and a["제목"].strip()):
            raise ValueError(f"'{name}' 앵커의 '제목'이 비어있거나 잘못되었습니다: {path}")
        if not isinstance(a.get("표안"), list):
            raise ValueError(f"'{name}' 앵커의 '표안'은 문구 목록이어야 합니다: {path}")
    return anchors


# ──────────────────────────── hwpx 파싱 ────────────────────────────


def _cell_text(tc):
    # 셀 안의 모든 텍스트 조각(hp:t)을 이어붙인다
    return "".join(t.text or "" for t in tc.iter(f"{HP}t")).strip()


def _walk_document(root):
    """본문을 문서 순서대로 순회하며 ("text", 문단텍스트) / ("table", tbl요소) 목록 반환

    표 앞에 나오는 문단(장 제목, "1. 활동보조" 등)을 문맥으로 쓰기 위해
    표 밖의 텍스트와 표를 구분해서 수집한다.
    """
    items = []

    def rec(el):
        for child in el:
            if child.tag == f"{HP}tbl":
                items.append(("table", child))
            elif child.tag == f"{HP}p":
                if child.find(f".//{HP}tbl") is not None:
                    rec(child)  # 문단 안에 표가 있으면 더 내려가서 표를 찾는다
                else:
                    txt = "".join(t.text or "" for t in child.iter(f"{HP}t")).strip()
                    if txt:
                        items.append(("text", txt))
            else:
                rec(child)

    rec(root)
    return items


def parse_tables(hwpx_path):
    """hwpx의 모든 표를 문서 순서대로 반환

    문서를 순서대로 걸으면서 "지금 어느 장(제N장), 어느 번호 항목(1. xxx) 아래인지"를
    추적하다가 표를 만나면 그 시점의 위치를 문맥으로 붙인다.
    -> 사람이 고시를 열어 '제3장 > 1. 활동보조'로 찾아갈 수 있는 구조적 좌표

    반환: [{"섹션": xml파일명, "표번호": n, "격자": [[셀텍스트]],
            "문맥": {"장": "제3장 ...", "항목": "1. 활동보조"}}]
    """
    tables = []
    with zipfile.ZipFile(hwpx_path) as z:
        section_names = sorted(
            n for n in z.namelist() if re.match(r"Contents/section\d+\.xml", n)
        )
        if not section_names:
            raise ValueError(f"hwpx 안에 본문(section*.xml)이 없습니다: {hwpx_path}")

        table_no = 0
        chapter = None  # 현재 장 제목 (예: "제3장 급여비용 및 산정기준", "부칙")
        item = None     # 현재 번호 항목 (예: "1. 활동보조")
        for name in section_names:
            root = ET.fromstring(z.read(name))
            for kind, payload in _walk_document(root):
                if kind == "text":
                    if re.match(r"^(제\s*\d+\s*장|부\s*칙)", payload):
                        chapter = payload
                        item = None  # 장이 바뀌면 항목도 처음부터
                    elif re.match(r"^\d+\.", payload):
                        # 항목이 긴 문장인 경우 앞부분만 저장 (위치 표지로는 충분)
                        item = payload if len(payload) <= 40 else payload[:40] + "…"
                else:
                    grid = [
                        [_cell_text(tc) for tc in tr.iter(f"{HP}tc")]
                        for tr in payload.iter(f"{HP}tr")
                    ]
                    tables.append({
                        "섹션": name,
                        "표번호": table_no,
                        "격자": grid,
                        "문맥": {"장": chapter, "항목": item},
                    })
                    table_no += 1
    return tables


def _heading(table):
    """표의 제목 = 문맥(장 + 항목)을 하나의 문자열로"""
    ctx = table["문맥"]
    return " ".join(x for x in [ctx.get("장"), ctx.get("항목")] if x)


def find_table(tables, service, anchors):
    """서비스 이름(핵심 키워드)으로 단가표를 찾는다.

    고시의 장/항목 제목에 앵커 설정의 [제목] 키워드가 나타나는 표를 찾는다.
    한 표의 제목에 여러 키워드가 걸리면 더 긴 키워드의 표로 본다
    (예: '방문간호지시서' 표의 제목에는 '방문간호'도 들어있다).
    0개/2개 이상이면 발견된 표 목록과 수정 지점을 안내하며 에러 (조용히 틀리지 않기).
    """
    def owner(t):
        # 이 표의 제목에 걸리는 앵커 키워드 중 가장 긴 것이 이 표의 주인
        matched = [n for n, a in anchors.items() if a["제목"] in _heading(t)]
        return max(matched, key=len) if matched else None

    extra = anchors[service]["표안"]
    hits = []
    for t in tables:
        if owner(t) != service:
            continue
        flat = " ".join(c for row in t["격자"] for c in row)
        if all(k in flat for k in extra):
            hits.append(t)

    if len(hits) != 1:
        listing = "\n".join(
            "  표{}: {} | 첫 행: {}".format(
                t["표번호"],
                _heading(t) or "(위치 불명)",
                " / ".join(c for c in t["격자"][0] if c)[:60] if t["격자"] else "(빈 표)",
            )
            for t in tables
        )
        raise ValueError(
            f"'{service}' 단가표가 {len(hits)}개 발견되었습니다 (1개여야 함).\n"
            f"이 고시에서 발견된 표 목록:\n{listing}\n"
            f"조치: 위 목록과 고시 원문을 대조하세요. 서비스 명칭이 개정되었다면 "
            f"{ANCHORS_PATH} 파일에서 '{service}'의 '제목' 키워드를 새 명칭으로 바꾸고, "
            "같은 제목 아래 표가 여러 개라면 '표안'에 그 표에만 있는 문구를 추가해 "
            "좁히세요. 코드는 열 필요 없습니다."
        )
    return hits[0]


def _amounts(text):
    # "17,270원" 같은 금액을 전부 int로 추출 (천 단위 콤마 포함 4자리 이상)
    return [int(m.replace(",", "")) for m in re.findall(r"([\d,]{4,})원", text)]


def _make_source(hwpx_path, table, row, col, label, raw):
    """금액의 출처 정보. 이후 근거 확인 시 이 정보로 고시 원문 위치를 찾아간다."""
    return {
        "파일": os.path.basename(hwpx_path),
        "섹션": table["섹션"],
        "표번호": table["표번호"],
        "행": row,
        "열": col,
        "라벨": label,
        "원문": raw,
        "문맥": table["문맥"],  # {"장": ..., "항목": ...} 사람이 고시에서 찾아갈 좌표
    }


# ──────────────────────────── 표별 추출 ────────────────────────────


def _read_hwaldong(hwpx_path, table):
    """활동보조 표: 행마다 (분류 라벨, 시간당 금액 + 가산수당)"""
    result = {}
    for r, row in enumerate(table["격자"]):
        label = row[0] if row else ""
        raw = " ".join(row)
        nums = _amounts(raw)
        if not nums:
            continue  # 헤더 행
        if len(nums) != 2:
            raise ValueError(f"활동보조 행에서 금액 2개(시간당+가산수당)를 기대했으나 {nums}: {raw}")
        # 라벨의 키워드로 분류 (조항 번호 ①②③에 의존하지 않음)
        if "심야" in label:
            key = "심야"
        elif "공휴일" in label or "근로자의 날" in label:
            key = "공휴일"
        else:
            key = "일반"
        if key in result:
            raise ValueError(f"활동보조 '{key}' 분류가 중복 추출되었습니다: {label}")
        result[key] = {
            "금액": nums[0],
            "가산수당": nums[1],
            "출처": _make_source(hwpx_path, table, r, 1, label, raw),
        }
    if set(result) != {"일반", "심야", "공휴일"}:
        raise ValueError(f"활동보조에서 일반/심야/공휴일을 기대했으나: {sorted(result)}")
    return result


def _read_mokyok(hwpx_path, table):
    """방문목욕 표: 차량내입욕 / 가정내입욕"""
    result = {}
    for r, row in enumerate(table["격자"]):
        label = row[0] if row else ""
        raw = " ".join(row)
        nums = _amounts(raw)
        if not nums:
            continue
        if len(nums) != 1:
            raise ValueError(f"방문목욕 행에서 금액 1개를 기대했으나 {nums}: {raw}")
        # 주의: 가정내입욕 라벨에도 "차량 내 온수"라는 문구가 들어있으므로
        #       "차량"이 아니라 목욕 장소를 나타내는 특징 문구로 구분한다
        squeezed = label.replace(" ", "")
        if "이동목욕용" in label and "차량내에서" in squeezed:
            key = "차량내입욕"
        elif "가정내에서" in squeezed:
            key = "가정내입욕"
        else:
            raise ValueError(f"방문목욕 분류를 인식할 수 없습니다: {label}")
        if key in result:
            raise ValueError(f"방문목욕 '{key}' 분류가 중복 추출되었습니다: {label}")
        result[key] = {
            "금액": nums[0],
            "출처": _make_source(hwpx_path, table, r, 1, label, raw),
        }
    if set(result) != {"차량내입욕", "가정내입욕"}:
        raise ValueError(f"방문목욕에서 차량내/가정내입욕을 기대했으나: {sorted(result)}")
    return result


def _read_ganho(hwpx_path, table):
    """방문간호 표: 시간 구간(30분 미만 / 30~60분 / 60분 이상)별 금액"""
    result = {}
    for r, row in enumerate(table["격자"]):
        label = row[0] if row else ""
        raw = " ".join(row)
        nums = _amounts(raw)
        if not nums:
            continue
        if len(nums) != 1:
            raise ValueError(f"방문간호 행에서 금액 1개를 기대했으나 {nums}: {raw}")
        squeezed = label.replace(" ", "")
        # "30분 미만" / "30분 이상 ~ 60분 미만" / "60분 이상"을 키워드로 분류
        if "30분미만" in squeezed:
            key = "30분미만"
        elif "60분미만" in squeezed:
            key = "30분이상60분미만"
        elif "60분이상" in squeezed:
            key = "60분이상"
        else:
            raise ValueError(f"방문간호 시간 구간을 인식할 수 없습니다: {label}")
        if key in result:
            raise ValueError(f"방문간호 '{key}' 구간이 중복 추출되었습니다: {label}")
        result[key] = {
            "금액": nums[0],
            "출처": _make_source(hwpx_path, table, r, 1, label, raw),
        }
    if set(result) != {"30분미만", "30분이상60분미만", "60분이상"}:
        raise ValueError(f"방문간호에서 시간 구간 3개를 기대했으나: {sorted(result)}")
    return result


def _read_jisiseo(hwpx_path, table):
    """방문간호지시서 표: 기관(의료/보건) x 방문방식(대상자 방문/의사 내방)

    이 표는 한 행의 분류 셀에 '가. 대상자가 방문 / 나. 의사가 가정 방문' 두 경우가,
    금액 셀에 금액 2개가 함께 들어있다. 문서에 적힌 순서(가->나)와
    금액 순서가 대응한다고 보고 짝짓되, 개수가 다르면 에러를 낸다.
    """
    result = {}
    for r, row in enumerate(table["격자"]):
        label = row[0] if row else ""
        raw = " ".join(row)
        nums = _amounts(raw)
        if not nums:
            continue
        # 기관 구분
        if "의료법" in label or "의료기관" in label:
            org = "의료기관"
        elif "지역보건법" in label or "보건소" in label:
            org = "보건기관"
        else:
            raise ValueError(f"지시서 발급기관을 인식할 수 없습니다: {label}")
        # 방문방식: 라벨에 등장하는 순서대로 금액과 짝짓는다
        ways = []
        for m in re.finditer(r"(대상자가.*?방문|의사가\s*가정을\s*방문)", label):
            ways.append("방문" if m.group().startswith("대상자") else "의사내방")
        if len(ways) != len(nums):
            raise ValueError(
                f"지시서 '{org}'에서 경우 {len(ways)}개와 금액 {len(nums)}개가 "
                f"일치하지 않습니다: {raw}"
            )
        for way, amount in zip(ways, nums):
            key = f"{org}_{way}"
            if key in result:
                raise ValueError(f"지시서 '{key}' 분류가 중복 추출되었습니다: {label}")
            result[key] = {
                "금액": amount,
                "출처": _make_source(hwpx_path, table, r, 1, label, raw),
            }
    expected = {"의료기관_방문", "의료기관_의사내방", "보건기관_방문", "보건기관_의사내방"}
    if set(result) != expected:
        raise ValueError(f"지시서에서 {sorted(expected)}를 기대했으나: {sorted(result)}")
    return result


# ──────────────────────────── 자가 검증 ────────────────────────────


def show_price(item):
    """값과 그 법령상 위치를 담당자가 읽을 수 있는 한 줄로 만든다 (에러 메시지용)

    예: 42,880원 <- 제3장 급여비용 및 산정기준 > 3. 방문간호 > 표7 행1 "① 30분 미만"
    """
    s = item["출처"]
    ctx = s.get("문맥", {}) or {}
    where = " > ".join(x for x in [ctx.get("장"), ctx.get("항목")] if x)
    return f'{item["금액"]:,}원 <- {where} > 표{s["표번호"]} 행{s["행"]} "{s["라벨"]}"'


# 검증 실패 시 담당자에게 안내할 공통 조치 문구
_FIX_GUIDE = (
    "\n조치: 위 출처를 고시 원문에서 찾아(한글에서 Ctrl+F로 장/항목 검색) 값을 확인한 뒤, "
    "params/단가_파라미터.json의 해당 '금액'만 수정하세요. 코드는 고칠 필요 없습니다."
)


def _validate(prices):
    """추출값이 상식적 불변식을 만족하는지 확인. 어긋나면 즉시 에러.

    에러 메시지에 '문제의 값 + 법령상 위치(출처)'를 함께 띄워서,
    담당자가 코드를 열지 않고 고시 원문과 파라미터 모음집만으로
    확인/수정할 수 있게 한다.
    """
    # 모든 금액은 양수
    for service, items in prices.items():
        for key, item in items.items():
            if item["금액"] <= 0:
                raise ValueError(
                    f"{service}/{key} 금액이 0 이하입니다.\n"
                    f"  현재 값: {show_price(item)}{_FIX_GUIDE}"
                )

    # 방문간호: 제공시간이 길수록 금액이 커야 한다
    g = prices["방문간호"]
    if not (g["30분미만"]["금액"] < g["30분이상60분미만"]["금액"] < g["60분이상"]["금액"]):
        raise ValueError(
            "방문간호 금액이 시간 구간 순으로 증가하지 않습니다.\n"
            f"  30분미만        : {show_price(g['30분미만'])}\n"
            f"  30분이상60분미만: {show_price(g['30분이상60분미만'])}\n"
            f"  60분이상        : {show_price(g['60분이상'])}{_FIX_GUIDE}"
        )

    # 활동보조: 심야/공휴일은 일반보다 높아야 한다
    h = prices["활동보조"]
    if not (h["일반"]["금액"] < h["심야"]["금액"] and h["일반"]["금액"] < h["공휴일"]["금액"]):
        raise ValueError(
            "활동보조 심야/공휴일 금액이 일반보다 높지 않습니다.\n"
            f"  일반  : {show_price(h['일반'])}\n"
            f"  심야  : {show_price(h['심야'])}\n"
            f"  공휴일: {show_price(h['공휴일'])}{_FIX_GUIDE}"
        )


# ──────────────────────────── 공개 함수 ────────────────────────────


def read_gosi_prices(hwpx_path):
    """고시 hwpx에서 서비스별 단가를 추출한다.

    반환 구조 (모든 금액에 출처 포함):
    {
        "활동보조": {"일반"|"심야"|"공휴일": {"금액", "가산수당", "출처"}},
        "방문목욕": {"차량내입욕"|"가정내입욕": {"금액", "출처"}},
        "방문간호": {"30분미만"|"30분이상60분미만"|"60분이상": {"금액", "출처"}},
        "방문간호지시서": {"의료기관_방문"|"의료기관_의사내방"|
                          "보건기관_방문"|"보건기관_의사내방": {"금액", "출처"}},
    }
    """
    tables = parse_tables(hwpx_path)
    anchors = load_anchors()  # params/앵커_키워드.json (없으면 기본값으로 생성)
    prices = {
        "활동보조": _read_hwaldong(hwpx_path, find_table(tables, "활동보조", anchors)),
        "방문목욕": _read_mokyok(hwpx_path, find_table(tables, "방문목욕", anchors)),
        "방문간호": _read_ganho(hwpx_path, find_table(tables, "방문간호", anchors)),
        "방문간호지시서": _read_jisiseo(
            hwpx_path, find_table(tables, "방문간호지시서", anchors)
        ),
    }
    _validate(prices)
    return prices


if __name__ == "__main__":
    import sys

    from app.param_store import DEFAULT_PATH, save_params

    if len(sys.argv) < 2:
        print("사용법: python -m app.gosi_reader <고시파일.hwpx> [저장경로.json]")
        sys.exit(1)

    # 윈도우 콘솔 인코딩과 무관하게 한글이 깨지지 않도록
    sys.stdout.reconfigure(encoding="utf-8")

    prices = read_gosi_prices(sys.argv[1])
    # 추출 결과는 항상 파라미터 모음집 파일로 이어진다
    # -> 이후 단계는 고시가 아니라 이 파일을 읽고, 담당자 수정도 이 파일에서 한다
    out = sys.argv[2] if len(sys.argv) >= 3 else DEFAULT_PATH
    path = save_params(prices, os.path.basename(sys.argv[1]), out)
    print(f"단가 파라미터 모음집 저장 완료: {path}")
