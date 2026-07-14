# -*- coding: utf-8 -*-
"""단가 파라미터 모음집(JSON) 저장/불러오기

고시에서 추출한 단가는 항상 이 모음집 파일로 저장되고,
이후 단계(결제단가 생성 등)는 고시가 아니라 이 파일을 읽는다.

-> 담당자가 값을 수정해야 할 때는 이 파일 하나만 열어 '금액'을 고치면 되고,
   수정한 값은 재추출(gosi_reader 실행)을 하지 않는 한 그대로 유지된다.
-> 각 값의 근거(고시의 어느 표/행/문구에서 왔는지)는 값 옆의 '출처'에 있다.
"""

import json
import os
from datetime import datetime

# 파라미터 모음집의 기본 위치 (프로젝트 루트 기준)
DEFAULT_PATH = "params/단가_파라미터.json"

_GUIDE = (
    "이 파일은 고시에서 추출한 단가 파라미터 모음집입니다. "
    "값이 틀렸다면 이 파일에서 '금액'을 직접 수정하세요. "
    "단, gosi_reader를 다시 실행하면 고시 기준으로 다시 생성되어 수정 내용이 사라집니다. "
    "각 값의 근거는 바로 옆 '출처'에서 확인할 수 있습니다."
)


def save_params(prices, gosi_filename, path=DEFAULT_PATH):
    """추출된 단가를 파라미터 모음집 파일로 저장하고 경로를 반환"""
    data = {
        "_안내": _GUIDE,
        "원본_고시": gosi_filename,
        "생성시각": datetime.now().isoformat(timespec="seconds"),
        "단가": prices,
    }
    folder = os.path.dirname(path)
    if folder:
        os.makedirs(folder, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return path


def load_params(path=DEFAULT_PATH):
    """파라미터 모음집을 읽어 단가 dict를 반환 (담당자가 수정한 값 포함)

    불러올 때도 gosi_reader와 같은 불변식 검증을 수행해서
    수동 수정 실수(0 입력, 자릿수 오타 등)를 바로 잡아낸다.
    """
    from app.gosi_reader import _validate

    if not os.path.exists(path):
        raise FileNotFoundError(
            f"파라미터 모음집이 없습니다: {path}\n"
            "먼저 고시에서 추출하세요: python -m app.gosi_reader <고시파일.hwpx>"
        )
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    prices = data["단가"]
    _validate(prices)  # 수정 실수도 조용히 통과시키지 않는다
    return prices
