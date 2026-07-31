import os

from app.jogyeon_reader import get_basic_df_headers  # 실제 조견표 대응 리더 (jh만 보완)
from app.param_store import DEFAULT_PATH
from app.payment_writer import write_payment_table
from app.table_writer import write_basic_table
from app.table_reader import get_basic_df_headers, read_ij_sj_table_extra
from app.table_writer import write_basic_table, write_add_table
from app.config import YEAR, CHASU, DANGA_FILENAME, ADD_DANGA_FILENAME, PAYMENT_FILENAME 
#년도, 차수, 생성되는 단가표 파일들 이름 정의는 config.py 로 따로 빼 둠

JOGYEON_FILENAME = f"xlsx_files/{YEAR}_조견표.xlsx"  # 실제(내부등록용) 조견표
JOGYEON_SHEET_NAMES = ["인정조사", "산정특례", "종합조사"]

# openpyxl로 열면 특히 한셀인 경우 외부 링크를 참조하고 있어서 안 열리는 문제 발생 pandas로 열어 DataFrame 형태로 읽기


# 1) 기본급여단가표
ij, jh, sj = get_basic_df_headers(JOGYEON_FILENAME, JOGYEON_SHEET_NAMES)

write_basic_table(DANGA_FILENAME, YEAR, CHASU, ij=ij, jh=jh, sj=sj)

# 2) 기본급여 단가표 x 고시 단가 -> 결제단가표
#    (고시에서 추출한 단가 파라미터가 있어야 생성 가능)
if os.path.exists(DEFAULT_PATH):
    write_payment_table(DANGA_FILENAME, PAYMENT_FILENAME, YEAR, CHASU)
else:
    print(f"[안내] {DEFAULT_PATH} 가 없어 결제단가표 생성을 건너뜁니다.")
    print(
        "       먼저 고시에서 단가를 추출하세요: python -m app.gosi_reader <고시파일.hwpx>"
    )

# 3) 추가급여단가표
GAGU_AMOUNTS, ADD_PAYPERCENT = read_ij_sj_table_extra(
    JOGYEON_FILENAME, "산정특례", "인정조사"
)
write_add_table(ADD_DANGA_FILENAME, GAGU_AMOUNTS, ADD_PAYPERCENT)
