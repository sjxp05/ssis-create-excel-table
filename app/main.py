from app.table_reader import get_basic_df_headers
from app.table_writer import write_basic_table

JOGYEON_FILENAME = "xlsx_files/조견표_샘플.xlsx"
JOGYEON_SHEET_NAMES = ["인정조사", "산정특례", "종합조사"]

DANGA_FILENAME = "xlsx_files/생성된_단가표.xlsx"


"""
openpyxl로 열면 특히 한셀인 경우 외부 링크를 참조하고 있어서 안 열리는 문제 발생
pandas로 열어 DataFrame 형태로 읽기
"""

ij_basic_df_headers = get_basic_df_headers(JOGYEON_FILENAME, JOGYEON_SHEET_NAMES)

write_basic_table(DANGA_FILENAME, ij=ij_basic_df_headers)
