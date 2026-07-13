from app.table_reader import get_basic_df_headers, read_ij_sj_table_extra
from app.table_writer import write_basic_table, write_add_table

JOGYEON_FILENAME = "xlsx_files/조견표_샘플.xlsx"
ALL_JOGYEON_FILENAME = "xlsx_files/download.xlsx"
JOGYEON_SHEET_NAMES = ["인정조사", "산정특례", "종합조사"]

DANGA_FILENAME = "xlsx_files/생성된_단가표.xlsx"
TK_DANGA_FILENAME = "xlsx_files/특례_단가표.xlsx"
ADD_DANGA_FILENAME="xlsx_files/추가급여_단가표.xlsx"
"""
openpyxl로 열면 특히 한셀인 경우 외부 링크를 참조하고 있어서 안 열리는 문제 발생
pandas로 열어 DataFrame 형태로 읽기
"""

ij_basic_df_headers = get_basic_df_headers(JOGYEON_FILENAME, JOGYEON_SHEET_NAMES)

write_basic_table(DANGA_FILENAME, ij=ij_basic_df_headers)

#특례
# table = read_sj_table(ALL_JOGYEON_FILENAME, "산정특례")
# tk_danga = write_tk_talbe(table)
# tk_danga.to_excel(TK_DANGA_FILENAME, index=False)

#추가급여단가표
GAGU_AMOUNTS,ADD_PAYPERCENT=read_ij_sj_table_extra(ALL_JOGYEON_FILENAME,"산정특례","인정조사")
write_add_table(ADD_DANGA_FILENAME,GAGU_AMOUNTS,ADD_PAYPERCENT)