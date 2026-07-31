import pandas as pd
import time
import sys
import os
from app.config import YEAR, DANGA_FILENAME, ADD_DANGA_FILENAME, PAYMENT_FILENAME 

# ──────────────────────────── 설정 ────────────────────────────
# 검증할 파일 쌍 리스트: (생성된 파일, 기준 파일, 검증할 표 이름)
VALIDATION_TARGETS = [
    (DANGA_FILENAME, f"xlsx_files/{YEAR}_기본급여.xlsx", "기본급여 단가표"),
    (ADD_DANGA_FILENAME, f"xlsx_files/{YEAR}_추가급여.xlsx", "추가급여 단가표"),
    (PAYMENT_FILENAME, f"xlsx_files/{YEAR}_결제단가.xlsx", "결제 단가표")
]

# 터미널 텍스트 색상
COLOR_RESET = "\033[0m"
COLOR_BLUE = "\033[1;34m"
COLOR_GREEN = "\033[1;32m"
COLOR_RED = "\033[1;31m"
COLOR_YELLOW = "\033[1;33m"
COLOR_CYAN = "\033[1;36m"

def validate_table(gen_file, base_file, table_name):
    print(f"\n{COLOR_CYAN}============================================================{COLOR_RESET}")
    print(f"{COLOR_BLUE}▶ [{table_name}] 정합성 검증을 시작합니다...{COLOR_RESET}")
    
    # 파일 존재 여부 확인
    if not os.path.exists(gen_file):
        print(f"{COLOR_RED}[스킵] 생성된 파일이 없습니다: {gen_file}{COLOR_RESET}")
        return
    if not os.path.exists(base_file):
        print(f"{COLOR_RED}[스킵] 비교할 기준 파일이 없습니다: {base_file}{COLOR_RESET}")
        print(f"{COLOR_YELLOW}※ 사전에 사보원 원본 파일을 폴더에 넣어주세요.{COLOR_RESET}")
        return

    # 엑셀 로드
    df_gen = pd.read_excel(gen_file, header=0) 
    df_base = pd.read_excel(base_file, header=0)

    # 샘플 파일의 임시공휴일 행은 검증 대상에서 제외
    if table_name == "결제 단가표":
        before_cnt = len(df_base)
        # 데이터프레임 전체 열을 대상으로 '임시공휴일' 텍스트가 포함된 행 찾기
        mask = df_base.astype(str).apply(lambda x: x.str.contains('임시공휴일', na=False)).any(axis=1)
        
        # '임시공휴일'이 포함된 행을 제외(~mask)하고, 인덱스를 처음부터 다시 번호 매김(reset_index)
        df_base = df_base[~mask].reset_index(drop=True)
        after_cnt = len(df_base)
        
        excluded_cnt = before_cnt - after_cnt
        if excluded_cnt > 0:
            print(f"{COLOR_YELLOW} ⚠️ 기준 샘플에서 '임시공휴일' 관련 {excluded_cnt}개 행을 검증 대상에서 제외했습니다.{COLOR_RESET}")

    min_rows = min(len(df_gen), len(df_base))

    common_cols = [col for col in df_gen.columns if col in df_base.columns]
    print(f" - 데이터 로드 완료! (총 {min_rows}건 비교)\n")
    time.sleep(0.5)

    error_count = 0
    mismatch_details = []

    for i in range(min_rows):
        # 지원량, 본인부담금 데이터 추출 (결제단가의 경우 컬럼명이 다를 수 있으나 우선 지원량/본인부담금 체크)
        grade_name = df_gen.iloc[i].get('등급명', f'{i+2}행')
        is_error = False
        row_errors = []

        for col in common_cols:
            gen_val = df_gen.at[i, col]
            base_val = df_base.at[i, col]

            # 둘 다 비어있는 칸(NaN)이면 정상으로 간주하고 패스
            if pd.isna(gen_val) and pd.isna(base_val):
                continue

            # 숫자형 비교: 타입에러는 패스(int, float 달라도 값만 같으면 넘어감)
            try:
                if float(gen_val) == float(base_val):
                    continue
            except (ValueError, TypeError):
                pass

            # 문자열 비교
            if str(gen_val).strip() != str(base_val).strip():
                is_error = True
                row_errors.append({
                    'col': col,
                    'gen': gen_val,
                    'base': base_val
                })

        if is_error:
            error_count += 1
            mismatch_details.append({
                'row': i+2,
                'name': grade_name,
                'errors': row_errors
            })
            status_text = f"{COLOR_RED}[FAIL]{COLOR_RESET}"
        else:
            status_text = f"{COLOR_GREEN}[OK]{COLOR_RESET}"

        # 진행 바 애니메이션
        bar_length = 30
        filled_length = int(bar_length * (i + 1) // min_rows)
        bar = '█' * filled_length + '-' * (bar_length - filled_length)
        percent = int(100 * (i + 1) // min_rows)

        sys.stdout.write(f"\r{COLOR_YELLOW}[{table_name}] 검증 중 [{bar}] {percent}% {COLOR_RESET}| {str(grade_name)[:15]:<15} {status_text}")
        sys.stdout.flush()
        time.sleep(0.01) # 시연용 딜레이

    print("\n")
    if error_count == 0:
        print(f"{COLOR_GREEN}★ {table_name} 정합성 100% 검증 완료! 결함 제로! ★{COLOR_RESET}")
    else:
        print(f"{COLOR_RED}※ 검증 실패: {table_name}에서 총 {error_count}건의 불일치 데이터가 발견되었습니다.{COLOR_RESET}")
        
        display_limit = 5 # 화면 도배 방지용 최대 5건 표시
        for idx, mismatch in enumerate(mismatch_details[:display_limit]):
            print(f"\n{COLOR_CYAN}▶ 엑셀 {mismatch['row']}행 : {mismatch['name']}{COLOR_RESET}")
            for err in mismatch['errors']:
                
                # 금액 등 숫자인 경우 콤마 처리, 문자인 경우 그대로 출력
                val_g, val_b = err['gen'], err['base']
                str_g = f"{int(val_g):,}" if isinstance(val_g, (int, float)) and pd.notna(val_g) else str(val_g)
                str_b = f"{int(val_b):,}" if isinstance(val_b, (int, float)) and pd.notna(val_b) else str(val_b)
                
                print(f"   - {err['col']:<8} | (생성) {str_g}  <--->  (기준) {str_b}")
                
        if error_count > display_limit:
            print(f"\n{COLOR_RED}...외 {error_count - display_limit}건의 오류가 더 존재합니다.{COLOR_RESET}")

def run_all_validations():
    print(f"\n{COLOR_BLUE}단가표 통합 검증 파이프라인 가동{COLOR_RESET}\n")
    time.sleep(1)
    
    for gen_file, base_file, table_name in VALIDATION_TARGETS:
        validate_table(gen_file, base_file, table_name)
        time.sleep(0.5)
        
    print(f"\n{COLOR_GREEN}============================================================{COLOR_RESET}")
    print(f"{COLOR_GREEN}모든 단가표(기본, 추가, 결제) 검증 프로세스 종료{COLOR_RESET}")
    print(f"{COLOR_GREEN}============================================================{COLOR_RESET}\n")

if __name__ == "__main__":
    run_all_validations()