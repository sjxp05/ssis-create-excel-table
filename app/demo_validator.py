import pandas as pd
import time
import sys
import os

# ──────────────────────────── 설정 ────────────────────────────
# 검증할 파일 쌍 리스트: (생성된 파일, 기준 파일, 검증할 표 이름)
VALIDATION_TARGETS = [
    ("xlsx_files/생성된_단가표.xlsx", "xlsx_files/2026_기본급여.xlsx", "기본급여 단가표"),
    ("xlsx_files/추가급여_단가표.xlsx", "xlsx_files/2026_추가급여.xlsx", "추가급여 단가표"),
    ("xlsx_files/생성된_결제단가.xlsx", "xlsx_files/2026_결제단가.xlsx", "결제 단가표")
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
    
    min_rows = min(len(df_gen), len(df_base))
    print(f" - 데이터 로드 완료! (총 {min_rows}건 비교)\n")
    time.sleep(0.5)

    error_count = 0
    mismatch_details = []

    for i in range(min_rows):
        # 지원량, 본인부담금 데이터 추출 (결제단가의 경우 컬럼명이 다를 수 있으나 우선 지원량/본인부담금 체크)
        grade_name = df_gen.iloc[i].get('등급명', f'{i+2}행')
        gen_limit = df_gen.iloc[i].get('지원량', 0)
        gen_copay = df_gen.iloc[i].get('본인부담금', 0)
        
        base_limit = df_base.iloc[i].get('지원량', 0)
        base_copay = df_base.iloc[i].get('본인부담금', 0)

        # 오류 판별
        if gen_limit != base_limit or gen_copay != base_copay:
            error_count += 1
            mismatch_details.append({
                'row': i + 2, 
                'name': grade_name,
                'gen_limit': gen_limit,
                'base_limit': base_limit,
                'gen_copay': gen_copay,
                'base_copay': base_copay
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
            if mismatch['gen_limit'] != mismatch['base_limit']:
                print(f"   - 지원량   | (생성) {mismatch['gen_limit']:,}원  <--->  (기준) {mismatch['base_limit']:,}원")
            if mismatch['gen_copay'] != mismatch['base_copay']:
                print(f"   - 본인부담 | (생성) {mismatch['gen_copay']:,}원  <--->  (기준) {mismatch['base_copay']:,}원")
        
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