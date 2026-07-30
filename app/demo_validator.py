import pandas as pd
import time
import sys
import os

# ──────────────────────────── 설정 ────────────────────────────
GENERATED_FILE = "xlsx_files/생성된_단가표.xlsx"
BASELINE_FILE = "xlsx_files/2026_기본급여.xlsx" # 사보원 원본 데이터

# 터미널 텍스트 색상 (ANSI Escape Codes)
COLOR_RESET = "\033[0m"
COLOR_BLUE = "\033[1;34m"
COLOR_GREEN = "\033[1;32m"
COLOR_RED = "\033[1;31m"
COLOR_YELLOW = "\033[1;33m"
COLOR_CYAN = "\033[1;36m"

def run_validation_demo():
    print(f"\n{COLOR_BLUE}▶ [시스템 시작] 2026년 본인부담금 단가표 정합성 자동 검증을 시작합니다...{COLOR_RESET}\n")
    time.sleep(1)

    if not os.path.exists(GENERATED_FILE) or not os.path.exists(BASELINE_FILE):
        print(f"{COLOR_RED}[오류] 비교할 엑셀 파일이 존재하지 않습니다. 경로를 확인해주세요.{COLOR_RESET}")
        return

    print(" - 생성된 단가표 로드 중...")
    df_gen = pd.read_excel(GENERATED_FILE, header=0) 
    print(" - 기존 기준단가표 로드 중...")
    df_base = pd.read_excel(BASELINE_FILE, header=0)
    
    # 두 파일의 행 개수가 다를 경우의 예외 처리
    min_rows = min(len(df_gen), len(df_base))
    print(f" - 데이터 로드 완료! (총 {min_rows}건 비교)\n")
    time.sleep(1)

    error_count = 0
    mismatch_details = [] # 불일치 상세 내역을 저장할 리스트

    # 2. 데이터 순차 비교 및 진행 바 출력
    for i in range(min_rows):
        grade_name = df_gen.iloc[i].get('등급명', f'{i}번째 행')
        gen_limit = df_gen.iloc[i].get('지원량', 0)
        gen_copay = df_gen.iloc[i].get('본인부담금', 0)
        
        base_limit = df_base.iloc[i].get('지원량', 0)
        base_copay = df_base.iloc[i].get('본인부담금', 0)

        # 불일치 발생 시 상세 내역 기록
        is_error = False
        if gen_limit != base_limit or gen_copay != base_copay:
            error_count += 1
            is_error = True
            mismatch_details.append({
                'row': i + 2, # 엑셀 기준 실제 행 번호 (헤더 1행 + 인덱스 0부터 시작하므로 +2)
                'name': grade_name,
                'gen_limit': gen_limit,
                'base_limit': base_limit,
                'gen_copay': gen_copay,
                'base_copay': base_copay
            })
            status_text = f"{COLOR_RED}[FAIL]{COLOR_RESET}"
        else:
            status_text = f"{COLOR_GREEN}[OK]{COLOR_RESET}"

        # 터미널 진행 바 생성
        bar_length = 30
        filled_length = int(bar_length * (i + 1) // min_rows)
        bar = '█' * filled_length + '-' * (bar_length - filled_length)
        percent = int(100 * (i + 1) // min_rows)

        sys.stdout.write(f"\r{COLOR_YELLOW}검증 진행 중 [{bar}] {percent}% {COLOR_RESET}| {grade_name:<15} {status_text}")
        sys.stdout.flush()
        time.sleep(0.015) 

    # 3. 최종 결과 리포트 및 불일치 데이터 출력
    print("\n\n" + "="*65)
    if error_count == 0:
        print(f"{COLOR_GREEN}★ 시스템 정합성 100% 검증 완료! (총 {min_rows}건) ★{COLOR_RESET}")
        print(f"{COLOR_GREEN}★ 기존 기준단가표와 단 1원의 오차도 없이 완벽하게 일치합니다. ★{COLOR_RESET}")
    else:
        print(f"{COLOR_RED}※ 검증 실패: 총 {error_count}건의 불일치 데이터가 발견되었습니다.{COLOR_RESET}")
        print("-" * 65)
        print(f"{COLOR_YELLOW}[불일치 상세 리포트]{COLOR_RESET}")
        
        # 불일치 내역 순회하며 출력 (최대 10개까지만 출력하여 터미널 도배 방지)
        display_limit = 10
        for idx, mismatch in enumerate(mismatch_details[:display_limit]):
            print(f"\n{COLOR_CYAN}▶ 엑셀 {mismatch['row']}행 : {mismatch['name']}{COLOR_RESET}")
            
            if mismatch['gen_limit'] != mismatch['base_limit']:
                print(f"   - 지원량   | (생성) {mismatch['gen_limit']:,}원  <--->  (기준) {mismatch['base_limit']:,}원")
            if mismatch['gen_copay'] != mismatch['base_copay']:
                print(f"   - 본인부담 | (생성) {mismatch['gen_copay']:,}원  <--->  (기준) {mismatch['base_copay']:,}원")
        
        if error_count > display_limit:
            print(f"\n{COLOR_RED}...외 {error_count - display_limit}건의 오류가 더 존재합니다.{COLOR_RESET}")
            
    print("="*65 + "\n")

if __name__ == "__main__":
    run_validation_demo()