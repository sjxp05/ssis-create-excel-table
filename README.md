### 실행 방법

1. venv 환경 세팅

    ```pwsh
    python -m venv .venv
    ```

2. venv 가상환경 활성화\
   (\* 3, 4는 활성화 시킨 상태에서 실행해야 됨!)

    ```bash
     .venv\Scripts\Activate # Windows
     source .venv/bin/activate # Mac
    ```

3. 필요한 라이브러리 다운로드

    ```pwsh
    pip install -r requirements.txt
    ```

4. (결제단가표까지 만들 경우) 고시에서 단가 추출\
   고시 hwpx 파일에서 서비스별 단가를 추출해 `params/단가_파라미터.json`으로 저장한다.\
   (이 단계를 건너뛰면 main.py는 기본 단가표까지만 생성)

    ```pwsh
    python -m app.gosi_reader "<고시파일.hwpx>"
    ```

5. main.py 실행
    ```pwsh
    python -m app.main
    ```

### 파이프라인 구조

```
조견표 엑셀 ──(table_reader/writer)──> 기본급여 단가표 (950행)
                                              x
고시 hwpx ──(gosi_reader)──> params/단가_파라미터.json ──(payment_master)──> 서비스 단가 25조합
                                              |
                                              v
                              결제단가표 (23,725행 + 단가근거 시트)
```

### 담당자 수정 지점 (코드 수정 없이 파일만)

| 상황 | 수정할 파일 |
| --- | --- |
| 단가 금액이 틀렸을 때 | `params/단가_파라미터.json` 의 해당 `금액` (근거는 옆의 `출처` 참고) |
| 고시에서 서비스 명칭이 바뀌었을 때 | `params/앵커_키워드.json` 의 해당 `제목` |

값이 이상하면 실행 시 검증에서 멈추고, 문제의 값과 고시상 위치(장 > 항목 > 표 > 행)를 안내한다.

### 검증

생성한 결제단가를 공식 파일과 전수 대조하려면:

```pwsh
python -m app.payment_writer <기본단가표.xlsx> <출력.xlsx> <연도> <차수> [검증용 공식파일.xlsx]
```

### 알려진 이슈

- `xlsx_files/조견표_샘플.xlsx` 는 월 한도액은 2026년 값이지만 종합조사 본인부담금
  그리드가 이전 연도 값이라, 종합조사 38개 등급의 본인부담금이 공식 2026 단가표와 다르다.
  (실제 2026 조견표로 교체 필요)
- 실제 2026 조견표(내부등록용)는 시트 레이아웃이 조견표_샘플과 달라
  현재 `read_jh_table` 이 읽지 못한다. (조견표 파싱 로직 보완 필요)
