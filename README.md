# 월례회의 인원보고서 자동화

GitHub: https://github.com/cozynestgpt-alt/hr_report

## 폴더 구조

```
hr_report/
├── main.py                  ← 실행 스크립트
├── README.md
├── input/                   ← 매월 갱신 파일 넣는 폴더 (.gitkeep)
│   ├── 사원명단YYYYMMDDYYYYMMDD.xlsx
│   └── 사원현황YYYYMMDDYYYYMMDD.xlsx
├── template/                ← 고정 템플릿 (수정 금지)
│   └── 월례회의_인원보고_템플릿.xlsx
├── output/                  ← 결과물 자동 생성 (.gitkeep)
│   ├── 월례회의_인원보고_YYYYMM.xlsx
│   └── 인원변동요약_YYYYMM.txt
├── DB/                      ← 참조용 DB 파일 (.gitkeep)
├── 참고/                    ← 집계 방법 문서·과거 산출물 (.gitkeep)
└── scripts/                 ← 내부 유틸 (수정 금지)
    └── recalc.py
```

> `input/`, `output/`, `DB/`, `참고/` 폴더는 개인정보가 포함된 실데이터라 `.gitignore`로 제외됩니다.
> 폴더 자체는 `.gitkeep` 파일로 저장소에 유지되므로, 새로 clone하면 빈 폴더 구조는 그대로 생기고 그 안에 매월 파일만 채워 넣으면 됩니다.

## 사용법

### 기본 실행 (당월 자동 적용)
```bash
python main.py
```

### 특정 월 지정
```bash
python main.py 202605
python main.py 202604
```

## 매월 작업 순서

1. **input 폴더**에 최신 사원명단·사원현황 xlsx 파일 복사
2. 터미널에서 실행:
   ```bash
   python main.py 202605
   ```
3. **output 폴더**에서 결과물 확인
   - `월례회의_인원보고_202605.xlsx` → 보고용 엑셀
   - `인원변동요약_202605.txt` → 변동사항 요약

## 주의사항

- `template/` 폴더의 템플릿 파일은 수정하지 마세요 (시트 구조가 바뀌면 스크립트 수정 필요)
- input 파일명은 `사원명단`으로 시작하면 어떤 형식이든 자동 인식됩니다
- 같은 월 재실행 시 output 파일을 덮어씁니다

## 요구 라이브러리
```bash
pip install pandas openpyxl
```
