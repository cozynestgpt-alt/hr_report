# 월례회의 인원보고서 자동화 (부서별인원현황)

## 프로젝트 구조
```
부서별인원현황/
├── main.py                          ← 실행 스크립트 (python main.py [YYYYMM])
├── 월례회의_인원보고_실행.bat        ← 팀원용 더블클릭 실행 파일 (NAS 배포)
├── CLAUDE.md
├── README.md
├── NAS_팀원_실행_매뉴얼.md          ← 팀원용 실행 가이드
├── 작업이력.md                      ← 코드·문서·구조 변경 이력
├── input/                           ← 매월 갱신되는 부서별인원현황_YYYYMMDD.xlsx,
│                                       ★월례회의_인원보고_양식.xlsx (.gitignore)
├── template/                        ← 고정 템플릿 (수정 금지)
├── output/                          ← 결과물 자동 생성 (.gitignore)
├── DB/                              ← 참조용 DB 파일 (.gitignore)
├── 참고/                            ← 집계 방법 문서·과거 산출물 (.gitignore)
└── scripts/
    └── recalc.py                    ← 수식 재계산 유틸 (수정 금지)
```

## 실행 방법
```bash
python main.py 202605   # 기준월 지정
python main.py          # 당월 자동 적용
```

## 경로 정보
- GitHub: https://github.com/cozynestgpt-alt/hr_report (저장소명은 hr_report로 유지)
- 담당자 작업 PC: `C:\Users\azmang\Documents\결산\★월별손익\부서별인원현황`
- NAS(팀 공유): `\\nas\CN_AMD\월례회의_자료\부서별인원현황`

세 환경 모두 같은 git 저장소를 origin으로 공유한다. 로컬 PC와 NAS 각각에서 `git clone`/`git pull`로 연결되어 있다.

## 매월 작업 흐름
1. 담당자 PC(`input/`)에 기준월 말일자 `부서별인원현황_YYYYMMDD.xlsx` 추가
2. 담당자가 `input/★월례회의_인원보고_양식.xlsx`의 중간관리(롯데/현대/신세계/갤러리아)·물류센터용역·평택점용역 값과 L2(기준일)를 당월 기준으로 직접 갱신
   - L2가 당월 말일과 일치해야 자동 반영됨. 갱신을 안 하면 스크립트가 감지해 수동 입력으로 안내함
3. 코드나 템플릿 수정이 필요하면 로컬 PC에서 수정 후 GitHub에 `git push`
4. NAS 폴더에서 `git pull`로 최신 코드 반영
5. 팀원은 NAS의 `월례회의_인원보고_실행.bat`을 더블클릭 → 기준월 입력 → `output/` 폴더에서 결과 확인

## 개인정보 처리 원칙
`input/`, `output/`, `DB/`, `참고/` 폴더의 실데이터와 템플릿을 제외한 모든 `*.xlsx`는 `.gitignore`로 GitHub에 올라가지 않는다. 폴더 구조 자체는 `.gitkeep`으로 유지되므로 새로 clone/pull해도 폴더는 비어 있는 채로 생성된다 — 실데이터는 NAS/로컬에 각자 직접 채워 넣어야 한다.

## 문서
- `README.md`: 사용법 전반, 폴더 구조
- `NAS_팀원_실행_매뉴얼.md`: 팀원 입장의 실행 가이드, 오류 해결
- `작업이력.md`: 코드·문서·구조에 실질적인 변경이 있을 때마다 기록하는 이력. 스크립트나 문서를 고칠 때는 이 파일에도 항목을 추가할 것
