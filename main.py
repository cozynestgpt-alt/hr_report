"""
월례회의 인원보고서 자동화 스크립트
사용법: python main.py [YYYYMM]
예시:  python main.py 202605
       python main.py           ← 당월 자동 적용
"""

import sys
import os
import shutil
import subprocess
import json
from datetime import datetime, date
from calendar import monthrange
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


# ── 경로 설정 ─────────────────────────────────────────────
BASE_DIR    = Path(__file__).parent
INPUT_DIR   = BASE_DIR / "input"
OUTPUT_DIR  = BASE_DIR / "output"
TEMPLATE    = BASE_DIR / "template" / "월례회의_인원보고_템플릿.xlsx"
RECALC      = BASE_DIR / "scripts" / "recalc.py"

REPORT_SHEET = "전년대비 인원증감현황"   # 템플릿 고정 시트명
COMPARE_SHEET = "인원비교보고서"


# ── 인원 집계 함수 ────────────────────────────────────────
def load_employee_list(input_dir: Path) -> pd.DataFrame:
    """input 폴더에서 사원명단 파일 자동 탐색 후 로드"""
    files = sorted(input_dir.glob("사원명단*.xlsx"))
    if not files:
        raise FileNotFoundError(f"input 폴더에 '사원명단*.xlsx' 파일이 없습니다: {input_dir}")
    path = files[-1]  # 가장 최신 파일
    print(f"  [사원명단] {path.name}")
    df = pd.read_excel(path)
    df["입사일"] = pd.to_datetime(df["입사일"], errors="coerce")
    df["퇴사일"] = pd.to_datetime(df["퇴사일"], errors="coerce")
    return df


def is_active(row: pd.Series, ref_date: pd.Timestamp) -> bool:
    """기준일 말일 재직 여부"""
    if pd.isna(row["입사일"]) or row["입사일"] > ref_date:
        return False
    if not pd.isna(row["퇴사일"]) and row["퇴사일"] < ref_date:
        return False
    return True


def classify_type(row: pd.Series) -> str:
    """사원구분 → 일반 / 판매"""
    return "일반" if row["사원구분"] in ("연봉직", "호봉직") else "판매"


def build_headcount(df: pd.DataFrame, ref_date: pd.Timestamp) -> dict:
    """기준일 기준 부서별 인원수 dict 반환
    key: (구분1, 구분2), value: {"일반": n, "판매": n}

    집계 전략:
    - 구분2가 있는 경우 → (구분1, 구분2) 키로 집계
    - 구분2가 없는 경우 → (구분1, 구분1) 키로 집계  ← 임원/온라인사업부/AMD/CXD/ETC 등
    """
    active = df[df.apply(lambda r: is_active(r, ref_date), axis=1)].copy()
    active["타입"] = active.apply(classify_type, axis=1)

    counts: dict = {}
    for _, row in active.iterrows():
        g1 = str(row.get("구분1", "") or "").strip()
        g2_raw = row.get("구분2", "")
        g2 = str(g2_raw).strip() if (g2_raw and str(g2_raw) not in ("nan", "None", "")) else g1
        t  = row["타입"]
        key = (g1, g2)
        counts.setdefault(key, {"일반": 0, "판매": 0})
        counts[key][t] += 1
    return counts


# ── 인원비교보고서 시트 읽기 ──────────────────────────────
def read_compare_sheet(wb) -> dict:
    """인원비교보고서 시트 → {(구분1,구분2): {전년일반, 전년판매, 당해일반, 당해판매}}"""
    ws = wb[COMPARE_SHEET]
    data = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:
            key = (str(row[0]).strip(), str(row[1]).strip())
            data[key] = {
                "전년일반": row[2] or 0,
                "전년판매": row[3] or 0,
                "당해일반": row[4] or 0,
                "당해판매": row[5] or 0,
            }
    return data


# ── 전년대비 시트 행→데이터 매핑 ──────────────────────────
# (행번호, 구분1, 구분2, "일반" or "판매")
ROW_MAP = [
    # 브랜드사업부문 (일반)
    (5,  "임원",        "임원",        "일반"),
    (6,  "영업본부",    "총괄",        "일반"),
    (7,  "영업본부",    "백화점",      "일반"),
    (8,  "영업본부",    "SMD",         "일반"),
    (9,  "영업본부",    "영업지원",    "일반"),
    (10, "온라인사업부","온라인사업부","일반"),
    (11, "상품본부",    "총괄",        "일반"),
    (12, "상품본부",    "PDD",         "일반"),
    (13, "상품본부",    "MMD",         "일반"),
    (14, "RND",         "총괄",        "일반"),
    (15, "RND",         "RND",         "일반"),
    (16, "RND",         "VMD",         "일반"),
    (17, "AMD",         "AMD",         "일반"),
    (18, "CXD",         "CXD",         "일반"),
    (19, "ETC",         "ETC",         "일반"),
    # 직영점 (판매)
    (21, "직영점",      "양재점",                  "판매"),
    (22, "직영점",      "일산점",                  "판매"),
    (23, "직영점",      "경기광주점",              "판매"),
    (24, "직영점",      "전주점",                  "판매"),
    (25, "직영점",      "NC충장점",                "판매"),
    (26, "직영점",      "NC불광점",                "판매"),
    (27, "직영점",      "청주점",                  "판매"),
    (28, "직영점",      "NC해운대점",              "판매"),
    (29, "직영점",      "NC일산점",                "판매"),
    (30, "직영점",      "고양터미널점(롯데아울렛)","판매"),
    (31, "직영점",      "가든5점_현대아울렛",      "판매"),
    (32, "직영점",      "현대커넥트부산점",        "판매"),
]

# 인원비교보고서에 없는 키의 고정값 처리
# 고양터미널점: 이윤정 1명(일산점 소속 파견)으로 사원명단 직접 집계
SPECIAL_ROWS = {
    (30, "직영점", "고양터미널점(롯데아울렛)", "판매"): "by_dept",  # 이윤정: 일산점 AR1800
}


def get_special_count(g1, g2, year_key, df_all, ref_date):
    """인원비교보고서에 없는 행 직접 계산 (고양터미널 등)"""
    # 현재는 이윤정(일산점) 1명 고정 → 재직 여부만 확인
    if g2 == "고양터미널점(롯데아울렛)":
        yoon = df_all[(df_all["성명"] == "이윤정") & (df_all["사번"] == 465)]
        if yoon.empty:
            return 1  # fallback
        row = yoon.iloc[0]
        return 1 if is_active(row, ref_date) else 0
    return 0


# ── 변동사항 분석 ─────────────────────────────────────────
def analyze_changes(compare_data: dict, target_year: str, prev_year: str) -> list[str]:
    """증감 현황을 텍스트 리스트로 반환"""
    lines = []
    for (g1, g2), v in compare_data.items():
        if g1 == "인원수합계":
            continue
        curr_gen = v.get("당해일반", 0)
        prev_gen = v.get("전년일반", 0)
        curr_sal = v.get("당해판매", 0)
        prev_sal = v.get("전년판매", 0)
        diff_gen = curr_gen - prev_gen
        diff_sal = curr_sal - prev_sal
        if diff_gen != 0:
            sign = "▲" if diff_gen > 0 else "▼"
            lines.append(f"  {g1}/{g2} 일반 {sign}{abs(diff_gen)}명 ({prev_gen}→{curr_gen})")
        if diff_sal != 0:
            sign = "▲" if diff_sal > 0 else "▼"
            lines.append(f"  {g1}/{g2} 판매 {sign}{abs(diff_sal)}명 ({prev_sal}→{curr_sal})")
    return lines


# ── 메인 처리 ─────────────────────────────────────────────
def main():
    # 기준월 파싱
    if len(sys.argv) >= 2:
        ym = sys.argv[1].strip()
        if len(ym) != 6 or not ym.isdigit():
            print("❌ 기준월 형식 오류: YYYYMM 형태로 입력하세요 (예: 202605)")
            sys.exit(1)
        year, month = int(ym[:4]), int(ym[4:])
    else:
        today = date.today()
        year, month = today.year, today.month
        ym = f"{year}{month:02d}"
        print(f"  기준월 미입력 → 당월 적용: {ym}")

    last_day    = monthrange(year, month)[1]
    ref_date    = pd.Timestamp(year, month, last_day)          # 기준년월 말일
    prev_year   = year - 1
    prev_date   = pd.Timestamp(prev_year, month, monthrange(prev_year, month)[1])

    print(f"\n{'='*55}")
    print(f"  월례회의 인원보고서 자동화  |  기준: {year}년 {month:02d}월 말일")
    print(f"  비교: {prev_year}년 {month:02d}월 말일 대비")
    print(f"{'='*55}\n")

    # ── 1. 사원명단 로드 ──────────────────────────────────
    print("[1/5] 사원명단 로드 중...")
    df = load_employee_list(INPUT_DIR)
    print(f"  총 {len(df)}명 (전체 이력 포함)")

    # ── 2. 인원비교보고서 업데이트 ────────────────────────
    print("[2/5] 인원비교보고서 집계 중...")
    cnt_curr = build_headcount(df, ref_date)
    cnt_prev = build_headcount(df, prev_date)

    target_year_str = str(year)
    prev_year_str   = str(prev_year)

    # 인원비교보고서 헤더 갱신 (날짜 표시용)
    # C1~F1은 플레이스홀더 → 실제 연월 텍스트로 교체


    output_path = OUTPUT_DIR / f"월례회의_인원보고_{ym}.xlsx"
    shutil.copy(TEMPLATE, output_path)

    wb = load_workbook(output_path)

    # 인원비교보고서 헤더 연월 갱신
    ws_cmp = wb[COMPARE_SHEET]
    ws_cmp['C1'] = f'{prev_year}년 {month:02d}월_일반'
    ws_cmp['D1'] = f'{prev_year}년 {month:02d}월_판매'
    ws_cmp['E1'] = f'{year}년 {month:02d}월_일반'
    ws_cmp['F1'] = f'{year}년 {month:02d}월_판매'

    # 인원비교보고서 시트 갱신
    for row in ws_cmp.iter_rows(min_row=2):
        g1 = str(row[0].value or "").strip()
        g2 = str(row[1].value or "").strip()
        if not g1 or not g2 or g1 == "인원수합계":
            continue
        key = (g1, g2)
        p_gen = cnt_prev.get(key, {}).get("일반", 0)
        p_sal = cnt_prev.get(key, {}).get("판매", 0)
        c_gen = cnt_curr.get(key, {}).get("일반", 0)
        c_sal = cnt_curr.get(key, {}).get("판매", 0)
        row[2].value = p_gen   # C: 전년 일반
        row[3].value = p_sal   # D: 전년 판매
        row[4].value = c_gen   # E: 당해 일반
        row[5].value = c_sal   # F: 당해 판매
        row[6].value = c_gen - p_gen  # 일반 증감
        row[7].value = round((c_gen - p_gen) / p_gen, 6) if p_gen else None
        row[8].value = c_sal - p_sal  # 판매 증감
        row[9].value = round((c_sal - p_sal) / p_sal, 6) if p_sal else None

    # 합계행 갱신
    total_p_gen = sum(v.get("일반", 0) for v in cnt_prev.values())
    total_p_sal = sum(v.get("판매", 0) for v in cnt_prev.values())
    total_c_gen = sum(v.get("일반", 0) for v in cnt_curr.values())
    total_c_sal = sum(v.get("판매", 0) for v in cnt_curr.values())

    # 인원비교보고서 합계행도 갱신

    for row in ws_cmp.iter_rows(min_row=2):
        if str(row[0].value or "").strip() == "인원수합계":
            row[2].value = total_p_gen
            row[3].value = total_p_sal
            row[4].value = total_c_gen
            row[5].value = total_c_sal
            row[6].value = total_c_gen - total_p_gen
            row[8].value = total_c_sal - total_p_sal
            break

    # ── 3. 전년대비 시트 L2(기준일) 갱신 ────────────────
    print("[3/5] 전년대비 시트 채우는 중...")
    ws_rep = wb[REPORT_SHEET]
    import datetime as dt
    ws_rep["L2"] = dt.date(year, month, last_day)

    compare_data = read_compare_sheet(wb)  # 방금 갱신된 시트 재읽기

    for row_num, g1, g2, type_ in ROW_MAP:
        key = (g1, g2)
        if key in compare_data:
            p_val = compare_data[key]["전년일반" if type_ == "일반" else "전년판매"]
            c_val = compare_data[key]["당해일반" if type_ == "일반" else "당해판매"]
        else:
            # 인원비교보고서에 없는 행(고양터미널 등) 직접 계산
            p_val = get_special_count(g1, g2, prev_year_str, df, prev_date)
            c_val = get_special_count(g1, g2, target_year_str, df, ref_date)

        if type_ == "일반":
            ws_rep.cell(row_num, 5).value = p_val   # E: 전년 일반
            ws_rep.cell(row_num, 7).value = c_val   # G: 당해 일반
        else:
            ws_rep.cell(row_num, 6).value = p_val   # F: 전년 판매
            ws_rep.cell(row_num, 8).value = c_val   # H: 당해 판매

    # 백화점 행 판매칸 0 고정 (기존 #REF! 수식 제거)
    ws_rep["F7"] = 0
    ws_rep["H7"] = 0

    # 증감율 수식에서 분모=0일 때 DIV/0! 방지 (J12, J13 등 IF 수식 → IFERROR로 교체)
    for r in range(5, 40):
        cell = ws_rep.cell(r, 10)  # J열
        if cell.value and str(cell.value).startswith("=IF("):
            # =IF(Ix/Ex=0,"",Ix/Ex) → =IFERROR(IF(Ix/Ex=0,"",Ix/Ex),"")
            cell.value = f'=IFERROR({cell.value[1:]},"")' 

    # ── 4. 저장 및 수식 재계산 ────────────────────────────
    print("[4/5] 파일 저장 및 수식 재계산 중...")
    wb.save(output_path)

    result = subprocess.run(
        [sys.executable, str(RECALC), str(output_path), "60"],
        capture_output=True, text=True
    )
    recalc_info = json.loads(result.stdout) if result.stdout else {}
    if recalc_info.get("status") == "errors_found":
        print(f"  ⚠️  수식 오류 {recalc_info['total_errors']}건: {recalc_info.get('error_summary', {})}")
    else:
        print(f"  ✅ 수식 재계산 완료 (수식 {recalc_info.get('total_formulas', '?')}개)")

    # ── 5. 변동사항 요약 텍스트 생성 ─────────────────────
    print("[5/5] 변동사항 요약 작성 중...")
    changes = analyze_changes(compare_data, target_year_str, prev_year_str)

    summary_lines = [
        f"전년동월대비 인원증감 현황 요약",
        f"기준: {year}년 {month:02d}월 말일 / 비교: {prev_year}년 {month:02d}월 말일",
        f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "=" * 50,
        f"【 전체 현황 】",
        f"  구분    | {prev_year}.{month:02d} | {year}.{month:02d} | 증감",
        f"  --------|-------|-------|-----",
        f"  일반직  | {total_p_gen:5}명 | {total_c_gen:5}명 | {total_c_gen-total_p_gen:+}명",
        f"  판매직  | {total_p_sal:5}명 | {total_c_sal:5}명 | {total_c_sal-total_p_sal:+}명",
        f"  합  계  | {total_p_gen+total_p_sal:5}명 | {total_c_gen+total_c_sal:5}명 | {(total_c_gen+total_c_sal)-(total_p_gen+total_p_sal):+}명",
        "=" * 50,
    ]

    if changes:
        summary_lines += ["【 부서별 증감 내역 】"] + changes
    else:
        summary_lines += ["【 부서별 증감 내역 】", "  전년 동월 대비 변동 없음"]

    summary_lines += [
        "=" * 50,
        f"출력 파일: {output_path.name}",
    ]

    summary_text = "\n".join(summary_lines)
    summary_path = OUTPUT_DIR / f"인원변동요약_{ym}.txt"
    summary_path.write_text(summary_text, encoding="utf-8-sig")

    # ── 완료 출력 ──────────────────────────────────────────
    print(f"\n{'='*55}")
    print(f"  ✅ 완료!")
    print(f"  📊 엑셀  : output/{output_path.name}")
    print(f"  📄 요약  : output/{summary_path.name}")
    print(f"{'='*55}\n")
    print(summary_text)


if __name__ == "__main__":
    main()
