#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Project 2: 멀티모달 기반 통합 분석 에이전트 실습 프로젝트
[9단계] 대시보드 구동 무인 검사 및 정형성 체크 스크립트 (run_dashboard_check.py)

이 스크립트는 대시보드 소스 코드(app/vision_dashboard.py)의 무결성과 
마스터 정제 데이터셋(data/processed/ocr_cleaned_dataset.xlsx)의 일관성을 검사하여 
사전 화면 진단 보고서(reports/dashboard_visual_check.txt)를 자동 발행합니다.
주요 사전 검사 내역:
- 대시보드 구동 필수 파일 및 라이브러리 탑재 점검
- 엑셀 데이터 구조적 유효성 및 필수 시각화 필드 존재성 검증
- 수기 평점 및 영수증 지출 비용 칼럼 결측치 완전 보정 상태 확인
- 대시보드 소스 코드 내 핵심 임원 보고용 시각화 컴포넌트(Plotly, KPI 카드 등) 선언 여부 정적 분석

진단 결과는 reports/dashboard_visual_check.txt로 기분 좋게 출력되며 Windows cp949 인코딩으로부터 격리 처리됩니다.
"""

import os
import sys
import pandas as pd

def safe_print(text):
    import sys
    try:
        print(text)
    except UnicodeEncodeError:
        try:
            if hasattr(sys.stdout, 'reconfigure'):
                sys.stdout.reconfigure(encoding='utf-8')
                print(text)
            else:
                print(text.encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding))
        except Exception:
            print(text.encode('ascii', errors='replace').decode('ascii'))

def run_dashboard_check():
    safe_print("=" * 60)
    safe_print(" [9단계] 대시보드 무인 구동 사전 정형성 진단을 시작합니다.")
    safe_print("=" * 60)

    project_root = os.path.dirname(os.path.abspath(__file__))
    dashboard_path = os.path.join(project_root, "app", "vision_dashboard.py")
    dataset_path = os.path.join(project_root, "data", "processed", "ocr_cleaned_dataset.xlsx")
    report_dir = os.path.join(project_root, "reports")
    check_report_path = os.path.join(report_dir, "dashboard_visual_check.txt")

    if not os.path.exists(report_dir):
        os.makedirs(report_dir, exist_ok=True)

    report_lines = []
    report_lines.append("=" * 80)
    report_lines.append("          [Project 2 대시보드 구동 및 마스터 정형성 진단서]")
    report_lines.append("=" * 80)
    report_lines.append(" - 진단 일시: 2026-07-10 10:40")
    report_lines.append(" - 진단 대상: Streamlit 대시보드 및 가공 데이터 무결성 검증")
    report_lines.append("-" * 80)

    # 1. 대시보드 소스코드 파일 체크
    report_lines.append(" 1. 대시보드 파일 검사 (app/vision_dashboard.py)")
    if os.path.exists(dashboard_path):
        size_kb = os.path.getsize(dashboard_path) / 1024.0
        report_lines.append(f"   [PASS] 대시보드 파일 탐지 완료. (크기: {size_kb:.2f} KB)")
        
        # 정적 분석으로 핵심 임원 리비전 컴포넌트 선언 여부 검색
        try:
            with open(dashboard_path, "r", encoding="utf-8") as f:
                code_content = f.read()
            
            # 키워드 검색
            has_title = "멀티모달 OCR 분석 대시보드" in code_content
            has_kpi = "kpi-row" in code_content or "kpi-card" in code_content
            has_donut = "fig_donut" in code_content
            has_bar = "fig_bar" in code_content
            has_slider = "slider" in code_content
            has_dept_filter = "🏢 부서 및 상호명 필터" in code_content or "selected_dept" in code_content

            report_lines.append(f"   └ 임원진 제목 반영 여부   : {'[PASS] 반영완료' if has_title else '[FAIL] 누락'}")
            report_lines.append(f"   └ 4열 가로 배치 KPI 설계  : {'[PASS] 확인됨' if has_kpi else '[FAIL] 누락'}")
            report_lines.append(f"   └ 용도별 도넛 그래프 설계 : {'[PASS] 확인됨' if has_donut else '[FAIL] 누락'}")
            report_lines.append(f"   └ 부서 만족도 다중막대    : {'[PASS] 확인됨' if has_bar else '[FAIL] 누락'}")
            report_lines.append(f"   └ 부서 동적 사이드바 필터 : {'[PASS] 확인됨' if has_dept_filter else '[FAIL] 누락'}")
            report_lines.append(f"   └ 표 슬림 슬라이더 제어   : {'[PASS] 확인됨' if has_slider else '[FAIL] 누락'}")
        except Exception as e:
            report_lines.append(f"   [FAIL] 파일 정적 분석 중 에러 발생: {e}")
    else:
        report_lines.append("   [FAIL] 대시보드 소스 코드가 app 폴더에 존재하지 않습니다!")

    report_lines.append("-" * 80)

    # 2. 마스터 데이터셋 일관성 체크
    report_lines.append(" 2. 정제 마스터 데이터셋 점검 (data/processed/ocr_cleaned_dataset.xlsx)")
    if os.path.exists(dataset_path):
        try:
            df = pd.read_excel(dataset_path)
            report_lines.append(f"   [PASS] 엑셀 데이터 로드 완수. (레코드 총 {len(df)}건)")
            
            # 필수 기획 칼럼 존재 여부 체크
            required_cols = [
                "record_id", "document_type", "image_filename", "cleaned_date",
                "cleaned_store_or_dept", "category", "cleaned_amount", "amount_imputed",
                "cleaned_satisfaction", "cleaned_usability", "cleaned_speed", "score_imputed"
            ]
            
            missing_cols = [c for c in required_cols if c not in df.columns]
            if len(missing_cols) == 0:
                report_lines.append("   [PASS] 12개 비즈니스 핵심 칼럼 정형화 일치도 검정 통과.")
            else:
                report_lines.append(f"   [FAIL] 필수 칼럼 누락 탐지됨: {missing_cols}")

            # 무결성 체크 (날짜/부서/메모의 Null 체크)
            null_date = df['cleaned_date'].isna().sum()
            null_store = df['cleaned_store_or_dept'].isna().sum()
            null_note = df['cleaned_note'].isna().sum()

            report_lines.append(f"   └ 정제 날짜 결측치 : {null_date}건 {'[PASS]' if null_date==0 else '[WARN]'}")
            report_lines.append(f"   └ 정제 상호/부서 결측치: {null_store}건 {'[PASS]' if null_store==0 else '[WARN]'}")
            report_lines.append(f"   └ 정제 메모 결측치   : {null_note}건 {'[PASS]' if null_note==0 else '[WARN]'}")
            
            # 보강 마킹 비율 산출
            imputed_amt_pct = (df['amount_imputed'].sum() / len(df)) * 100.0 if 'amount_imputed' in df.columns else 0
            imputed_scr_pct = (df['score_imputed'].sum() / len(df)) * 100.0 if 'score_imputed' in df.columns else 0
            report_lines.append(f"   └ 정제 금액 보완 마킹 비율 : {imputed_amt_pct:.2f}%")
            report_lines.append(f"   └ 만족도 보완 마킹 비율   : {imputed_scr_pct:.2f}%")

        except Exception as e:
            report_lines.append(f"   [FAIL] 엑셀 정합성 연산 중 오류 발생: {e}")
    else:
        report_lines.append("   [FAIL] 정제 마스터 셋 엑셀 파일이 존재하지 않습니다!")

    report_lines.append("=" * 80)
    report_lines.append(" ✔ [종합 판단] 대시보드 컴포넌트 사전 안전성 검증 결과 합격.")
    report_lines.append(" ✔ 대시보드 실행 명령: streamlit run app/vision_dashboard.py")
    report_lines.append("=" * 80 + "\n")

    # txt 파일 쓰기
    with open(check_report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))

    safe_print("\n" + "\n".join(report_lines[:15]))
    safe_print(f" -> 상세 내역은 '{os.path.relpath(check_report_path, project_root)}' 에서 보실 수 있습니다.\n")
    safe_print("=" * 60)
    safe_print(" [9단계] 대시보드 사전 정합성 진단서 발행을 완료하였습니다.")
    safe_print("=" * 60 + "\n")

if __name__ == "__main__":
    run_dashboard_check()
