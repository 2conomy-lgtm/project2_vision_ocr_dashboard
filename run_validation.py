#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Project 2: 멀티모달 기반 통합 분석 에이전트 실습 프로젝트
[1단계] 원본 데이터와 이미지 매칭 검증 스크립트 (run_validation.py)

이 스크립트는 정답 정형 데이터(CSV)에 기록된 이미지 파일 경로가 실제 로컬 폴더에 존재하는지 검사하고,
동시에 중복된 record_id나 잘못된 확장자가 있는지 확인하여 품질 리포트를 작성합니다.
"""

import os
import sys
import pandas as pd

def run_validation():
    print("=" * 60)
    print(" [1단계] 원본 데이터와 이미지 매칭 검증을 시작합니다.")
    print("=" * 60)

    # 1. 파일 경로 설정
    project_root = os.path.dirname(os.path.abspath(__file__))
    csv_path = os.path.join(project_root, "data", "source_structured", "ground_truth_multimodal_240.csv")
    report_dir = os.path.join(project_root, "reports")
    report_path = os.path.join(report_dir, "data_image_validation_report.txt")

    # reports 폴더 자동 생성
    if not os.path.exists(report_dir):
        os.makedirs(report_dir, exist_ok=True)
        print(f"-> [생성 완료] 보고서 폴더: {report_dir}")

    # CSV 파일 존재 여부 확인
    if not os.path.exists(csv_path):
        print(f"[오류] 정답 CSV 파일을 찾을 수 없습니다: {csv_path}")
        sys.exit(1)

    # 2. 데이터 로드
    try:
        df = pd.read_csv(csv_path)
        print(f"-> [로드 완료] {os.path.basename(csv_path)} - 총 {len(df)}개 레코드 발견.")
    except Exception as e:
        print(f"[오류] CSV 파일을 읽는 도중 에러가 발생했습니다: {e}")
        sys.exit(1)

    # 3. 검증 항목 정의
    total_records = len(df)
    missing_images = []
    invalid_extensions = []
    duplicate_records = []
    
    # record_id 중복 검증
    id_counts = df['record_id'].value_counts()
    duplicates = id_counts[id_counts > 1]
    for rec_id, count in duplicates.items():
        duplicate_records.append({
            "record_id": rec_id,
            "count": count,
            "indices": df[df['record_id'] == rec_id].index.tolist()
        })

    # 허용할 이미지 확장자
    allowed_extensions = {'.jpg', '.jpeg', '.png'}

    # 이미지 파일 존재 여부 및 확장자 검증
    for idx, row in df.iterrows():
        rec_id = row.get('record_id', f"Unknown_Row_{idx}")
        img_rel_path = row.get('image_filename', '')
        
        if pd.isna(img_rel_path) or not img_rel_path.strip():
            missing_images.append({
                "record_id": rec_id,
                "row": idx + 2, # CSV의 1-indexed, 헤더 포함 라인 보정
                "path": "N/A (결측치)",
                "reason": "이미지 파일명이 비어 있습니다."
            })
            continue

        # 상대 경로를 절대 경로로 보정
        img_abs_path = os.path.normpath(os.path.join(project_root, img_rel_path))
        
        # 1) 확장자 체크
        _, ext = os.path.splitext(img_rel_path.lower())
        if ext not in allowed_extensions:
            invalid_extensions.append({
                "record_id": rec_id,
                "row": idx + 2,
                "path": img_rel_path,
                "ext": ext if ext else "확장자 없음"
            })

        # 2) 파일 존재 체크
        if not os.path.exists(img_abs_path):
            missing_images.append({
                "record_id": rec_id,
                "row": idx + 2,
                "path": img_rel_path,
                "reason": "파일이 지정된 로컬 경로에 존재하지 않습니다."
            })

    # 4. 검증 결과 요약 및 표 작성
    validation_passed = (len(missing_images) == 0) and (len(invalid_extensions) == 0) and (len(duplicate_records) == 0)
    
    # 리포트 텍스트 생성
    report_content = []
    report_content.append("=" * 80)
    report_content.append("           [데이터 & 이미지 매칭 정밀 검증 결과 보고서]")
    report_content.append("=" * 80)
    report_content.append(f" - 검증 일시: 2026-07-10 09:40")
    report_content.append(f" - 대상 CSV: data/source_structured/ground_truth_multimodal_240.csv")
    report_content.append(f" - 총 검사 레코드 수: {total_records}건")
    report_content.append(f" - 종합 검증 상태: {'[PASS] 이상 없음' if validation_passed else '[WARNING] 이슈 발견'}")
    report_content.append("-" * 80)
    report_content.append(f" 1. 누락 이미지 파일 수: {len(missing_images)}건")
    report_content.append(f" 2. 허용되지 않은 확장자 파일 수: {len(invalid_extensions)}건")
    report_content.append(f" 3. 중복된 record_id 수: {len(duplicate_records)}건")
    report_content.append("=" * 80 + "\n")

    # 세부 표 추가: 누락 이미지
    if missing_images:
        report_content.append("[표 1] 누락 이미지 목록 (Missing Images)")
        report_content.append("-" * 90)
        report_content.append(f"{'레코드 ID':<12} | {'행 번호':<8} | {'설명':<35} | {'이미지 상대 경로'}")
        report_content.append("-" * 90)
        for item in missing_images:
            report_content.append(f"{item['record_id']:<12} | {item['row']:<8} | {item['reason']:<35} | {item['path']}")
        report_content.append("-" * 90 + "\n")
    else:
        report_content.append("[OK] 누락된 이미지 파일이 발견되지 않았습니다.\n")

    # 세부 표 추가: 잘못된 확장자
    if invalid_extensions:
        report_content.append("[표 2] 허용되지 않은 확장자 목록 (Invalid Extensions - 허용: JPG, JPEG, PNG)")
        report_content.append("-" * 90)
        report_content.append(f"{'레코드 ID':<12} | {'행 번호':<8} | {'확장자':<15} | {'이미지 상대 경로'}")
        report_content.append("-" * 90)
        for item in invalid_extensions:
            report_content.append(f"{item['record_id']:<12} | {item['row']:<8} | {item['ext']:<15} | {item['path']}")
        report_content.append("-" * 90 + "\n")
    else:
        report_content.append("[OK] 잘못된 파일 확장자가 발견되지 않았습니다.\n")

    # 세부 표 추가: 중복 레코드
    if duplicate_records:
        report_content.append("[표 3] 중복 record_id 목록 (Duplicate record_ids)")
        report_content.append("-" * 90)
        report_content.append(f"{'레코드 ID':<15} | {'중복 빈도(건)':<15} | {'중복 위치 (CSV Row Indices)'}")
        report_content.append("-" * 90)
        for item in duplicate_records:
            indices_str = ", ".join(map(lambda x: str(x+2), item['indices']))
            report_content.append(f"{item['record_id']:<15} | {item['count']:<15} | Rows: {indices_str}")
        report_content.append("-" * 90 + "\n")
    else:
        report_content.append("[OK] 중복된 record_id가 발견되지 않았습니다.\n")

    # 파일에 기록
    try:
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("\n".join(report_content))
        print(f"-> [작성 완료] 검사 리포트 저장 위치: {report_path}")
    except Exception as e:
        print(f"[오류] 리포트 파일 생성 도중 에러가 발생했습니다: {e}")

    # 터미널 화면에 결과 출력 (Windows 콘솔 인코딩 에러 방지용 safe_print 구현)
    def safe_print(text):
        try:
            print(text)
        except UnicodeEncodeError:
            try:
                # utf-8 터미널 세팅 시도
                if hasattr(sys.stdout, 'reconfigure'):
                    sys.stdout.reconfigure(encoding='utf-8')
                    print(text)
                else:
                    print(text.encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding))
            except Exception:
                print(text.encode('ascii', errors='replace').decode('ascii'))

    safe_print("\n" + "\n".join(report_content[:15]))
    safe_print(f" -> 상세 내역은 '{os.path.relpath(report_path, project_root)}' 파일에서 보실 수 있습니다.\n")
    safe_print("=" * 60)
    safe_print(" [1단계] 이미지 매칭 검증 처리가 성공적으로 종료되었습니다.")
    safe_print("=" * 60 + "\n")
    
    return validation_passed

if __name__ == "__main__":
    run_validation()
