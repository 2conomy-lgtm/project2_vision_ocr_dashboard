#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Project 2: 멀티모달 기반 통합 분석 에이전트 실습 프로젝트
[3단계] OCR 추출기 스크립트 (run_ocr_extractor.py)

이 스크립트는 data/input_images 폴더 안의 이미지들을 순회하며 OCR을 구동합니다.
실제 EasyOCR, Tesseract 등 무거운 패키지가 설치되지 않은 실습 환경에서도 원활히 동작할 수 있도록,
정답 데이터(ground_truth_multimodal_240.csv)에 기반한 "고성능 하이브리드 OCR 시뮬레이션 엔진"을 기본 지원합니다.

특히, 이미지의 노이즈(has_noise) 및 저해상도(is_low_resolution) 상태에 따라 고의적으로 
오타/결측치/낮은 신뢰도(Confidence)를 발생시키며, '이미지 전처리' 단계(2단계)를 거친 경우 
품질과 신뢰도가 대폭 교정 및 회복되는 고수준 시뮬레이션이 연동되어 실습 효과를 극대화합니다.
"""

import os
import sys
import random
import pandas as pd
import numpy as np

def run_ocr():
    print("=" * 60)
    print(" [3단계] OCR 텍스트 및 데이터 추출을 시작합니다.")
    print("=" * 60)

    # 1. 경로 설정
    project_root = os.path.dirname(os.path.abspath(__file__))
    gt_csv_path = os.path.join(project_root, "data", "source_structured", "ground_truth_multimodal_240.csv")
    ocr_dir = os.path.join(project_root, "data", "ocr")
    preprocessed_dir = os.path.join(ocr_dir, "preprocessed_images")
    output_csv_path = os.path.join(ocr_dir, "ocr_extracted_raw.csv")

    # output 폴더 자동 생성
    if not os.path.exists(ocr_dir):
        os.makedirs(ocr_dir, exist_ok=True)
        print(f"-> [생성 완료] OCR 결과 폴더: {ocr_dir}")

    # 정답 CSV 로딩
    if not os.path.exists(gt_csv_path):
        print(f"[오류] 정답 기준 데이터를 찾을 수 없습니다: {gt_csv_path}")
        sys.exit(1)

    try:
        gt_df = pd.read_csv(gt_csv_path)
    except Exception as e:
        print(f"[오류] CSV 로드 실패: {e}")
        sys.exit(1)

    # 2. 이미지 전처리 수행 여부 동적 체크
    # data/ocr/preprocessed_images 폴더 내에 변환된 이미지가 다수 존재하면 전처리가 완료된 것으로 판별
    preprocessed_files_count = 0
    if os.path.exists(preprocessed_dir):
        preprocessed_files_count = len([f for f in os.listdir(preprocessed_dir) if f.lower().endswith(('.jpg', '.jpeg', '.png'))])
    
    # 전처리 완료 조건: 전체 240장 중 최소 100장 이상 전처리된 경우 활성화로 판정
    preprocessing_active = preprocessed_files_count >= 100
    if preprocessing_active:
        print(f"-> [상태 감지] 이미지 전처리 적용됨 확인! (전처리된 이미지 수: {preprocessed_files_count}장)")
        print("   ※ 전처리 보정 효과가 적용되어 OCR 인식 신뢰도와 추출 성공률이 대폭 상승합니다.")
    else:
        print("-> [상태 감지] 이미지 전처리 미적용 (원본 이미지 기준 추출)")
        print("   ※ 노이즈 및 저해상도 이미지의 경우 텍스트 오차, 결측치 및 낮은 Confidence 점수가 발생합니다.")

    # 3. 하이브리드 OCR 처리 루프
    extracted_data = []

    print("\n-> OCR 추출기 가동 중... (240개 레코드 처리)")
    
    # 임의성 제어를 위한 시드 설정 (매번 일정한 실습 결과를 제공하기 위함)
    random.seed(42)
    np.random.seed(42)

    for idx, row in gt_df.iterrows():
        rec_id = row['record_id']
        doc_type = row['document_type']
        img_rel_path = row['image_filename']
        has_noise = row['has_noise']
        is_low_res = row['is_low_resolution']
        
        # 실제 이미지 절대 경로 확인
        img_abs_path = os.path.normpath(os.path.join(project_root, img_rel_path))
        img_exists = os.path.exists(img_abs_path)

        # 초기값 세팅 (정답 데이터 기준)
        orig_date = row['doc_date']
        orig_store_or_dept = row['organization_or_store'] if doc_type == 'receipt' else row['respondent_dept']
        orig_amount = row['total_amount']
        
        # 설문 점수 (satisfaction, usability, speed)
        if doc_type == 'survey':
            orig_scores = f"{int(row['satisfaction_score'])},{int(row['usability_score'])},{int(row['speed_score'])}"
        else:
            orig_scores = ""
            
        orig_note = row['handwritten_note']

        # 에러 처리
        if not img_exists:
            extracted_data.append({
                "record_id": rec_id,
                "document_type": doc_type,
                "image_filename": img_rel_path,
                "extracted_date": "",
                "extracted_store_or_dept": "",
                "extracted_amount": "",
                "extracted_scores": "",
                "extracted_note": "",
                "confidence": 0.0,
                "error_message": "Image file not found on disk.",
                "preprocessing_used": preprocessing_active
            })
            continue

        # 품질 계수 설정
        # 기본적으로 노이즈나 저해상도가 있으면 추출 에러율 상승
        is_impaired = has_noise or is_low_res
        
        # 4. OCR 인식 시뮬레이터의 물리 오차 반영 핵심 로직
        confidence = 0.0
        error_msg = ""
        
        ext_date = orig_date
        ext_store_or_dept = orig_store_or_dept
        ext_amount = orig_amount
        ext_scores = orig_scores
        ext_note = orig_note

        if is_impaired:
            if preprocessing_active:
                # 전처리를 거친 경우: 장애가 있던 이미지도 품질이 대폭 복구됨
                confidence = round(random.uniform(0.90, 0.97), 3)
                
                # 미세한 오차가 남을 수 있지만 결측치는 거의 복구됨
                if has_noise and random.random() < 0.05:
                    # 아주 극소수로 상호명이나 부서 오타 시뮬레이션
                    ext_store_or_dept = orig_store_or_dept[:-1] + "*"
                if is_low_res and random.random() < 0.05:
                    # 극소수로 금액 끝전 오차 시뮬레이션
                    if doc_type == 'receipt' and pd.notna(orig_amount):
                        ext_amount = int(orig_amount) + random.choice([-50, 50, 100])
            else:
                # 전처리 미적용 + 오염/저해상도 이미지: 최악의 품질 발생
                confidence = round(random.uniform(0.40, 0.62), 3)
                
                # 1) 날짜 유실 발생 (15% 확률로 결측)
                if random.random() < 0.15:
                    ext_date = ""
                    
                # 2) 상호명/부서 오타 및 훼손 (30% 확률로 훼손)
                if random.random() < 0.30:
                    if pd.notna(orig_store_or_dept) and len(str(orig_store_or_dept)) > 2:
                        ext_store_or_dept = str(orig_store_or_dept)[:2] + "**"
                    else:
                        ext_store_or_dept = ""
                
                # 3) 금액 오류 및 결측 (영수증의 경우 25% 확률로 결측, 15% 확률로 파싱 오차)
                if doc_type == 'receipt':
                    rand_val = random.random()
                    if rand_val < 0.25:
                        ext_amount = np.nan  # 금액 완전 누락
                    elif rand_val < 0.40:
                        # 0을 누락하거나 자릿수가 밀림 (예: 95150 -> 9515)
                        if pd.notna(orig_amount):
                            ext_amount = int(str(int(orig_amount))[:-1]) if len(str(int(orig_amount))) > 2 else np.nan
                
                # 4) 설문 만족도 점수 결측 (설문지의 경우 20% 확률로 특정 점수 결측/누락)
                if doc_type == 'survey':
                    if random.random() < 0.20:
                        # 특정 항목 결측 시뮬레이션 (콤마 사이에 값 비움)
                        score_list = orig_scores.split(',')
                        idx_to_missing = random.choice([0, 1, 2])
                        score_list[idx_to_missing] = "" # 비워버림
                        ext_scores = ",".join(score_list)
                
                # 5) 수기 메모 누락 (수기 메모는 OCR 인식 한계가 크므로 전처리 없으면 45% 확률로 완전 유실)
                if random.random() < 0.45:
                    ext_note = ""
        else:
            # 상태가 아주 깨끗한 원본 이미지: 매우 높은 품질로 정상 추출
            confidence = round(random.uniform(0.94, 0.99), 3)
            # 결측치나 오차 전혀 없음

        # 5. 리스트에 저장
        # pandas 저장을 위해 numeric 값들이 NaN일 경우 빈 값 처리
        ext_amount_val = "" if pd.isna(ext_amount) or ext_amount == "" else str(int(ext_amount))
        
        extracted_data.append({
            "record_id": rec_id,
            "document_type": doc_type,
            "image_filename": img_rel_path,
            "extracted_date": ext_date if pd.notna(ext_date) else "",
            "extracted_store_or_dept": ext_store_or_dept if pd.notna(ext_store_or_dept) else "",
            "extracted_amount": ext_amount_val,
            "extracted_scores": ext_scores,
            "extracted_note": ext_note if pd.notna(ext_note) else "",
            "confidence": confidence,
            "error_message": error_msg,
            "preprocessing_used": preprocessing_active
        })

    # 6. CSV 저장
    out_df = pd.DataFrame(extracted_data)
    
    # 정렬 순서 보장 (R- 계열 먼저, S- 계열 나중에)
    out_df = out_df.sort_values(by=['document_type', 'record_id']).reset_index(drop=True)
    
    out_df.to_csv(output_csv_path, index=False, encoding="utf-8-sig")
    print(f"-> [추출 완료] OCR 데이터가 '{os.path.relpath(output_csv_path, project_root)}'에 저장되었습니다.")

    # 7. 터미널 요약 보고
    total_count = len(out_df)
    low_confidence_count = len(out_df[out_df['confidence'] < 0.70])
    
    # 결측 건수 대략 측정
    missing_date = len(out_df[out_df['extracted_date'] == ""])
    missing_store = len(out_df[out_df['extracted_store_or_dept'] == ""])
    missing_amount = len(out_df[(out_df['document_type'] == 'receipt') & (out_df['extracted_amount'] == "")])
    missing_scores = len(out_df[(out_df['document_type'] == 'survey') & (out_df['extracted_scores'].apply(lambda x: "" in str(x).split(',') if str(x) else True))])
    missing_notes = len(out_df[out_df['extracted_note'] == ""])

    print("\n" + "=" * 50)
    print("           [OCR 1차 원본 결과 요약 통계]")
    print("=" * 50)
    print(f" - 총 처리 이미지 수  : {total_count}장")
    print(f" - 전처리 적용 여부   : {'적용 완료 [YES]' if preprocessing_active else '미적용 [NO]'}")
    print(f" - 저신뢰도(Confidence < 70%) 건수: {low_confidence_count}건")
    print("-" * 50)
    print("   [필드별 결측치(Missing) 발생 집계]")
    print(f"   * 날짜(extracted_date) 누락          : {missing_date}건")
    print(f"   * 상호/부서(store_or_dept) 누락      : {missing_store}건")
    print(f"   * 영수증 합계금액(amount) 누락       : {missing_amount}건")
    print(f"   * 설문 세부 점수(scores) 부분 누락    : {missing_scores}건")
    print(f"   * 수기 메모(extracted_note) 누락     : {missing_notes}건")
    print("=" * 50)
    print(" ※ 이미지 전처리 전이므로 저품질 이미지에서 다량의 결측치가 시뮬레이션되었습니다.")
    print(" ※ 4단계 전처리기를 구동하면 위 결측치와 신뢰도가 눈에 띄게 개선됩니다.\n")
    
    print("=" * 60)
    print(" [3단계] OCR 텍스트 추출이 성공적으로 완료되었습니다.")
    print("=" * 60 + "\n")

if __name__ == "__main__":
    run_ocr()
