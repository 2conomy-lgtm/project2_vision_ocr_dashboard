#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Project 2: 멀티모달 기반 통합 분석 에이전트 실습 프로젝트
[4단계] 이미지 전처리 스크립트 (run_image_preprocessing.py)

이 스크립트는 원본 영수증 및 설문지 이미지 전체(240장)를 로드하여 OpenCV 기반 전처리를 적용합니다.
전처리 프로세스는 다음 4단계 핵심 파이프라인으로 구성됩니다.
1. Grayscale 변환 (색상 정보 제거하여 명암 대비에 집중)
2. Bilateral Filter 기반 노이즈 제거 (글자 에지를 보존하며 배경 노이즈 제거)
3. CLAHE (Contrast Limited Adaptive Histogram Equalization) 대비 향상 (균일하지 않은 조명 보정)
4. Otsu 이진화 (Thresholding) (배경과 글자를 흑백으로 뚜렷이 분리)

전처리된 이미지는 data/ocr/preprocessed_images 폴더에 플랫하게 저장됩니다.
"""

import os
import sys
import glob
import cv2

def preprocess_images():
    print("=" * 60)
    print(" [4단계] OpenCV 기반 이미지 전처리를 시작합니다.")
    print("=" * 60)

    # 1. 경로 설정
    project_root = os.path.dirname(os.path.abspath(__file__))
    input_base_dir = os.path.join(project_root, "data", "input_images")
    output_dir = os.path.join(project_root, "data", "ocr", "preprocessed_images")

    # 출력 디렉터리 자동 생성
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
        print(f"-> [생성 완료] 전처리 이미지 저장 경로: {output_dir}")

    # 입력 이미지 경로 확인
    receipts_dir = os.path.join(input_base_dir, "receipts")
    surveys_dir = os.path.join(input_base_dir, "surveys")

    if not os.path.exists(receipts_dir) or not os.path.exists(surveys_dir):
        print("[오류] 입력 이미지 폴더를 찾을 수 없습니다. 경로를 확인해주세요.")
        sys.exit(1)

    # 이미지 목록 수집 (JPG, JPEG, PNG 대소문자 무관 수집)
    receipt_files = []
    survey_files = []
    
    for ext in ['*.jpg', '*.jpeg', '*.png', '*.JPG', '*.JPEG', '*.PNG']:
        receipt_files.extend(glob.glob(os.path.join(receipts_dir, ext)))
        survey_files.extend(glob.glob(os.path.join(surveys_dir, ext)))

    # Windows 대소문자 무구분으로 인한 중복 경로 수집 방지 (set 및 정규화 적용)
    receipt_files = sorted(list(set(os.path.normpath(p) for p in receipt_files)))
    survey_files = sorted(list(set(os.path.normpath(p) for p in survey_files)))

    total_images = len(receipt_files) + len(survey_files)
    print(f"-> [수집 완료] 총 {total_images}장의 원본 이미지 발견.")
    print(f"   * 영수증 이미지 : {len(receipt_files)}장")
    print(f"   * 설문지 이미지 : {len(survey_files)}장")
    print("-" * 60)

    if total_images == 0:
        print("[경고] 전처리할 이미지가 발견되지 않았습니다.")
        return False

    # 2. 이미지 전처리 수행 루프
    print("-> OpenCV 전처리 파이프라인 가동 중... (실시간 처리 및 정형화)")
    
    processed_count = 0
    
    # 두 카테고리의 모든 이미지 순회
    all_image_paths = [('receipt', p) for p in receipt_files] + [('survey', p) for p in survey_files]

    for cat_type, img_path in all_image_paths:
        file_name = os.path.basename(img_path)
        
        # 1) 이미지 읽기
        img = cv2.imread(img_path)
        if img is None:
            print(f"   [오류] 이미지를 로드할 수 없습니다: {file_name}")
            continue

        try:
            # 2) Grayscale 변환
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

            # 3) 노이즈 제거 (Bilateral Filter: 에지를 선명하게 유지하며 주변 노이즈만 제거)
            # d=9, sigmaColor=75, sigmaSpace=75는 OCR 대비용 표준 최적 파라미터입니다.
            denoised = cv2.bilateralFilter(gray, 9, 75, 75)

            # 4) 대비 향상 (CLAHE - 조도 불균형 해소)
            # clipLimit=2.0, tileGridSize=(8,8)로 제한적 대비 적응을 주어 하이라이트 깨짐을 방지합니다.
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            enhanced = clahe.apply(denoised)

            # 5) 이진화 (Otsu Thresholding - 흑백 대비 뚜렷화)
            # 임계값을 Otsu 알고리즘으로 자동 계산하여 바이너리 0과 255로 매핑합니다.
            _, thresholded = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

            # 6) 전처리 이미지 저장 (출력 폴더에 파일명 동일하게 플랫 저장)
            out_file_path = os.path.join(output_dir, file_name)
            cv2.imwrite(out_file_path, thresholded)

            processed_count += 1
            
            # 40장 단위로 처리 진행 상태 모니터링 로그 출력
            if processed_count % 40 == 0 or processed_count == total_images:
                print(f"   [Progress] {processed_count}/{total_images} 장 처리 완료... ({int(processed_count/total_images*100)}%)")

        except Exception as e:
            print(f"   [오류] {file_name} 전처리 실패: {e}")

    # 3. 전처리 완료 리포트 및 다음 단계 가이드
    print("-" * 60)
    print(f"-> [전처리 완료] 총 {processed_count}장의 이미지 전처리 완료 및 저장.")
    print(f"   * 저장 경로: {os.path.relpath(output_dir, project_root)}")
    print("=" * 60)
    print(" [4단계] OpenCV 이미지 전처리 파이프라인이 성공적으로 완료되었습니다.")
    print("=" * 60 + "\n")
    
    return True

if __name__ == "__main__":
    preprocess_images()
