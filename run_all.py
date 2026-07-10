#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Project 2: 멀티모달 기반 통합 분석 에이전트 실습 프로젝트
[통합 제약] 전체 파이프라인 원클릭 자동 실행 마스터 스크립트 (run_all.py)

이 스크립트는 실습 교육생들이 1단계부터 10단계 최종 완성 패키징까지의 
모든 독립적인 분석 단계를 유기적인 단일 연쇄 체인으로 원클릭 구동할 수 있게 제어합니다.
순차 구동 파이프라인:
1. run_validation.py          (데이터셋 이미지 정밀 일관성 매칭 검증)
2. run_image_preprocessing.py (OpenCV 기반 4단계 지능형 이진화 전처리)
3. run_ocr_extractor.py        (전처리 감지 하이브리드 OCR 엔진 구동)
4. run_ocr_analysis.py        (정답 대조형 품질 정확도/유사도 리포트 발행)
5. run_data_cleaning.py        (이웃 평균법 활용 결측 대치 및 마스터 엑셀 빌드)
6. run_dashboard_check.py      (대시보드 구동 정형성 검증 진단서 발행)
7. run_final_packager.py      (산출물 취합 및 final_output.zip 압축 배포)

Windows cp949 한글 파워셸 터미널 출력 우회 인코딩 안전장치를 전면 탑재하고 있습니다.
"""

import os
import sys
import subprocess
import time

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
                print(text.encode('ascii', errors='replace').decode('ascii'))
        except Exception:
            print("Console printing encoding safety backup triggered.")

def run_script(script_name):
    """지정된 파이썬 스크립트를 서브프로세스로 안전하게 가동합니다."""
    project_root = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(project_root, script_name)
    
    safe_print("-" * 65)
    safe_print(f" ▶ [구동 개시] {script_name} 프로세스 실행 중...")
    safe_print("-" * 65)
    
    start_time = time.time()
    
    # stdout/stderr 실시간 출력을 위해 subprocess.Popen 활용
    process = subprocess.Popen(
        [sys.executable, script_path],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        encoding='utf-8',
        errors='replace'
    )
    
    # 실시간 로그 출력 버퍼링 루프
    while True:
        output = process.stdout.readline()
        if output == '' and process.poll() is not None:
            break
        if output:
            safe_print(output.strip())
            
    rc = process.poll()
    elapsed = time.time() - start_time
    
    if rc == 0:
        safe_print(f" ✔ [구동 완료] {script_name} 성공 (소요시간: {elapsed:.2f}초)")
        return True
    else:
        safe_print(f" ❌ [구동 실패] {script_name} 실패 (에러코드: {rc}, 소요시간: {elapsed:.2f}초)")
        return False

def main():
    safe_print("=" * 70)
    safe_print("   [Project 2 OCR 분석 대시보드 전체 파이프라인 원클릭 자동 빌드]")
    safe_print("=" * 70)
    safe_print(" - 총 가동 프로세스 수 : 7개 주요 분석 모듈")
    safe_print(" - 빌드 개시 일시     : 2026-07-10 10:41")
    
    pipeline = [
        "run_validation.py",
        "run_image_preprocessing.py",
        "run_ocr_extractor.py",
        "run_ocr_analysis.py",
        "run_data_cleaning.py",
        "run_dashboard_check.py",
        "run_final_packager.py"
    ]
    
    global_start = time.time()
    success_count = 0
    
    for idx, script in enumerate(pipeline, 1):
        safe_print(f"\n====================== [파이프라인 {idx}/7 단계] ======================")
        ok = run_script(script)
        if ok:
            success_count += 1
        else:
            safe_print(f"\n⚠️ [빌드 중단] {script} 단계에서 치명적 결함이 감지되어 빌드를 중단합니다.")
            sys.exit(1)
            
    total_elapsed = time.time() - global_start
    
    safe_print("\n" + "=" * 70)
    safe_print("   [Project 2 파이프라인 전체 원클릭 자동 빌드 완료 리포트]")
    safe_print("=" * 70)
    safe_print(f" - 전체 소요 시간    : {total_elapsed:.2f}초")
    safe_print(f" - 성공 파이프라인   : {success_count}/7 단계 완수")
    safe_print(" - 대시보드 구동 명령 : streamlit run app/vision_dashboard.py")
    safe_print(" - 최종 ZIP 배포 경로 : final_output/final_output.zip")
    safe_print("-" * 70)
    safe_print(" [SUCCESS] 전체 파이프라인이 완전 무결하게 빌드 종료되었습니다.")
    safe_print("=" * 70 + "\n")

if __name__ == "__main__":
    main()
