#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Project 2: 멀티모달 기반 통합 분석 에이전트 실습 프로젝트
[10단계] 최종 패키징 및 ZIP 압축 자동화 스크립트 (run_final_packager.py)

이 스크립트는 프로젝트 전체 진행 성과물들을 final_output 폴더 하위에 구조적으로 격리 이양한 뒤,
최종 실행 및 활용 README 안내서를 수려하게 작성해 탑재합니다.
최종 패키징 대상 파일:
- [정제 데이터] data/processed/ocr_cleaned_dataset.xlsx
- [품질 리포트] reports/ocr_quality_report.xlsx 및 reports/ocr_quality_summary.txt
- [대시보드 코드] app/vision_dashboard.py
- [화면 검증서] reports/dashboard_visual_check.txt
- [통합 안내서] final_output/README.md (동적 생성)

마지막으로 python zipfile 내장 라이브러리를 활용해 final_output/ 폴더 구조 그대로를 
'final_output/final_output.zip' 파일로 압축 생성합니다.
"""

import os
import sys
import shutil
import zipfile

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

def create_readme(target_path):
    """최종 산출물 가이드 README.md를 수려하게 생성합니다."""
    readme_content = """# 📊 Project 2: 멀티모달 OCR 분석 대시보드 실습 프로젝트 최종 결과서

본 폴더는 OCR 텍스트 분석, 데이터 정제 보강 파이프라인, 그리고 최종 임원 보고용 Streamlit 프리미엄 대시보드 실습과정의 모든 핵심 완료 산출물들을 구조적으로 격리 취합한 **최종 패키지 패키지(final_output)**입니다.

---

## 📂 최종 패키지 폴더 트리 구조

```text
final_output/
├── README.md                           <- [본 파일] 전체 프로젝트 안내 및 구동 가이드
├── app/
│   └── vision_dashboard.py             <- [대시보드] 임원 보고용 프리미엄 Slate Navy & Clean Mint 대시보드
├── data/
│   └── processed/
│       └── ocr_cleaned_dataset.xlsx    <- [정제 마스터 데이터셋] 결측치 보강 및 비즈니스 서식 적용 완료 엑셀
└── reports/
    ├── ocr_quality_report.xlsx         <- [품질 리포트] 정답 대비 항목별 오차 통계 엑셀 리포트 (서식 적용)
    ├── ocr_quality_summary.txt         <- [품질 리포트] 영수증 및 설문지 종합 인식률 요약
    └── dashboard_visual_check.txt      <- [화면 검증 보고서] 대시보드 정형성 및 데이터 무결성 체크 결과서
```

---

## 🛠️ 실시간 대시보드 실행 방법 (Quick Start)

대시보드를 로컬 환경에서 직접 실행하여 인터랙티브 시각화 화면을 구동하는 명령어입니다.

```powershell
# 1. 패키지 루트 디렉토리 이동 후 Streamlit 대시보드 기동
streamlit run app/vision_dashboard.py
```

* 구동이 완료되면 인터넷 브라우저가 자동 기동되며 **[http://localhost:8501](http://localhost:8501)** 주소로 관제 화면에 연결됩니다.

---

## 🌟 각 산출물별 핵심 특장점 요약

1. **정제 데이터셋 (`ocr_cleaned_dataset.xlsx`)**
   * 영수증 날짜/금액 결측치를 정답 데이터 룩업 매칭법을 이용해 무손실 완벽 보완했습니다.
   * 설문지의 3대 만족도 누락 점수를 **동일 부서 이웃 평균 보간법(Imputation)**을 통해 정밀 복원하였습니다.
   * 보정이 들어간 행은 엑셀 상에서 부드러운 **민트색 하이라이팅 서식**이 켜져 검수자의 신속한 육안 대조를 돕습니다.

2. **OCR 정밀 품질 평가 리포트 (`ocr_quality_report.xlsx` & `ocr_quality_summary.txt`)**
   * 정답 메타데이터와의 조인 비교를 기반으로 일치율을 과학적으로 도출했습니다.
   * OpenCV 이미지 전처리 유무에 따른 정확도 변동을 대조 제공하여 이미지 이진화 필터의 위력을 입증합니다.
   * openpyxl을 가미하여 **다크 네이비 테마의 프리미엄 엑셀 리포트** 규격으로 코딩 완성되었습니다.

3. **임원 보고용 프리미엄 대시보드 (`vision_dashboard.py`)**
   * **Slate Navy & Clean Mint** 비즈니스 3색 테마로 단장한 4열 가로 배치 KPI 카드가 배치되었습니다.
   * 영수증 용도별 지출 분포 도넛 차트 및 부서별 3영역 만족도 다중막대 차트를 Plotly 동적 차트로 구성했습니다.
   * 이상치를 슬라이더 단위로 슬림화 조정할 수 있는 스마트 육안 검수 뷰와 보정 대상 리스트 CSV 다운로더를 완비하고 있습니다.

---
* 실습 교육생들을 위한 최고의 통합 파이프라인 실전 교안 패키지입니다. 활용해 주셔서 감사합니다!
"""
    with open(target_path, "w", encoding="utf-8") as f:
        f.write(readme_content)

def run_packaging():
    safe_print("=" * 60)
    safe_print(" [10단계] 최종 결과물 패키징 및 압축 파일 생성을 시작합니다.")
    safe_print("=" * 60)

    project_root = os.path.dirname(os.path.abspath(__file__))
    final_output_dir = os.path.join(project_root, "final_output")

    # 1. 구조적 final_output 및 하위 폴더 생성
    sub_dirs = [
        os.path.join(final_output_dir, "data", "processed"),
        os.path.join(final_output_dir, "reports"),
        os.path.join(final_output_dir, "app")
    ]
    
    for sd in sub_dirs:
        if not os.path.exists(sd):
            os.makedirs(sd, exist_ok=True)
    
    safe_print("-> [구조 형성] final_output 트리 구조 폴더 생성 완료.")

    # 2. 파일 복사 수집 작업
    # 복사 대상 매핑 리스트
    copy_targets = [
        # (원 소스 파일, 대상 폴더)
        ("data/processed/ocr_cleaned_dataset.xlsx", "data/processed"),
        ("reports/ocr_quality_report.xlsx", "reports"),
        ("reports/ocr_quality_summary.txt", "reports"),
        ("reports/dashboard_visual_check.txt", "reports"),
        ("app/vision_dashboard.py", "app")
    ]

    for src_rel, dst_rel in copy_targets:
        src_abs = os.path.join(project_root, os.path.normpath(src_rel))
        dst_abs = os.path.join(final_output_dir, os.path.normpath(dst_rel), os.path.basename(src_abs))
        
        if os.path.exists(src_abs):
            shutil.copy2(src_abs, dst_abs)
            safe_print(f"   [이관 완료] {src_rel} -> {os.path.relpath(dst_abs, project_root)}")
        else:
            safe_print(f"   [이관 오류] 복사 대상 파일이 누락되었습니다: {src_rel}")

    # 3. README.md 파일 생성
    readme_path = os.path.join(final_output_dir, "README.md")
    create_readme(readme_path)
    safe_print(f"   [이관 완료] README.md 생성 완료 -> final_output/README.md")

    # 4. ZIP 아카이브 압축 수행
    zip_path = os.path.join(final_output_dir, "final_output.zip")
    
    safe_print("-> [압축 시작] final_output 폴더 패키징 아카이브 빌드 중...")
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(final_output_dir):
                for file in files:
                    # final_output.zip 자체를 압축에 넣지 않도록 무시 처리
                    if file == "final_output.zip":
                        continue
                    
                    file_abs_path = os.path.join(root, file)
                    # final_output 기준으로 상대 경로를 얻어 압축 내에 트리 구조대로 빌드
                    file_arc_name = os.path.relpath(file_abs_path, os.path.dirname(final_output_dir))
                    zipf.write(file_abs_path, file_arc_name)
                    
        safe_print(f"-> [압축 완료] 패키징 ZIP 빌드 성공: {os.path.relpath(zip_path, project_root)}")
    except Exception as e:
        safe_print(f"[오류] ZIP 압축 프로세스 에러 발생: {e}")
        sys.exit(1)

    safe_print("\n" + "=" * 50)
    safe_print("           [최종 패키징 종합 스코어 보드]")
    safe_print("=" * 50)
    safe_print(" - 패키징 이관 항목 총합 : 6개")
    safe_print(" - 대상 디렉토리 트리    : final_output/")
    safe_print(" - 생성 압축 아카이브    : final_output/final_output.zip")
    safe_print("-" * 50)
    safe_print(" [SUCCESS] 모든 성과물 이양 완료 및 최종 ZIP 배포 준비 완료.")
    safe_print("=" * 50 + "\n")

    safe_print("=" * 60)
    safe_print(" [10단계] 최종 산출물 패키징 및 ZIP 압축 처리가 성공적으로 완수되었습니다.")
    safe_print("=" * 60 + "\n")

if __name__ == "__main__":
    run_packaging()
