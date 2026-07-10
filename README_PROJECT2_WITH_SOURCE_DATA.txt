# Antigravity Project 2 OCR 실습 데이터 패키지

이 패키지는 Project 2: 멀티모달 기반 통합 분석 에이전트 실습용 원본 데이터입니다.
이번 실습은 크롤링이 아니라, 제공된 이미지/정형 데이터를 기반으로 OCR 추출, 결측치 보간, Streamlit 대시보드 생성을 진행합니다.

포함 데이터:
- 영수증 이미지 120장: data/input_images/receipts
- 수기 설문지 이미지 120장: data/input_images/surveys
- 원본 정답 데이터 CSV/XLSX/JSONL: data/source_structured
- 프롬프트 순서 파일: prompts
- 체크리스트와 오류 대응 프롬프트: checklists, teacher

수강생 진행 방식:
1. 이 폴더 전체를 Antigravity 프로젝트 폴더에 넣습니다.
2. prompts/antigravity_prompt_sequence_project2_with_source_data.txt를 엽니다.
3. 프롬프트 1번부터 순서대로 Antigravity에 입력합니다.
4. Antigravity가 OCR 스크립트, 정제 스크립트, 대시보드 앱을 생성하게 합니다.
5. final_output ZIP이 생성되면 실습을 완료합니다.
