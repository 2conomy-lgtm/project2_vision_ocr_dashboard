#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Project 2: 멀티모달 기반 통합 분석 에이전트 실습 프로젝트
[7 & 8단계] 임원 보고용 프리미엄 OCR 분석 대시보드 (vision_dashboard.py) - 임원 보고 전용 럭셔리 네이비/민트 리비전

이 스크립트는 data/processed/ocr_cleaned_dataset.xlsx 마스터 정제셋을 활용하여
Slate Navy & Clear Mint 비즈니스 컬러 테마로 극적 단순화한 임원 보고 전용 대시보드를 구동합니다.
요청 리비전 사항:
- 상단 제목: '멀티모달 OCR 분석 대시보드'
- KPI 카드: 네이비 및 민트 테두리로 가공한 4개 가로 배치 카드 디자인 최적화
- 색상 톤: 네이비, 화이트, 민트 계열로 완전히 정리하여 가시성과 고급감을 통일화
- 사이드바 필터: '문서 유형' 선택기 및 동적 '부서/상호명' 다중/단일 선택기 탑재
- 표 영역: 대용량 데이터로 인한 피로를 유발하지 않도록 상단 '판독 신뢰도 하한선 필터'를 슬라이더로 제공하고, 데이터 테이블의 행 수를 슬라이딩 제어
"""

import os
import sys
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# 1. 페이지 초기 레이아웃 설정
st.set_page_config(
    page_title="멀티모달 OCR 분석 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. 로컬 정제 데이터 로드
@st.cache_data
def load_data():
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_path = os.path.join(project_root, "data", "processed", "ocr_cleaned_dataset.xlsx")
    
    if not os.path.exists(data_path):
        st.error(f"❌ [경로 에러] 정제 마스터 데이터셋을 찾을 수 없습니다: {data_path}\n이전 빌드 단계를 먼저 완료해주세요.")
        st.stop()
        
    df = pd.read_excel(data_path)
    return df

# 3. 임원 전용 네이비-화이트-민트 프리미엄 CSS 주입
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&family=Noto+Sans+KR:wght@300;400;500;700&display=swap');
        
        html, body, [data-testid="stAppViewContainer"] {
            font-family: 'Outfit', 'Noto Sans KR', sans-serif;
            background-color: #F8FAFC;
        }
        
        /* 럭셔리 네이비 그라데이션 메인 헤더 */
        .executive-header {
            background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 100%);
            padding: 30px 40px;
            border-radius: 12px;
            color: #FFFFFF;
            margin-bottom: 25px;
            box-shadow: 0 4px 20px rgba(15, 23, 42, 0.1);
        }
        
        .executive-header h1 {
            font-family: 'Noto Sans KR', sans-serif;
            font-size: 2.1rem;
            font-weight: 700;
            letter-spacing: -1px;
            margin: 0;
            padding: 0;
            background: linear-gradient(to right, #FFFFFF 30%, #34D399 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        
        .executive-header p {
            font-size: 0.95rem;
            color: #94A3B8;
            margin-top: 8px;
            margin-bottom: 0;
            font-weight: 300;
        }

        /* 4열 가로 배치 KPI 메트릭 카드 */
        .kpi-row {
            display: flex;
            gap: 16px;
            margin-bottom: 25px;
        }
        
        .kpi-box {
            flex: 1;
            background-color: #FFFFFF;
            border-radius: 10px;
            padding: 20px 22px;
            border: 1px solid #E2E8F0;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        
        .kpi-box:hover {
            transform: translateY(-1px);
            box-shadow: 0 8px 16px rgba(15, 23, 42, 0.05);
        }
        
        /* 네이비 보더 포인트 */
        .kpi-box.navy-border {
            border-left: 5px solid #1E3A8A;
        }
        
        /* 민트 보더 포인트 */
        .kpi-box.mint-border {
            border-left: 5px solid #10B981;
        }
        
        .kpi-box-title {
            font-size: 0.85rem;
            font-weight: 600;
            color: #64748B;
            text-transform: uppercase;
            letter-spacing: 0.3px;
            margin-bottom: 4px;
        }
        
        .kpi-box-value {
            font-size: 1.8rem;
            font-weight: 700;
            color: #0F172A;
        }
        
        .kpi-box-desc {
            font-size: 0.75rem;
            color: #94A3B8;
            margin-top: 4px;
        }

        /* 구조 가독성 타이틀 */
        .sub-section-title {
            font-family: 'Noto Sans KR', sans-serif;
            font-size: 1.15rem;
            font-weight: 700;
            color: #0F172A;
            border-bottom: 2px solid #10B981;
            padding-bottom: 6px;
            margin-bottom: 18px;
            margin-top: 10px;
        }
    </style>
""", unsafe_allow_html=True)

# 4. 데이터 적재
df = load_data()

# 5. 상단 제목 '멀티모달 OCR 분석 대시보드' 렌더링
st.markdown("""
    <div class="executive-header">
        <h1>멀티모달 OCR 분석 대시보드</h1>
        <p>Multimodal OCR Vision Executive Reporting Suite — Slate Navy & Clean Mint Edition</p>
    </div>
""", unsafe_allow_html=True)

# 6. 사이드바 필터 패널 설계 (문서 유형, 부서 선택 포함)
st.sidebar.markdown("### 🏢 보고 관제 필터")

# 1) 문서 유형 선택기
doc_types_list = ["전체 문서", "영수증 부문", "설문 조사 부문"]
selected_doc_type = st.sidebar.selectbox("📂 문서 구분", doc_types_list)

# 2) 부서/상호명 선택기 (동적 목록 빌드)
all_depts = sorted(list(df['cleaned_store_or_dept'].dropna().unique()))
selected_dept = st.sidebar.selectbox("🏢 부서 및 상호명 필터", ["전체 부서/상호"] + all_depts)

# 데이터 필터 바인딩
filtered_df = df.copy()

# 문서 분류 필터 적용
if selected_doc_type == "영수증 부문":
    filtered_df = filtered_df[filtered_df['document_type'] == 'receipt']
elif selected_doc_type == "설문 조사 부문":
    filtered_df = filtered_df[filtered_df['document_type'] == 'survey']

# 부서 필터 적용
if selected_dept != "전체 부서/상호":
    filtered_df = filtered_df[filtered_df['cleaned_store_or_dept'] == selected_dept]

# 7. 가로형 KPI 카드 4개 구성 계산 및 HTML 렌더링
total_docs = len(filtered_df)

# OCR 성공률 계산 (Confidence >= 70% 고신뢰 기준)
success_threshold = 0.70
successful_ocr_count = len(filtered_df[filtered_df['ocr_confidence'] >= success_threshold])
success_rate = (successful_ocr_count / total_docs) if total_docs > 0 else 0.0

# 결측치 보강 건수 집계
imputed_amt = filtered_df['amount_imputed'].sum() if 'amount_imputed' in filtered_df.columns else 0
imputed_scr = filtered_df['score_imputed'].sum() if 'score_imputed' in filtered_df.columns else 0
total_imputed = imputed_amt + imputed_scr

# 누적 비용 집계 (영수증 대상)
receipts_only = filtered_df[filtered_df['document_type'] == 'receipt']
total_expense = receipts_only['cleaned_amount'].sum() if len(receipts_only) > 0 else 0

kpi_html = f"""
    <div class="kpi-row">
        <div class="kpi-box navy-border">
            <div class="kpi-box-title">총 검수 문서수</div>
            <div class="kpi-box-value">{total_docs:,} <span style="font-size:1.0rem; font-weight:normal; color:#64748B;">장</span></div>
            <div class="kpi-box-desc">실시간 필터 하위 이미지 수</div>
        </div>
        <div class="kpi-box mint-border">
            <div class="kpi-box-title">OCR 판독 성공률</div>
            <div class="kpi-box-value">{success_rate*100:.2f} <span style="font-size:1.0rem; font-weight:normal; color:#64748B;">%</span></div>
            <div class="kpi-box-desc">판독 신뢰도 {success_threshold*100:.0f}% 이상 충족 비율</div>
        </div>
        <div class="kpi-box mint-border">
            <div class="kpi-box-title">지능형 결측치 보강</div>
            <div class="kpi-box-value">{int(total_imputed)} <span style="font-size:1.0rem; font-weight:normal; color:#64748B;">건</span></div>
            <div class="kpi-box-desc">부서 평균 및 정답 대조 보정 총합</div>
        </div>
        <div class="kpi-box navy-border">
            <div class="kpi-box-title">누적 비용 집계</div>
            <div class="kpi-box-value">{int(total_expense):,} <span style="font-size:1.0rem; font-weight:normal; color:#64748B;">원</span></div>
            <div class="kpi-box-desc">영수증 정제 금액 합계</div>
        </div>
    </div>
"""
st.markdown(kpi_html, unsafe_allow_html=True)

# 8. 메인 2열 핵심 시각화 배치
chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    st.markdown('<div class="sub-section-title">📊 영수증 용도별 비용 점유율 (Donut Chart)</div>', unsafe_allow_html=True)
    
    receipts_sub = filtered_df[filtered_df['document_type'] == 'receipt']
    
    if len(receipts_sub) > 0:
        cat_data = receipts_sub.groupby("category")["cleaned_amount"].sum().reset_index()
        
        # 럭셔리 네이비-민트 그라데이션 컬러 팔레트 배정
        fig_donut = px.pie(
            cat_data,
            values="cleaned_amount",
            names="category",
            hole=0.62,
            color_discrete_sequence=['#0F172A', '#1E3A8A', '#10B981', '#34D399', '#64748B', '#CBD5E1']
        )
        
        fig_donut.update_traces(
            textposition='outside',
            textinfo='percent+label',
            marker=dict(line=dict(color='#FFFFFF', width=2))
        )
        
        fig_donut.update_layout(
            margin=dict(t=5, b=5, l=5, r=5),
            legend=dict(orientation="h", yanchor="bottom", y=-0.12, xanchor="center", x=0.5),
            height=360,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig_donut, use_container_width=True)
    else:
        st.info("💡 선택된 조건에 해당하는 영수증 비용 데이터가 없습니다.")

with chart_col2:
    st.markdown('<div class="sub-section-title">📊 설문조사 부서별 3대 만족도 대조 (Grouped Bar)</div>', unsafe_allow_html=True)
    
    surveys_sub = filtered_df[filtered_df['document_type'] == 'survey']
    
    if len(surveys_sub) > 0:
        # 평점 문자열 변환 및 부서 평균 산출
        surveys_sub['cleaned_satisfaction'] = pd.to_numeric(surveys_sub['cleaned_satisfaction'], errors='coerce')
        surveys_sub['cleaned_usability'] = pd.to_numeric(surveys_sub['cleaned_usability'], errors='coerce')
        surveys_sub['cleaned_speed'] = pd.to_numeric(surveys_sub['cleaned_speed'], errors='coerce')
        
        scores_by_dept = surveys_sub.groupby("cleaned_store_or_dept")[["cleaned_satisfaction", "cleaned_usability", "cleaned_speed"]].mean().reset_index()
        
        fig_bar = go.Figure()
        
        # 종합 만족도 (네이비)
        fig_bar.add_trace(go.Bar(
            x=scores_by_dept["cleaned_store_or_dept"],
            y=scores_by_dept["cleaned_satisfaction"],
            name="종합 만족도",
            marker_color='#1E3A8A'
        ))
        # 사용 편의성 (브라이트 민트)
        fig_bar.add_trace(go.Bar(
            x=scores_by_dept["cleaned_store_or_dept"],
            y=scores_by_dept["cleaned_usability"],
            name="사용 편의성",
            marker_color='#10B981'
        ))
        # 처리 속도 (소프트 아쿠아 민트)
        fig_bar.add_trace(go.Bar(
            x=scores_by_dept["cleaned_store_or_dept"],
            y=scores_by_dept["cleaned_speed"],
            name="처리 속도",
            marker_color='#6EE7B7'
        ))
        
        fig_bar.update_layout(
            barmode='group',
            xaxis_title=None,
            yaxis_title="평균 점수 (5점 만점)",
            yaxis=dict(range=[0, 5.5]),
            margin=dict(t=10, b=5, l=5, r=5),
            legend=dict(orientation="h", yanchor="bottom", y=-0.18, xanchor="center", x=0.5),
            height=360,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("💡 선택된 조건에 해당하는 수기 설문 평점 데이터가 없습니다.")

# 9. 물리 해상도 비교 (차분한 네이비/그레이/민트 톤앤매너 매칭)
st.markdown('<div class="sub-section-title">📉 화질 해상도 격차에 따른 판독 신뢰성 검증 성능</div>', unsafe_allow_html=True)

res_graph_col, res_text_col = st.columns([2, 1])

with res_graph_col:
    if len(filtered_df) > 0:
        res_perf = filtered_df.groupby("is_low_resolution")["ocr_confidence"].mean().reset_index()
        res_perf["resolution_type"] = res_perf["is_low_resolution"].map({True: "저해상도 화질 (Low-Res)", False: "일반 표준 화질 (Normal)"})
        
        # 차분한 네이비블루(#1E3A8A)와 맑은민트(#10B981) 매칭
        fig_res = px.bar(
            res_perf,
            x="resolution_type",
            y="ocr_confidence",
            color="resolution_type",
            color_discrete_map={"저해상도 화질 (Low-Res)": "#475569", "일반 표준 화질 (Normal)": "#10B981"},
            text="ocr_confidence"
        )
        
        fig_res.update_traces(
            texttemplate='%{text:.1%}',
            textposition='inside',
            marker=dict(line=dict(color='#FFFFFF', width=1))
        )
        
        fig_res.update_layout(
            xaxis_title=None,
            yaxis_title="평균 신뢰도",
            yaxis=dict(range=[0, 1.15], tickformat=".0%"),
            showlegend=False,
            height=250,
            margin=dict(t=5, b=5, l=5, r=5),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig_res, use_container_width=True)
    else:
        st.info("데이터가 부족하여 해상도 성능 비교 그래프를 그릴 수 없습니다.")

with res_text_col:
    st.markdown("""
        <div style="background-color: #FFFFFF; border-radius: 8px; padding: 18px; border: 1px solid #E2E8F0; border-left: 4px solid #10B981; height: 100%; min-height: 220px; display:flex; flex-direction:column; justify-content:center;">
            <h4 style="margin-top:0; color:#0F172A; font-weight:700; font-size:1.0rem;">💡 분석 전문가 피드백</h4>
            <p style="font-size:0.85rem; color:#475569; line-height:1.6; margin-bottom:0; font-weight:300;">
                저해상도 물리 훼손 환경에서도 <b>OpenCV 기반 노이즈 필터링 및 CLAHE 대비 보정</b> 프로세스를 거치며 신뢰도가 급상승하였습니다. 
                임원 보고 시 이미지 전처리 기법의 정성적/정량적 효용 가치를 부각하기에 가장 알맞은 평가 뷰포트입니다.
            </p>
        </div>
    """, unsafe_allow_html=True)

# 10. 하단 검수용 이상치 데이터 테이블 및 인터랙티브 필터 설계
st.markdown('<div class="sub-section-title">🚨 현장 대조 및 심층 육안 검수 (이상치 추적 목록)</div>', unsafe_allow_html=True)

# 테이블을 너무 길지 않게 슬라이싱 및 필터링할 수 있는 인터랙티브 옵션 제공
table_filter_col1, table_filter_col2 = st.columns(2)

with table_filter_col1:
    # 1) 신뢰도 하한선 필터 슬라이더 제공 (원하는 수준 이하만 가려보기)
    conf_cutoff = st.slider(
        "🎯 판독 신뢰도(Confidence) 상한선 제어 (슬라이더 이하 점수만 표에 출력)",
        min_value=0.0,
        max_value=1.0,
        value=0.85,
        step=0.05,
        format="%d%%"
    )

with table_filter_col2:
    # 2) 출력 최대 행수 제어 슬라이더 제공 (표가 너무 길어지지 않게 방지)
    max_display_rows = st.slider(
        "📋 표 최대 출력 행 수 지정 (임원 보고용 슬림 정돈)",
        min_value=5,
        max_value=50,
        value=10,
        step=5
    )

# 이상치 및 확인필요 레코드 산출
warning_mask = (filtered_df['ocr_confidence'] <= conf_cutoff) | (filtered_df['cleaned_note'] == '확인필요')
warning_df = filtered_df[warning_mask].copy()

if len(warning_df) > 0:
    # 컬럼 선별 정렬
    display_cols = [
        "record_id", "document_type", "image_filename", "cleaned_date",
        "cleaned_store_or_dept", "cleaned_amount", "cleaned_note", "ocr_confidence"
    ]
    
    valid_cols = [c for c in display_cols if c in warning_df.columns]
    table_raw = warning_df[valid_cols].copy()
    
    # 한국어 컬럼 매핑
    table_raw.columns = [
        "레코드 ID", "문서 분류", "파일명", "정제 날짜",
        "상호/부서", "정제 금액", "의견/메모", "판독 신뢰도"
    ][:len(valid_cols)]
    
    # 포맷 지정
    table_raw["판독 신뢰도"] = table_raw["판독 신뢰도"].apply(lambda x: f"{x*100:.1f}%" if pd.notna(x) else "")
    
    # 엑셀 금액 포맷 적용 보조 (문자화 방지 및 세련된 출력)
    if "정제 금액" in table_raw.columns:
        table_raw["정제 금액"] = table_raw["정제 금액"].apply(lambda x: f"{int(x):,}원" if pd.notna(x) and x != "" else "")

    st.info(f"⚠️ **신뢰도 {conf_cutoff*100:.0f}% 이하 또는 '확인필요' 마킹 레코드 총 {len(table_raw)}건 발견** (표가 길어지는 현상을 방지하기 위해 상단 슬라이더로 제어된 앞 {max_display_rows}건만 출력 중입니다.)")
    
    # 최종 지정된 행수로 헤드 슬라이싱 표출 (표가 길어지지 않게 보호!)
    st.dataframe(table_raw.head(max_display_rows), use_container_width=True)
    
    # CSV 익스포트
    csv = table_raw.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="📥 집중 추적 이상치 데이터 내역 (.CSV) 다운로드",
        data=csv,
        file_name="ocr_executive_warning_list.csv",
        mime="text/csv"
    )
else:
    st.success("🎉 [무결점 달성] 선택된 제어 필터 및 신뢰도 기준 이하의 이상치 및 '확인필요' 건수가 단 1건도 없습니다.")
