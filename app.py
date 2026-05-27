
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns
import os
import unicodedata
from datetime import datetime
from io import BytesIO

# 통계 분석
from scipy import stats
from scipy.stats import chi2_contingency
from sklearn.linear_model import LinearRegression

# python-pptx 임포트 (PPT 생성용)
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# ──────────────────────────────────────────────────────────────────────────
# 페이지 설정
# ──────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SQA 펌웨어 분석 대시보드",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ──────────────────────────────────────────────────────────────────────────
# 한글 폰트
# ──────────────────────────────────────────────────────────────────────────
@st.cache_resource
def setup_korean_font():
    font_path = '/usr/share/fonts/truetype/nanum/NanumGothic.ttf'
    if os.path.exists(font_path):
        fm.fontManager.addfont(font_path)
        return 'NanumGothic'
    return 'DejaVu Sans'

font_name = setup_korean_font()
plt.rcParams['font.family'] = font_name
plt.rcParams['axes.unicode_minus'] = False
sns.set_style("whitegrid")


# ──────────────────────────────────────────────────────────────────────────
# 상수
# ──────────────────────────────────────────────────────────────────────────
KEYWORD_DOMAIN_MAP = {
    'line broken':     {'area': '좌표 보간/샘플링 레이트',     'priority': 'HIGH'},
    'jitter':          {'area': '노이즈 필터/디바운싱',         'priority': 'MID'},
    'ghost touch':     {'area': '노이즈 임계값/그라운드',       'priority': 'HIGH'},
    'no touch':        {'area': '감도 보정/터치 임계값',         'priority': 'CRITICAL'},
    '2 point로 인식':  {'area': '멀티터치 분리 알고리즘',         'priority': 'MID'},
    'edge 과밀착':     {'area': '경계영역 보정 (강하게 반응)',    'priority': 'MID'},
    'edge 미밀착':     {'area': '경계영역 보정 (약하게 반응)',    'priority': 'MID'},
    'touch delay':     {'area': '응답속도/인터럽트 처리',         'priority': 'CRITICAL'},
}

SEVERITY_MAP = {
    'no touch': 5, 'touch delay': 4, 'ghost touch': 4, 'line broken': 3,
    '2 point로 인식': 3, 'edge 과밀착': 2, 'edge 미밀착': 2, 'jitter': 2,
}

TEST_AREA_MAP = {
    'Wet': '방수·방습', 'Linearity': '좌표 보간', 'Filter': '노이즈 필터링',
    'A/P/L': '인식 정확도', 'Separation': '멀티터치 분리',
    'Sensitivity': '터치 감도', 'Palm': 'Palm Rejection',
}

KEYWORDS_ALL = list(KEYWORD_DOMAIN_MAP.keys())

# 보고서 컬러
RPT_HEADER_BG = RGBColor(0x06, 0x5A, 0x82)
RPT_HEADER_FG = RGBColor(0xFF, 0xFF, 0xFF)
RPT_ALT_ROW   = RGBColor(0xF4, 0xF8, 0xFB)
RPT_CRITICAL  = RGBColor(0xD6, 0x28, 0x28)
RPT_HIGH      = RGBColor(0xF7, 0x7F, 0x00)
RPT_MID       = RGBColor(0x06, 0x5A, 0x82)
RPT_TITLE     = RGBColor(0x06, 0x5A, 0x82)
PPT_FONT = "맑은 고딕"


# ──────────────────────────────────────────────────────────────────────────
# 데이터 로드 (캐시)
# ──────────────────────────────────────────────────────────────────────────
@st.cache_data
def load_base_data():
    df = pd.read_csv(
        'SQA_test_results_clean.csv',
        parse_dates=['Test_Date'],
        dtype={'IC_Code': str, 'Model_Minor': str}
    )
    df['IC_Code'] = df['IC_Code'].str.zfill(2)
    df['Model_Minor'] = df['Model_Minor'].str.zfill(2)
    df['Keyword'] = df['Keyword'].fillna('')
    return df


def get_video_link(video_id):
    if pd.isna(video_id) or video_id is None or video_id == '':
        return None
    return f"https://drive.google.com/file/d/{video_id}/view"


# ──────────────────────────────────────────────────────────────────────────
# 세션 상태 초기화
# ──────────────────────────────────────────────────────────────────────────
if 'uploaded_data' not in st.session_state:
    st.session_state.uploaded_data = []  # 누적된 업로드 데이터들
if 'data_version' not in st.session_state:
    st.session_state.data_version = 0  # 데이터 갱신 추적용


def get_current_df():
    """기본 데이터 + 업로드된 데이터 모두 합쳐 반환"""
    base_df = load_base_data()
    if st.session_state.uploaded_data:
        all_dfs = [base_df] + st.session_state.uploaded_data
        return pd.concat(all_dfs, ignore_index=True)
    return base_df


# ==============================================================================
# 🎯 메뉴 1: 필터 분석
# ==============================================================================

def render_filter_analysis(df, fail_df):
    plt.rcParams['font.family'] = font_name

    # 상단 KPI
    total_tests_all = len(df)
    total_fails_all = len(fail_df)
    fail_rate_all = total_fails_all / total_tests_all * 100

    st.markdown("### 📊 전체 데이터 현황")
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown("##### 📋 전체 테스트")
        st.markdown(f"### :blue[**{total_tests_all:,}건**]")
    with k2:
        st.markdown("##### ❌ 전체 Fail")
        st.markdown(f"### :red[**{total_fails_all:,}건**]")
    with k3:
        st.markdown("##### 📈 전체 Fail율")
        st.markdown(f"### :orange[**{fail_rate_all:.1f}%**]")
    with k4:
        st.markdown("##### 💻 분석 모델")
        st.markdown(f"### :green[**{df['Model'].nunique()}개**]")

    st.markdown("---")

    # 필터 UI (다중 선택 가능)
    st.markdown("### 🎯 필터 선택")
    st.caption("각 필터에서 여러 값을 동시에 선택할 수 있습니다. (비어두면 전체)")

    row1_col1, row1_col2, row1_col3 = st.columns(3)
    with row1_col1:
        ic_filter = st.multiselect(
            "🔌 IC", 
            options=sorted(df['IC'].unique().tolist()),
            placeholder="전체 IC (비어두면 모두 선택)"
        )
    with row1_col2:
        model_filter = st.multiselect(
            "💻 Model",
            options=sorted(df['Model'].unique().tolist()),
            placeholder="전체 모델 (비어두면 모두 선택)"
        )
    with row1_col3:
        build_filter = st.multiselect(
            "🔄 Build",
            options=sorted(df['Build_Num'].unique().tolist()),
            placeholder="전체 빌드 (비어두면 모두 선택)"
        )
    
    row2_col1, row2_col2, row2_col3 = st.columns(3)
    with row2_col1:
        test_item_filter = st.multiselect(
            "🧪 Test_Item",
            options=sorted(df['Test_Item'].unique().tolist()),
            placeholder="전체 Test_Item (비어두면 모두 선택)"
        )
    with row2_col2:
        kw_filter = st.multiselect(
            "🐛 Fail_Type",
            options=KEYWORDS_ALL,
            placeholder="전체 Fail_Type (비어두면 모두 선택)"
        )
    with row2_col3:
        st.write("")

    # 필터 적용 (multiselect는 리스트)
    filtered_all = df.copy()
    filtered_fail = fail_df.copy()
    active_filters = []
    if ic_filter:
        filtered_all = filtered_all[filtered_all['IC'].isin(ic_filter)]
        filtered_fail = filtered_fail[filtered_fail['IC'].isin(ic_filter)]
        active_filters.append(f"IC={','.join(ic_filter)}")
    if model_filter:
        filtered_all = filtered_all[filtered_all['Model'].isin(model_filter)]
        filtered_fail = filtered_fail[filtered_fail['Model'].isin(model_filter)]
        if len(model_filter) <= 3:
            active_filters.append(f"Model={','.join(model_filter)}")
        else:
            active_filters.append(f"Model={len(model_filter)}개")
    if build_filter:
        filtered_all = filtered_all[filtered_all['Build_Num'].isin(build_filter)]
        filtered_fail = filtered_fail[filtered_fail['Build_Num'].isin(build_filter)]
        if len(build_filter) <= 3:
            active_filters.append(f"Build={','.join(map(str, build_filter))}")
        else:
            active_filters.append(f"Build={len(build_filter)}개")
    if test_item_filter:
        filtered_all = filtered_all[filtered_all['Test_Item'].isin(test_item_filter)]
        filtered_fail = filtered_fail[filtered_fail['Test_Item'].isin(test_item_filter)]
        if len(test_item_filter) <= 3:
            active_filters.append(f"Test_Item={','.join(test_item_filter)}")
        else:
            active_filters.append(f"Test_Item={len(test_item_filter)}개")
    if kw_filter:
        filtered_fail = filtered_fail[filtered_fail['Keyword'].isin(kw_filter)]
        if len(kw_filter) <= 3:
            active_filters.append(f"Fail_Type={','.join(kw_filter)}")
        else:
            active_filters.append(f"Fail_Type={len(kw_filter)}개")

    st.markdown("---")
    if active_filters:
        st.markdown(f"### 🎯 적용된 필터: `{' · '.join(active_filters)}`")
    else:
        st.markdown("### 🎯 전체 데이터 분석 (필터 미적용)")

    if len(filtered_fail) == 0:
        st.warning("⚠️ 선택한 조건에 해당하는 Fail 케이스가 없습니다.")
        return

    # 필터 결과 KPI
    st.markdown("### 📊 ① 필터 결과 요약")
    total_tests = len(filtered_all)
    total_fails = len(filtered_fail)
    fail_rate = (total_fails / total_tests * 100) if total_tests > 0 else 0
    n_models = filtered_fail['Model'].nunique()

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown("##### 📋 전체 테스트")
        st.markdown(f"### :blue[**{total_tests:,}건**]")
    with k2:
        st.markdown("##### ❌ Fail 건수")
        st.markdown(f"### :red[**{total_fails:,}건**]")
    with k3:
        st.markdown("##### 📈 Fail율")
        st.markdown(f"### :orange[**{fail_rate:.1f}%**]")
    with k4:
        st.markdown("##### 💻 영향 모델")
        st.markdown(f"### :green[**{n_models}개**]")

    st.markdown("---")

    # Fail_Type 순위
    if not kw_filter:
        st.markdown("### 🏆 ② Fail_Type 순위 (TOP 5)")
        kw_rank = filtered_fail['Keyword'].value_counts().head(5).reset_index()
        kw_rank.columns = ['Fail_Type', '건수']
        kw_rank['비중'] = (kw_rank['건수'] / kw_rank['건수'].sum() * 100).round(1)
        kw_rank['검토영역'] = kw_rank['Fail_Type'].apply(
            lambda k: KEYWORD_DOMAIN_MAP.get(k, {}).get('area', '?'))
        kw_rank.insert(0, '순위', range(1, len(kw_rank) + 1))

        col_left, col_right = st.columns([1, 1])
        with col_left:
            fig, ax = plt.subplots(figsize=(6, 5))
            colors = plt.cm.Set3(np.linspace(0, 1, len(kw_rank)))
            ax.pie(kw_rank['건수'], labels=kw_rank['Fail_Type'],
                   autopct='%1.1f%%', wedgeprops=dict(width=0.4),
                   startangle=90, colors=colors,
                   textprops={'fontsize': 10})
            ax.set_title(f'Fail_Type 분포 (총 {len(filtered_fail)}건)',
                         fontsize=12, fontweight='bold')
            plt.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
        with col_right:
            st.dataframe(kw_rank, use_container_width=True, hide_index=True)
        st.markdown("---")

    # 빌드별 추이
    if not build_filter:
        st.markdown("### 📈 ③ 빌드별 Fail 추이")
        build_trend = filtered_fail.groupby('Build_Num').size().reset_index(name='count')
        build_trend = build_trend.sort_values('Build_Num')
        if len(build_trend) > 1:
            fig, ax = plt.subplots(figsize=(12, 4))
            ax.plot(build_trend['Build_Num'], build_trend['count'],
                    marker='o', markersize=8, linewidth=2, color='#1C7293')
            for _, row in build_trend.iterrows():
                ax.annotate(f"{row['count']}", (row['Build_Num'], row['count']),
                            textcoords="offset points", xytext=(0, 10),
                            ha='center', fontsize=10, fontweight='bold')
            ax.set_xlabel('빌드 번호')
            ax.set_ylabel('Fail 건수')
            ax.grid(True, alpha=0.3)
            ax.set_xticks(build_trend['Build_Num'])
            ax.set_xticklabels(build_trend['Build_Num'], rotation=45)
            plt.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
        st.markdown("---")

    # 모델별 순위
    if not model_filter:
        st.markdown("### 🔥 ④ 모델별 Fail 순위 (TOP 10)")
        model_rank = filtered_fail['Model'].value_counts().head(10).reset_index()
        model_rank.columns = ['Model', 'Fail 건수']
        if len(model_rank) > 0:
            fig, ax = plt.subplots(figsize=(10, max(4, len(model_rank) * 0.4)))
            bars = ax.barh(model_rank['Model'], model_rank['Fail 건수'],
                           color=plt.cm.YlOrRd(np.linspace(0.4, 0.9, len(model_rank))))
            for bar, cnt in zip(bars, model_rank['Fail 건수']):
                ax.text(bar.get_width() + max(model_rank['Fail 건수']) * 0.01,
                        bar.get_y() + bar.get_height()/2,
                        f'{cnt}건', va='center', fontsize=10, fontweight='bold')
            ax.set_xlabel('Fail 건수')
            ax.invert_yaxis()
            plt.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
        st.markdown("---")

    # Test_Item별 순위
    if not test_item_filter:
        st.markdown("### 🧪 ⑤ Test_Item별 Fail 순위 (TOP 10)")
        item_rank = filtered_fail['Test_Item'].value_counts().head(10).reset_index()
        item_rank.columns = ['Test_Item', 'Fail 건수']
        if len(item_rank) > 0:
            fig, ax = plt.subplots(figsize=(10, max(4, len(item_rank) * 0.4)))
            bars = ax.barh(item_rank['Test_Item'], item_rank['Fail 건수'],
                           color=plt.cm.Blues(np.linspace(0.4, 0.9, len(item_rank))))
            for bar, cnt in zip(bars, item_rank['Fail 건수']):
                ax.text(bar.get_width() + max(item_rank['Fail 건수']) * 0.01,
                        bar.get_y() + bar.get_height()/2,
                        f'{cnt}건', va='center', fontsize=10, fontweight='bold')
            ax.set_xlabel('Fail 건수')
            ax.invert_yaxis()
            plt.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
        st.markdown("---")

    # 상세 표
    st.markdown(f"### 📋 ⑥ 상세 Fail 케이스 ({len(filtered_fail):,}건)")
    st.caption("표의 영상 링크를 클릭하면 새 탭에서 결함 영상이 재생됩니다.")

    display_df = filtered_fail[[
        'No', 'IC', 'Model', 'Keyword', 'Build_Num',
        'Customer', 'Category', 'Test_Item', 'video_id'
    ]].copy()
    display_df['영상'] = display_df['video_id'].apply(get_video_link)
    display_df = display_df.drop(columns=['video_id'])
    # Fail_Type 건수 기준 정렬 (많은 것이 위로)
    kw_counts = filtered_fail['Keyword'].value_counts().to_dict()
    display_df['_sort'] = display_df['Keyword'].map(kw_counts).fillna(0)
    display_df = display_df.sort_values('_sort', ascending=False).drop(columns=['_sort'])

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "영상": st.column_config.LinkColumn("영상 보기", display_text="🎬 재생"),
            "Keyword": st.column_config.TextColumn("Fail_Type"),
        }
    )


# ==============================================================================
# 📊 메뉴 2: 전체 인사이트 (생략 - 기존과 동일)
# ==============================================================================
def render_full_insights(df, fail_df):
    plt.rcParams['font.family'] = font_name

    # 빌드별 ∩ 곡선
    st.markdown("### 1️⃣ 빌드별 ∩(역U) 안정화 곡선")
    build_summary = df.groupby('Build_Num').agg(
        total=('Result', 'count'),
        fail=('Result', lambda x: (x == 'FAIL').sum())
    ).reset_index()
    build_summary['fail_rate'] = (build_summary['fail'] / build_summary['total'] * 100).round(2)
    build_summary = build_summary.sort_values('Build_Num').reset_index(drop=True)

    fig, ax = plt.subplots(figsize=(14, 5))
    ax.plot(build_summary['Build_Num'], build_summary['fail_rate'],
            marker='o', markersize=10, linewidth=2, color='#e74c3c')
    for i, row in build_summary.iterrows():
        if i == 0:
            offset_y, va = 15, 'bottom'
        else:
            prev_rate = build_summary.loc[i-1, 'fail_rate']
            offset_y = -20 if row['fail_rate'] < prev_rate else 15
            va = 'top' if row['fail_rate'] < prev_rate else 'bottom'
        ax.annotate(f"{row['fail_rate']}%",
                    (row['Build_Num'], row['fail_rate']),
                    textcoords="offset points", xytext=(0, offset_y),
                    ha='center', va=va, fontsize=11, fontweight='bold',
                    bbox=dict(boxstyle='round,pad=0.3', facecolor='white',
                              edgecolor='#cccccc', alpha=0.9))
    ax.set_title('빌드별 전체 Fail율 추이', fontsize=14, fontweight='bold')
    ax.set_xlabel('빌드 번호')
    ax.set_ylabel('Fail율 (%)')
    ax.grid(True, alpha=0.3)
    ax.set_xticks(build_summary['Build_Num'])
    ax.set_xticklabels(build_summary['Build_Num'], rotation=45)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)
    st.markdown("---")

    # IC별 + Donut
    st.markdown("### 2️⃣ IC별 Fail율 + Fail_Type 분포")
    ic_summary = df.groupby('IC').agg(
        total=('Result', 'count'),
        fail=('Result', lambda x: (x == 'FAIL').sum())
    ).reset_index()
    ic_summary['fail_rate'] = (ic_summary['fail'] / ic_summary['total'] * 100).round(2)
    ic_summary = ic_summary.sort_values('fail_rate', ascending=False)

    fig1, ax1 = plt.subplots(figsize=(8, 4))
    bars = ax1.bar(ic_summary['IC'], ic_summary['fail_rate'],
                   color=['#e74c3c', '#f39c12', '#3498db'])
    ax1.set_title('IC별 Fail율', fontsize=14, fontweight='bold')
    ax1.set_ylabel('Fail율 (%)')
    ax1.set_ylim(0, max(ic_summary['fail_rate']) * 1.2)
    for bar, rate in zip(bars, ic_summary['fail_rate']):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3,
                 f'{rate}%', ha='center', fontsize=12, fontweight='bold')
    plt.tight_layout()
    st.pyplot(fig1)
    plt.close(fig1)

    ic_keyword = fail_df.groupby(['IC', 'Keyword']).size().reset_index(name='count')
    colors_kw = plt.cm.Set3(np.linspace(0, 1, 8))
    color_dict = dict(zip(KEYWORDS_ALL, colors_kw))

    fig2, axes2 = plt.subplots(1, 3, figsize=(14, 5))
    fig2.suptitle('IC별 Fail_Type 분포', fontsize=14, fontweight='bold')
    for ax, ic in zip(axes2, ['G7500', 'GT1T0A', 'GT9XS']):
        ic_data = ic_keyword[ic_keyword['IC'] == ic].sort_values('count', ascending=False)
        total_cnt = ic_data['count'].sum()
        colors_this = [color_dict[kw] for kw in ic_data['Keyword']]
        ax.pie(ic_data['count'], labels=None,
               autopct=lambda pct: f'{pct:.1f}%' if pct >= 5 else '',
               wedgeprops=dict(width=0.4), startangle=90, colors=colors_this,
               textprops={'fontsize': 10, 'fontweight': 'bold'})
        ax.text(0, 0.1, ic, ha='center', va='center', fontsize=14, fontweight='bold')
        ax.text(0, -0.15, f'Fail {total_cnt}건', ha='center', va='center',
                fontsize=10, color='#666666')
    legend_patches = [plt.Rectangle((0,0),1,1, color=color_dict[kw]) for kw in KEYWORDS_ALL]
    fig2.legend(legend_patches, KEYWORDS_ALL, loc='lower center', ncol=4,
                fontsize=9, bbox_to_anchor=(0.5, -0.05))
    plt.tight_layout()
    st.pyplot(fig2)
    plt.close(fig2)
    st.markdown("---")

    # 카테고리 히트맵
    st.markdown("### 3️⃣ 카테고리 × Fail_Type 히트맵")
    pivot_cat = fail_df.pivot_table(
        index='Category', columns='Keyword', values='No', aggfunc='count', fill_value=0)
    pivot_cat = pivot_cat[[k for k in KEYWORDS_ALL if k in pivot_cat.columns]]
    fig, ax = plt.subplots(figsize=(12, 6))
    sns.heatmap(pivot_cat, annot=True, fmt='d', cmap='YlOrRd',
                cbar_kws={'label': 'Fail 건수'}, ax=ax,
                linewidths=0.5, linecolor='white')
    ax.set_title('카테고리 × Fail_Type 히트맵', fontsize=14, fontweight='bold', pad=15)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)
    st.markdown("---")

    # 모델 히트맵
    st.markdown("### 4️⃣ 모델 × Fail_Type 히트맵 ⭐")
    pivot_model = fail_df.pivot_table(
        index='Model', columns='Keyword', values='No', aggfunc='count', fill_value=0)
    pivot_model = pivot_model.loc[pivot_model.sum(axis=1).sort_values(ascending=False).index]
    pivot_model = pivot_model[[k for k in KEYWORDS_ALL if k in pivot_model.columns]]
    fig, ax = plt.subplots(figsize=(12, max(8, len(pivot_model) * 0.3)))
    sns.heatmap(pivot_model, annot=True, fmt='d', cmap='YlOrRd',
                cbar_kws={'label': 'Fail 건수'}, ax=ax,
                linewidths=0.5, linecolor='white')
    ax.set_title('모델 × Fail_Type 히트맵', fontsize=14, fontweight='bold', pad=15)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)
    st.markdown("---")

    # 발생 횟수 × 영향 모델 수 버블 차트
    st.markdown("### 5️⃣ Fail_Type 영향 분석 (발생 횟수 × 영향 모델 수)")
    st.caption("오른쪽 위에 있을수록 자주 발생하고 많은 모델에 영향 → 우선 검토 대상")
    
    # 각 Fail_Type별 통계
    bubble_data = []
    for kw in KEYWORDS_ALL:
        kw_fails = fail_df[fail_df['Keyword'] == kw]
        if len(kw_fails) > 0:
            bubble_data.append({
                'keyword': kw,
                'count': len(kw_fails),
                'n_models': kw_fails['Model'].nunique()
            })
    
    if bubble_data:
        max_count = max(b['count'] for b in bubble_data)
        max_models = max(b['n_models'] for b in bubble_data)
        
        fig, ax = plt.subplots(figsize=(12, 7))
        colors = plt.cm.YlOrRd([b['count']/max_count for b in bubble_data])
        
        for i, b in enumerate(bubble_data):
            size = 300 + (b['count'] / max_count) * 2700
            ax.scatter(b['count'], b['n_models'], s=size, c=[colors[i]],
                       alpha=0.65, edgecolors='black', linewidths=1.5, zorder=3)
        
        # 라벨 (adjust_text 없이 수동으로 위치 분산)
        # y값 기준 정렬해서 가까운 것끼리 위치 어긋나게
        # 라벨 위치 분산 (각 버블마다 다른 방향)
        sorted_bubbles = sorted(bubble_data, key=lambda b: (b['n_models'], b['count']))
        # 8개 방향 분산
        directions = [
            (25, 25), (25, -25), (-25, 25), (-25, -25),
            (35, 5), (-35, 5), (5, 35), (5, -35)
        ]
        for i, b in enumerate(sorted_bubbles):
            offset_x, offset_y = directions[i % len(directions)]
            ha = 'left' if offset_x > 0 else 'right' 
            
            ax.annotate(
                f"{b['keyword']}\n({b['count']}건, {b['n_models']}개 모델)",
                (b['count'], b['n_models']),
                xytext=(offset_x, offset_y),
                textcoords='offset points',
                fontsize=9, fontweight='bold', ha=ha,
                bbox=dict(boxstyle='round,pad=0.3', facecolor='white',
                          edgecolor='#cccccc', alpha=0.9),
                zorder=4,
                arrowprops=dict(arrowstyle='-', color='#888888', lw=0.5, alpha=0.5)
            )
        
        # 평균선
        avg_count = sum(b['count'] for b in bubble_data) / len(bubble_data)
        avg_models = sum(b['n_models'] for b in bubble_data) / len(bubble_data)
        ax.axhline(y=avg_models, color='gray', linestyle='--', alpha=0.4,
                   label=f'평균 영향 모델 ({avg_models:.0f}개)')
        ax.axvline(x=avg_count, color='gray', linestyle='--', alpha=0.4,
                   label=f'평균 발생 횟수 ({avg_count:.0f}건)')
        
        ax.set_xlabel('발생 횟수 (Fail 건수)', fontsize=12, fontweight='bold')
        ax.set_ylabel('영향 모델 수', fontsize=12, fontweight='bold')
        ax.set_title('Fail_Type 영향 분석',
                     fontsize=14, fontweight='bold', pad=15)
        ax.set_xlim(0, max_count * 1.3)
        ax.set_ylim(0, max_models * 1.3)
        ax.grid(True, alpha=0.3)
        ax.legend(loc='lower right', fontsize=9)
        
        # 사분면 안내
        ax.text(max_count * 1.25, max_models * 1.25, '★ 자주 + 광범위\n우선 검토',
                fontsize=10, ha='right', va='top', color='#c0392b', fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.4', facecolor='#FFF5F5', edgecolor='#e74c3c'))
        
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)
        
        # 표
        bubble_df = pd.DataFrame(bubble_data).sort_values('count', ascending=False).reset_index(drop=True)
        bubble_df.insert(0, '순위', range(1, len(bubble_df) + 1))
        bubble_df.columns = ['순위', 'Fail_Type', '발생 횟수', '영향 모델 수']
        st.dataframe(bubble_df, use_container_width=True, hide_index=True)
        
        st.info("💡 **활용**: 자주 발생하면서 많은 모델에 영향을 주는 Fail_Type을 우선 검토하세요. 오른쪽 위 사분면이 가장 시급한 영역입니다.")


# ==============================================================================
# 📊 메뉴 추가: 전문 통계 분석 (카이제곱 + 신뢰구간 + 시계열 예측)
# ==============================================================================
def render_advanced_statistics(df, fail_df):
    """전문 분석가 시선의 통계 분석"""
    plt.rcParams['font.family'] = font_name
    
    st.markdown("### 🎯 전문 통계 분석")
    st.caption("scipy + sklearn으로 통계적 유의성을 검증하고 미래 예측까지 수행합니다.")
    
    # ===== ① 카이제곱 검정 (Chi-square Test) =====
    st.markdown("---")
    st.markdown("### 1️⃣ 카이제곱 독립성 검정 (Chi-square Test)")
    st.caption("두 범주형 변수가 통계적으로 유의미한 관계가 있는지 검정합니다.")
    
    # 추천 분석 시나리오 콤보 (의미 있는 조합만)
    scenario = st.selectbox(
        "🎯 분석 시나리오 선택",
        options=[
            "IC × Keyword (IC별 Fail 패턴 분석) ⭐ 추천",
            "Customer × Keyword (고객사별 결함 패턴)",
            "Category × Keyword (테스트 카테고리별 결함)",
            "Test_Item × Keyword (테스트 항목별 결함)",
            "Model × Keyword (모델별 결함 패턴)",
            "IC × Result (IC별 합격률 비교)",
            "Customer × Result (고객사별 합격률 비교)",
            "Category × Result (카테고리별 합격률)",
        ],
        key='chi_scenario',
        help="의미 있는 분석 조합을 미리 준비했습니다."
    )
    
    # 시나리오 → 변수 자동 매핑
    scenario_map = {
        "IC × Keyword (IC별 Fail 패턴 분석) ⭐ 추천": ('IC', 'Keyword'),
        "Customer × Keyword (고객사별 결함 패턴)": ('Customer', 'Keyword'),
        "Category × Keyword (테스트 카테고리별 결함)": ('Category', 'Keyword'),
        "Test_Item × Keyword (테스트 항목별 결함)": ('Test_Item', 'Keyword'),
        "Model × Keyword (모델별 결함 패턴)": ('Model', 'Keyword'),
        "IC × Result (IC별 합격률 비교)": ('IC', 'Result'),
        "Customer × Result (고객사별 합격률 비교)": ('Customer', 'Result'),
        "Category × Result (카테고리별 합격률)": ('Category', 'Result'),
    }
    
    cat_var1, cat_var2 = scenario_map[scenario]
    
    # 변수 정보 표시
    col_info1, col_info2 = st.columns(2)
    with col_info1:
        st.info(f"📊 **변수 1**: `{cat_var1}`")
    with col_info2:
        st.info(f"📊 **변수 2**: `{cat_var2}`")
    
    if cat_var1 != cat_var2:
        try:
            # 분석용 데이터
            analysis_df = fail_df if cat_var2 == 'Keyword' else df
            analysis_df = analysis_df[analysis_df[cat_var2] != '']
            
            contingency = pd.crosstab(analysis_df[cat_var1], analysis_df[cat_var2])
            
            if contingency.size > 0:
                chi2, p_value, dof, expected = chi2_contingency(contingency)
                
                # Cramér's V (효과 크기)
                n = contingency.sum().sum()
                cramers_v = np.sqrt(chi2 / (n * (min(contingency.shape) - 1)))
                
                # 결과 표시
                k1, k2, k3, k4 = st.columns(4)
                with k1:
                    st.metric("χ² (카이제곱)", f"{chi2:.2f}")
                with k2:
                    st.metric("자유도 (df)", f"{dof}")
                with k3:
                    if p_value < 0.001:
                        st.metric("p-value", "< 0.001", delta="유의함", delta_color="normal")
                    elif p_value < 0.05:
                        st.metric("p-value", f"{p_value:.4f}", delta="유의함", delta_color="normal")
                    else:
                        st.metric("p-value", f"{p_value:.4f}", delta="유의 X", delta_color="inverse")
                with k4:
                    # Cramér's V 해석
                    if cramers_v < 0.1:
                        strength = "매우 약함"
                    elif cramers_v < 0.3:
                        strength = "약함"
                    elif cramers_v < 0.5:
                        strength = "중간"
                    else:
                        strength = "강함"
                    st.metric("Cramér's V", f"{cramers_v:.3f}", delta=strength)
                
                # 해석
                if p_value < 0.05:
                    st.success(
                        f"✅ **통계적으로 유의미한 관계** (p={p_value:.4f} < 0.05)\n\n"
                        f"`{cat_var1}`과(와) `{cat_var2}` 사이에 우연이 아닌 관계가 존재합니다. "
                        f"효과 크기는 **{strength}** 수준입니다 (Cramér's V = {cramers_v:.3f})."
                    )
                else:
                    st.warning(
                        f"⚠️ **통계적으로 유의미하지 않음** (p={p_value:.4f} ≥ 0.05)\n\n"
                        f"두 변수 사이의 관계가 우연으로 설명될 수 있습니다."
                    )
                
                # 관측값 vs 기대값 비교 히트맵
                st.markdown("##### 📊 관측값 vs 기대값 차이 (잔차 분석)")
                residuals = (contingency.values - expected) / np.sqrt(expected)
                resid_df = pd.DataFrame(residuals, index=contingency.index, columns=contingency.columns)
                
                fig, ax = plt.subplots(figsize=(12, max(4, len(resid_df) * 0.5)))
                sns.heatmap(resid_df, annot=True, fmt='.2f', cmap='RdBu_r', center=0,
                            cbar_kws={'label': '표준화 잔차'}, ax=ax,
                            linewidths=0.5, linecolor='white')
                ax.set_title(f'{cat_var1} × {cat_var2} 표준화 잔차 (|값| > 2: 유의)',
                             fontsize=13, fontweight='bold')
                plt.tight_layout()
                st.pyplot(fig)
                plt.close(fig)
                
                st.info(
                    "💡 **해석 가이드**:\n"
                    "- 빨간색 (양수) = 기대보다 많이 발생\n"
                    "- 파란색 (음수) = 기대보다 적게 발생\n"
                    "- |값| > 2 = 통계적으로 의미 있는 차이"
                )
        except Exception as e:
            st.error(f"검정 중 오류: {e}")
    else:
        st.warning("두 변수를 다르게 선택해주세요.")
    
    # ===== ② 신뢰구간 분석 (Confidence Interval) =====
    st.markdown("---")
    st.markdown("### 2️⃣ Fail율 95% 신뢰구간 (Confidence Interval)")
    st.caption("표본의 Fail율로 모집단의 진짜 Fail율 범위를 추정합니다.")
    
    # 전체 Fail율 신뢰구간
    n_total = len(df)
    n_fail = len(fail_df)
    p_hat = n_fail / n_total
    
    # 윌슨 신뢰구간 (Wilson score interval) - 비율에 적합
    z = 1.96  # 95% 신뢰수준
    denom = 1 + z**2/n_total
    center = (p_hat + z**2/(2*n_total)) / denom
    margin = z * np.sqrt(p_hat*(1-p_hat)/n_total + z**2/(4*n_total**2)) / denom
    ci_low = max(0, center - margin)
    ci_high = min(1, center + margin)
    
    # 표시
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("전체 Fail율 (점추정)", f"{p_hat*100:.2f}%")
    with col2:
        st.metric("95% 신뢰구간 하한", f"{ci_low*100:.2f}%")
    with col3:
        st.metric("95% 신뢰구간 상한", f"{ci_high*100:.2f}%")
    
    st.success(
        f"✅ **해석**: 95% 신뢰수준에서 모집단의 진짜 Fail율은 "
        f"**[{ci_low*100:.2f}%, {ci_high*100:.2f}%]** 범위 안에 있습니다.\n\n"
        f"표본 크기 n={n_total:,}로 충분히 크므로 신뢰구간이 좁고(±{margin*100:.2f}%p) 신뢰도가 높습니다."
    )
    
    # IC별 신뢰구간
    st.markdown("##### 📊 IC별 Fail율 + 95% 신뢰구간")
    
    ic_ci = []
    for ic in df['IC'].unique():
        ic_df = df[df['IC'] == ic]
        n_ic = len(ic_df)
        n_ic_fail = (ic_df['Result'] == 'FAIL').sum()
        p_ic = n_ic_fail / n_ic if n_ic > 0 else 0
        
        # 윌슨 CI
        denom = 1 + z**2/n_ic
        center = (p_ic + z**2/(2*n_ic)) / denom
        margin = z * np.sqrt(p_ic*(1-p_ic)/n_ic + z**2/(4*n_ic**2)) / denom
        ci_low_ic = max(0, center - margin)
        ci_high_ic = min(1, center + margin)
        
        ic_ci.append({
            'IC': ic,
            '표본수': n_ic,
            'Fail': n_ic_fail,
            'Fail율(%)': round(p_ic*100, 2),
            'CI하한(%)': round(ci_low_ic*100, 2),
            'CI상한(%)': round(ci_high_ic*100, 2),
            '오차범위(±%p)': round(margin*100, 2)
        })
    
    ic_ci_df = pd.DataFrame(ic_ci).sort_values('Fail율(%)', ascending=False)
    st.dataframe(ic_ci_df, use_container_width=True, hide_index=True)
    
    # 신뢰구간 시각화
    fig, ax = plt.subplots(figsize=(10, max(4, len(ic_ci_df) * 0.6)))
    y_pos = range(len(ic_ci_df))
    
    for i, row in ic_ci_df.iterrows():
        # 신뢰구간 (가로 막대)
        ax.plot([row['CI하한(%)'], row['CI상한(%)']], 
                [y_pos[list(ic_ci_df.index).index(i)], y_pos[list(ic_ci_df.index).index(i)]],
                color='#3498db', linewidth=4, alpha=0.5)
        # 점추정값
        ax.scatter(row['Fail율(%)'], y_pos[list(ic_ci_df.index).index(i)],
                  s=150, color='#e74c3c', zorder=5, edgecolor='black', linewidth=1.5)
        # 라벨
        ax.text(row['CI상한(%)'] + 0.5, y_pos[list(ic_ci_df.index).index(i)],
               f"{row['Fail율(%)']}% [{row['CI하한(%)']:.1f}, {row['CI상한(%)']:.1f}]",
               va='center', fontsize=10)
    
    ax.set_yticks(list(y_pos))
    ax.set_yticklabels(ic_ci_df['IC'].tolist())
    ax.set_xlabel('Fail율 (%)')
    ax.set_title('IC별 Fail율 + 95% 신뢰구간', fontsize=13, fontweight='bold')
    ax.grid(True, alpha=0.3, axis='x')
    ax.set_xlim(0, max(ic_ci_df['CI상한(%)']) * 1.3)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)
    
    # ===== ③ 시계열 예측 (Linear Regression) =====
    st.markdown("---")
    st.markdown("### 3️⃣ 다음 빌드 Fail율 예측 (시계열 회귀)")
    st.caption("과거 빌드별 Fail율을 학습해서 다음 빌드의 Fail율을 예측합니다.")
    
    # 빌드별 Fail율 추출
    build_summary = df.groupby('Build_Num').agg(
        total=('Result', 'count'),
        fail=('Result', lambda x: (x == 'FAIL').sum())
    ).reset_index()
    build_summary['fail_rate'] = (build_summary['fail'] / build_summary['total'] * 100).round(2)
    build_summary = build_summary.sort_values('Build_Num').reset_index(drop=True)
    
    if len(build_summary) >= 3:
        # 모델 선택
        col1, col2 = st.columns(2)
        with col1:
            n_recent = st.slider(
                "최근 몇 개 빌드로 학습?", 
                min_value=3, 
                max_value=len(build_summary),
                value=min(5, len(build_summary)),
                help="최근 데이터일수록 미래 예측에 유리합니다."
            )
        with col2:
            next_build_input = st.number_input(
                "예측할 빌드 번호",
                min_value=int(build_summary['Build_Num'].max()) + 1,
                value=int(build_summary['Build_Num'].max()) + 100,
                step=10
            )
        
        # 학습 데이터
        train_data = build_summary.tail(n_recent).copy()
        X = train_data[['Build_Num']].values
        y = train_data['fail_rate'].values
        
        # 선형 회귀
        model = LinearRegression()
        model.fit(X, y)
        
        # 예측
        next_build_pred = model.predict([[next_build_input]])[0]
        
        # R² 점수
        r2 = model.score(X, y)
        
        # 예측 결과
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("예측 Fail율", f"{next_build_pred:.2f}%")
        with col2:
            current = build_summary['fail_rate'].iloc[-1]
            delta = next_build_pred - current
            st.metric("최신 빌드 대비", f"{delta:+.2f}%p",
                      delta=f"{'증가' if delta > 0 else '감소'}",
                      delta_color="inverse" if delta > 0 else "normal")
        with col3:
            st.metric("모델 적합도 (R²)", f"{r2:.3f}",
                      delta=f"{'좋음' if r2 > 0.5 else '낮음'}",
                      delta_color="normal" if r2 > 0.5 else "off")
        
        # 추세 차트
        fig, ax = plt.subplots(figsize=(13, 6))
        
        # 전체 빌드 (회색 점)
        not_used = build_summary[~build_summary['Build_Num'].isin(train_data['Build_Num'])]
        if len(not_used) > 0:
            ax.scatter(not_used['Build_Num'], not_used['fail_rate'],
                       s=120, c='lightgray', edgecolors='gray', linewidth=1.5,
                       zorder=3, label='과거 빌드 (학습 미사용)')
        
        # 학습에 사용된 빌드 (파란 점)
        ax.scatter(train_data['Build_Num'], train_data['fail_rate'],
                   s=120, c='#3498db', edgecolors='black', linewidth=1.5,
                   zorder=4, label=f'학습 데이터 (최근 {n_recent}개)')
        
        # 회귀선 (실제 학습 범위)
        X_line = np.linspace(train_data['Build_Num'].min(), 
                              train_data['Build_Num'].max(), 100).reshape(-1, 1)
        y_line = model.predict(X_line)
        ax.plot(X_line.flatten(), y_line, '-', color='#3498db', 
                linewidth=2.5, alpha=0.7, label='회귀선 (학습 구간)')
        
        # 예측선 (미래)
        X_pred = np.linspace(train_data['Build_Num'].max(), 
                              next_build_input, 100).reshape(-1, 1)
        y_pred = model.predict(X_pred)
        ax.plot(X_pred.flatten(), y_pred, '--', color='#e74c3c', 
                linewidth=2.5, label='예측선 (미래)')
        
        # 예측점 (별표)
        ax.scatter([next_build_input], [next_build_pred], 
                   s=400, c='#e74c3c', marker='*', edgecolors='black', linewidth=2,
                   zorder=10, label=f'예측: 빌드 {next_build_input} = {next_build_pred:.2f}%')
        
        ax.annotate(f'예측 {next_build_pred:.2f}%',
                    (next_build_input, next_build_pred),
                    xytext=(15, 15), textcoords='offset points',
                    fontsize=12, fontweight='bold', color='#e74c3c',
                    bbox=dict(boxstyle='round,pad=0.4', facecolor='white',
                              edgecolor='#e74c3c', linewidth=2))
        
        ax.set_xlabel('빌드 번호', fontsize=12, fontweight='bold')
        ax.set_ylabel('Fail율 (%)', fontsize=12, fontweight='bold')
        ax.set_title(f'다음 빌드 Fail율 예측 (선형 회귀, 최근 {n_recent}개 학습)',
                     fontsize=13, fontweight='bold', pad=15)
        ax.grid(True, alpha=0.3)
        ax.legend(loc='upper right', fontsize=10)
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)
        
        # 해석
        if r2 > 0.7:
            interp = "강한 추세"
            color = "✅"
        elif r2 > 0.4:
            interp = "중간 추세"
            color = "📊"
        else:
            interp = "약한 추세 (예측 신뢰도 낮음)"
            color = "⚠️"
        
        st.info(
            f"{color} **모델 해석**:\n"
            f"- 최근 {n_recent}개 빌드의 Fail율 추이는 **{interp}** (R² = {r2:.3f})\n"
            f"- 다음 빌드 {next_build_input}의 예측 Fail율: **{next_build_pred:.2f}%**\n"
            f"- 최신 빌드 대비 **{'증가' if delta > 0 else '감소'}** {abs(delta):.2f}%p 예상\n\n"
            f"💡 R²가 1에 가까울수록 모델이 데이터를 잘 설명합니다. "
            f"R² < 0.4면 예측을 참고로만 활용하세요."
        )
    else:
        st.warning("빌드가 3개 이상 필요합니다.")




# ==============================================================================
# 🚨 메뉴 3: 회귀 알람
# ==============================================================================
def render_regression_alert(df, fail_df):
    builds = sorted(df['Build_Num'].unique())
    latest_build = builds[-1]
    prev_build = builds[-2] if len(builds) > 1 else builds[-1]

    latest_fail = fail_df[fail_df['Build_Num'] == latest_build]
    prev_fail = fail_df[fail_df['Build_Num'] == prev_build]

    latest_combo = latest_fail.groupby(['Model', 'Keyword']).size().reset_index(name='latest')
    prev_combo = prev_fail.groupby(['Model', 'Keyword']).size().reset_index(name='prev')

    merged = latest_combo.merge(prev_combo, on=['Model', 'Keyword'], how='left')
    merged['prev'] = merged['prev'].fillna(0).astype(int)
    merged['상태'] = merged.apply(
        lambda r: 'NEW' if r['prev'] == 0 else ('WORSE' if r['latest'] > r['prev'] else ''),
        axis=1)
    alerts = merged[merged['상태'].isin(['NEW', 'WORSE'])].copy()
    alerts['검토영역'] = alerts['Keyword'].apply(
        lambda k: KEYWORD_DOMAIN_MAP.get(k, {}).get('area', '?'))

    # Fail 건수 기준 정렬 (많은 게 위로)
    alerts = alerts.sort_values('latest', ascending=False)

    n_new = (alerts['상태'] == 'NEW').sum()
    n_worse = (alerts['상태'] == 'WORSE').sum()
    total_fail_increase = (alerts['latest'] - alerts['prev']).sum()

    st.markdown(f"**최신 빌드 {latest_build}  vs  이전 빌드 {prev_build}** 비교")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown("##### 🚨 전체 알람")
        st.markdown(f"### :red[**{len(alerts)}건**]")
    with col2:
        st.markdown("##### 🆕 NEW (신규)")
        st.markdown(f"### :red[**{n_new}건**]")
    with col3:
        st.markdown("##### 📈 WORSE (악화)")
        st.markdown(f"### :orange[**{n_worse}건**]")
    with col4:
        st.markdown("##### ⚠️ Fail 증가")
        st.markdown(f"### :blue[**+{total_fail_increase}건**]")
    st.markdown("---")

    alert_rows = []
    for _, row in alerts.iterrows():
        sample = latest_fail[(latest_fail['Model'] == row['Model']) &
                             (latest_fail['Keyword'] == row['Keyword'])].iloc[0]
        video_link = get_video_link(sample['video_id'])
        alert_rows.append({
            '상태': row['상태'],
            'Model': row['Model'],
            'Fail_Type': row['Keyword'],
            '이전 건수': row['prev'],
            '현재 건수': row['latest'],
            '증가': int(row['latest']) - int(row['prev']),
            '검토 영역': row['검토영역'],
            '영상': video_link
        })
    alert_df = pd.DataFrame(alert_rows)

    st.dataframe(
        alert_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "영상": st.column_config.LinkColumn("영상 보기", display_text="🎬 재생")
        }
    )


# ==============================================================================
# 📈 메뉴 새로 추가: 전체 통계
# ==============================================================================
def render_full_statistics(df, fail_df):
    """전체 데이터 통계 - 월별, Customer, Test_Item 매트릭스, 메타 정보"""
    plt.rcParams['font.family'] = font_name
    
    # ===== 종합 KPI =====
    st.markdown("### 🎯 종합 KPI")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📋 전체 테스트", f"{len(df):,}건")
    with col2:
        fail_rate = len(fail_df) / len(df) * 100 if len(df) > 0 else 0
        st.metric("❌ Fail율", f"{fail_rate:.1f}%", f"{len(fail_df):,}건")
    with col3:
        st.metric("💻 모델 수", f"{df['Model'].nunique()}개")
    with col4:
        st.metric("🔄 빌드 수", f"{df['Build_Num'].nunique()}개")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("🔌 IC 종류", f"{df['IC'].nunique()}종")
    with col2:
        st.metric("🏢 고객사", f"{df['Customer'].nunique()}개")
    with col3:
        st.metric("🧪 Test_Item", f"{df['Test_Item'].nunique()}종")
    with col4:
        st.metric("📁 카테고리", f"{df['Category'].nunique()}개")
    
    st.markdown("---")
    
    # ===== 1. 월별 테스트 추이 =====
    st.markdown("### 📅 1. 월별 테스트 추이")
    st.caption("시간 흐름에 따른 테스트량과 Fail율 변화")
    
    df_dated = df.copy()
    df_dated['Test_Date'] = pd.to_datetime(df_dated['Test_Date'], errors='coerce')
    df_dated = df_dated.dropna(subset=['Test_Date'])
    df_dated['year_month'] = df_dated['Test_Date'].dt.to_period('M').astype(str)
    
    monthly = df_dated.groupby('year_month').agg(
        total=('Result', 'count'),
        fail=('Result', lambda x: (x == 'FAIL').sum())
    ).reset_index()
    monthly['fail_rate'] = (monthly['fail'] / monthly['total'] * 100).round(1)
    
    fig, ax1 = plt.subplots(figsize=(14, 5))
    
    # 막대: 테스트 건수
    bars = ax1.bar(monthly['year_month'], monthly['total'],
                    color='#3498db', alpha=0.6, label='테스트 건수')
    ax1.set_xlabel('월', fontsize=11)
    ax1.set_ylabel('테스트 건수', fontsize=11, color='#3498db')
    ax1.tick_params(axis='y', labelcolor='#3498db')
    ax1.tick_params(axis='x', rotation=45)
    
    for bar, val in zip(bars, monthly['total']):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 20,
                 f'{val}', ha='center', fontsize=9, color='#3498db')
    
    # 라인: Fail율
    ax2 = ax1.twinx()
    ax2.plot(monthly['year_month'], monthly['fail_rate'],
             marker='o', markersize=8, linewidth=2, color='#e74c3c', label='Fail율 (%)')
    ax2.set_ylabel('Fail율 (%)', fontsize=11, color='#e74c3c')
    ax2.tick_params(axis='y', labelcolor='#e74c3c')
    ax2.set_ylim(0, max(monthly['fail_rate']) * 1.3)
    
    for i, val in enumerate(monthly['fail_rate']):
        ax2.text(i, val + 1, f'{val}%', ha='center', fontsize=9, color='#e74c3c', fontweight='bold')
    
    ax1.set_title('월별 테스트 건수 + Fail율 추이', fontsize=14, fontweight='bold')
    ax1.grid(True, alpha=0.3)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)
    
    # 표
    monthly_display = monthly.copy()
    monthly_display.columns = ['월', '전체', 'Fail', 'Fail율(%)']
    st.dataframe(monthly_display, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    
    # ===== 2. Customer (고객사) 분석 =====
    st.markdown("### 🏢 2. 고객사별 분석")
    st.caption("어떤 고객사가 어떤 IC를 어떻게 테스트하는지")
    
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.markdown("##### 고객사별 테스트 건수 및 Fail율")
        cust = df.groupby('Customer').agg(
            total=('Result', 'count'),
            fail=('Result', lambda x: (x == 'FAIL').sum())
        ).reset_index()
        cust['fail_rate'] = (cust['fail'] / cust['total'] * 100).round(1)
        cust = cust.sort_values('total', ascending=False)
        
        fig, ax = plt.subplots(figsize=(8, 5))
        bars = ax.barh(cust['Customer'], cust['total'],
                       color=plt.cm.Set2(np.linspace(0, 1, len(cust))))
        for bar, t, fr in zip(bars, cust['total'], cust['fail_rate']):
            ax.text(bar.get_width() + max(cust['total']) * 0.01,
                    bar.get_y() + bar.get_height()/2,
                    f'{t}건 ({fr}%)', va='center', fontsize=9, fontweight='bold')
        ax.set_xlabel('테스트 건수')
        ax.invert_yaxis()
        ax.set_title('고객사별 테스트 분포', fontsize=12, fontweight='bold')
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)
    
    with col_right:
        st.markdown("##### IC × 고객사 매트릭스")
        ic_cust = df.groupby(['IC', 'Customer']).size().unstack(fill_value=0)
        
        fig, ax = plt.subplots(figsize=(8, 5))
        sns.heatmap(ic_cust, annot=True, fmt='d', cmap='Blues',
                    cbar_kws={'label': '테스트 건수'}, ax=ax,
                    linewidths=0.5, linecolor='white')
        ax.set_title('IC × 고객사 분포', fontsize=12, fontweight='bold')
        ax.set_xlabel('Customer')
        ax.set_ylabel('IC')
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)
    
    st.markdown("---")
    
    # ===== 3. Test_Item × Fail_Type 매트릭스 =====
    st.markdown("### 🧪 3. Test_Item × Fail_Type 매트릭스")
    st.caption("어떤 테스트 항목에서 어떤 결함이 자주 발생하는지")
    
    if len(fail_df) > 0:
        # Top 15 Test_Item × Fail_Type
        top_items = fail_df['Test_Item'].value_counts().head(15).index.tolist()
        filtered = fail_df[fail_df['Test_Item'].isin(top_items)]
        
        item_kw = filtered.pivot_table(
            index='Test_Item', columns='Keyword', values='No',
            aggfunc='count', fill_value=0
        )
        item_kw = item_kw.loc[top_items]  # 순서 유지
        item_kw = item_kw[[k for k in KEYWORDS_ALL if k in item_kw.columns]]
        
        fig, ax = plt.subplots(figsize=(14, max(6, len(item_kw) * 0.4)))
        sns.heatmap(item_kw, annot=True, fmt='d', cmap='YlOrRd',
                    cbar_kws={'label': 'Fail 건수'}, ax=ax,
                    linewidths=0.5, linecolor='white')
        ax.set_title('Test_Item (TOP 15) × Fail_Type 매트릭스',
                     fontsize=14, fontweight='bold', pad=15)
        ax.set_xlabel('Fail_Type')
        ax.set_ylabel('Test_Item')
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)
        
        st.info("💡 **활용**: 매트릭스의 진한 칸을 보면 어떤 테스트 항목이 어떤 결함에 취약한지 즉시 파악 가능합니다.")
    
    st.markdown("---")
    
    # ===== 4. 빌드 출시 정보 =====
    st.markdown("### 🔄 4. 빌드 출시 정보")
    st.caption("각 빌드별 테스트량과 평균 Fail율")
    
    build_info = df.groupby('Build_Num').agg(
        total=('Result', 'count'),
        fail=('Result', lambda x: (x == 'FAIL').sum()),
        n_models=('Model', 'nunique'),
        n_ics=('IC', 'nunique')
    ).reset_index()
    build_info['fail_rate'] = (build_info['fail'] / build_info['total'] * 100).round(1)
    build_info.columns = ['빌드', '전체 Test', 'Fail', '모델 수', 'IC 수', 'Fail율(%)']
    
    st.dataframe(build_info, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    
    # ===== 5. 데이터 메타 정보 =====
    st.markdown("### 📋 5. 데이터 메타 정보")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("##### 📅 수집 기간")
        if len(df_dated) > 0:
            min_date = df_dated['Test_Date'].min()
            max_date = df_dated['Test_Date'].max()
            duration = (max_date - min_date).days
            st.markdown(f"""
            - **시작**: {min_date.strftime('%Y-%m-%d')}
            - **종료**: {max_date.strftime('%Y-%m-%d')}
            - **기간**: 약 {duration}일 ({duration//30}개월)
            - **수집 월**: {df_dated['year_month'].nunique()}개월
            """)
    
    with col2:
        st.markdown("##### 🔌 IC 종류")
        for ic in sorted(df['IC'].unique()):
            ic_data = df[df['IC'] == ic]
            ic_fail = ic_data[ic_data['Result'] == 'FAIL']
            top_kw = ic_fail['Keyword'].value_counts().head(1)
            top_text = f"{top_kw.index[0]} ({top_kw.iloc[0]}건)" if len(top_kw) > 0 else "Fail 없음"
            st.markdown(f"- **{ic}**: {len(ic_data):,}건 · Top Fail: {top_text}")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("##### 📊 결과 분포")
        result_counts = df['Result'].value_counts()
        for r, c in result_counts.items():
            pct = c / len(df) * 100
            st.markdown(f"- **{r}**: {c:,}건 ({pct:.1f}%)")
    
    with col2:
        st.markdown("##### 🐛 주요 Fail_Type (TOP 5)")
        top_kw = fail_df['Keyword'].value_counts().head(5)
        for i, (kw, c) in enumerate(top_kw.items(), 1):
            st.markdown(f"- **{i}위 {kw}**: {c}건")




# ==============================================================================
# 📁 메뉴 4: 데이터 업로드 (선택 삭제 + 팝업 UX)
# ==============================================================================
def render_data_upload():
    st.markdown("### 📁 주간 체크리스트 업로드")
    st.caption("회사 SQA 시스템에서 받은 엑셀(.xlsx) 또는 CSV 파일을 업로드하면 자동으로 분석 데이터에 추가됩니다.")
    
    # 현재 데이터 상태
    base_df = load_base_data()
    current_total = len(base_df) + sum(len(d) for d in st.session_state.uploaded_data)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📋 기본 데이터", f"{len(base_df):,}건")
    with col2:
        st.metric("➕ 추가된 데이터", f"{sum(len(d) for d in st.session_state.uploaded_data):,}건")
    with col3:
        st.metric("🔵 현재 총합", f"{current_total:,}건")
    
    st.markdown("---")
    
    # 이전 결과 팝업 표시 (한 번만)
    if 'upload_result' in st.session_state and st.session_state.upload_result:
        result = st.session_state.upload_result
        if result['status'] == 'success':
            st.success(f"✅ **업로드 성공** — {result['message']}")
            st.balloons()
        elif result['status'] == 'error':
            st.error(f"❌ **업로드 실패** — {result['message']}")
        elif result['status'] == 'delete':
            st.success(f"🗑️ **삭제 완료** — {result['message']}")
        st.session_state.upload_result = None
    
    # 업로드 이력 추적용 세션
    if 'uploaded_files_info' not in st.session_state:
        st.session_state.uploaded_files_info = []
    
    if 'uploader_key' not in st.session_state:
        st.session_state.uploader_key = 0
    
    # ===== 파일 업로드 =====
    uploaded_file = st.file_uploader(
        "📤 파일 선택 (xlsx, csv)",
        type=['xlsx', 'csv'],
        help="컬럼 형식: 기본 CSV와 동일 (No, IC, Model, Build_Num, Result, Keyword 등)",
        key=f"file_uploader_{st.session_state.uploader_key}"
    )
    
    if uploaded_file is not None:
        try:
            # 파일 읽기
            if uploaded_file.name.endswith('.csv'):
                new_df = pd.read_csv(uploaded_file)
            else:
                new_df = pd.read_excel(uploaded_file)
            
            # 필수 컬럼 체크
            required_cols = ['IC', 'Model', 'Build_Num', 'Result', 'Test_Item', 'Keyword']
            missing = [c for c in required_cols if c not in new_df.columns]
            if missing:
                st.session_state.upload_result = {
                    'status': 'error',
                    'message': f"필수 컬럼 누락: {', '.join(missing)}"
                }
                st.session_state.uploader_key += 1
                st.rerun()
            
            # 데이터 전처리
            if 'Test_Date' in new_df.columns:
                new_df['Test_Date'] = pd.to_datetime(new_df['Test_Date'], errors='coerce')
            if 'IC_Code' in new_df.columns:
                new_df['IC_Code'] = new_df['IC_Code'].astype(str).str.zfill(2)
            if 'Model_Minor' in new_df.columns:
                new_df['Model_Minor'] = new_df['Model_Minor'].astype(str).str.zfill(2)
            new_df['Keyword'] = new_df['Keyword'].fillna('')
            
            # ===== 중복 검사 =====
            
            # [1] 같은 파일
            file_signature = (uploaded_file.name, len(new_df))
            already_uploaded_files = [
                (info['filename'], info['n_rows']) 
                for info in st.session_state.uploaded_files_info
            ]
            if file_signature in already_uploaded_files:
                st.session_state.upload_result = {
                    'status': 'error',
                    'message': f"'{uploaded_file.name}' ({len(new_df)}건)은 이미 업로드된 파일입니다."
                }
                st.session_state.uploader_key += 1
                st.rerun()
            
            # [2] No 컬럼 중복
            if 'No' in new_df.columns:
                new_nos = set(new_df['No'].dropna().astype(int).tolist())
                existing_nos = set(base_df['No'].dropna().astype(int).tolist())
                for d in st.session_state.uploaded_data:
                    if 'No' in d.columns:
                        existing_nos.update(d['No'].dropna().astype(int).tolist())
                
                overlap_nos = new_nos & existing_nos
                if overlap_nos:
                    overlap_sample = sorted(list(overlap_nos))[:5]
                    sample_str = ', '.join(map(str, overlap_sample))
                    extra = f" 등 총 {len(overlap_nos)}건" if len(overlap_nos) > 5 else ""
                    st.session_state.upload_result = {
                        'status': 'error',
                        'message': f"No 컬럼 중복 — 이미 존재하는 번호: {sample_str}{extra}"
                    }
                    st.session_state.uploader_key += 1
                    st.rerun()
            
            # [3] 행 단위 중복
            check_cols = ['Build_Num', 'Model', 'Test_Item', 'Result']
            if all(c in new_df.columns for c in check_cols):
                all_existing = pd.concat([base_df] + st.session_state.uploaded_data, ignore_index=True) \
                    if st.session_state.uploaded_data else base_df
                existing_keys = set(
                    all_existing[check_cols].apply(lambda r: tuple(r), axis=1).tolist()
                )
                new_keys = new_df[check_cols].apply(lambda r: tuple(r), axis=1).tolist()
                n_overlap = sum(1 for k in new_keys if k in existing_keys)
                overlap_rate = n_overlap / len(new_df) * 100 if len(new_df) > 0 else 0
                
                if overlap_rate >= 50:
                    st.session_state.upload_result = {
                        'status': 'error',
                        'message': f"데이터 중복 — {n_overlap}건 ({overlap_rate:.1f}%)이 기존 데이터와 동일합니다."
                    }
                    st.session_state.uploader_key += 1
                    st.rerun()
            
            # 통과 → 추가
            no_set = set(new_df['No'].dropna().astype(int).tolist()) if 'No' in new_df.columns else set()
            st.session_state.uploaded_files_info.append({
                'filename': uploaded_file.name,
                'n_rows': len(new_df),
                'no_set': no_set,
                'uploaded_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })
            st.session_state.uploaded_data.append(new_df)
            st.session_state.data_version += 1
            
            new_builds = sorted(new_df['Build_Num'].unique())
            existing_builds = sorted(base_df['Build_Num'].unique())
            truly_new_builds = [b for b in new_builds if b not in existing_builds]
            
            pass_n = (new_df['Result'] == 'PASS').sum()
            fail_n = (new_df['Result'] == 'FAIL').sum()
            
            success_msg = f"`{uploaded_file.name}` 파일에서 **{len(new_df)}건** 추가됨 (PASS {pass_n} · FAIL {fail_n})"
            if truly_new_builds:
                success_msg += f"  🆕 신규 빌드: {truly_new_builds}"
            
            st.session_state.upload_result = {
                'status': 'success',
                'message': success_msg
            }
            st.session_state.uploader_key += 1
            st.rerun()
        
        except Exception as e:
            st.session_state.upload_result = {
                'status': 'error',
                'message': f"파일 처리 중 오류 발생: {str(e)}"
            }
            st.session_state.uploader_key += 1
            st.rerun()
    
    # ===== 📜 업로드 이력 (체크박스 + 선택 삭제) =====
    if st.session_state.uploaded_files_info:
        st.markdown("---")
        st.markdown("#### 📜 업로드 이력")
        st.caption("체크박스로 삭제할 파일을 선택할 수 있습니다.")
        
        # 각 파일에 대한 체크박스 (세션에 선택 상태 저장)
        if 'delete_selection' not in st.session_state or \
           len(st.session_state.delete_selection) != len(st.session_state.uploaded_files_info):
            st.session_state.delete_selection = [False] * len(st.session_state.uploaded_files_info)
        
        # 헤더
        col_h1, col_h2, col_h3, col_h4 = st.columns([0.7, 3, 1.5, 2])
        with col_h1:
            st.markdown("**선택**")
        with col_h2:
            st.markdown("**파일명**")
        with col_h3:
            st.markdown("**건수**")
        with col_h4:
            st.markdown("**업로드 시각**")
        
        # 각 항목 표시
        for i, info in enumerate(st.session_state.uploaded_files_info):
            col1, col2, col3, col4 = st.columns([0.7, 3, 1.5, 2])
            with col1:
                checked = st.checkbox(
                    f"선택 {i+1}",
                    key=f"file_check_{i}_{st.session_state.uploader_key}",
                    value=st.session_state.delete_selection[i],
                    label_visibility="collapsed"
                )
                st.session_state.delete_selection[i] = checked
            with col2:
                st.markdown(f"`{info['filename']}`")
            with col3:
                st.markdown(f"{info['n_rows']:,}건")
            with col4:
                st.markdown(info['uploaded_at'])
        
        st.markdown("")
        
        # 선택된 개수
        n_selected = sum(st.session_state.delete_selection)
        n_total = len(st.session_state.uploaded_files_info)
        
        # 액션 버튼 2개
        btn_col1, btn_col2 = st.columns(2)
        
        with btn_col1:
            if st.button("☑️ 전체 선택", use_container_width=True):
                st.session_state.delete_selection = [True] * n_total
                st.rerun()
        
        with btn_col2:
            disabled = n_selected == 0
            label = f"🗑️ 선택 삭제 ({n_selected}개)" if n_selected > 0 else "🗑️ 선택 삭제 (0개)"
            if st.button(label, type="primary" if not disabled else "secondary",
                         use_container_width=True, disabled=disabled):
                # 선택된 파일들의 인덱스 (역순으로 정렬해서 삭제)
                indices_to_delete = sorted(
                    [i for i, sel in enumerate(st.session_state.delete_selection) if sel],
                    reverse=True
                )
                deleted_names = []
                deleted_count = 0
                for idx in indices_to_delete:
                    deleted_names.append(st.session_state.uploaded_files_info[idx]['filename'])
                    deleted_count += st.session_state.uploaded_files_info[idx]['n_rows']
                    del st.session_state.uploaded_files_info[idx]
                    del st.session_state.uploaded_data[idx]
                    del st.session_state.delete_selection[idx]
                
                st.session_state.data_version += 1
                st.session_state.uploader_key += 1
                st.session_state.upload_result = {
                    'status': 'delete',
                    'message': f"{len(deleted_names)}개 파일 ({deleted_count}건) 삭제됨: {', '.join(deleted_names[:3])}{'...' if len(deleted_names) > 3 else ''}"
                }
                st.rerun()


# ==============================================================================
# 📥 메뉴 5: 보고서 생성 (이력 저장 + 선택 삭제)
# ==============================================================================
def render_report_generator(df, fail_df):
    st.markdown("### 📥 주간 보고서 자동 생성")
    st.caption("월 단위로 데이터를 선택해 회사 표준 SQA Quality Review PPT를 자동 생성합니다.")
    
    # 보고서 이력 세션 초기화
    if 'report_history' not in st.session_state:
        st.session_state.report_history = []  # [{filename, month, n_data, generated_at, ppt_bytes}, ...]
    if 'report_delete_selection' not in st.session_state:
        st.session_state.report_delete_selection = []
    if 'report_result' not in st.session_state:
        st.session_state.report_result = None
    
    # 이전 결과 팝업 표시
    if st.session_state.report_result:
        result = st.session_state.report_result
        if result['status'] == 'success':
            st.success(f"✅ **보고서 생성 완료** — {result['message']}")
            st.balloons()
        elif result['status'] == 'delete':
            st.success(f"🗑️ **삭제 완료** — {result['message']}")
        elif result['status'] == 'error':
            st.error(f"❌ **생성 실패** — {result['message']}")
        st.session_state.report_result = None
    
    # 데이터에서 사용 가능한 월 목록 추출
    df_copy = df.copy()
    df_copy['Test_Date'] = pd.to_datetime(df_copy['Test_Date'], errors='coerce')
    df_copy['year_month'] = df_copy['Test_Date'].dt.to_period('M')
    available_months = sorted(df_copy['year_month'].dropna().unique(), reverse=True)
    
    if len(available_months) == 0:
        st.error("⚠️ 분석 가능한 데이터가 없습니다.")
        return
    
    # 월 선택 UI
    month_options = [f"{m.year}년 {m.month}월" for m in available_months]
    month_map = {f"{m.year}년 {m.month}월": m for m in available_months}
    
    col1, col2 = st.columns(2)
    with col1:
        selected_month_str = st.selectbox(
            "📅 보고서 기준 월",
            options=month_options,
            help="이 월의 데이터로 보고서를 생성합니다."
        )
    with col2:
        report_type = st.selectbox("📄 보고서 형식", ["SQA-SW Quality Review (10장)"])
    
    selected_month = month_map[selected_month_str]
    
    # 선택한 월의 데이터 필터링
    monthly_df = df_copy[df_copy['year_month'] == selected_month].copy()
    monthly_fail = monthly_df[monthly_df['Result'] == 'FAIL'].copy()
    
    # 미리보기 정보
    st.markdown("---")
    st.markdown(f"#### 📋 {selected_month_str} 보고서 정보")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📋 분석 데이터", f"{len(monthly_df):,}건")
    with col2:
        st.metric("❌ Fail 건수", f"{len(monthly_fail):,}건")
    with col3:
        st.metric("💻 모델 수", f"{monthly_df['Model'].nunique()}개")
    with col4:
        st.metric("🔄 빌드 수", f"{monthly_df['Build_Num'].nunique()}개")
    
    if len(monthly_df) == 0:
        st.warning("⚠️ 선택한 월에 데이터가 없습니다.")
        return
    
    st.markdown("#### 📑 포함될 슬라이드 (10장)")
    st.markdown(f"""
    1. 표지 ({selected_month_str} 자동)
    2. 목차
    3. 기본정보
    4. **월간 Test 요약** (Test_Item별 + PASS/FAIL)
    5. **IC × 빌드별 상세** ({selected_month_str} 빌드 한정)
    6. **Fail/Issue Review** (CRITICAL/HIGH 우선)
    7. **🎯 Action Item Review** (도메인 dict 자동 매칭)
    8. **회귀 알람** (월 내 최신 빌드 NEW/WORSE 자동 검출)
    9. **회의 요약 및 향후 계획** (자동 요약)
    10. Thank you
    """)
    
    st.markdown("---")
    
    # 생성 버튼
    if st.button("🚀 보고서 생성하기", type="primary", use_container_width=True):
        with st.spinner("📝 PPT 생성 중... (약 3~5초)"):
            try:
                ppt_buffer = generate_ppt_report(monthly_df, monthly_fail, selected_month)
                ppt_bytes = ppt_buffer.getvalue()
                
                filename = f"SQA_Quality_Review_{selected_month.year}_{selected_month.month:02d}.pptx"
                
                # 보고서 이력에 저장
                st.session_state.report_history.append({
                    'filename': filename,
                    'month_str': selected_month_str,
                    'n_data': len(monthly_df),
                    'n_fail': len(monthly_fail),
                    'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'ppt_bytes': ppt_bytes
                })
                st.session_state.report_delete_selection = [False] * len(st.session_state.report_history)
                
                st.session_state.report_result = {
                    'status': 'success',
                    'message': f"`{filename}` 생성 완료 (분석 데이터: {len(monthly_df):,}건). 아래 이력에서 다운로드하세요."
                }
                st.rerun()
            except Exception as e:
                st.session_state.report_result = {
                    'status': 'error',
                    'message': f"보고서 생성 실패: {str(e)}"
                }
                import traceback
                st.code(traceback.format_exc())
                st.rerun()
    
    # ===== 📜 보고서 이력 (체크박스 + 선택 삭제 + 다운로드) =====
    if st.session_state.report_history:
        st.markdown("---")
        st.markdown("#### 📜 생성한 보고서 이력")
        st.caption("다운로드 또는 체크박스로 삭제할 수 있습니다.")
        
        # 선택 상태 동기화
        if len(st.session_state.report_delete_selection) != len(st.session_state.report_history):
            st.session_state.report_delete_selection = [False] * len(st.session_state.report_history)
        
        # 헤더
        col_h1, col_h2, col_h3, col_h4, col_h5 = st.columns([0.7, 2.5, 1.2, 1.8, 1.5])
        with col_h1:
            st.markdown("**선택**")
        with col_h2:
            st.markdown("**파일명**")
        with col_h3:
            st.markdown("**기준 월**")
        with col_h4:
            st.markdown("**생성 시각**")
        with col_h5:
            st.markdown("**다운로드**")
        
        # 각 항목
        for i, info in enumerate(st.session_state.report_history):
            col1, col2, col3, col4, col5 = st.columns([0.7, 2.5, 1.2, 1.8, 1.5])
            with col1:
                checked = st.checkbox(
                    f"선택 {i+1}",
                    key=f"report_check_{i}",
                    value=st.session_state.report_delete_selection[i],
                    label_visibility="collapsed"
                )
                st.session_state.report_delete_selection[i] = checked
            with col2:
                st.markdown(f"`{info['filename']}`")
            with col3:
                st.markdown(info['month_str'])
            with col4:
                st.markdown(info['generated_at'])
            with col5:
                st.download_button(
                    label="📥 다운로드",
                    data=info['ppt_bytes'],
                    file_name=info['filename'],
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key=f"download_report_{i}",
                    use_container_width=True
                )
        
        st.markdown("")
        
        # 선택된 개수
        n_selected = sum(st.session_state.report_delete_selection)
        n_total = len(st.session_state.report_history)
        
        # 액션 버튼 2개
        btn_col1, btn_col2 = st.columns(2)
        
        with btn_col1:
            if st.button("☑️ 전체 선택", use_container_width=True, key="report_select_all"):
                st.session_state.report_delete_selection = [True] * n_total
                st.rerun()
        
        with btn_col2:
            disabled = n_selected == 0
            label = f"🗑️ 선택 삭제 ({n_selected}개)" if n_selected > 0 else "🗑️ 선택 삭제 (0개)"
            if st.button(label, type="primary" if not disabled else "secondary",
                         use_container_width=True, disabled=disabled, key="report_delete"):
                indices_to_delete = sorted(
                    [i for i, sel in enumerate(st.session_state.report_delete_selection) if sel],
                    reverse=True
                )
                deleted_names = []
                for idx in indices_to_delete:
                    deleted_names.append(st.session_state.report_history[idx]['filename'])
                    del st.session_state.report_history[idx]
                    del st.session_state.report_delete_selection[idx]
                
                st.session_state.report_result = {
                    'status': 'delete',
                    'message': f"{len(deleted_names)}개 보고서 삭제됨: {', '.join(deleted_names[:3])}{'...' if len(deleted_names) > 3 else ''}"
                }
                st.rerun()




def generate_ppt_report(df, fail_df, selected_month):
    """PPT 보고서 생성 → BytesIO 반환 (페이지 맞춤 버전)"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    month_str = f"{selected_month.year}년 {selected_month.month}월"
    
    # 헬퍼 함수
    def set_cell(cell, text, bold=False, size=10, color=None, bg=None, align=None):
        cell.text = ""
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        if align:
            p.alignment = align
        run = p.add_run()
        run.text = str(text)
        run.font.name = PPT_FONT
        run.font.size = Pt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = color
        if bg:
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
    
    def add_text(slide, text, x, y, w, h, size=14, bold=False, color=None, align=None):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        if align:
            p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.name = PPT_FONT
        run.font.size = Pt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = color
    
    def add_header(slide, page_num, title):
        box = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.3), Inches(0.3), Inches(0.4), Inches(0.4))
        box.fill.solid()
        box.fill.fore_color.rgb = RPT_HEADER_BG
        box.line.fill.background()
        add_text(slide, str(page_num), 0.3, 0.32, 0.4, 0.4,
                 size=14, bold=True, color=RPT_HEADER_FG, align=PP_ALIGN.CENTER)
        add_text(slide, "SQA-SW Quality Review", 0.8, 0.3, 5, 0.4,
                 size=18, bold=True, color=RPT_TITLE)
        add_text(slide, f"❑ {title}", 0.4, 0.85, 12, 0.4,
                 size=16, bold=True, color=RPT_TITLE)
    
    # ===== Slide 1: 표지 =====
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RPT_HEADER_BG
    bg.line.fill.background()
    box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(1.5), Inches(3), Inches(10.3), Inches(1.5))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    box.line.fill.background()
    add_text(s1, "SQA-SW Quality Review", 1.5, 3.4, 10.3, 0.8,
             size=32, bold=True, color=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
    add_text(s1, month_str, 5, 5.5, 3.3, 0.5,
             size=22, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    
    # ===== Slide 2: 목차 =====
    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s2, 1, "List")
    items = [
        "1.   회의 개요",
        "2.   월간 TEST 결과 요약 (SQA)",
        "3.   Fail 항목 Review (SQA)",
        "4.   Action Item Review (SW)",
        "5.   회의 요약 및 내부 공유",
    ]
    y_pos = 2.0
    for item in items:
        add_text(s2, item, 1.5, y_pos, 10, 0.5, size=18, color=RGBColor(0x33, 0x33, 0x33))
        y_pos += 0.7
    
    # ===== Slide 3: 기본정보 =====
    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s3, 2, "기본정보")
    table = s3.shapes.add_table(5, 2, Inches(1.5), Inches(2),
                                 Inches(10.3), Inches(3.5)).table
    table.columns[0].width = Inches(2.5)
    table.columns[1].width = Inches(7.8)
    info = [
        ("항 목", "내 용"),
        ("회의명", "SQA-SW 주간 Quality Review"),
        ("보고서 기준 월", month_str),
        ("참석자", "SQA 팀, SW 개발팀"),
        ("회의목적", "SQA 펌웨어 테스트 결과 공유 및 Action Item Review"),
    ]
    for r, (c1, c2) in enumerate(info):
        if r == 0:
            set_cell(table.cell(r, 0), c1, bold=True, size=12,
                     color=RPT_HEADER_FG, bg=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 1), c2, bold=True, size=12,
                     color=RPT_HEADER_FG, bg=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
        else:
            set_cell(table.cell(r, 0), c1, bold=True, size=11,
                     bg=RPT_ALT_ROW, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 1), c2, size=11)
    
    # ===== Slide 4: 월간 Test 요약 =====
    # 페이지 맞춤: 최대 10개 항목까지만 (페이지 넘침 방지)
    s4 = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s4, 3, "월간 Test 요약")
    add_text(s4, f"{month_str} Test Summary", 0.4, 1.4, 5, 0.4,
             size=14, bold=True, color=RPT_TITLE)
    
    # Top 10개만 표시 (페이지 맞춤)
    item_summary = df['Test_Item'].value_counts().head(10).reset_index()
    item_summary.columns = ['Test_Item', 'count']
    
    rows_t1 = len(item_summary) + 2
    # 페이지 높이 맞춤: 행 높이 자동 조절
    available_height = 5.0  # y=2부터 7까지
    row_height = min(0.4, available_height / rows_t1)
    
    table1 = s4.shapes.add_table(rows_t1, 2, Inches(0.4), Inches(2),
                                  Inches(5.5), Inches(row_height * rows_t1)).table
    table1.columns[0].width = Inches(3.5)
    table1.columns[1].width = Inches(2.0)
    set_cell(table1.cell(0, 0), "Test Item", bold=True, size=10,
             color=RPT_HEADER_FG, bg=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
    set_cell(table1.cell(0, 1), month_str, bold=True, size=10,
             color=RPT_HEADER_FG, bg=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
    total_count = 0
    for r, row in enumerate(item_summary.itertuples(), 1):
        bg = RPT_ALT_ROW if r % 2 == 0 else None
        set_cell(table1.cell(r, 0), row.Test_Item, size=9, bg=bg)
        set_cell(table1.cell(r, 1), row.count, size=9, bg=bg, align=PP_ALIGN.CENTER)
        total_count += row.count
    set_cell(table1.cell(rows_t1-1, 0), "Total", bold=True, size=10,
             bg=RPT_HEADER_BG, color=RPT_HEADER_FG, align=PP_ALIGN.CENTER)
    set_cell(table1.cell(rows_t1-1, 1), total_count, bold=True, size=10,
             bg=RPT_HEADER_BG, color=RPT_HEADER_FG, align=PP_ALIGN.CENTER)
    
    pass_count = (df['Result'] == 'PASS').sum()
    fail_count = (df['Result'] == 'FAIL').sum()
    total = pass_count + fail_count
    fail_rate = fail_count / total * 100 if total > 0 else 0
    
    table2 = s4.shapes.add_table(4, 3, Inches(6.5), Inches(2),
                                  Inches(6), Inches(1.6)).table
    table2.columns[0].width = Inches(1.5)
    table2.columns[1].width = Inches(1.8)
    table2.columns[2].width = Inches(2.7)
    set_cell(table2.cell(0, 0), "구분", bold=True, size=11,
             color=RPT_HEADER_FG, bg=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
    set_cell(table2.cell(0, 1), month_str, bold=True, size=11,
             color=RPT_HEADER_FG, bg=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
    set_cell(table2.cell(0, 2), "Fail율", bold=True, size=11,
             color=RPT_HEADER_FG, bg=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
    set_cell(table2.cell(1, 0), "PASS", bold=True, size=11, bg=RPT_ALT_ROW, align=PP_ALIGN.CENTER)
    set_cell(table2.cell(1, 1), f"{pass_count:,}", size=11, bg=RPT_ALT_ROW, align=PP_ALIGN.CENTER)
    set_cell(table2.cell(1, 2), "", size=11, bg=RPT_ALT_ROW)
    set_cell(table2.cell(2, 0), "FAIL", bold=True, size=11, color=RPT_CRITICAL, align=PP_ALIGN.CENTER)
    set_cell(table2.cell(2, 1), f"{fail_count:,}", size=11, color=RPT_CRITICAL, bold=True, align=PP_ALIGN.CENTER)
    set_cell(table2.cell(2, 2), f"{fail_rate:.1f}%", size=12, color=RPT_CRITICAL, bold=True, align=PP_ALIGN.CENTER)
    set_cell(table2.cell(3, 0), "Total", bold=True, size=11, bg=RPT_HEADER_BG, color=RPT_HEADER_FG, align=PP_ALIGN.CENTER)
    set_cell(table2.cell(3, 1), f"{total:,}", bold=True, size=11, bg=RPT_HEADER_BG, color=RPT_HEADER_FG, align=PP_ALIGN.CENTER)
    set_cell(table2.cell(3, 2), "", bg=RPT_HEADER_BG)
    
    add_text(s4, "💡 자동 인사이트", 6.5, 4.0, 6, 0.4, size=12, bold=True, color=RPT_TITLE)
    if len(fail_df) > 0:
        top_fail_kw = fail_df['Keyword'].value_counts().head(3)
        insight = f"• Top Fail_Type: {', '.join(top_fail_kw.index[:3])}\n"
        insight += f"• 영향 모델: {fail_df['Model'].nunique()}개\n"
        insight += f"• 영향 IC: {fail_df['IC'].nunique()}종\n"
        insight += f"• 분석 빌드: {df['Build_Num'].nunique()}개"
    else:
        insight = "• 이번 달 Fail 없음 ✅"
    add_text(s4, insight, 6.5, 4.4, 6, 2.5, size=11, color=RGBColor(0x33, 0x33, 0x33))
    
    # ===== Slide 5: IC × 빌드별 상세 (페이지 맞춤) =====
    s5 = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s5, 3, "월간 Test 요약")
    add_text(s5, f"{month_str} IC × 빌드별 Test 결과 상세", 0.4, 1.4, 10, 0.4,
             size=14, bold=True, color=RPT_TITLE)
    
    summary = df.groupby(['IC', 'Customer', 'Build_Num']).agg(
        total=('Result', 'count'),
        pass_n=('Result', lambda x: (x == 'PASS').sum()),
        fail_n=('Result', lambda x: (x == 'FAIL').sum())
    ).reset_index()
    summary['fail_rate'] = (summary['fail_n'] / summary['total'] * 100).round(1)
    # 최대 8개 행으로 제한 (페이지 맞춤)
    summary = summary.sort_values('fail_rate', ascending=False).head(8)
    
    if len(summary) > 0:
        rows = len(summary) + 1
        available_height = 5.0
        row_height = min(0.5, available_height / rows)
        
        table = s5.shapes.add_table(rows, 7, Inches(0.4), Inches(2),
                                     Inches(12.6), Inches(row_height * rows)).table
        widths = [1.0, 1.5, 2.5, 1.3, 1.5, 1.5, 1.3]
        for i, w in enumerate(widths):
            table.columns[i].width = Inches(w)
        headers = ['IC', '고객사', '빌드 번호', '총 Test', 'Pass', 'Fail', 'Fail율(%)']
        for c, h in enumerate(headers):
            set_cell(table.cell(0, c), h, bold=True, size=10,
                     color=RPT_HEADER_FG, bg=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
        for r, row in enumerate(summary.itertuples(), 1):
            bg = RPT_ALT_ROW if r % 2 == 0 else None
            set_cell(table.cell(r, 0), row.IC, bold=True, size=9, bg=bg, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 1), row.Customer, size=9, bg=bg, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 2), f"R00.0.{row.Build_Num}", size=9, bg=bg, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 3), row.total, size=9, bg=bg, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 4), row.pass_n, size=9, bg=bg, align=PP_ALIGN.CENTER)
            fail_color = RPT_CRITICAL if row.fail_n > 0 else None
            set_cell(table.cell(r, 5), row.fail_n, size=9, bold=row.fail_n > 0, bg=bg, color=fail_color, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 6), f"{row.fail_rate}%", size=9, bold=row.fail_rate > 5, bg=bg,
                     color=fail_color if row.fail_rate > 5 else None, align=PP_ALIGN.CENTER)
    else:
        add_text(s5, "이번 달 데이터가 없습니다.", 0.4, 3, 10, 0.4, size=14)
    
    # ===== Slide 6: Fail Review (페이지 맞춤) =====
    s6 = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s6, 4, "Fail / Issue Review")
    add_text(s6, "주요 Fail 케이스 (CRITICAL/HIGH 우선)", 0.4, 1.4, 10, 0.4,
             size=14, bold=True, color=RPT_TITLE)
    
    fail_df_copy = fail_df.copy()
    if len(fail_df_copy) > 0:
        # Fail 건수 기준 정렬 (많은 게 위로)
        kw_counts = fail_df['Keyword'].value_counts().to_dict()
        fail_df_copy['_sort'] = fail_df_copy['Keyword'].map(kw_counts).fillna(0)
        fail_df_copy = fail_df_copy.sort_values(['_sort', 'Build_Num'], ascending=[False, False]).head(7)
        
        rows = len(fail_df_copy) + 1
        available_height = 4.8
        row_height = min(0.55, available_height / rows)
        
        table = s6.shapes.add_table(rows, 6, Inches(0.4), Inches(2),
                                     Inches(12.6), Inches(row_height * rows)).table
        widths = [0.6, 1.2, 1.5, 2.0, 1.5, 5.8]
        for i, w in enumerate(widths):
            table.columns[i].width = Inches(w)
        headers = ['No.', 'Test 항목', '발생 빌드', '모델 / IC', 'Fail Type', '주요 현상']
        for c, h in enumerate(headers):
            set_cell(table.cell(0, c), h, bold=True, size=10,
                     color=RPT_HEADER_FG, bg=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
        for r, row in enumerate(fail_df_copy.itertuples(), 1):
            bg = RPT_ALT_ROW if r % 2 == 0 else None
            set_cell(table.cell(r, 0), r, size=9, bg=bg, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 1), row.Test_Item, size=8, bg=bg)
            set_cell(table.cell(r, 2), f"R00.0.{row.Build_Num}", size=8, bg=bg, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 3), f"{row.Model} / {row.IC}", size=8, bg=bg, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 4), row.Keyword, size=8, bold=True, bg=bg, color=RPT_CRITICAL, align=PP_ALIGN.CENTER)
            desc = str(row.Fail_Description)[:80] if pd.notna(row.Fail_Description) else "-"
            set_cell(table.cell(r, 5), desc, size=7, bg=bg)
    else:
        add_text(s6, "이번 달 Fail 케이스가 없습니다. ✅", 0.4, 3, 10, 0.4,
                 size=16, bold=True, color=RPT_MID)
    
    # ===== Slide 7: Action Item (페이지 맞춤) ⭐ =====
    s7 = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s7, 5, "Action Item Review (자동 생성)")
    add_text(s7, "📌 도메인 dict 기반 Action Item 자동 매칭", 0.4, 1.4, 10, 0.4,
             size=14, bold=True, color=RPT_TITLE)
    
    if len(fail_df) > 0:
        fail_df_copy2 = fail_df.copy()
        fail_df_copy2['action'] = fail_df_copy2['Keyword'].apply(
            lambda k: KEYWORD_DOMAIN_MAP.get(k, {}).get('area', '?'))
        # Fail 건수 기준 정렬 (많은 게 위로)
        kw_counts = fail_df['Keyword'].value_counts().to_dict()
        fail_df_copy2['_sort'] = fail_df_copy2['Keyword'].map(kw_counts).fillna(0)
        # 최대 7개로 제한
        unique_actions = fail_df_copy2.drop_duplicates(subset=['Model', 'Keyword']) \
                                      .sort_values(['_sort', 'Build_Num'], ascending=[False, False]) \
                                      .head(7)
        
        rows = len(unique_actions) + 1
        available_height = 4.8
        row_height = min(0.55, available_height / rows)
        
        table = s7.shapes.add_table(rows, 5, Inches(0.4), Inches(2),
                                     Inches(12.6), Inches(row_height * rows)).table
        widths = [0.6, 1.5, 2.5, 2.0, 6.0]
        for i, w in enumerate(widths):
            table.columns[i].width = Inches(w)
        headers = ['No.', '발생 빌드', '모델', 'Fail Type', 'Action Item']
        for c, h in enumerate(headers):
            set_cell(table.cell(0, c), h, bold=True, size=10,
                     color=RPT_HEADER_FG, bg=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
        for r, row in enumerate(unique_actions.itertuples(), 1):
            bg = RPT_ALT_ROW if r % 2 == 0 else None
            set_cell(table.cell(r, 0), r, size=9, bg=bg, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 1), f"R00.0.{row.Build_Num}", size=9, bg=bg, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 2), row.Model, size=9, bg=bg, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 3), row.Keyword, size=9, bold=True, bg=bg, color=RPT_CRITICAL, align=PP_ALIGN.CENTER)
            set_cell(table.cell(r, 4), row.action + " 검토", size=9, bg=bg, bold=True)
    else:
        add_text(s7, "이번 달 Action Item이 없습니다. ✅", 0.4, 3, 10, 0.4,
                 size=16, bold=True, color=RPT_MID)
    
    # ===== Slide 8: 회귀 알람 (페이지 맞춤) =====
    s8 = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s8, 6, "회귀 알람 (월 내 최신 빌드)")
    
    builds = sorted(df['Build_Num'].unique())
    if len(builds) >= 2 and len(fail_df) > 0:
        latest_build = builds[-1]
        prev_build = builds[-2]
        latest = fail_df[fail_df['Build_Num'] == latest_build]
        prev = fail_df[fail_df['Build_Num'] == prev_build]
        
        latest_combo = latest.groupby(['Model', 'Keyword']).size().reset_index(name='latest')
        prev_combo = prev.groupby(['Model', 'Keyword']).size().reset_index(name='prev')
        merged = latest_combo.merge(prev_combo, on=['Model', 'Keyword'], how='left')
        merged['prev'] = merged['prev'].fillna(0).astype(int)
        merged['상태'] = merged.apply(
            lambda r: 'NEW' if r['prev'] == 0 else ('WORSE' if r['latest'] > r['prev'] else ''), axis=1)
        alerts = merged[merged['상태'].isin(['NEW', 'WORSE'])].copy()
        # Fail 건수 기준 정렬 (많은 게 위로)
        alerts = alerts.sort_values('latest', ascending=False).head(8)
        
        add_text(s8, f"빌드 R00.0.{latest_build}에서 신규/악화 결함 자동 검출", 0.4, 1.4, 10, 0.4,
                 size=14, bold=True, color=RPT_TITLE)
        n_new = (alerts['상태'] == 'NEW').sum()
        n_worse = (alerts['상태'] == 'WORSE').sum()
        add_text(s8, f"🚨 전체 {len(alerts)}건  ·  신규 {n_new}건  ·  악화 {n_worse}건", 0.4, 1.85, 12, 0.4,
                 size=11, color=RPT_CRITICAL)
        
        if len(alerts) > 0:
            rows = len(alerts) + 1
            available_height = 4.5
            row_height = min(0.5, available_height / rows)
            
            table = s8.shapes.add_table(rows, 6, Inches(0.4), Inches(2.4),
                                         Inches(12.6), Inches(row_height * rows)).table
            widths = [1.0, 2.0, 1.8, 1.2, 1.2, 5.4]
            for i, w in enumerate(widths):
                table.columns[i].width = Inches(w)
            headers = ['상태', '모델', 'Fail Type', '이전', '현재', '검토 영역']
            for c, h in enumerate(headers):
                set_cell(table.cell(0, c), h, bold=True, size=10,
                         color=RPT_HEADER_FG, bg=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
            for r, row in enumerate(alerts.itertuples(), 1):
                bg = RPT_ALT_ROW if r % 2 == 0 else None
                state_color = RPT_CRITICAL if row.상태 == 'NEW' else RPT_HIGH
                set_cell(table.cell(r, 0), row.상태, size=9, bold=True, color=state_color, bg=bg, align=PP_ALIGN.CENTER)
                set_cell(table.cell(r, 1), row.Model, size=9, bg=bg, align=PP_ALIGN.CENTER)
                set_cell(table.cell(r, 2), row.Keyword, size=9, bold=True, color=RPT_CRITICAL, bg=bg, align=PP_ALIGN.CENTER)
                set_cell(table.cell(r, 3), row.prev, size=9, bg=bg, align=PP_ALIGN.CENTER)
                set_cell(table.cell(r, 4), row.latest, size=9, bold=True, bg=bg, align=PP_ALIGN.CENTER)
                area = KEYWORD_DOMAIN_MAP.get(row.Keyword, {}).get('area', '?')
                set_cell(table.cell(r, 5), area, size=9, bg=bg)
        else:
            add_text(s8, "이번 빌드에서 회귀 알람 없음 ✅", 0.4, 3, 10, 0.4,
                     size=16, color=RPT_MID)
    else:
        add_text(s8, "이 월에 빌드가 1개 이하라 회귀 비교 불가", 0.4, 3, 10, 0.4,
                 size=14, color=RPT_MID)
    
    # ===== Slide 9: 회의 요약 =====
    s9 = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s9, 7, "회의 요약 및 향후 계획")
    
    table = s9.shapes.add_table(5, 2, Inches(0.4), Inches(2),
                                 Inches(12.6), Inches(0.55 * 5)).table
    table.columns[0].width = Inches(3.0)
    table.columns[1].width = Inches(9.6)
    set_cell(table.cell(0, 0), "구분", bold=True, size=12,
             color=RPT_HEADER_FG, bg=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
    set_cell(table.cell(0, 1), "내용", bold=True, size=12,
             color=RPT_HEADER_FG, bg=RPT_HEADER_BG, align=PP_ALIGN.CENTER)
    
    set_cell(table.cell(1, 0), "주요 결과", bold=True, size=11, bg=RPT_ALT_ROW, align=PP_ALIGN.CENTER)
    if len(df) > 0:
        result_text = f"• {month_str} Test {len(df):,}건 中 FAIL {len(fail_df):,}건 ({len(fail_df)/len(df)*100:.1f}%)\n"
        if len(builds) > 0:
            latest_b = builds[-1]
            latest_fail_n = len(fail_df[fail_df['Build_Num'] == latest_b]) if len(fail_df) > 0 else 0
            result_text += f"• 최신 빌드 R00.0.{latest_b}: FAIL {latest_fail_n}건\n"
        if len(fail_df) > 0:
            result_text += f"• 영향 모델: {fail_df['Model'].nunique()}개  ·  IC: {fail_df['IC'].nunique()}종"
    else:
        result_text = "• 이번 달 데이터 없음"
    set_cell(table.cell(1, 1), result_text, size=10)
    
    set_cell(table.cell(2, 0), "주요 이슈", bold=True, size=11, bg=RPT_ALT_ROW, align=PP_ALIGN.CENTER)
    if len(fail_df) > 0:
        top_kw = fail_df['Keyword'].value_counts().head(3)
        issue_text = "Top 3 Fail Type (건수 기준):\n"
        for i, (kw, cnt) in enumerate(top_kw.items(), 1):
            issue_text += f"• {i}위 {kw}: {cnt}건\n"
    else:
        issue_text = "이번 달 Fail 이슈 없음 ✅"
    set_cell(table.cell(2, 1), issue_text.strip(), size=10)
    
    set_cell(table.cell(3, 0), "Action Item", bold=True, size=11, bg=RPT_ALT_ROW, align=PP_ALIGN.CENTER)
    if len(fail_df) > 0:
        top_3 = fail_df['Keyword'].value_counts().head(3)
        action_text = "주요 검토 항목 (Fail 빈도 TOP 3):\n"
        for i, (kw, cnt) in enumerate(top_3.items(), 1):
            area = KEYWORD_DOMAIN_MAP.get(kw, {}).get('area', '?')
            action_text += f"• {i}위 {kw} ({cnt}건): {area} 검토\n"
        action_text += "• 다음 빌드에서 회귀 모니터링 필수"
    else:
        action_text = "별도 Action Item 없음"
    set_cell(table.cell(3, 1), action_text, size=10)
    
    set_cell(table.cell(4, 0), "차기 회의", bold=True, size=11, bg=RPT_ALT_ROW, align=PP_ALIGN.CENTER)
    set_cell(table.cell(4, 1), "다음 빌드 출시 후 1주 이내", size=10)
    
    # ===== Slide 10: Thank you =====
    s10 = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RPT_HEADER_BG
    bg.line.fill.background()
    add_text(s10, "Thank you", 0, 3, 13.3, 1.5,
             size=60, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    
    # BytesIO로 저장
    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


# ==============================================================================
# 메인 앱
# ==============================================================================

st.title("🔍 SQA 터치패널 펌웨어 분석 대시보드")
st.caption("서강대학교 AI·SW대학원 | 생성형 AI와 파이썬 데이터 분석 | A74072 조희주")
st.markdown("---")

try:
    df = get_current_df()
    fail_df = df[df['Result'] == 'FAIL'].copy()
except Exception as e:
    st.error(f"⚠️ 데이터 로드 실패: {e}")
    st.stop()

st.sidebar.title("📋 메뉴")
menu = st.sidebar.radio(
    "분석 메뉴 선택",
    ["🎯 필터 분석", "📊 전체 인사이트", "📈 전체 통계", "🔬 전문 통계 분석", "🚨 회귀 알람", "📁 데이터 업로드", "📥 보고서 생성"]
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📦 데이터 정보")
st.sidebar.markdown(f"- 전체 테스트: **{len(df):,}**건")
st.sidebar.markdown(f"- Fail 건수: **{len(fail_df):,}**건")
st.sidebar.markdown(f"- 분석 모델: **{df['Model'].nunique()}**개")
st.sidebar.markdown(f"- 분석 FW: **{df['FW_Version'].nunique()}**개")

if st.session_state.uploaded_data:
    st.sidebar.success(f"➕ 추가 데이터: {sum(len(d) for d in st.session_state.uploaded_data)}건")

st.sidebar.markdown("---")
st.sidebar.markdown("### 💡 메뉴 가이드")
st.sidebar.markdown("""
- **🎯 필터 분석**  
  5가지 필터 조합 동적 분석
- **📊 전체 인사이트**  
  핵심 차트 5개
- **📈 전체 통계**  
  월별·고객사·매트릭스 분석
- **🔬 전문 통계 분석** ⭐ NEW  
  카이제곱·신뢰구간·시계열 예측
- **🚨 회귀 알람**  
  최신 빌드 자동 검출
- **📁 데이터 업로드**  
  엑셀/CSV 추가 → 즉시 분석
- **📥 보고서 생성**  
  월 단위 PPT 자동 생성
""")


if menu == "🎯 필터 분석":
    st.markdown("## 🎯 필터 분석")
    st.markdown("**회의 중 즉시 사용 가능한 동적 분석 화면**")
    st.markdown("---")
    render_filter_analysis(df, fail_df)

elif menu == "📊 전체 인사이트":
    st.markdown("## 📊 전체 인사이트")
    st.markdown("---")
    render_full_insights(df, fail_df)

elif menu == "📈 전체 통계":
    st.markdown("## 📈 전체 통계")
    st.markdown("**전체 데이터의 통계 정보를 종합적으로 보여줍니다 (월별·고객사·매트릭스·메타)**")
    st.markdown("---")
    render_full_statistics(df, fail_df)

elif menu == "🔬 전문 통계 분석":
    st.markdown("## 🔬 전문 통계 분석")
    st.markdown("**카이제곱 검정 + 신뢰구간 + 시계열 예측 (scipy + sklearn 활용)**")
    st.markdown("---")
    render_advanced_statistics(df, fail_df)

elif menu == "🚨 회귀 알람":
    st.markdown("## 🚨 회귀 알람")
    st.markdown("**최신 빌드에서 새로 등장하거나 악화된 결함을 자동 검출**")
    st.markdown("---")
    render_regression_alert(df, fail_df)

elif menu == "📁 데이터 업로드":
    st.markdown("## 📁 데이터 업로드")
    st.markdown("**새 주간 체크리스트를 업로드하면 자동으로 분석 데이터에 추가됩니다**")
    st.markdown("---")
    render_data_upload()

elif menu == "📥 보고서 생성":
    st.markdown("## 📥 주간 보고서 생성")
    st.markdown("**현재 데이터로 회사 표준 SQA Quality Review PPT를 자동 생성합니다**")
    st.markdown("---")
    render_report_generator(df, fail_df)
