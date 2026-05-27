
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns
import os
import unicodedata

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
# 상수 (도메인 dict)
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


# ──────────────────────────────────────────────────────────────────────────
# 데이터 로드
# ──────────────────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
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


# ==============================================================================
# 차트 함수 (기존 유지)
# ==============================================================================

def draw_ic_bar_donut(df, fail_df):
    plt.rcParams['font.family'] = font_name
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

    st.info("💡 IC별 Fail율 편차는 약 2%p로 작지만, Fail_Type 분포는 IC마다 완전히 다릅니다.")


def draw_category_bar(df):
    plt.rcParams['font.family'] = font_name
    cat_summary = df.groupby('Category').agg(
        total=('Result', 'count'),
        fail=('Result', lambda x: (x == 'FAIL').sum())
    ).reset_index()
    cat_summary['fail_rate'] = (cat_summary['fail'] / cat_summary['total'] * 100).round(2)
    cat_summary = cat_summary.sort_values('fail_rate', ascending=False)

    fig, ax = plt.subplots(figsize=(12, 5))
    labels = [f"{c} ({TEST_AREA_MAP.get(c, '?')})" for c in cat_summary['Category']]
    colors = plt.cm.RdYlGn_r(cat_summary['fail_rate'] / cat_summary['fail_rate'].max())
    bars = ax.barh(labels, cat_summary['fail_rate'], color=colors)
    for bar, rate, total in zip(bars, cat_summary['fail_rate'], cat_summary['total']):
        ax.text(bar.get_width() + 0.2, bar.get_y() + bar.get_height()/2,
                f'{rate}%  ({total:,}건)', va='center', fontsize=10)
    ax.set_title('카테고리별 Fail율', fontsize=14, fontweight='bold')
    ax.set_xlabel('Fail율 (%)')
    ax.set_xlim(0, cat_summary['fail_rate'].max() * 1.3)
    ax.invert_yaxis()
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)


def draw_category_heatmap(fail_df):
    plt.rcParams['font.family'] = font_name
    pivot = fail_df.pivot_table(
        index='Category', columns='Keyword', values='No', aggfunc='count', fill_value=0
    )
    pivot = pivot[[k for k in KEYWORDS_ALL if k in pivot.columns]]

    fig, ax = plt.subplots(figsize=(12, 6))
    sns.heatmap(pivot, annot=True, fmt='d', cmap='YlOrRd',
                cbar_kws={'label': 'Fail 건수'}, ax=ax,
                linewidths=0.5, linecolor='white')
    ax.set_title('카테고리 × Fail_Type 히트맵',
                 fontsize=14, fontweight='bold', pad=15)
    ax.set_xlabel('Fail_Type')
    ax.set_ylabel('Category')
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)
    st.info("💡 카테고리와 실제 Fail_Type의 검토 영역이 일치하지 않습니다.")


def draw_build_curve(df):
    plt.rcParams['font.family'] = font_name
    build_summary = df.groupby('Build_Num').agg(
        total=('Result', 'count'),
        fail=('Result', lambda x: (x == 'FAIL').sum())
    ).reset_index()
    build_summary['fail_rate'] = (build_summary['fail'] / build_summary['total'] * 100).round(2)
    build_summary = build_summary.sort_values('Build_Num').reset_index(drop=True)

    fig, ax = plt.subplots(figsize=(14, 6))
    ax.plot(build_summary['Build_Num'], build_summary['fail_rate'],
            marker='o', markersize=10, linewidth=2, color='#e74c3c', label='Fail율')

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

    ax.set_title('빌드별 전체 Fail율 추이 — ∩(역U) 안정화 곡선',
                 fontsize=14, fontweight='bold', pad=15)
    ax.set_xlabel('빌드 번호')
    ax.set_ylabel('Fail율 (%)')
    ax.grid(True, alpha=0.3)
    ax.set_xticks(build_summary['Build_Num'])
    ax.set_xticklabels(build_summary['Build_Num'], rotation=45)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)


def draw_model_heatmap(fail_df):
    plt.rcParams['font.family'] = font_name
    pivot = fail_df.pivot_table(
        index='Model', columns='Keyword', values='No', aggfunc='count', fill_value=0
    )
    pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=False).index]
    pivot = pivot[[k for k in KEYWORDS_ALL if k in pivot.columns]]

    fig, ax = plt.subplots(figsize=(12, max(8, len(pivot) * 0.3)))
    sns.heatmap(pivot, annot=True, fmt='d', cmap='YlOrRd',
                cbar_kws={'label': 'Fail 건수'}, ax=ax,
                linewidths=0.5, linecolor='white')
    ax.set_title('모델 × Fail_Type 히트맵', fontsize=14, fontweight='bold', pad=15)
    ax.set_xlabel('Fail_Type')
    ax.set_ylabel('Model')
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)


def draw_priority_matrix():
    plt.rcParams['font.family'] = font_name
    fig, ax = plt.subplots(figsize=(10, 7))

    for kw in KEYWORDS_ALL:
        severity = SEVERITY_MAP[kw]
        priority = KEYWORD_DOMAIN_MAP[kw]['priority']
        color = {'CRITICAL': '#e74c3c', 'HIGH': '#f39c12', 'MID': '#3498db'}[priority]
        size = severity * 200

        ax.scatter(severity, severity, s=size, c=color, alpha=0.6,
                   edgecolors='black', linewidths=1)
        ax.annotate(kw, (severity, severity),
                    xytext=(10, 10), textcoords='offset points', fontsize=10)

    ax.axhline(y=3.5, color='gray', linestyle='--', alpha=0.5)
    ax.axvline(x=3.5, color='gray', linestyle='--', alpha=0.5)
    ax.set_xlabel('Severity (심각도)', fontsize=12)
    ax.set_ylabel('Priority (우선순위)', fontsize=12)
    ax.set_title('Severity × Priority 우선순위 매트릭스', fontsize=14, fontweight='bold')
    ax.set_xlim(0, 6)
    ax.set_ylim(0, 6)
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)


# ==============================================================================
# 🆕 필터 분석 함수 (NEW)
# ==============================================================================

def render_filter_analysis(df, fail_df):
    """복합 필터 기반 동적 분석 화면"""
    plt.rcParams['font.family'] = font_name
    
    # ===== 필터 UI =====
    st.markdown("### 🎯 필터 선택")
    st.caption("여러 조건을 조합해서 원하는 데이터만 분석할 수 있습니다.")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        ic_filter = st.selectbox(
            "🔌 IC",
            ["전체"] + sorted(df['IC'].unique().tolist()),
            help="칩셋 종류 선택"
        )
    with col2:
        model_filter = st.selectbox(
            "💻 Model",
            ["전체"] + sorted(df['Model'].unique().tolist()),
            help="노트북 모델 선택"
        )
    with col3:
        build_filter = st.selectbox(
            "🔄 Build",
            ["전체"] + sorted(df['Build_Num'].unique().tolist()),
            help="펌웨어 빌드 번호 선택"
        )
    with col4:
        kw_filter = st.selectbox(
            "🐛 Fail_Type",
            ["전체"] + KEYWORDS_ALL,
            help="결함 유형 선택"
        )
    
    # ===== 필터 적용 =====
    filtered_all = df.copy()
    filtered_fail = fail_df.copy()
    
    active_filters = []
    if ic_filter != "전체":
        filtered_all = filtered_all[filtered_all['IC'] == ic_filter]
        filtered_fail = filtered_fail[filtered_fail['IC'] == ic_filter]
        active_filters.append(f"IC={ic_filter}")
    if model_filter != "전체":
        filtered_all = filtered_all[filtered_all['Model'] == model_filter]
        filtered_fail = filtered_fail[filtered_fail['Model'] == model_filter]
        active_filters.append(f"Model={model_filter}")
    if build_filter != "전체":
        filtered_all = filtered_all[filtered_all['Build_Num'] == build_filter]
        filtered_fail = filtered_fail[filtered_fail['Build_Num'] == build_filter]
        active_filters.append(f"Build={build_filter}")
    if kw_filter != "전체":
        # Fail_Type 필터는 fail_df에만 적용
        filtered_fail = filtered_fail[filtered_fail['Keyword'] == kw_filter]
        active_filters.append(f"Fail_Type={kw_filter}")
    
    # 활성 필터 표시
    st.markdown("---")
    if active_filters:
        filter_text = " · ".join(active_filters)
        st.markdown(f"#### 🎯 적용된 필터: `{filter_text}`")
    else:
        st.markdown("#### 🎯 전체 데이터 분석 (필터 미적용)")
    
    # 데이터 없음 처리
    if len(filtered_fail) == 0:
        st.warning("⚠️ 선택한 조건에 해당하는 Fail 케이스가 없습니다. 필터를 조정해주세요.")
        return
    
    # ===== ① KPI 카드 =====
    st.markdown("### 📊 ① 요약 통계")
    
    total_tests = len(filtered_all) if len(filtered_all) > 0 else 0
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
    
    # ===== ② Fail_Type 순위 (TOP 5) =====
    if kw_filter == "전체":  # Fail_Type 필터 안 한 경우만 표시
        st.markdown("### 🏆 ② Fail_Type 순위 (TOP 5)")
        
        kw_rank = filtered_fail['Keyword'].value_counts().head(5).reset_index()
        kw_rank.columns = ['Fail_Type', '건수']
        kw_rank['비중'] = (kw_rank['건수'] / kw_rank['건수'].sum() * 100).round(1)
        kw_rank['우선순위'] = kw_rank['Fail_Type'].apply(
            lambda k: KEYWORD_DOMAIN_MAP.get(k, {}).get('priority', '?')
        )
        kw_rank['검토영역'] = kw_rank['Fail_Type'].apply(
            lambda k: KEYWORD_DOMAIN_MAP.get(k, {}).get('area', '?')
        )
        kw_rank.insert(0, '순위', range(1, len(kw_rank) + 1))
        
        col_left, col_right = st.columns([1, 1])
        with col_left:
            # 도넛 차트
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
            # 순위 표
            st.dataframe(
                kw_rank,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "순위": st.column_config.NumberColumn("순위", width=60),
                    "Fail_Type": st.column_config.TextColumn("Fail_Type", width=120),
                    "건수": st.column_config.NumberColumn("건수", format="%d"),
                    "비중": st.column_config.NumberColumn("비중 (%)", format="%.1f"),
                    "우선순위": st.column_config.TextColumn("우선순위", width=80),
                    "검토영역": st.column_config.TextColumn("검토 영역"),
                }
            )
        
        st.markdown("---")
    
    # ===== ③ 빌드별 추이 =====
    if build_filter == "전체":  # 빌드 필터 안 한 경우만 표시
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
        else:
            st.info(f"단일 빌드({build_trend['Build_Num'].iloc[0]})만 해당하여 추이 차트는 생략됩니다.")
        
        st.markdown("---")
    
    # ===== ④ 모델별 Fail 순위 =====
    if model_filter == "전체":  # 모델 필터 안 한 경우만 표시
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
    
    # ===== ⑤ 상세 케이스 + 영상 =====
    st.markdown(f"### 📋 ⑤ 상세 Fail 케이스 ({len(filtered_fail):,}건)")
    st.caption("표의 영상 링크를 클릭하면 새 탭에서 결함 영상이 재생됩니다.")
    
    display_df = filtered_fail[[
        'No', 'IC', 'Model', 'Keyword', 'Build_Num',
        'Customer', 'Category', 'Test_Item', 'video_id'
    ]].copy()
    display_df['우선순위'] = display_df['Keyword'].apply(
        lambda k: KEYWORD_DOMAIN_MAP.get(k, {}).get('priority', '?')
    )
    display_df['영상'] = display_df['video_id'].apply(get_video_link)
    display_df = display_df.drop(columns=['video_id'])
    
    # CRITICAL 우선 정렬
    priority_order = {'CRITICAL': 0, 'HIGH': 1, 'MID': 2}
    display_df['_sort'] = display_df['우선순위'].map(priority_order).fillna(99)
    display_df = display_df.sort_values('_sort').drop(columns=['_sort'])
    
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "영상": st.column_config.LinkColumn(
                "영상 보기",
                display_text="🎬 재생"
            ),
            "우선순위": st.column_config.TextColumn("우선순위", width=90),
            "Keyword": st.column_config.TextColumn("Fail_Type"),
        }
    )


# ==============================================================================
# 메인 앱
# ==============================================================================

st.title("🔍 SQA 터치패널 펌웨어 분석 대시보드")
st.caption("서강대학교 AI·SW대학원 | 생성형 AI와 파이썬 데이터 분석 | A74072 조희주")
st.markdown("---")

try:
    with st.spinner('데이터 로드 중...'):
        df = load_data()
    fail_df = df[df['Result'] == 'FAIL'].copy()
    st.success(f"✅ 데이터 로드 완료 ({len(df):,}건)")
except Exception as e:
    st.error(f"⚠️ 데이터 로드 실패: {e}")
    st.stop()

# 사이드바
st.sidebar.title("📋 메뉴")
menu = st.sidebar.radio(
    "분석 메뉴 선택",
    ["🏠 대시보드 (홈)", "📊 Visualization", "🔬 Analysis", "🎯 필터 분석", "🚨 회귀 알람"]
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📦 데이터 정보")
st.sidebar.markdown(f"- 전체 테스트: **{len(df):,}**건")
st.sidebar.markdown(f"- Fail 건수: **{len(fail_df):,}**건")
st.sidebar.markdown(f"- 분석 모델: **{df['Model'].nunique()}**개")
st.sidebar.markdown(f"- 분석 FW: **{df['FW_Version'].nunique()}**개")


# ============= 메뉴 1. 홈 =============
if menu == "🏠 대시보드 (홈)":
    st.markdown("## 📌 전체 분석 현황")
    total_tests = len(df)
    total_fails = len(fail_df)
    fail_rate = total_fails / total_tests * 100

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown("##### 📊 전체 테스트")
        st.markdown(f"### :blue[**{total_tests:,}**]")
    with col2:
        st.markdown("##### ❌ Fail율")
        st.markdown(f"### :red[**{fail_rate:.1f}%**]")
        st.caption(f"({total_fails:,}건)")
    with col3:
        st.markdown("##### 📱 분석 모델")
        st.markdown(f"### :green[**{df['Model'].nunique()}개**]")
    with col4:
        st.markdown("##### 🔄 분석 FW")
        st.markdown(f"### :orange[**{df['FW_Version'].nunique()}개**]")

    st.markdown("---")
    st.markdown("### 🎯 프로젝트 소개")
    st.markdown("""
    이 대시보드는 **터치패널 펌웨어 SQA 테스트 결과를 자동 분석**합니다.
    매주 분석가가 수기로 30분~1시간 걸리던 작업을 5초 만에 자동화하고,
    회귀 알람·우선순위 분류·영상 확인까지 한 화면에서 제공합니다.
    """)
    
    st.markdown("---")
    st.markdown("### 💡 사용 가이드")
    st.markdown("""
    - **📊 Visualization**: 전체 시각화 차트 4개 (IC/카테고리/빌드별)
    - **🔬 Analysis**: 모델별 히트맵, 드릴다운, 우선순위 매트릭스
    - **🎯 필터 분석** ⭐ NEW: IC/모델/빌드/Fail_Type 조합 필터 → 동적 분석
    - **🚨 회귀 알람**: 최신 빌드의 신규/악화 결함 자동 검출
    """)


# ============= 메뉴 2. Visualization =============
elif menu == "📊 Visualization":
    st.markdown("## 📊 Visualization")
    st.markdown("---")
    st.markdown("### 1️⃣ IC별 Fail율 + Fail_Type 분포")
    draw_ic_bar_donut(df, fail_df)
    st.markdown("---")
    st.markdown("### 2️⃣ 카테고리별 Fail율")
    draw_category_bar(df)
    st.markdown("---")
    st.markdown("### 3️⃣ 카테고리 × Fail_Type 히트맵")
    draw_category_heatmap(fail_df)
    st.markdown("---")
    st.markdown("### 4️⃣ 빌드별 ∩ 곡선")
    draw_build_curve(df)


# ============= 메뉴 3. Analysis =============
elif menu == "🔬 Analysis":
    st.markdown("## 🔬 Analysis")
    st.markdown("---")

    tab1, tab2, tab3 = st.tabs([
        "🔥 모델 히트맵", "🔍 모델 드릴다운", "🎯 우선순위 매트릭스"
    ])

    with tab1:
        st.markdown("### 모델 × Fail_Type 히트맵 (시그니처)")
        draw_model_heatmap(fail_df)

    with tab2:
        st.markdown("### 단일 모델 드릴다운")
        model_fails = fail_df['Model'].value_counts()
        models_with_fails = model_fails.index.tolist()
        selected_model = st.selectbox(
            "🎯 분석할 모델 선택 (Fail 건수 많은 순)",
            options=models_with_fails,
            index=0
        )
        if selected_model:
            target_data = df[df['Model'] == selected_model]
            target_fail = target_data[target_data['Result'] == 'FAIL']
            if len(target_fail) > 0:
                # 간단한 드릴다운 (기존 코드 재사용)
                fig, axes = plt.subplots(2, 2, figsize=(14, 10))
                fig.suptitle(f'모델 드릴다운: {selected_model}', fontsize=16, fontweight='bold')
                kw_dist = target_fail['Keyword'].value_counts()
                axes[0, 0].pie(kw_dist.values, labels=kw_dist.index, autopct='%1.1f%%',
                               wedgeprops=dict(width=0.4), startangle=90, colors=plt.cm.Set3.colors)
                axes[0, 0].set_title('Fail_Type 분포')
                fw_trend = target_fail.groupby('Build_Num').size().reset_index(name='count')
                axes[0, 1].plot(fw_trend['Build_Num'], fw_trend['count'],
                                marker='o', markersize=10, linewidth=2, color='#e74c3c')
                axes[0, 1].set_title('빌드별 Fail 추이')
                axes[0, 1].set_xticks(fw_trend['Build_Num'])
                axes[0, 1].set_xticklabels(fw_trend['Build_Num'], rotation=45)
                axes[0, 1].grid(True, alpha=0.3)
                cat_stats = target_data.groupby('Category').agg(
                    total=('Result', 'count'),
                    fail=('Result', lambda x: (x == 'FAIL').sum())
                ).reset_index()
                cat_stats['fail_rate'] = cat_stats['fail'] / cat_stats['total'] * 100
                cat_stats = cat_stats.sort_values('fail_rate', ascending=True)
                axes[1, 0].barh(cat_stats['Category'], cat_stats['fail_rate'], color='#3498db')
                axes[1, 0].set_title('카테고리별 Fail율')
                axes[1, 1].axis('off')
                info_text = (
                    f"전체: {len(target_data)}건\n"
                    f"Fail: {len(target_fail)}건 ({len(target_fail)/len(target_data)*100:.1f}%)\n\n"
                    f"우세 Fail_Type: {kw_dist.idxmax()}\n"
                    f"비중: {kw_dist.max()/kw_dist.sum()*100:.1f}%"
                )
                axes[1, 1].text(0.05, 0.95, info_text, fontsize=12, verticalalignment='top',
                                bbox=dict(boxstyle='round', facecolor='#f0f0f0'))
                plt.tight_layout()
                st.pyplot(fig)
                plt.close(fig)

    with tab3:
        st.markdown("### Severity × Priority 우선순위 매트릭스")
        draw_priority_matrix()


# ============= 메뉴 4. 🎯 필터 분석 (NEW) =============
elif menu == "🎯 필터 분석":
    st.markdown("## 🎯 필터 분석")
    st.markdown("**회의 중 즉시 사용 가능한 동적 분석 화면**")
    st.markdown("IC, 모델, 빌드, Fail_Type 4가지 필터를 조합해 원하는 데이터만 분석합니다.")
    st.markdown("---")
    
    render_filter_analysis(df, fail_df)


# ============= 메뉴 5. 회귀 알람 =============
elif menu == "🚨 회귀 알람":
    st.markdown("## 🚨 빌드 17920 회귀 알람")
    st.markdown("최신 빌드에서 새로 등장하거나 악화된 결함을 자동 검출합니다.")
    st.markdown("---")

    latest_build = 17920
    prev_build = 17751

    if latest_build in df['Build_Num'].values:
        latest_fail = fail_df[fail_df['Build_Num'] == latest_build]
        prev_fail = fail_df[fail_df['Build_Num'] == prev_build]

        latest_combo = latest_fail.groupby(['Model', 'Keyword']).size().reset_index(name='latest')
        prev_combo = prev_fail.groupby(['Model', 'Keyword']).size().reset_index(name='prev')

        merged = latest_combo.merge(prev_combo, on=['Model', 'Keyword'], how='left')
        merged['prev'] = merged['prev'].fillna(0).astype(int)
        merged['상태'] = merged.apply(
            lambda r: 'NEW' if r['prev'] == 0 else ('WORSE' if r['latest'] > r['prev'] else ''),
            axis=1
        )
        alerts = merged[merged['상태'].isin(['NEW', 'WORSE'])].copy()
        alerts['우선순위'] = alerts['Keyword'].apply(
            lambda k: KEYWORD_DOMAIN_MAP.get(k, {}).get('priority', '?')
        )
        alerts['검토영역'] = alerts['Keyword'].apply(
            lambda k: KEYWORD_DOMAIN_MAP.get(k, {}).get('area', '?')
        )

        priority_order = {'CRITICAL': 0, 'HIGH': 1, 'MID': 2}
        alerts['_sort'] = alerts['우선순위'].map(priority_order).fillna(99)
        alerts = alerts.sort_values(['_sort', 'latest'], ascending=[True, False])

        n_critical = (alerts['우선순위'] == 'CRITICAL').sum()
        n_high = (alerts['우선순위'] == 'HIGH').sum()
        n_mid = (alerts['우선순위'] == 'MID').sum()

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown("##### 🚨 전체 알람")
            st.markdown(f"### :red[**{len(alerts)}건**]")
        with col2:
            st.markdown("##### 🔴 CRITICAL")
            st.markdown(f"### :red[**{n_critical}건**]")
        with col3:
            st.markdown("##### 🟠 HIGH")
            st.markdown(f"### :orange[**{n_high}건**]")
        with col4:
            st.markdown("##### 🟡 MID")
            st.markdown(f"### :blue[**{n_mid}건**]")

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
                '우선순위': row['우선순위'],
                '검토 영역': row['검토영역'],
                '영상': video_link
            })
        alert_df = pd.DataFrame(alert_rows)

        st.dataframe(
            alert_df,
            use_container_width=True,
            column_config={
                "영상": st.column_config.LinkColumn("영상 보기", display_text="🎬 재생")
            }
        )
