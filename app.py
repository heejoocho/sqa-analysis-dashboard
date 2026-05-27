
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
# 🎯 메뉴 1: 필터 분석 (Test_Item 추가)
# ==============================================================================

def render_filter_analysis(df, fail_df):
    plt.rcParams['font.family'] = font_name

    # ===== 상단 KPI (전체 데이터) =====
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

    # ===== 필터 UI (5개로 확장) =====
    st.markdown("### 🎯 필터 선택")
    st.caption("여러 조건을 조합해서 원하는 데이터만 분석할 수 있습니다.")

    # 5개 필터를 2행으로 배치
    row1_col1, row1_col2, row1_col3 = st.columns(3)
    with row1_col1:
        ic_filter = st.selectbox(
            "🔌 IC",
            ["전체"] + sorted(df['IC'].unique().tolist()),
            help="칩셋 종류 선택"
        )
    with row1_col2:
        model_filter = st.selectbox(
            "💻 Model",
            ["전체"] + sorted(df['Model'].unique().tolist()),
            help="노트북 모델 선택"
        )
    with row1_col3:
        build_filter = st.selectbox(
            "🔄 Build",
            ["전체"] + sorted(df['Build_Num'].unique().tolist()),
            help="펌웨어 빌드 번호 선택"
        )
    
    row2_col1, row2_col2, row2_col3 = st.columns(3)
    with row2_col1:
        test_item_filter = st.selectbox(
            "🧪 Test_Item",
            ["전체"] + sorted(df['Test_Item'].unique().tolist()),
            help="테스트 항목 선택"
        )
    with row2_col2:
        kw_filter = st.selectbox(
            "🐛 Fail_Type",
            ["전체"] + KEYWORDS_ALL,
            help="결함 유형 선택"
        )
    with row2_col3:
        st.write("")  # 공간

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
    if test_item_filter != "전체":
        filtered_all = filtered_all[filtered_all['Test_Item'] == test_item_filter]
        filtered_fail = filtered_fail[filtered_fail['Test_Item'] == test_item_filter]
        active_filters.append(f"Test_Item={test_item_filter}")
    if kw_filter != "전체":
        filtered_fail = filtered_fail[filtered_fail['Keyword'] == kw_filter]
        active_filters.append(f"Fail_Type={kw_filter}")

    st.markdown("---")
    if active_filters:
        filter_text = " · ".join(active_filters)
        st.markdown(f"### 🎯 적용된 필터: `{filter_text}`")
    else:
        st.markdown("### 🎯 전체 데이터 분석 (필터 미적용)")

    if len(filtered_fail) == 0:
        st.warning("⚠️ 선택한 조건에 해당하는 Fail 케이스가 없습니다. 필터를 조정해주세요.")
        return

    # ===== ① 필터 결과 KPI =====
    st.markdown("### 📊 ① 필터 결과 요약")

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
    if kw_filter == "전체":
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
    if build_filter == "전체":
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
    if model_filter == "전체":
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

    # ===== ⑤ Test_Item별 Fail 순위 (NEW) =====
    if test_item_filter == "전체":
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

    # ===== ⑥ 상세 케이스 + 영상 =====
    st.markdown(f"### 📋 ⑥ 상세 Fail 케이스 ({len(filtered_fail):,}건)")
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
# 📊 메뉴 2: 전체 인사이트
# ==============================================================================

def render_full_insights(df, fail_df):
    plt.rcParams['font.family'] = font_name

    # ===== ① 빌드별 ∩ 곡선 =====
    st.markdown("### 1️⃣ 빌드별 ∩(역U) 안정화 곡선")
    st.caption("FW 빌드 진행에 따른 전체 Fail율 추이")

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

    # ===== ② IC별 Bar + Donut =====
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

    # ===== ③ 카테고리 × Fail_Type 히트맵 =====
    st.markdown("### 3️⃣ 카테고리 × Fail_Type 히트맵")

    pivot_cat = fail_df.pivot_table(
        index='Category', columns='Keyword', values='No', aggfunc='count', fill_value=0
    )
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

    # ===== ④ 모델 히트맵 =====
    st.markdown("### 4️⃣ 모델 × Fail_Type 히트맵 ⭐")

    pivot_model = fail_df.pivot_table(
        index='Model', columns='Keyword', values='No', aggfunc='count', fill_value=0
    )
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

    # ===== ⑤ 우선순위 매트릭스 =====
    st.markdown("### 5️⃣ Severity × Priority 매트릭스")

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
    ax.set_title('Severity × Priority 매트릭스', fontsize=14, fontweight='bold')
    ax.set_xlim(0, 6)
    ax.set_ylim(0, 6)
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)


# ==============================================================================
# 🚨 메뉴 3: 회귀 알람
# ==============================================================================

def render_regression_alert(df, fail_df):
    latest_build = 17920
    prev_build = 17751

    if latest_build not in df['Build_Num'].values:
        st.warning("최신 빌드 데이터를 찾을 수 없습니다.")
        return

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
        hide_index=True,
        column_config={
            "영상": st.column_config.LinkColumn("영상 보기", display_text="🎬 재생")
        }
    )

    st.info(
        f"💡 **인사이트**: 빌드 17920에서 총 {len(alerts)}건 회귀 알람 검출. "
        f"CRITICAL {n_critical}건은 즉시 패치 검토 필요."
    )


# ==============================================================================
# 메인 앱
# ==============================================================================

st.title("🔍 SQA 터치패널 펌웨어 분석 대시보드")
st.caption("서강대학교 AI·SW대학원 | 생성형 AI와 파이썬 데이터 분석 | A74072 조희주")
st.markdown("---")

try:
    df = load_data()
    fail_df = df[df['Result'] == 'FAIL'].copy()
except Exception as e:
    st.error(f"⚠️ 데이터 로드 실패: {e}")
    st.stop()

st.sidebar.title("📋 메뉴")
menu = st.sidebar.radio(
    "분석 메뉴 선택",
    ["🎯 필터 분석", "📊 전체 인사이트", "🚨 회귀 알람"]
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📦 데이터 정보")
st.sidebar.markdown(f"- 전체 테스트: **{len(df):,}**건")
st.sidebar.markdown(f"- Fail 건수: **{len(fail_df):,}**건")
st.sidebar.markdown(f"- 분석 모델: **{df['Model'].nunique()}**개")
st.sidebar.markdown(f"- 분석 FW: **{df['FW_Version'].nunique()}**개")

st.sidebar.markdown("---")
st.sidebar.markdown("### 💡 메뉴 가이드")
st.sidebar.markdown("""
- **🎯 필터 분석** (메인)  
  IC/모델/빌드/Test_Item/Fail_Type 조합 필터
- **📊 전체 인사이트**  
  전체 데이터 핵심 차트 5개
- **🚨 회귀 알람**  
  최신 빌드 신규/악화 결함 자동 검출
""")


if menu == "🎯 필터 분석":
    st.markdown("## 🎯 필터 분석")
    st.markdown("**회의 중 즉시 사용 가능한 동적 분석 화면**")
    st.markdown("5가지 필터(IC / Model / Build / Test_Item / Fail_Type)를 조합해 원하는 데이터만 분석합니다.")
    st.markdown("---")
    render_filter_analysis(df, fail_df)

elif menu == "📊 전체 인사이트":
    st.markdown("## 📊 전체 인사이트")
    st.markdown("**전체 데이터의 핵심 패턴과 발견 사항을 5개 차트로 정리**")
    st.markdown("---")
    render_full_insights(df, fail_df)

elif menu == "🚨 회귀 알람":
    st.markdown("## 🚨 빌드 17920 회귀 알람")
    st.markdown("**최신 빌드에서 새로 등장하거나 악화된 결함을 자동 검출**")
    st.markdown("---")
    render_regression_alert(df, fail_df)
