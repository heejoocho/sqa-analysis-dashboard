
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
DRIVE_FOLDER_ID = "1_lwEa6AI1jVdGxXnqedNXqJPlFRUfXgl"

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
    
    # 영상 파일명 추출 + NFC 정규화
    def extract_filename(link):
        if pd.isna(link) or link == '':
            return None
        return unicodedata.normalize('NFC', os.path.basename(str(link).strip()))
    
    df['video_filename'] = df['Video_Link'].apply(extract_filename)
    return df


def get_video_url(filename):
    """파일명으로 Drive 검색 URL 생성"""
    if pd.isna(filename) or filename is None or filename == '':
        return None
    # Drive 폴더 안에서 파일명으로 검색 (자동 재생 가능)
    return f"https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}"


# ==============================================================================
# 차트 함수
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
    
    max_idx = build_summary['fail_rate'].idxmax()
    min_idx = build_summary['fail_rate'].idxmin()
    ax.scatter(build_summary.loc[max_idx, 'Build_Num'], build_summary.loc[max_idx, 'fail_rate'],
               s=300, color='red', zorder=5, label='최고 (악화 정점)',
               edgecolor='darkred', linewidth=2)
    ax.scatter(build_summary.loc[min_idx, 'Build_Num'], build_summary.loc[min_idx, 'fail_rate'],
               s=300, color='green', zorder=5, label='최저 (안정화)',
               edgecolor='darkgreen', linewidth=2)
    
    ax.set_title('빌드별 전체 Fail율 추이 — ∩(역U) 안정화 곡선',
                 fontsize=14, fontweight='bold', pad=15)
    ax.set_xlabel('빌드 번호')
    ax.set_ylabel('Fail율 (%)')
    ax.legend(loc='upper right', fontsize=10)
    ax.grid(True, alpha=0.3)
    ax.set_xticks(build_summary['Build_Num'])
    ax.set_xticklabels(build_summary['Build_Num'], rotation=45)
    ax.set_ylim(build_summary['fail_rate'].min() - 2, build_summary['fail_rate'].max() + 2)
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
    ax.set_title('모델 × Fail_Type 히트맵 (시그니처 차트)',
                 fontsize=14, fontweight='bold', pad=15)
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
    ax.set_title('Severity × Priority 우선순위 매트릭스',
                 fontsize=14, fontweight='bold')
    ax.set_xlim(0, 6)
    ax.set_ylim(0, 6)
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)


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
    ["🏠 대시보드 (홈)", "📊 Visualization", "🔬 Analysis", "🚨 회귀 알람"]
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📦 데이터 정보")
st.sidebar.markdown(f"- 전체 테스트: **{len(df):,}**건")
st.sidebar.markdown(f"- Fail 건수: **{len(fail_df):,}**건")
st.sidebar.markdown(f"- 분석 모델: **{df['Model'].nunique()}**개")
st.sidebar.markdown(f"- 분석 FW: **{df['FW_Version'].nunique()}**개")


# ============= 메뉴 1. 홈 =============
if menu == "🏠 대시보드 (홈)":
    st.markdown("## 📌 전체 분석 현황 한눈에 보기")
    
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
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("### 🔍 분석 단계")
        st.markdown("""
        - **묶음 1**: 5축 기본 집계
        - **묶음 2**: 다축 패턴 관찰
        - **묶음 3**: 변화·알람·우선순위
        """)
    with col_b:
        st.markdown("### 🎯 분석 차원")
        st.markdown(f"""
        - **IC**: {df['IC'].nunique()}종
        - **Fail_Type**: 8종
        - **Build**: {df['Build_Num'].nunique()}개
        - **Customer**: {df['Customer'].nunique()}개
        """)
    
    st.markdown("---")
    st.markdown("### 📂 전체 데이터")
    st.dataframe(df.head(50), use_container_width=True)


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
    
    st.markdown("### 3️⃣ 빌드별 ∩ 곡선")
    draw_build_curve(df)


# ============= 메뉴 3. Analysis =============
elif menu == "🔬 Analysis":
    st.markdown("## 🔬 Analysis")
    st.markdown("---")
    
    tab1, tab2, tab3 = st.tabs(["🔥 모델 히트맵", "🎯 우선순위 매트릭스", "📋 Fail 검색"])
    
    with tab1:
        st.markdown("### 모델 × Fail_Type 히트맵")
        draw_model_heatmap(fail_df)
    
    with tab2:
        st.markdown("### Severity × Priority 우선순위 매트릭스")
        draw_priority_matrix()
    
    with tab3:
        st.markdown("### 🔎 Fail 케이스 검색 + 영상 보기")
        
        # 필터
        col1, col2, col3 = st.columns(3)
        with col1:
            ic_filter = st.selectbox("IC 선택", ["전체"] + sorted(df['IC'].unique().tolist()))
        with col2:
            keyword_filter = st.selectbox("Fail_Type 선택", ["전체"] + KEYWORDS_ALL)
        with col3:
            model_filter = st.selectbox("Model 선택", ["전체"] + sorted(df['Model'].unique().tolist()))
        
        # 필터링
        filtered = fail_df.copy()
        if ic_filter != "전체":
            filtered = filtered[filtered['IC'] == ic_filter]
        if keyword_filter != "전체":
            filtered = filtered[filtered['Keyword'] == keyword_filter]
        if model_filter != "전체":
            filtered = filtered[filtered['Model'] == model_filter]
        
        st.markdown(f"**검색 결과: {len(filtered):,}건**")
        
        # 영상 링크 추가
        display_df = filtered[['No', 'IC', 'Model', 'Keyword', 'Build_Num',
                                'Customer', 'Test_Item', 'video_filename']].copy()
        display_df['영상'] = display_df['video_filename'].apply(
            lambda x: f"https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}" if pd.notna(x) else None
        )
        display_df = display_df.drop(columns=['video_filename'])
        
        st.dataframe(
            display_df,
            use_container_width=True,
            column_config={
                "영상": st.column_config.LinkColumn(
                    "영상 보기",
                    display_text="🎬 Drive에서 보기"
                )
            }
        )
        
        st.info("💡 '영상 보기' 컬럼의 링크를 클릭하면 Drive 폴더가 새 탭으로 열립니다. 검색창에 파일명 입력으로 영상 찾기 가능.")


# ============= 메뉴 4. 회귀 알람 =============
elif menu == "🚨 회귀 알람":
    st.markdown("## 🚨 빌드 17920 회귀 알람")
    st.markdown("최신 빌드에서 새로 등장하거나 악화된 결함을 자동 검출합니다.")
    st.markdown("---")
    
    # 17920 빌드 데이터
    latest_build = 17920
    prev_build = 17751
    
    if latest_build in df['Build_Num'].values:
        latest_fail = fail_df[fail_df['Build_Num'] == latest_build]
        prev_fail = fail_df[fail_df['Build_Num'] == prev_build]
        
        # 모델 + Fail_Type 조합으로 그룹핑
        latest_combo = latest_fail.groupby(['Model', 'Keyword']).size().reset_index(name='latest')
        prev_combo = prev_fail.groupby(['Model', 'Keyword']).size().reset_index(name='prev')
        
        # 비교
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
        
        # 정렬 (CRITICAL → HIGH → MID)
        priority_order = {'CRITICAL': 0, 'HIGH': 1, 'MID': 2}
        alerts['_sort'] = alerts['우선순위'].map(priority_order).fillna(99)
        alerts = alerts.sort_values(['_sort', 'latest'], ascending=[True, False])
        
        # KPI
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
        
        # 알람 표
        display = alerts[['상태', 'Model', 'Keyword', 'prev', 'latest', '우선순위', '검토영역']].copy()
        display.columns = ['상태', 'Model', 'Fail_Type', '이전', '현재', '우선순위', '검토영역']
        
        # 영상 링크
        alert_with_video = []
        for _, row in alerts.iterrows():
            sample = latest_fail[(latest_fail['Model'] == row['Model']) &
                                 (latest_fail['Keyword'] == row['Keyword'])].iloc[0]
            alert_with_video.append({
                '상태': row['상태'],
                'Model': row['Model'],
                'Fail_Type': row['Keyword'],
                '이전': row['prev'],
                '현재': row['latest'],
                '우선순위': row['우선순위'],
                '검토영역': row['검토영역'],
                '영상': f"https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}" if pd.notna(sample['video_filename']) else None
            })
        alert_df = pd.DataFrame(alert_with_video)
        
        st.dataframe(
            alert_df,
            use_container_width=True,
            column_config={
                "영상": st.column_config.LinkColumn(
                    "영상 보기",
                    display_text="🎬 Drive 보기"
                )
            }
        )
        
        st.info(
            f"💡 **인사이트**: 빌드 17920에서 총 {len(alerts)}건 회귀 알람 검출. "
            f"CRITICAL {n_critical}건은 즉시 패치 검토 필요."
        )
