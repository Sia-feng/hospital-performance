import pandas as pd
import streamlit as st
from io import BytesIO

# ==================== 页面配置（手机端优化）====================
st.set_page_config(
    page_title="西山医院绩效系统",
    page_icon="🏥",
    layout="centered",  # centered 在手机端更友好
    initial_sidebar_state="collapsed"  # 默认收起侧边栏
)

# 自定义CSS，优化手机触摸体验
st.markdown("""
<style>
    /* 手机端按钮变大 */
    .stButton>button {
        width: 100%;
        height: 3em;
        font-size: 1.1em;
        border-radius: 10px;
    }

    /* 文件上传区域优化 */
    .stFileUploader {
        padding: 1em;
    }

    /* 表格字体调大 */
    .stDataFrame {
        font-size: 1.1em;
    }

    /* 指标卡片样式 */
    div[data-testid="metric-container"] {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 10px;
    }

    /* 隐藏默认的Streamlit页头（更简洁） */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* 手机端标题居中 */
    h1 {
        text-align: center;
        font-size: 1.5em !important;
    }

    /* 成功提示框 */
    .stSuccess {
        border-radius: 10px;
        padding: 1em;
    }
</style>
""", unsafe_allow_html=True)

st.title("🏥 西山医院绩效系统")

# ==================== 侧边栏参数（手机端展开时显示）====================
with st.sidebar:
    st.subheader("⚙️ 绩效设置")

    TOTAL_PERFORMANCE = st.number_input(
        "💰 总绩效（元）",
        value=3447,
        step=1,
        help="每月胸痛中心总绩效"
    )

    DATABASE_PERFORMANCE = round(TOTAL_PERFORMANCE * 0.75, 2)
    FOLLOWUP_PERFORMANCE = round(TOTAL_PERFORMANCE * 0.25, 2)

    st.markdown(f"**数据库绩效**: `{DATABASE_PERFORMANCE}` 元")
    st.markdown(f"**随访绩效**: `{FOLLOWUP_PERFORMANCE}` 元")

    st.markdown("---")
    st.caption("📱 提示：点击左上角 > 展开设置")
    st.caption("v1.0 西山医院")

# ==================== 主界面：模式选择（大按钮，适合手机）====================
st.markdown("### 👇 选择计算类型")

# 使用session_state跟踪模式
if 'calc_mode' not in st.session_state:
    st.session_state.calc_mode = None

# 两列大按钮
col1, col2 = st.columns(2)
with col1:
    if st.button("📊 数据库绩效\n(75%)", use_container_width=True):
        st.session_state.calc_mode = 'database'
        st.rerun()
with col2:
    if st.button("📋 随访绩效\n(25%)", use_container_width=True):
        st.session_state.calc_mode = 'followup'
        st.rerun()

# 显示当前选择
if st.session_state.calc_mode == 'database':
    st.info("📊 当前模式：数据库绩效计算")
elif st.session_state.calc_mode == 'followup':
    st.info("📋 当前模式：随访绩效计算")


# ==================== 数据库绩效函数（原有逻辑不变）====================
def calculate_performance(row, a, base_price):
    门诊登记 = float(row['门诊登记']) if pd.notna(row['门诊登记']) else 0
    门诊收入院 = float(row['门诊收入院']) if pd.notna(row['门诊收入院']) else 0
    门诊建档 = float(row['门诊建档']) if pd.notna(row['门诊建档']) else 0
    急诊建档 = float(row['急诊建档']) if pd.notna(row['急诊建档']) else 0
    急诊未入院 = float(row['急诊未入院']) if pd.notna(row['急诊未入院']) else 0
    典型病例 = float(row['典型病例']) if pd.notna(row['典型病例']) else 0
    工作条目 = str(row['工作条目']) if pd.notna(row['工作条目']) else ''

    work_score = (
            门诊登记 * 0.1 +
            门诊收入院 * 1 +
            门诊建档 * 0.5 +
            急诊建档 * 1 +
            急诊未入院 * 1
    )
    work_performance = work_score * base_price

    qc_performance = 0
    fixed_bonus = 0

    if '数据录入' in 工作条目:
        qc_performance = a * 1.7 * base_price
        fixed_bonus = 300
    elif '审核' in 工作条目:
        qc_performance = a * 0.2 * base_price
    elif '归档' in 工作条目:
        qc_performance = a * 0.2 * 0.8 * base_price

    case_bonus = 0
    if 典型病例 > 0 and '数据录入' not in 工作条目:
        case_bonus = 300

    return round(work_performance + qc_performance + fixed_bonus + case_bonus, 2)


def find_header_row(uploaded_file):
    for skip in range(0, 5):
        try:
            df = pd.read_excel(uploaded_file, skiprows=skip, header=0)
            cols = [str(c).strip() for c in df.columns]
            required = ['姓名', '门诊登记']
            if all(any(r in c for c in cols) for r in required):
                return skip, df
        except:
            continue
    return None, None


def process_database_excel(uploaded_file, database_performance):
    skip_rows, df = find_header_row(uploaded_file)
    if df is None:
        raise Exception("无法找到表头，请检查Excel格式")

    df.columns = [str(c).strip().replace('\n', '').replace(' ', '') for c in df.columns]

    col_mapping = {}
    for std_col in ['工作条目', '姓名', '门诊登记', '门诊收入院', '门诊建档',
                    '急诊建档', '急诊未入院', '典型病例']:
        for actual_col in df.columns:
            if std_col in actual_col:
                col_mapping[std_col] = actual_col
                break

    df = df.rename(columns=col_mapping)

    a = df['门诊建档'].fillna(0).astype(float).sum() + df['急诊建档'].fillna(0).astype(float).sum()

    work_base_sum = (
            df['门诊登记'].fillna(0).astype(float).sum() * 0.1 +
            df['门诊收入院'].fillna(0).astype(float).sum() * 1 +
            df['门诊建档'].fillna(0).astype(float).sum() * 0.5 +
            df['急诊建档'].fillna(0).astype(float).sum() * 1 +
            df['急诊未入院'].fillna(0).astype(float).sum() * 1
    )

    qc_base_sum = 0
    if any('数据录入' in str(x) for x in df['工作条目'].fillna('')):
        qc_base_sum += a * 1.7
    if any('审核' in str(x) for x in df['工作条目'].fillna('')):
        qc_base_sum += a * 0.2
    if any('归档' in str(x) for x in df['工作条目'].fillna('')):
        qc_base_sum += a * 0.2 * 0.8

    total_base = work_base_sum + qc_base_sum
    base_price = (database_performance - 300) / total_base if total_base > 0 else 0

    df['总绩效'] = df.apply(lambda row: calculate_performance(row, a, base_price), axis=1)
    df['个人实发'] = df['总绩效'].round(0).astype(int)

    return df, a, base_price, work_base_sum, qc_base_sum, total_base


# ==================== 随访绩效函数（从Excel读取）====================
def process_followup_excel(uploaded_file, followup_performance):
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip().replace('\n', '').replace(' ', '') for c in df.columns]

    col_mapping = {}
    for std_col in ['姓名', '工作条目', '随访病例', '微信入群']:
        for actual_col in df.columns:
            if std_col in actual_col:
                col_mapping[std_col] = actual_col
                break

    df = df.rename(columns=col_mapping)
    df['随访病例'] = pd.to_numeric(df['随访病例'], errors='coerce').fillna(0)
    df['微信入群'] = pd.to_numeric(df['微信入群'], errors='coerce').fillna(0)
    df['工作条目'] = df['工作条目'].fillna('').astype(str)

    total_wechat = df['微信入群'].sum()
    remaining = followup_performance - total_wechat
    mgmt_ratio = 0.13
    mgmt_total = remaining * mgmt_ratio
    case_total = remaining * (1 - mgmt_ratio)
    total_case_base = df['随访病例'].sum()
    case_base_price = case_total / total_case_base if total_case_base > 0 else 0

    df['微信入群绩效'] = df['微信入群']
    df['随访管理绩效'] = 0.0
    mgmt_mask = df['工作条目'].str.contains('随访管理', na=False)
    if mgmt_mask.sum() > 0:
        df.loc[mgmt_mask, '随访管理绩效'] = mgmt_total / mgmt_mask.sum()

    df['随访病例绩效'] = df['随访病例'] * case_base_price
    df['随访绩效'] = (df['微信入群绩效'] + df['随访管理绩效'] + df['随访病例绩效']).round(2)
    df['随访实发'] = df['随访绩效'].round(0).astype(int)

    def get_detail(row):
        parts = []
        if row['微信入群绩效'] > 0:
            parts.append(f"微信入群:{row['微信入群绩效']:.0f}")
        if row['随访管理绩效'] > 0:
            parts.append(f"管理:{row['随访管理绩效']:.2f}")
        if row['随访病例绩效'] > 0:
            parts.append(f"病例{row['随访病例']:.0f}份={row['随访病例绩效']:.2f}")
        return ' | '.join(parts) if parts else '无'

    df['绩效明细'] = df.apply(get_detail, axis=1)
    return df, case_base_price, mgmt_total, total_wechat


# ==================== 数据库绩效模式 ====================
if st.session_state.calc_mode == 'database':
    st.markdown("---")
    st.subheader("📤 上传数据库绩效数据")

    data = st.file_uploader(
        "选择Excel文件（需包含：姓名、工作条目、门诊登记等列）",
        type=["xlsx", "xls"],
        key="db_upload"
    )

    if data:
        try:
            df, a, base_price, work_base, qc_base, total_base = process_database_excel(data, DATABASE_PERFORMANCE)
            st.session_state["df_db"] = df

            st.success("✅ 计算完成！")

            # 关键指标（手机端用metric，一目了然）
            st.subheader("📊 关键指标")
            c1, c2 = st.columns(2)
            c3, c4 = st.columns(2)
            with c1:
                st.metric("总人数", f"{len(df)}人")
            with c2:
                st.metric("总绩效", f"{df['总绩效'].sum():.0f}元")
            with c3:
                st.metric("实发总数", f"{df['个人实发'].sum()}元")
            with c4:
                st.metric("人均", f"{df['总绩效'].mean():.0f}元")

            # 展开查看参数
            with st.expander("🔧 查看计算参数"):
                st.write(f"数据库绩效: {DATABASE_PERFORMANCE}元")
                st.write(f"工作量基数: {work_base:.2f}")
                st.write(f"质控基数: {qc_base:.2f}")
                st.write(f"绩效基数单价: {base_price:.4f}元")

            # 绩效明细（手机端只显示关键列）
            st.subheader("📋 绩效明细")
            display_cols = ['姓名', '工作条目', '总绩效', '个人实发']
            available_cols = [c for c in display_cols if c in df.columns]

            # 手机端用表格，桌面端可以展开看全部
            show_all = st.toggle("显示全部列", value=False)
            if show_all:
                all_cols = ['姓名', '工作条目', '门诊登记', '门诊收入院', '门诊建档',
                            '急诊建档', '急诊未入院', '典型病例', '总绩效', '个人实发']
                available_cols = [c for c in all_cols if c in df.columns]

            st.dataframe(df[available_cols], use_container_width=True, hide_index=True)

            # 排名（手机端友好）
            st.subheader("🏆 绩效排名")
            top5 = df.nlargest(5, '总绩效')[['姓名', '总绩效', '个人实发']]
            for i, row in top5.iterrows():
                st.write(f"**{row['姓名']}**: {row['总绩效']:.2f}元 (实发{row['个人实发']}元)")

            # 导出
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='数据库绩效', index=False)
            st.download_button(
                "📥 下载结果Excel",
                output.getvalue(),
                "数据库绩效.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ 错误：{str(e)}")


# ==================== 随访绩效模式 ====================
elif st.session_state.calc_mode == 'followup':
    st.markdown("---")
    st.subheader("📤 上传随访绩效数据")

    data = st.file_uploader(
        "选择Excel文件（需包含：姓名、工作条目、随访病例、微信入群）",
        type=["xlsx", "xls"],
        key="followup_upload"
    )

    if data:
        try:
            df, case_base_price, mgmt_total, total_wechat = process_followup_excel(data, FOLLOWUP_PERFORMANCE)
            st.session_state["df_followup"] = df

            st.success("✅ 计算完成！")

            # 关键指标
            st.subheader("📊 关键指标")
            c1, c2 = st.columns(2)
            c3, c4 = st.columns(2)
            with c1:
                st.metric("参与人数", f"{len(df[df['随访绩效'] > 0])}人")
            with c2:
                st.metric("随访总额", f"{df['随访绩效'].sum():.0f}元")
            with c3:
                st.metric("应发总额", f"{FOLLOWUP_PERFORMANCE}元")
            with c4:
                st.metric("病例单价", f"{case_base_price:.2f}元")

            # 校验
            diff = abs(df['随访绩效'].sum() - FOLLOWUP_PERFORMANCE)
            if diff > 1:
                st.warning(f"⚠️ 差异: {diff:.2f}元（可忽略）")

            # 明细
            st.subheader("📋 随访绩效明细")
            show_df = df[df['随访绩效'] > 0][['姓名', '随访绩效', '随访实发', '绩效明细']]
            st.dataframe(show_df, use_container_width=True, hide_index=True)

            # 各类明细
            with st.expander("🔍 查看分类明细"):
                st.write("**微信入群**")
                st.dataframe(df[df['微信入群绩效'] > 0][['姓名', '微信入群', '微信入群绩效']], hide_index=True)

                st.write("**随访管理**")
                st.dataframe(df[df['随访管理绩效'] > 0][['姓名', '工作条目', '随访管理绩效']], hide_index=True)

                st.write("**随访病例**")
                st.dataframe(df[df['随访病例绩效'] > 0][['姓名', '随访病例', '随访病例绩效']], hide_index=True)

            # 导出
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df[df['随访绩效'] > 0].to_excel(writer, sheet_name='随访绩效汇总', index=False)
                df.to_excel(writer, sheet_name='全部人员', index=False)

            st.download_button(
                "📥 下载结果Excel",
                output.getvalue(),
                "随访绩效.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ 错误：{str(e)}")


# ==================== 默认状态 ====================
else:
    st.markdown("---")
    st.info("👆 请点击上方按钮选择计算类型")

    # 使用说明（折叠）
    with st.expander("📖 使用说明"):
        st.markdown("""
        **1. 设置总绩效**
        - 点击左上角 `>` 展开侧边栏
        - 输入每月总绩效金额（默认3447）

        **2. 数据库绩效（75%）**
        - 上传含门诊数据的Excel
        - 系统自动计算工作量+质控绩效

        **3. 随访绩效（25%）**
        - 上传含随访病例的Excel
        - 系统自动计算病例+管理+入群绩效

        **4. 导出结果**
        - 计算完成后点击下载按钮
        """)

    # 快速预览图（可选）
    st.caption("💡 提示：可以把此网页添加到手机桌面，像App一样使用")
    st.caption("iPhone: Safari分享→添加到主屏幕 | Android: Chrome菜单→添加到主屏幕")