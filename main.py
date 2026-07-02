
import pandas as pd
import streamlit as st
from io import BytesIO


# ==================== 数据库绩效函数 ====================

def calculate_performance(row, a, base_price):
    """计算个人总绩效 - 全部按基数分配"""
    门诊登记 = float(row['门诊登记']) if pd.notna(row['门诊登记']) else 0
    门诊收入院 = float(row['门诊收入院']) if pd.notna(row['门诊收入院']) else 0
    门诊建档 = float(row['门诊建档']) if pd.notna(row['门诊建档']) else 0
    急诊建档 = float(row['急诊建档']) if pd.notna(row['急诊建档']) else 0
    急诊未入院 = float(row['急诊未入院']) if pd.notna(row['急诊未入院']) else 0
    典型病例 = float(row['典型病例']) if pd.notna(row['典型病例']) else 0
    工作条目 = str(row['工作条目']) if pd.notna(row['工作条目']) else ''

    # 基础工作量绩效
    work_score = (
            门诊登记 * 0.1 +
            门诊收入院 * 1 +
            门诊建档 * 0.5 +
            急诊建档 * 1 +
            急诊未入院 * 1
    )
    work_performance = work_score * base_price

    # 质控绩效 - 全部按基数分配
    qc_performance = 0
    if '数据录入' in 工作条目:
        qc_performance = a * 1.7 * base_price
    elif '审核' in 工作条目:
        qc_performance = a * 0.2 * base_price
    elif '归档' in 工作条目:
        qc_performance = a * 0.2 * 0.8 * base_price

    # 典型病例 - 按基数算（如果典型病例列有值）
    case_bonus = 典型病例 * base_price if 典型病例 > 0 else 0

    return round(work_performance + qc_performance + case_bonus, 2)


def find_header_row(uploaded_file):
    """自动检测表头行位置"""
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


def process_database_excel(uploaded_file, database_performance, fixed_bonus):
    """读取Excel并计算数据库绩效"""
    skip_rows, df = find_header_row(uploaded_file)

    if df is None:
        df_raw = pd.read_excel(uploaded_file, header=None)
        raise Exception(f"无法找到表头。前5行原始数据：\n{df_raw.head().to_string()}")

    df.columns = [str(c).strip().replace('\n', '').replace(' ', '') for c in df.columns]

    col_mapping = {}
    for std_col in ['工作条目', '姓名', '门诊登记', '门诊收入院', '门诊建档',
                    '急诊建档', '急诊未入院', '典型病例']:
        for actual_col in df.columns:
            if std_col in actual_col:
                col_mapping[std_col] = actual_col
                break

    df = df.rename(columns=col_mapping)

    # a = 门诊建档 + 急诊建档
    a = df['门诊建档'].fillna(0).astype(float).sum() + df['急诊建档'].fillna(0).astype(float).sum()

    # 工作量基数
    work_base_sum = (
            df['门诊登记'].fillna(0).astype(float).sum() * 0.1 +
            df['门诊收入院'].fillna(0).astype(float).sum() * 1 +
            df['门诊建档'].fillna(0).astype(float).sum() * 0.5 +
            df['急诊建档'].fillna(0).astype(float).sum() * 1 +
            df['急诊未入院'].fillna(0).astype(float).sum() * 1
    )

    # 质控基数
    qc_base_sum = 0
    if any('数据录入' in str(x) for x in df['工作条目'].fillna('')):
        qc_base_sum += a * 1.7
    if any('审核' in str(x) for x in df['工作条目'].fillna('')):
        qc_base_sum += a * 0.2
    if any('归档' in str(x) for x in df['工作条目'].fillna('')):
        qc_base_sum += a * 0.2 * 0.8

    total_base = work_base_sum + qc_base_sum

    # 【修正】基数单价 = (数据库绩效 - 固定奖励) / 总基数
    # 原代码：base_price = database_performance / total_base
    # 修正后：base_price = (database_performance - fixed_bonus) / total_base
    base_price = (database_performance - fixed_bonus) / total_base if total_base > 0 else 0

    # 计算绩效（基数部分）
    df['基数绩效'] = df.apply(lambda row: calculate_performance(row, a, base_price), axis=1)

    # 郝芳额外加300元固定奖励（数据录入岗位）
    df['固定奖励'] = 0.0
    df.loc[df['姓名'] == '郝芳', '固定奖励'] = fixed_bonus

    # 最终绩效 = 基数绩效 + 固定奖励
    df['总绩效'] = df['基数绩效'] + df['固定奖励']
    df['个人实发'] = df['总绩效'].round(0).astype(int)

    return df, a, base_price, work_base_sum, qc_base_sum, total_base


# ==================== 随访绩效函数（从Excel读取数据）====================

def process_followup_excel(uploaded_file, followup_performance):
    """
    读取随访绩效Excel并计算
    Excel列：姓名、工作条目、随访病例、微信入群
    """
    df = pd.read_excel(uploaded_file)

    # 标准化列名
    df.columns = [str(c).strip().replace('\n', '').replace(' ', '') for c in df.columns]

    # 列名映射
    col_mapping = {}
    for std_col in ['姓名', '工作条目', '随访病例', '微信入群']:
        for actual_col in df.columns:
            if std_col in actual_col:
                col_mapping[std_col] = actual_col
                break

    df = df.rename(columns=col_mapping)

    # 确保数值列正确
    df['随访病例'] = pd.to_numeric(df['随访病例'], errors='coerce').fillna(0)
    df['微信入群'] = pd.to_numeric(df['微信入群'], errors='coerce').fillna(0)
    df['工作条目'] = df['工作条目'].fillna('').astype(str)

    # 计算随访病例总绩效基数
    total_case_base = df['随访病例'].sum()

    # 计算微信入群总额
    total_wechat = df['微信入群'].sum()

    # 剩余金额 = 随访绩效总额 - 微信入群总额
    remaining = followup_performance - total_wechat

    # 随访管理占剩余的13%
    mgmt_ratio = 0.13
    mgmt_total = remaining * mgmt_ratio

    # 病例绩效占剩余的87%
    case_total = remaining * (1 - mgmt_ratio)

    # 病例绩效基数单价
    case_base_price = case_total / total_case_base if total_case_base > 0 else 0

    # 先计算微信入群（固定值）
    df['微信入群绩效'] = df['微信入群']

    # 计算随访管理绩效
    df['随访管理绩效'] = 0.0
    mgmt_mask = df['工作条目'].str.contains('随访管理', na=False)
    if mgmt_mask.sum() > 0:
        df.loc[mgmt_mask, '随访管理绩效'] = mgmt_total / mgmt_mask.sum()

    # 计算随访病例绩效
    df['随访病例绩效'] = df['随访病例'] * case_base_price

    # 随访绩效合计
    df['随访绩效'] = df['微信入群绩效'] + df['随访管理绩效'] + df['随访病例绩效']
    df['随访绩效'] = df['随访绩效'].round(2)

    # 个人实发（四舍五入到整数）
    df['随访实发'] = df['随访绩效'].round(0).astype(int)

    # 添加绩效明细说明
    def get_detail(row):
        parts = []
        if row['微信入群绩效'] > 0:
            parts.append(f"微信入群:{row['微信入群绩效']:.2f}")
        if row['随访管理绩效'] > 0:
            parts.append(f"随访管理:{row['随访管理绩效']:.2f}")
        if row['随访病例绩效'] > 0:
            parts.append(f"病例{row['随访病例']:.0f}份×{case_base_price:.2f}={row['随访病例绩效']:.2f}")
        return ' | '.join(parts) if parts else '无'

    df['绩效明细'] = df.apply(get_detail, axis=1)

    return df, case_base_price, mgmt_total, total_wechat


# ==================== Streamlit 页面 ====================
st.set_page_config(page_title="西山医院绩效系统", page_icon="🏥", layout="wide")

st.title("💡 医院胸痛中心数据库绩效系统")

# ========== 初始化 session_state ==========
if "calc_mode" not in st.session_state:
    st.session_state.calc_mode = None

# ========== 侧边栏参数设置 ==========
st.sidebar.subheader("⚙️ 绩效参数设置")

# 总绩效输入
TOTAL_PERFORMANCE = st.sidebar.number_input(
    "💰 总绩效（元）",
    value=3447,
    step=1,
    help="每月胸痛中心总绩效，如3447"
)

# 固定扣除（郝芳做PPT等）
FIXED_DEDUCTION = st.sidebar.number_input(
    "固定扣除（元）",
    value=300,
    step=1,
    help="质控会、典型病例讨论、门诊幻灯、绩效统计、ACS汇总表等固定工作"
)

# 郝芳固定奖励（包含在数据库绩效内）
HAO_BONUS = st.sidebar.number_input(
    "郝芳固定奖励（元）",
    value=300,
    step=1,
    help="郝芳额外固定奖励，包含在数据库绩效总额内"
)

# 扣除后剩余
REMAINING = TOTAL_PERFORMANCE - FIXED_DEDUCTION

# 分配比例
DATABASE_PERFORMANCE = round(REMAINING * 0.75 + FIXED_DEDUCTION, 2)
FOLLOWUP_PERFORMANCE = round(REMAINING * 0.25, 2)

st.sidebar.markdown("---")
st.sidebar.markdown(f"**总绩效**: `{TOTAL_PERFORMANCE}` 元")
st.sidebar.markdown(f"**固定扣除**: `{FIXED_DEDUCTION}` 元")
st.sidebar.markdown(f"**剩余分配**: `{REMAINING}` 元")
st.sidebar.markdown(f"**数据库绩效(75%+固定)**: `{DATABASE_PERFORMANCE}` 元")
st.sidebar.markdown(f"**随访绩效(25%)**: `{FOLLOWUP_PERFORMANCE}` 元")
st.sidebar.markdown(f"**郝芳固定奖励**: `{HAO_BONUS}` 元")
st.sidebar.markdown("---")

# 两个计算按钮
st.sidebar.subheader("🚀 选择计算类型")

col1, col2 = st.sidebar.columns(2)
with col1:
    if st.button("📊 数据库绩效", use_container_width=True):
        st.session_state.calc_mode = 'database'
        st.rerun()
with col2:
    if st.button("📋 随访绩效", use_container_width=True):
        st.session_state.calc_mode = 'followup'
        st.rerun()

# ========== 数据库绩效模式 ==========
if st.session_state.calc_mode == 'database':
    st.header("📊 数据库绩效计算")

    data = st.file_uploader("上传数据库绩效数据（Excel格式）：", type=["xlsx", "xls"], key="db_upload")

    if data:
        try:
            df, a, base_price, work_base, qc_base, total_base = process_database_excel(data, DATABASE_PERFORMANCE, HAO_BONUS)
            st.session_state["df_db"] = df
            st.session_state["a"] = a
            st.session_state["base_price"] = base_price

            st.success("✅ 数据库绩效计算完成！")

            # 计算参数
            st.subheader("📊 计算参数")
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                st.metric("数据库绩效", f"{DATABASE_PERFORMANCE}元")
            with c2:
                st.metric("工作量基数", f"{work_base:.2f}")
            with c3:
                st.metric("质控基数", f"{qc_base:.2f}")
            with c4:
                st.metric("总基数", f"{total_base:.2f}")
            with c5:
                st.metric("绩效基数单价", f"{base_price:.2f}元")

            # 关键指标
            st.subheader("📈 关键指标")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("👥 总人数", f"{len(df)}人")
            with col2:
                st.metric("💰 基数绩效总和", f"{df['基数绩效'].sum():.2f}元")
            with col3:
                st.metric("💵 实发总数", f"{df['个人实发'].sum()}元")
            with col4:
                st.metric("📈 人均绩效", f"{df['总绩效'].mean():.2f}元")

            # 绩效明细
            st.subheader("📋 绩效明细")
            display_cols = ['姓名', '工作条目', '门诊登记', '门诊收入院', '门诊建档',
                            '急诊建档', '急诊未入院', '典型病例', '基数绩效', '固定奖励', '总绩效', '个人实发']
            available_cols = [c for c in display_cols if c in df.columns]


            def highlight_max(s):
                is_max = s == s.max()
                return ['background-color: #ffeb3b' if v else '' for v in is_max]


            styled_df = df[available_cols].style.apply(highlight_max, subset=['总绩效'])
            st.dataframe(styled_df, use_container_width=True, height=400)

            # 可视化
            st.subheader("📈 数据可视化")
            tab1, tab2, tab3, tab4 = st.tabs(["绩效排名", "岗位分布", "工作量分析", "绩效构成"])
            with tab1:
                top10 = df.nlargest(10, '总绩效')[['姓名', '总绩效']].set_index('姓名')
                st.bar_chart(top10)
            with tab2:
                position_stats = df.groupby('工作条目')['总绩效'].sum().sort_values(ascending=False)
                st.bar_chart(position_stats)
                st.write("岗位人数：", df['工作条目'].value_counts())
            with tab3:
                work_cols = ['门诊登记', '门诊收入院', '门诊建档', '急诊建档', '急诊未入院']
                work_sum = df[work_cols].sum()
                st.bar_chart(work_sum)
            with tab4:
                bins = [0, 50, 100, 200, 500, 1000, 2000, 5000]
                labels = ['0-50', '50-100', '100-200', '200-500', '500-1000', '1000-2000', '2000+']
                df['绩效区间'] = pd.cut(df['总绩效'], bins=bins, labels=labels)
                interval_count = df['绩效区间'].value_counts().sort_index()
                st.bar_chart(interval_count)

            # 导出
            st.subheader("💾 导出数据")
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df[available_cols].to_excel(writer, sheet_name='数据库绩效', index=False)
            st.download_button("📥 下载数据库绩效Excel", output.getvalue(), "数据库绩效.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"❌ 文件读取失败：{str(e)}")
            with st.expander("🔍 调试：查看原始数据前10行"):
                try:
                    df_debug = pd.read_excel(data, header=None)
                    st.write(df_debug.head(10).to_string())
                except Exception as e2:
                    st.write(f"无法读取原始数据：{e2}")


# ========== 随访绩效模式（从Excel读取数据）==========
elif st.session_state.calc_mode == 'followup':
    st.header("📋 随访绩效计算")

    data = st.file_uploader("上传随访绩效数据（Excel格式，需包含：姓名、工作条目、随访病例、微信入群）：",
                            type=["xlsx", "xls"], key="followup_upload")

    if data:
        try:
            df, case_base_price, mgmt_total, total_wechat = process_followup_excel(data, FOLLOWUP_PERFORMANCE)
            st.session_state["df_followup"] = df

            st.success("✅ 随访绩效计算完成！")

            # 关键指标
            st.subheader("📈 随访绩效关键指标")
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                st.metric("👥 总人数", f"{len(df)}人")
            with c2:
                st.metric("💰 随访绩效总额", f"{df['随访绩效'].sum():.2f}元")
            with c3:
                st.metric("💵 随访实发总数", f"{df['随访实发'].sum()}元")
            with c4:
                st.metric("📊 应发随访绩效", f"{FOLLOWUP_PERFORMANCE}元")
            with c5:
                st.metric("📋 病例基数单价", f"{case_base_price:.2f}元")

            # 校验
            diff = abs(df['随访绩效'].sum() - FOLLOWUP_PERFORMANCE)
            if diff > 1:
                st.warning(f"⚠️ 实发({df['随访绩效'].sum():.2f})与应发({FOLLOWUP_PERFORMANCE})差异: {diff:.2f}元")
            else:
                st.success(f"✅ 实发与应发基本吻合，差异: {diff:.2f}元")

            # 随访绩效明细表
            st.subheader("📋 随访绩效明细")
            display_cols = ['姓名', '工作条目', '随访病例', '微信入群',
                            '微信入群绩效', '随访管理绩效', '随访病例绩效',
                            '随访绩效', '随访实发', '绩效明细']
            available_cols = [c for c in display_cols if c in df.columns]

            # 筛选有绩效的人员显示
            df_show = df[df['随访绩效'] > 0].copy() if True else df


            def highlight_followup(s):
                is_max = s == s.max()
                return ['background-color: #90EE90' if v else '' for v in is_max]


            styled_summary = df_show[available_cols].style.apply(highlight_followup, subset=['随访绩效'])
            st.dataframe(styled_summary, use_container_width=True, height=400)

            # 全部人员（含0绩效）
            with st.expander("📋 查看全部人员（含0绩效）"):
                st.dataframe(df[['姓名', '工作条目', '随访病例', '微信入群', '随访绩效', '随访实发', '绩效明细']],
                             use_container_width=True)

            # 随访绩效构成分析
            st.subheader("📊 随访绩效构成分析")

            type_stats = {
                '微信入群': df['微信入群绩效'].sum(),
                '随访管理': df['随访管理绩效'].sum(),
                '随访病例': df['随访病例绩效'].sum()
            }
            st.bar_chart(pd.Series(type_stats))

            # 显示各类明细
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.markdown("**微信入群明细**")
                wechat_df = df[df['微信入群绩效'] > 0][['姓名', '微信入群', '微信入群绩效']]
                st.dataframe(wechat_df, use_container_width=True, hide_index=True)
            with col_b:
                st.markdown("**随访管理明细**")
                mgmt_df = df[df['随访管理绩效'] > 0][['姓名', '工作条目', '随访管理绩效']]
                st.dataframe(mgmt_df, use_container_width=True, hide_index=True)
            with col_c:
                st.markdown("**随访病例明细**")
                case_df = df[df['随访病例绩效'] > 0][['姓名', '随访病例', '随访病例绩效']]
                st.dataframe(case_df, use_container_width=True, hide_index=True)

            # 导出
            st.subheader("💾 导出数据")
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Sheet1: 有绩效人员汇总
                df_show[available_cols].to_excel(writer, sheet_name='随访绩效汇总', index=False)
                # Sheet2: 全部人员
                df[['姓名', '工作条目', '随访病例', '微信入群', '随访绩效', '随访实发', '绩效明细']].to_excel(
                    writer, sheet_name='全部人员', index=False)
                # Sheet3: 配置参数
                config_df = pd.DataFrame({
                    '项目': ['总绩效', '固定扣除', '剩余分配',
                             '数据库绩效(75%+固定)', '随访绩效(25%)',
                             '病例基数单价', '随访管理总额', '微信入群总额',
                             '随访绩效实发', '随访绩效应发', '差异'],
                    '数值': [
                        TOTAL_PERFORMANCE,
                        FIXED_DEDUCTION,
                        REMAINING,
                        DATABASE_PERFORMANCE,
                        FOLLOWUP_PERFORMANCE,
                        round(case_base_price, 4),
                        round(mgmt_total, 2),
                        round(total_wechat, 2),
                        round(df['随访绩效'].sum(), 2),
                        FOLLOWUP_PERFORMANCE,
                        round(abs(df['随访绩效'].sum() - FOLLOWUP_PERFORMANCE), 2)
                    ]
                })
                config_df.to_excel(writer, sheet_name='汇总统计', index=False)

            st.download_button("📥 下载随访绩效Excel", output.getvalue(), "随访绩效.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"❌ 文件读取失败：{str(e)}")
            with st.expander("🔍 调试：查看原始数据前10行"):
                try:
                    df_debug = pd.read_excel(data, header=None)
                    st.write(df_debug.head(10).to_string())
                except Exception as e2:
                    st.write(f"无法读取原始数据：{e2}")


# ========== 默认状态 ==========
else:
    st.info("👈 请在左侧边栏选择计算类型：「数据库绩效」或「随访绩效」")

    st.subheader("📖 使用说明")
    st.markdown("""
    **绩效分配规则：**
    1. 总绩效先扣除固定部分（郝芳做PPT、典型病例讨论、门诊幻灯、绩效统计、ACS汇总表）
    2. 剩余部分按75%/25%分配：
       - 数据库绩效 = 剩余 × 75% + 固定扣除
       - 随访绩效 = 剩余 × 25%

    **3. 数据库绩效计算（已修正）：**
    - 绩效基数 = 工作量基数 + 质控基数
    - 工作量基数 = 门诊登记×0.1 + 门诊收入院×1 + 门诊建档×0.5 + 急诊建档×1 + 急诊未入院×1
    - 质控基数：
      - 数据录入：a × 1.7
      - 审核：a × 0.2
      - 归档：a × 0.2 × 0.8
    - **基数单价 = (数据库绩效 - 郝芳固定奖励) / 总基数** ← 【关键修正】
    - 郝芳额外加固定奖励（包含在数据库绩效内）

    **4. 随访绩效（25%）**
    - 上传包含姓名、工作条目、随访病例、微信入群的Excel
    - 系统自动从Excel读取数据计算随访绩效
    """)
