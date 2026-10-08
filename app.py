import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import timedelta, date
import calendar

st.set_page_config(page_title="渠道销量 & GMV 看板", layout="wide")
st.title("📊 渠道销量 & GMV 分析看板")
st.markdown("""
**口径说明**：
- 私域 = 优选商城 + 企业采购
- 公域 = 小红书 + 微信小店 + 抖店
- 只统计 **商品名称包含“轻舟眠”** 的订单
- **轻舟眠销售数量** = 销量 × 规格系数（单只装=1，一对装=2）
- GMV 同时展示 **有效GMV** 和 **券后GMV**
- 私域航道分析使用 **销地航道**
""")

uploaded_file = st.file_uploader("上传 Excel 文件（.xlsx）", type=["xlsx"])

if uploaded_file:
    # ---- dtype=str 避免超大整数溢出 ----
    df = pd.read_excel(uploaded_file, sheet_name=0, dtype=str)

    # ---- 1. 数据清洗 ----
    df.columns = df.columns.str.strip()
    required = ['销售渠道', '销量', '有效GMV', '券后GMV', '销地航道', '下单时间', '商品名称', '规格']
    missing = [col for col in required if col not in df.columns]
    if missing:
        st.error(f"Excel 缺少必须列：{', '.join(missing)}")
        st.stop()

    df['销量'] = pd.to_numeric(df['销量'], errors='coerce').fillna(0)
    df['有效GMV'] = pd.to_numeric(df['有效GMV'], errors='coerce').fillna(0)
    df['券后GMV'] = pd.to_numeric(df['券后GMV'], errors='coerce').fillna(0)
    df['下单时间'] = pd.to_datetime(df['下单时间'], errors='coerce')

    # ---- 2. 过滤：商品名称包含“轻舟眠” ----
    df['商品名称'] = df['商品名称'].fillna('')
    df = df[df['商品名称'].str.contains('轻舟眠', case=False, na=False)]
    if df.empty:
        st.warning("没有包含“轻舟眠”的订单，请检查数据。")
        st.stop()

    # ---- 3. 计算“轻舟眠销售数量” = 销量 × 规格系数 ----
    def get_spec_coef(spec):
        if pd.isna(spec):
            return 0
        spec = str(spec)
        if '单只装' in spec:
            return 1
        elif '一对装' in spec and '2只' in spec:
            return 2
        else:
            return 0

    df['规格系数'] = df['规格'].apply(get_spec_coef)
    df['轻舟眠销售数量'] = df['销量'] * df['规格系数']

    # ---- 4. 公域/私域分类 ----
    private_list = ['优选商城', '企业采购']
    public_list = ['小红书', '微信小店', '抖店']
    df['销售渠道_clean'] = df['销售渠道'].str.strip()

    def classify_channel(channel):
        if channel in private_list:
            return '私域'
        elif channel in public_list:
            return '公域'
        else:
            return '忽略'

    df['渠道类型'] = df['销售渠道_clean'].apply(classify_channel)
    df_filtered = df[df['渠道类型'].isin(['私域', '公域'])].copy()

    if df_filtered.empty:
        st.warning("未找到「优选商城、企业采购、小红书、微信小店、抖店」的订单。")
        st.stop()

    # ---- 5. 提取日期 ----
    df_filtered['日期'] = df_filtered['下单时间'].dt.date

    # ---- 6. 全局时间范围 ----
    global_min = df_filtered['下单时间'].min().date()
    global_max = df_filtered['下单时间'].max().date()

    # ===================== 日期选择器 =====================
    st.markdown("---")
    st.subheader("📅 选择分析时间段")

    # ---- 快捷按钮 ----
    today = date.today()

    def first_day_of_month(d):
        return d.replace(day=1)

    def last_day_of_month(d):
        return d.replace(day=calendar.monthrange(d.year, d.month)[1])

    def clamp(d, lo, hi):
        return max(min(d, hi), lo)

    # 数据变化时重置 session_state
    range_key = f"{global_min}_{global_max}"
    if st.session_state.get('range_key') != range_key:
        st.session_state.start_date = global_min
        st.session_state.end_date = global_max
        st.session_state.range_key = range_key

    # 快捷按钮区
    col_q1, col_q2, col_q3, col_q4, col_q5 = st.columns(5)

    with col_q1:
        if st.button("📆 本月", use_container_width=True):
            st.session_state.start_date = clamp(first_day_of_month(today), global_min, global_max)
            st.session_state.end_date = clamp(last_day_of_month(today), global_min, global_max)
    with col_q2:
        if st.button("📅 上月", use_container_width=True):
            last_m = first_day_of_month(today) - timedelta(days=1)
            st.session_state.start_date = clamp(first_day_of_month(last_m), global_min, global_max)
            st.session_state.end_date = clamp(last_day_of_month(last_m), global_min, global_max)
    with col_q3:
        if st.button("🗓 近7天", use_container_width=True):
            st.session_state.start_date = clamp(today - timedelta(days=6), global_min, global_max)
            st.session_state.end_date = clamp(today, global_min, global_max)
    with col_q4:
        if st.button("🗓 近30天", use_container_width=True):
            st.session_state.start_date = clamp(today - timedelta(days=29), global_min, global_max)
            st.session_state.end_date = clamp(today, global_min, global_max)
    with col_q5:
        if st.button("📊 全部", use_container_width=True):
            st.session_state.start_date = global_min
            st.session_state.end_date = global_max

    # 日期选择器
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        start_date = st.date_input(
            "开始日期",
            min_value=global_min,
            max_value=global_max,
            key='start_date'
        )
    with col_d2:
        end_date = st.date_input(
            "结束日期",
            min_value=global_min,
            max_value=global_max,
            key='end_date'
        )

    if start_date > end_date:
        st.error("开始日期不能晚于结束日期，请重新选择。")
        st.stop()

    # ---- 7. 根据选择的时间段筛选数据 ----
    mask_period = (df_filtered['日期'] >= start_date) & (df_filtered['日期'] <= end_date)
    current_df = df_filtered[mask_period].copy()
    current_days = (end_date - start_date).days + 1

    if current_df.empty:
        st.warning("所选时间段内没有数据，请调整日期范围。")
        st.stop()

    st.caption(f"当前显示时间段：{start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')} （共 {current_days} 天，{len(current_df)} 条订单）")

    # ===================== 环比分析 =====================
    st.markdown("---")
    st.subheader("📈 环比分析（与上一对比周期对比）")

    # ---- 对比方式选择 ----
    compare_mode = st.radio(
        "选择对比方式：",
        ['智能（整月自动对比上月）', '按自然月', '按天数'],
        horizontal=True,
        index=0
    )

    def is_full_month(sd, ed):
        if sd.day != 1:
            return False
        last_day = calendar.monthrange(ed.year, ed.month)[1]
        if ed.day != last_day:
            return False
        if sd.year != ed.year or sd.month != ed.month:
            return False
        return True

    def get_prev_period(sd, ed, mode):
        if mode == '智能（整月自动对比上月）':
            actual_mode = 'month' if is_full_month(sd, ed) else 'days'
        elif mode == '按自然月':
            actual_mode = 'month'
        else:
            actual_mode = 'days'

        if actual_mode == 'month':
            first_of_this_month = sd.replace(day=1)
            prev_month_end = first_of_this_month - timedelta(days=1)
            prev_month_start = prev_month_end.replace(day=1)
            return prev_month_start, prev_month_end, '自然月'
        else:
            days = (ed - sd).days + 1
            prev_end = sd - timedelta(days=1)
            prev_start = prev_end - timedelta(days=days - 1)
            return prev_start, prev_end, '按天数'

    prev_start, prev_end, actual_mode = get_prev_period(start_date, end_date, compare_mode)

    mask_prev = (df_filtered['日期'] >= prev_start) & (df_filtered['日期'] <= prev_end)
    prev_df = df_filtered[mask_prev]

    def agg_metrics(data):
        if data.empty:
            return {'销量': 0, '有效GMV': 0, '券后GMV': 0, '轻舟眠销售数量': 0}
        return {
            '销量': data['销量'].sum(),
            '有效GMV': data['有效GMV'].sum(),
            '券后GMV': data['券后GMV'].sum(),
            '轻舟眠销售数量': data['轻舟眠销售数量'].sum()
        }

    current_metrics = agg_metrics(current_df)
    prev_metrics = agg_metrics(prev_df)

    def calc_mom(current, prev):
        if prev == 0:
            return float('inf') if current > 0 else 0
        return (current - prev) / prev * 100

    st.caption(f"对比周期（{actual_mode}）：{prev_start.strftime('%Y-%m-%d')} ~ {prev_end.strftime('%Y-%m-%d')}（若无数据则显示为0）")

    # ---- 四个指标卡 ----
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)

    delta_val = current_metrics['有效GMV'] - prev_metrics['有效GMV']
    pct = calc_mom(current_metrics['有效GMV'], prev_metrics['有效GMV'])
    delta_str = f"{delta_val:+,.2f} ({pct:+.2f}%)" if pct != float('inf') else "无上期数据"
    with col_m1:
        st.metric("💰 有效GMV", f"{current_metrics['有效GMV']:,.2f}", delta=delta_str)

    delta_val = current_metrics['券后GMV'] - prev_metrics['券后GMV']
    pct = calc_mom(current_metrics['券后GMV'], prev_metrics['券后GMV'])
    delta_str = f"{delta_val:+,.2f} ({pct:+.2f}%)" if pct != float('inf') else "无上期数据"
    with col_m2:
        st.metric("🎫 券后GMV", f"{current_metrics['券后GMV']:,.2f}", delta=delta_str)

    delta_val = current_metrics['销量'] - prev_metrics['销量']
    pct = calc_mom(current_metrics['销量'], prev_metrics['销量'])
    delta_str = f"{delta_val:+,.0f} ({pct:+.2f}%)" if pct != float('inf') else "无上期数据"
    with col_m3:
        st.metric("📦 销量（订单件数）", f"{current_metrics['销量']:,.0f}", delta=delta_str)

    delta_val = current_metrics['轻舟眠销售数量'] - prev_metrics['轻舟眠销售数量']
    pct = calc_mom(current_metrics['轻舟眠销售数量'], prev_metrics['轻舟眠销售数量'])
    delta_str = f"{delta_val:+,.0f} ({pct:+.2f}%)" if pct != float('inf') else "无上期数据"
    with col_m4:
        st.metric("🛏️ 轻舟眠销售数量", f"{current_metrics['轻舟眠销售数量']:,.0f}", delta=delta_str)

    # ---- 环比对比明细表 ----
    compare_df = pd.DataFrame({
        '指标': ['销量', '有效GMV', '券后GMV', '轻舟眠销售数量'],
        '当前周期': [current_metrics['销量'], current_metrics['有效GMV'], current_metrics['券后GMV'], current_metrics['轻舟眠销售数量']],
        '对比周期': [prev_metrics['销量'], prev_metrics['有效GMV'], prev_metrics['券后GMV'], prev_metrics['轻舟眠销售数量']],
    })
    compare_df['变化值'] = (compare_df['当前周期'] - compare_df['对比周期']).round(2)
    compare_df['环比数值'] = compare_df.apply(
        lambda row: (row['变化值'] / row['对比周期'] * 100) if row['对比周期'] != 0 else (float('inf') if row['当前周期'] > 0 else 0),
        axis=1
    )
    compare_df['环比(%)'] = compare_df['环比数值'].apply(lambda x: f"{x:+.2f}%" if x != float('inf') else "无上期数据")

    def color_change(val):
        if isinstance(val, (int, float)):
            if val > 0:
                return 'color: green'
            elif val < 0:
                return 'color: red'
        return ''

    def color_pct(val):
        if isinstance(val, str):
            if val.startswith('-'):
                return 'color: red'
            elif val != "无上期数据" and not val.startswith('-'):
                return 'color: green'
        return ''

    styled_compare = (
        compare_df.style
        .map(color_change, subset=['变化值'])
        .map(color_pct, subset=['环比(%)'])
        .format({
            '当前周期': '{:,.2f}',
            '对比周期': '{:,.2f}',
            '变化值': '{:,.2f}'
        })
    )

    st.caption("📋 当前周期 vs 对比周期 明细对比")
    st.dataframe(styled_compare, use_container_width=True, hide_index=True)

    st.markdown("---")

    # ===================== 以下所有展示均基于 current_df =====================

    # ---- 公域 vs 私域 总览 ----
    st.subheader("公域 vs 私域 总览")
    type_summary = current_df.groupby('渠道类型').agg({
        '销量': 'sum', '有效GMV': 'sum', '券后GMV': 'sum', '轻舟眠销售数量': 'sum'
    }).reset_index()

    col1, col2 = st.columns(2)
    with col1:
        fig1 = px.pie(type_summary, values='有效GMV', names='渠道类型', title='有效GMV占比')
        st.plotly_chart(fig1, use_container_width=True)
    with col2:
        fig2 = px.pie(type_summary, values='券后GMV', names='渠道类型', title='券后GMV占比')
        st.plotly_chart(fig2, use_container_width=True)

    st.dataframe(type_summary.style.format({'销量': '{:,.0f}', '有效GMV': '{:,.2f}', '券后GMV': '{:,.2f}', '轻舟眠销售数量': '{:,.0f}'}), use_container_width=True)
    st.markdown("---")

    # ---- 趋势图：公域 vs 私域 每日趋势 ----
    st.subheader("📈 公域 vs 私域 每日趋势")
    trend_overview = current_df.groupby(['日期', '渠道类型']).agg({
        '销量': 'sum', '有效GMV': 'sum', '券后GMV': 'sum', '轻舟眠销售数量': 'sum'
    }).reset_index()

    if not trend_overview.empty:
        metric_choice = st.radio("选择趋势图指标：", ['有效GMV', '券后GMV', '销量', '轻舟眠销售数量'], horizontal=True)
        fig_trend1 = px.line(trend_overview, x='日期', y=metric_choice, color='渠道类型',
                             title=f'每日 {metric_choice} 趋势（公域 vs 私域）', markers=True)
        st.plotly_chart(fig_trend1, use_container_width=True)
    else:
        st.info("所选时间段内无每日数据（可能只有一天），无法绘制趋势图。")
    st.markdown("---")

    # ---- 各渠道明细 ----
    st.subheader("各渠道明细")
    channel_summary = current_df.groupby('销售渠道_clean').agg({
        '销量': 'sum', '有效GMV': 'sum', '券后GMV': 'sum', '轻舟眠销售数量': 'sum'
    }).reset_index().rename(columns={'销售渠道_clean': '销售渠道'})
    fig3 = px.bar(channel_summary, x='销售渠道', y=['销量', '有效GMV', '券后GMV', '轻舟眠销售数量'], barmode='group')
    st.plotly_chart(fig3, use_container_width=True)
    st.dataframe(channel_summary.style.format({'销量': '{:,.0f}', '有效GMV': '{:,.2f}', '券后GMV': '{:,.2f}', '轻舟眠销售数量': '{:,.0f}'}), use_container_width=True)
    st.markdown("---")

    # ---- 私域航道分析 ----
    st.subheader("🔍 私域航道分析（按销地航道）")
    private_df = current_df[current_df['渠道类型'] == '私域']
    if not private_df.empty:
        channel_summary_private = private_df.groupby('销地航道').agg({
            '销量': 'sum', '有效GMV': 'sum', '券后GMV': 'sum', '轻舟眠销售数量': 'sum'
        }).reset_index()

        col3, col4 = st.columns(2)
        with col3:
            fig4 = px.pie(channel_summary_private, values='有效GMV', names='销地航道', title='私域-各航道有效GMV占比')
            st.plotly_chart(fig4, use_container_width=True)
        with col4:
            fig5 = px.pie(channel_summary_private, values='券后GMV', names='销地航道', title='私域-各航道券后GMV占比')
            st.plotly_chart(fig5, use_container_width=True)

        st.dataframe(channel_summary_private.style.format({'销量': '{:,.0f}', '有效GMV': '{:,.2f}', '券后GMV': '{:,.2f}', '轻舟眠销售数量': '{:,.0f}'}), use_container_width=True)

        # 私域航道趋势
        st.subheader("📈 私域各销地航道 每日趋势")
        trend_private = private_df.groupby(['日期', '销地航道']).agg({
            '销量': 'sum', '有效GMV': 'sum', '券后GMV': 'sum', '轻舟眠销售数量': 'sum'
        }).reset_index()
        if not trend_private.empty:
            metric_choice_private = st.radio("选择航道趋势指标：", ['有效GMV', '券后GMV', '销量', '轻舟眠销售数量'], horizontal=True, key="private_metric")
            fig_trend2 = px.line(trend_private, x='日期', y=metric_choice_private, color='销地航道',
                                 title=f'私域各航道每日 {metric_choice_private} 趋势', markers=True)
            st.plotly_chart(fig_trend2, use_container_width=True)
        else:
            st.info("所选时间段内无私域航道每日数据。")
    else:
        st.info("所选时间段内无私域数据。")

    # 调试信息（防止大整数溢出，统一转字符串）
    with st.expander("🔍 查看当前所选时间段数据预览（前10行）"):
        preview_df = current_df.head(10).copy()
        for col in preview_df.columns:
            if preview_df[col].dtype == 'object':
                preview_df[col] = preview_df[col].astype(str)
        st.dataframe(preview_df)

else:
    st.info("👆 请上传 Excel 文件以开始分析")
