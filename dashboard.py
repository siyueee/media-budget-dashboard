import streamlit as st
import pandas as pd
from pathlib import Path
import altair as alt
from datetime import timedelta
import base64

# --- 1. 页面配置 ---
st.set_page_config(page_title="媒体预算明细看板", layout="wide")

# 辅助函数：将本地图片转为 base64 以便在 HTML 中显示
def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

# 注入 CSS：增加卡片阴影、悬浮效果和商务蓝配色
st.markdown("""
    <style>
    /* 指标卡片容器样式 */
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 15px !important;
        border-radius: 10px !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06); 
        transition: transform 0.2s ease-in-out;
    }
    /* 鼠标悬停提升效果 */
    div[data-testid="stMetric"]:hover {
        transform: translateY(-4px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    /* 数字颜色改为商务蓝 */
    div[data-testid="stMetricValue"] > div {
        color: #1E40AF !important; 
    }
    /* 指标标题动态缩放效果 */
    .header-gif {
        transition: transform 0.3s ease;
        cursor: pointer;
    }
    .header-gif:hover {
        transform: scale(1.2) rotate(5deg);
    }
    /* 指标标签颜色加深 */
    div[data-testid="stMetricLabel"] > div > p {
        color: #475569 !important;
        font-weight: 600;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 标题与多个本地 GIF 整合 ---
GIF1_PATH = "吉伊bb.gif"
GIF2_PATH = "吉伊bb2.gif"

# 构造 HTML 内容
title_html = '<div style="display: flex; align-items: center; margin-bottom: 15px;">'
title_html += '<h1 style="margin: 0; font-size: 2.8rem;">🛰️媒体预算全平台明细看板</h1>'

# 检查并添加第一个 GIF
#if Path(GIF1_PATH).exists():
#   bin_str1 = get_base64_of_bin_file(GIF1_PATH)
 #   title_html += f'<img src="data:image/gif;base64,{bin_str1}" class="header-gif" width="150" style="margin-left: 25px;">'

# 检查并添加第二个 GIF
if Path(GIF2_PATH).exists():
    bin_str2 = get_base64_of_bin_file(GIF2_PATH)
    title_html += f'<img src="data:image/gif;base64,{bin_str2}" class="header-gif" width="150" style="margin-left: 15px;">'

title_html += '</div>'

# 渲染
st.markdown(title_html, unsafe_allow_html=True)

# --- 2. 数据加载与清洗 ---
@st.cache_data
def load_data(file_path):
    try:
        excel_file = pd.ExcelFile(file_path)
        all_dfs = [pd.read_excel(file_path, sheet_name=sn) for sn in excel_file.sheet_names]
        if not all_dfs: return pd.DataFrame()
        df = pd.concat(all_dfs, ignore_index=True)

        name_map = {
            '广告主激活量': '激活量', '唤醒数': '唤醒量', '次日回访量': '次留数',
            '2日留存量': '2留数', '二日留存': '2留数', '2日留存数': '2留数',
            '新增量': '新登数', '新登量': '新登数', '下单数': '下单量',
            '付费数': '付费数', '首购数': '首购量',
            '上报广告主曝光数': '曝光', '上报广告主次数': '点击'  # 文案统一映射
        }
        df.rename(columns={k: v for k, v in name_map.items() if k in df.columns}, inplace=True)

        def clean_val(v):
            if pd.isna(v): return 0.0
            if isinstance(v, (int, float)): return float(v)
            s = str(v).replace('%', '').replace('¥', '').replace(',', '').strip()
            try:
                val = float(s)
                return val / 100.0 if '%' in str(v) or val > 1.0 else val
            except:
                return 0.0

        num_cols = ['合作价格', '激活量', '次留数', '2留数', '唤醒量', '下单量', '付费数', '首购量', '新登数',
                    '考核结果', '考核数值', '点击', '曝光']
        for col in num_cols:
            if col in df.columns:
                df[col] = df[col].apply(clean_val).fillna(0.0)

        def get_target_conversion(row):
            dim = str(row.get('回传维度', ''))
            mapping = {'激活': '激活量', '唤醒': '唤醒量', '下单': '下单量', '付费': '付费数', '新登': '新登数',
                       '次留': '次留数', '首购': '首购量'}
            for key, col_name in mapping.items():
                if key in dim: return row.get(col_name, 0)
            return 0.0

        df['目标转化数'] = df.apply(get_target_conversion, axis=1)
        df['点击率'] = df.apply(
            lambda r: r['点击'] / r['曝光'] if r['曝光'] > 0 else 0.0, axis=1)

        int_cols = ['激活量', '唤醒量', '下单量', '付费数', '首购量', '新登数', '次留数', '2留数', '点击', '曝光']
        for col in int_cols:
            if col in df.columns: df[col] = df[col].astype(int)

        if '日期' in df.columns:
            df['_sort_date'] = pd.to_datetime(df['日期']).dt.normalize()

        def calc_settle(row):
            p, d = row.get('合作价格', 0), str(row.get('回传维度', ''))
            mapping = {'激活': '激活量', '唤醒': '唤醒量', '下单': '下单量', '付费': '付费数', '次留': '次留数',
                       '新登': '新登数', '首购': '首购量'}
            for k, v in mapping.items():
                if k in d: return p * row.get(v, 0)
            return 0.0

        df['结算金额'] = df.apply(calc_settle, axis=1)
        df['指标转化率'] = df.apply(lambda r: r['目标转化数'] / r['点击'] if r['点击'] > 0 else 0.0,
                                    axis=1)
        df['是否达标'] = df.apply(lambda r: r.get('考核结果', 0) >= r.get('考核数值', 0), axis=1)
        return df
    except Exception as e:
        st.error(f"加载失败: {e}");
        return pd.DataFrame()


def reset_filters():
    for key in ['甲方_filter','归属_filter' ,'产品_filter', '媒体平台_filter', '调度中心id_filter', '配置号_filter', '渠道号_filter']:
        if key in st.session_state: st.session_state[key] = []


# --- 4. 主程序界面 ---
FILE_PATH = "媒体预算日数据_附带明细.xlsx"
if Path(FILE_PATH).exists():
    df_raw = load_data(FILE_PATH)
    if not df_raw.empty:
        st.sidebar.header("🔍 维度筛选")
        st.sidebar.button("🧹 一键重置筛选", on_click=reset_filters)

        min_date_raw, max_date_raw = df_raw['_sort_date'].min().date(), df_raw['_sort_date'].max().date()
        date_sel = st.sidebar.date_input("日期范围", [min_date_raw, max_date_raw])

        if isinstance(date_sel, (list, tuple)) and len(date_sel) == 2:
            curr_start, curr_end = date_sel[0], date_sel[1]
            days_diff = (curr_end - curr_start).days + 1
            prev_start = curr_start - timedelta(days=days_diff)
            prev_end = curr_start - timedelta(days=1)
            curr_period_df = df_raw[
                (df_raw['_sort_date'].dt.date >= curr_start) & (df_raw['_sort_date'].dt.date <= curr_end)]
            prev_period_df = df_raw[
                (df_raw['_sort_date'].dt.date >= prev_start) & (df_raw['_sort_date'].dt.date <= prev_end)]
        else:
            curr_period_df = df_raw.copy()
            prev_period_df = pd.DataFrame()

        filtered_df = curr_period_df.copy()
        for col in ['甲方','归属', '产品', '媒体平台', '调度中心id', '配置号', '渠道号']:
            if col in filtered_df.columns:
                options = sorted(filtered_df[col].unique().astype(str))
                sel = st.sidebar.multiselect(f"选择{col}", options, key=f"{col}_filter")
                if sel:
                    filtered_df = filtered_df[filtered_df[col].astype(str).isin(sel)]
                    if not prev_period_df.empty:
                        prev_period_df = prev_period_df[prev_period_df[col].astype(str).isin(sel)]

        # --- 5. 指标卡片 ---
        st.markdown("---")
        c1, c2, c3, c4, c5, c6 = st.columns(6)


        def get_delta(curr_val, prev_val):
            if prev_val == 0: return None
            change = (curr_val - prev_val) / prev_val
            return f"{change:+.2%}"


        curr_settle = filtered_df['结算金额'].sum()
        prev_settle = prev_period_df['结算金额'].sum() if not prev_period_df.empty else 0
        c1.metric("总结算金额", f"¥{curr_settle:,.2f}", get_delta(curr_settle, prev_settle))

        curr_clicks = filtered_df['点击'].sum()
        prev_clicks = prev_period_df['点击'].sum() if not prev_period_df.empty else 0
        c2.metric("总点击", f"{int(curr_clicks):,}", get_delta(curr_clicks, prev_clicks))

        curr_exp = filtered_df['曝光'].sum()
        prev_exp = prev_period_df['曝光'].sum() if not prev_period_df.empty else 0
        curr_ctr = curr_clicks / curr_exp if curr_exp > 0 else 0
        prev_ctr = prev_clicks / prev_exp if prev_exp > 0 else 0
        c3.metric("点击率(CTR)", f"{curr_ctr:.2%}", get_delta(curr_ctr, prev_ctr))

        curr_conv = filtered_df['目标转化数'].sum()
        prev_conv = prev_period_df['目标转化数'].sum() if not prev_period_df.empty else 0
        c4.metric("总目标转化", f"{int(curr_conv):,}", get_delta(curr_conv, prev_conv))

        curr_cvr = curr_conv / curr_clicks if curr_clicks > 0 else 0
        prev_cvr = prev_conv / prev_clicks if prev_clicks > 0 else 0
        c5.metric("指标转化率", f"{curr_cvr:.2%}", get_delta(curr_cvr, prev_cvr))

        c6.metric("异常预警数", f"{len(filtered_df[filtered_df['是否达标'] == False])} 条")

        # --- 6. 图表逻辑 ---
        st.markdown("---")
        chart_col, rank_col = st.columns([2, 1])
        with chart_col:
            st.subheader("📈 数据趋势走势")
            trend_map = {"结算金额": "结算金额", "点击率": "点击率", "指标转化率": "指标转化率", "点击": "点击",
                         "目标转化数": "目标转化数"}
            target_label = st.selectbox("选择趋势指标：", list(trend_map.keys()))
            target_col = trend_map[target_label]
            if target_label in ["点击率", "指标转化率"]:
                chart_data = filtered_df.groupby('_sort_date')[target_col].mean().reset_index()
                y_fmt = ".2%"
            else:
                chart_data = filtered_df.groupby('_sort_date')[target_col].sum().reset_index()
                y_fmt = ",d"
            st.altair_chart(alt.Chart(chart_data).mark_line(point=True, color="#1E40AF").encode(
                x=alt.X('_sort_date:T', title='日期', axis=alt.Axis(format='%m-%d', labelAngle=-45)),
                y=alt.Y(f'{target_col}:Q', axis=alt.Axis(format=y_fmt), title=target_label),
                tooltip=[alt.Tooltip('_sort_date:T', format='%Y-%m-%d'), alt.Tooltip(f'{target_col}:Q', format=y_fmt)]
            ).properties(height=350).interactive(), use_container_width=True)

        with rank_col:
            st.subheader("🏆 结算排行 Top 10")
            rank_dim = st.radio("排行维度：", ["产品", "媒体平台", "甲方"], horizontal=True)
            rank_data = filtered_df.groupby(rank_dim)['结算金额'].sum().reset_index().sort_values('结算金额',
                                                                                                  ascending=False).head(
                10)
            st.altair_chart(alt.Chart(rank_data).mark_bar(color="#94A3B8").encode(
                x=alt.X('结算金额:Q', title='总结算金额'),
                y=alt.Y(f'{rank_dim}:N', sort='-x', title=None),
                tooltip=[alt.Tooltip(rank_dim), alt.Tooltip('结算金额:Q', format='~s')]
            ).properties(height=350), use_container_width=True)

        # --- 7. 数据明细列表 ---
        st.markdown("---")
        st.subheader("📋 数据明细列表")
        base_cols = ['日期', '甲方', '产品', '媒体平台', '配置号', '调度中心id', '回传维度', '考核结果', '考核数值',
                     '考核备注', '点击率', '指标转化率', '结算金额']
        metric_cols = ['曝光', '点击', '激活量', '新登数', '唤醒量', '下单量', '付费数', '首购量', '次留数', '2留数']
        all_display_cols = [c for c in base_cols + metric_cols if c in filtered_df.columns]
        display_df = filtered_df[all_display_cols + ['_sort_date']].copy()

        for rc in ['点击率', '指标转化率']:
            if rc in display_df.columns: display_df[rc] = display_df[rc] * 100
        for col in ['考核结果', '考核数值']:
            if col in display_df.columns: display_df[col] = display_df[col].apply(lambda x: f"{x * 100:.2f}%")
        if '_sort_date' in display_df.columns:
            display_df['日期'] = display_df['_sort_date'].dt.strftime('%Y-%m-%d')
            display_df = display_df.drop(columns=['_sort_date'])

        config = {col: st.column_config.Column(width="small") for col in all_display_cols}
        config.update({
            "点击率": st.column_config.NumberColumn("点击率", format="%.2f%%"),
            "指标转化率": st.column_config.NumberColumn("转化率", format="%.2f%%"),
            "结算金额": st.column_config.NumberColumn("结算", format="¥%.2f"),
            "曝光": st.column_config.NumberColumn("曝光", format="%d"),
            "点击": st.column_config.NumberColumn("点击", format="%d")
        })

        st.dataframe(display_df.style.apply(
            lambda x: ['color: #EF4444; font-weight: bold;' if not filtered_df.loc[i, '是否达标'] else '' for i in
                       x.index],
            subset=['考核结果'] if '考核结果' in display_df.columns else [], axis=0),
            use_container_width=True, hide_index=True, column_config=config)

else:
    st.error(f"⚠️ 找不到文件: {FILE_PATH}")