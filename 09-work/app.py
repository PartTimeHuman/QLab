import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# 设置页面配置
st.set_page_config(page_title="产业数据分析看板", layout="wide")


# ==========================================
# 1. 数据加载与预处理模块（已强化数据清洗）
# ==========================================
@st.cache_data
def load_data(filepath):
    # 读取原始数据
    raw_df = pd.read_csv(filepath, header=None, low_memory=False)
    series_names = raw_df.iloc[1, 1:].values
    df = raw_df.iloc[2:].copy()
    df.columns = ['Date'] + list(series_names)
    
    # 1. 转换日期格式，并使用 .dt.normalize() 剥离时分秒，只保留纯净的年月日
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.normalize()
    
    # 2. 剔除那些无法识别日期（NaT）的无效行
    df = df.dropna(subset=['Date'])
    
    # 3. 设置日期为索引
    df.set_index('Date', inplace=True)
    
    # 4. 【关键修复】剔除重复的日期行（同一个日期如果有多行，默认保留最后一行的数据）
    df = df[~df.index.duplicated(keep='last')]
    
    # 将数据转为数值型并去除空列
    df = df.apply(pd.to_numeric, errors='coerce')
    df = df.loc[:, df.columns.notna()]
    
    # 5. 【关键修复】按时间顺序把数据老老实实排好，避免时间序列发生错位倒流
    df = df.sort_index()
    
    return df

# ==========================================
# 2. 序列分类模块
# ==========================================
def categorize_columns(columns):
    """根据列名中的关键词进行简单分类"""
    categories = {'估值与利润': [], '供需': [], '库存': [], '价格与其他': []}
    for col in columns:
        col_str = str(col)
        if any(kw in col_str for kw in ['估值', '利润', '基差', '价差', '收益', '成本']):
            categories['估值与利润'].append(col)
        elif any(kw in col_str for kw in ['产量', '开工', '需求', '进出口', '消费', '销量']):
            categories['供需'].append(col)
        elif any(kw in col_str for kw in ['库存', '仓单']):
            categories['库存'].append(col)
        else:
            categories['价格与其他'].append(col)
            
    # 过滤掉空的分类
    return {k: v for k, v in categories.items() if v}
# ==========================================
# 3. 季节性画图与相关性分析模块（已修复断点逻辑）
# ==========================================
class DataAnalyzer:
    @staticmethod
    def plot_seasonality(df, col_name):
        # 提取有效数据
        temp_df = df[[col_name]].dropna().copy()
        if temp_df.empty:
            return None
            
        # 提取年份，用于分组画线
        temp_df['Year'] = temp_df.index.year.astype(str)
        
        # 【关键修复】将所有数据的年份强行映射到一个“闰年（2000年）”
        # 为什么选2000年？因为它包含02-29，这样能容纳所有闰年数据不出错
        # 这样 X 轴变成了连续的时间轴，彻底告别字符串导致的离散断点问题
        temp_df['Plot_Date'] = pd.to_datetime('2000-' + temp_df.index.strftime('%m-%d'))
        
        # 确保数据按时间先后连线
        temp_df = temp_df.sort_values('Plot_Date')
        
        fig = px.line(
            temp_df, 
            x='Plot_Date', 
            y=col_name, 
            color='Year',
            title=f"<b>{col_name}</b> - 季节性走势图",
            template='plotly_white'
        )
        
        # 【关键修复】告诉 Plotly：遇到周末/假日的缺失值时，请把两头的数据连起来，不要留空！
        fig.update_traces(connectgaps=True)
        
        # 重新格式化 X 轴的显示，隐藏2000年，只展示 月-日
        fig.update_xaxes(tickformat="%m-%d", nticks=12)
        fig.update_layout(hovermode="x unified", legend_title_text='年份')
        return fig

    @staticmethod
    def plot_rolling_corr(df, col1, col2, window):
        # （此处保持原样即可...）
        temp_df = df[[col1, col2]].dropna().copy()
        if temp_df.empty or len(temp_df) < window:
            return None
            
        corr_series = temp_df[col1].rolling(window=window).corr(temp_df[col2])
        
        fig = px.line(
            x=corr_series.index, 
            y=corr_series.values,
            title=f"<b>{col1}</b> 与 <b>{col2}</b> 的 {window}天滚动相关系数",
            labels={'x': '日期', 'y': '相关系数 (Correlation)'},
            template='plotly_white'
        )
        fig.update_yaxes(range=[-1.1, 1.1])
        fig.add_hline(y=0, line_dash="dash", line_color="gray")
        fig.update_layout(hovermode="x unified")
        return fig

    @staticmethod
    def plot_rolling_quantile(df, col_name, window):
        temp_df = df[[col_name]].dropna().copy()
        if temp_df.empty or len(temp_df) < window:
            return None
            
        # 计算滚动分位数 (%)，使用 rolling 和 rank
        temp_df['Quantile'] = temp_df[col_name].rolling(window=window).rank(pct=True) * 100
        
        # 画双Y轴图
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        
        fig.add_trace(
            go.Scatter(x=temp_df.index, y=temp_df[col_name], name="原始数值", mode='lines', line=dict(color='#1f77b4')),
            secondary_y=False,
        )
        
        fig.add_trace(
            go.Scatter(x=temp_df.index, y=temp_df['Quantile'], name=f"{window}日滚动分位数(%)", mode='lines', line=dict(color='#d62728', dash='dot')),
            secondary_y=True,
        )
        
        fig.update_layout(
            title=f"<b>{col_name}</b> 走势及 {window} 日滚动分位数",
            template='plotly_white',
            hovermode="x unified"
        )
        fig.update_yaxes(title_text="原始数值", secondary_y=False)
        fig.update_yaxes(title_text="分位数 (%)", range=[-5, 105], secondary_y=True)
        return fig
# ==========================================
# 4. Streamlit 页面布局
# ==========================================
def main():
    st.title("📈 产业数据库与量化分析看板")
    
    # 1. 加载数据
    file_path = "画图模板-日报.xlsm"
    try:
        with st.spinner('正在加载数据...'):
            df = load_data(file_path)
    except Exception as e:
        st.error(f"数据加载失败，请确保 {file_path} 与该脚本在同一目录下。错误信息：{e}")
        return

    st.success(f"数据加载成功！共包含 {df.shape[0]} 个交易日，{df.shape[1]} 条数据序列。")
    
    # 2. 获取分类
    categories = categorize_columns(df.columns)
    
    # 3. 创建Tab页面
    tab1, tab2 = st.tabs(["📊 季节性分析 (Seasonality)", "🔗 滚动相关性分析 (Rolling Correlation)"])
    tab1, tab2, tab3 = st.tabs(["📊 季节性分析 (Seasonality)", "🔗 滚动相关性分析 (Rolling Correlation)", "📈 滚动分位数 (Rolling Quantile)"])
    
    analyzer = DataAnalyzer()

    # --------------- 季节性分析 Tab ---------------
    with tab1:
        st.subheader("序列季节性规律展示")
        col1, col2 = st.columns([1, 3])
        
        with col1:
            # 第一级菜单：选择大类
            selected_cat = st.selectbox("选择指标类别：", list(categories.keys()), key='season_cat')
            # 第二级菜单：选择具体指标
            selected_col = st.selectbox("选择具体序列：", categories[selected_cat], key='season_col')
            
        with col2:
            if selected_col:
                fig_season = analyzer.plot_seasonality(df, selected_col)
                if fig_season:
                    st.plotly_chart(fig_season, use_container_width=True)
                else:
                    st.warning("该序列有效数据不足，无法绘制季节性图。")

    # --------------- 相关性分析 Tab ---------------
    with tab2:
        st.subheader("双序列滚动相关系数 (Rolling Correlation)")
        
        # 布局：左侧选择控制，右侧展示图表
        ctrl_col1, ctrl_col2, ctrl_col3 = st.columns(3)
        
        with ctrl_col1:
            st.markdown("#### 选择序列 A")
            cat_a = st.selectbox("类别 (A)：", list(categories.keys()), key='corr_cat_a')
            col_a = st.selectbox("序列 (A)：", categories[cat_a], key='corr_col_a')
            
        with ctrl_col2:
            st.markdown("#### 选择序列 B")
            cat_b = st.selectbox("类别 (B)：", list(categories.keys()), index=min(1, len(categories)-1), key='corr_cat_b')
            col_b = st.selectbox("序列 (B)：", categories[cat_b], key='corr_col_b')
            
        with ctrl_col3:
            st.markdown("#### 设置滚动窗口")
            window = st.number_input("滚动窗口天数：", min_value=5, max_value=720, value=60, step=10)
            st.info(f"当前设置：计算过去 {window} 个数据点的相关性。")

        # 画图
        if col_a and col_b:
            fig_corr = analyzer.plot_rolling_corr(df, col_a, col_b, window)
            if fig_corr:
                st.plotly_chart(fig_corr, use_container_width=True)
            else:
                st.warning("所选序列的重合有效数据过少，无法计算滚动相关性。")

    # --------------- 滚动分位数 Tab ---------------
    with tab3:
        st.subheader("指标当前值与滚动分位数 (Rolling Quantile)")
        
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            q_cat = st.selectbox("选择指标类别：", list(categories.keys()), key='q_cat')
        with col2:
            q_col = st.selectbox("选择具体序列：", categories[q_cat], key='q_col')
        with col3:
            q_window = st.number_input("计算分位数的滚动窗口（天）：", min_value=10, max_value=1000, value=60, step=10, key='q_window')
            
        if q_col:
            temp_df = df[[q_col]].dropna()
            if len(temp_df) >= q_window:
                # 计算最新分位数显示
                latest_val = temp_df[q_col].iloc[-1]
                latest_date = temp_df.index[-1].strftime('%Y-%m-%d')
                rolling_series = temp_df[q_col].iloc[-q_window:]
                current_quantile = rolling_series.rank(pct=True).iloc[-1] * 100
                
                st.metric(label=f"最新值 ({latest_date})", value=f"{latest_val:.4f}", delta=f"当前处于过去 {q_window} 天的 {current_quantile:.1f}% 分位")
                
                fig_q = analyzer.plot_rolling_quantile(df, q_col, q_window)
                if fig_q:
                    st.plotly_chart(fig_q, use_container_width=True)
            else:
                st.warning(f"该序列有效数据不足 {q_window} 个，无法计算。")

if __name__ == "__main__":
    main()