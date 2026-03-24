# %% [markdown]
# # 产业数据分析报告 - 估值与驱动体系对比
# 
# 该报告用于将核心指标输出为可视化视图。
# 在显示各个价格与供需走势的同时，通过双坐标轴一并绘制 **60日滚动分位数**。
#
# 数据分类说明：
# - **估值类**：价格、价差、基差、月差
# - **驱动类**：产量 (供给)、开工率、需求 (各项下游指标)

# %%
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import caustic_utils
from pathlib import Path

# %% [markdown]
# ## 1. 数据加载与分类定义

# %%
# 读取数据 (利用现成的 caustic_utils 处理逻辑)
file_path = Path("工作簿2.csv")
df = caustic_utils.load_database(file_path)
print(f"数据加载成功！总维度: {df.shape[0]} 行 × {df.shape[1]} 列。")

# 重新映射为 "估值" 和 "驱动" 两大类
REPORT_CATEGORIES = {
    "估值": {
        "价格": caustic_utils.COLUMN_GROUPS.get("现货价格", []),
        "基差": caustic_utils.COLUMN_GROUPS.get("基差", []),
        "月差": caustic_utils.COLUMN_GROUPS.get("月差", []),
        "价差(比价)": caustic_utils.COLUMN_GROUPS.get("比价", [])
    },
    "驱动": {
        "供给(产量)": caustic_utils.COLUMN_GROUPS.get("产量", []),
        "开工率": caustic_utils.COLUMN_GROUPS.get("开工率", []),
        "需求": (caustic_utils.COLUMN_GROUPS.get("下游-氧化铝", []) + 
                caustic_utils.COLUMN_GROUPS.get("下游-印染化纤", []) +
                caustic_utils.COLUMN_GROUPS.get("下游-MDI", []) +
                caustic_utils.COLUMN_GROUPS.get("下游-环氧丙烷", []))
    }
}

# %% [markdown]
# ## 2. 核心绘图函数 (引入 60日滚动分位数)

# %%
def plot_quantile_chart(df, col_name, window=60):
    """绘制原始序列及滚动分位数双轴图"""
    if col_name not in df.columns:
        return
    
    temp_df = df[[col_name]].dropna().copy()
    if len(temp_df) < window:
        print(f"[{col_name}] 有效数据不足 {window} 天，跳过绘制。")
        return
        
    # 计算滚动分位数 (%)
    temp_df['Quantile'] = temp_df[col_name].rolling(window=window).rank(pct=True) * 100
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    # 原始序列
    fig.add_trace(go.Scatter(x=temp_df.index, y=temp_df[col_name], name="指标数值", mode='lines', line=dict(color='#1f77b4')), secondary_y=False)
    # 分位数序列
    fig.add_trace(go.Scatter(x=temp_df.index, y=temp_df['Quantile'], name=f"{window}日滚动分位(%)", mode='lines', line=dict(color='#d62728', dash='dot')), secondary_y=True)
    
    fig.update_layout(title=f"<b>{col_name}</b> - 近期走势与分位数", template='plotly_white', hovermode="x unified", width=1000, height=450)
    fig.update_yaxes(title_text="原始数值", secondary_y=False)
    fig.update_yaxes(title_text="滚动分位数 (%)", range=[-5, 105], secondary_y=True)
    
    fig.show()

# %% [markdown]
# ## 3. 报告输出：【估值】体系分析

# %%
for sub_cat, cols in REPORT_CATEGORIES["估值"].items():
    print(f"\n" + "="*50 + f"\n[估值] - {sub_cat} 序列分析\n" + "="*50)
    for col in cols:
        plot_quantile_chart(df, col, window=60)

# %% [markdown]
# ## 4. 报告输出：【驱动】体系分析

# %%
for sub_cat, cols in REPORT_CATEGORIES["驱动"].items():
    print(f"\n" + "="*50 + f"\n[驱动] - {sub_cat} 序列分析\n" + "="*50)
    for col in cols:
        plot_quantile_chart(df, col, window=60)
