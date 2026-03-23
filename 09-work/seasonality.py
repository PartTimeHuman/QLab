import pandas as pd
import plotly.express as px

def plot_seasonality(df, date_col='日期', max_plots=5):
    """
    对DataFrame中的数值列绘制季节性图
    
    参数:
    df (pd.DataFrame): 包含数据的DataFrame
    date_col (str): 日期列的名称
    max_plots (int): 最大输出的图表数量（默认前5个）
    """
    # 1. 确保日期列为datetime格式
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    
    # 过滤掉日期为空的行
    df = df.dropna(subset=[date_col]).copy()
    
    # 2. 提取“年份”以及用于对齐X轴的“统一日期”
    # 将所有年份映射到同一个闰年（如2000年），这样可以在X轴上按月份/日期完美对齐
    df['Year'] = df[date_col].dt.year.astype(str) # 转换为字符串使颜色变为离散的类别型
    df['Month_Day'] = df[date_col].dt.strftime('2000-%m-%d')
    df['Month_Day'] = pd.to_datetime(df['Month_Day'])
    
    # 3. 筛选出所有数值类型的指标列
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    
    # 排除可能被误认的无用列（比如Unnamed或者序号列）
    target_cols = [col for col in numeric_cols if not col.startswith('Unnamed')]
    
    # 获取前 N 个指标
    cols_to_plot = target_cols[:max_plots]
    
    print(f"即将绘制以下 {len(cols_to_plot)} 个指标的季节性图: {cols_to_plot}")
    
    # 4. 循环绘制Plotly交互式折线图
    figs = []
    for col in cols_to_plot:
        # 去除当前指标的缺失值，避免断线
        plot_df = df.dropna(subset=[col]).sort_values(by='Month_Day')
        
        # 构建 Plotly Express 折线图
        fig = px.line(
            plot_df, 
            x='Month_Day', 
            y=col, 
            color='Year',
            title=f'{col} - 历年季节性对比图',
            labels={'Month_Day': '日期 (月-日)', col: '数值', 'Year': '年份'},
            markers=True # 如果数据点较少，可以开启marker
        )
        
        # 优化X轴显示，只显示月和日
        fig.update_xaxes(tickformat="%m-%d")
        fig.update_layout(
            hovermode="x unified", # 鼠标悬停时显示同一天所有年份的数据
            template="plotly_white",
            legend_title_text='年份'
        )
        
        # 显示图表
        fig.show()
        figs.append(fig)
        
    return figs

# ================= 使用示例 =================

# 读取数据，根据文件片段，真实的表头在第5行(索引为4)
# 请确保你的文件路径正确
file_path = '工作簿2.csv'
try:
    df = pd.read_csv(file_path, header=4)
    
    # 清理列名中的空格
    df.columns = df.columns.str.strip()
    
    # 调用函数，绘制前5个指标的季节性图
    figures = plot_seasonality(df, date_col='日期', max_plots=5)
    
except Exception as e:
    print(f"数据读取或处理时出错: {e}")