import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import math

def generate_seasonal_charts(df, date_col, value_cols):
    """
    核心绘图函数：根据给定的DataFrame生成季节性对比图(按年对比)。
    """
    # 1. 确保日期列是 datetime 格式，并提取年份
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df['Year'] = df[date_col].dt.year
    
    # 创建一个统一的虚拟年份（使用2024这个闰年，完美兼容2月29日）用于对齐X轴
    df['Dummy_Date'] = pd.to_datetime('2024-' + df[date_col].dt.strftime('%m-%d'), errors='coerce')
    
    # 2. 定义每一年的颜色和线宽（完全还原你 VBA 中的设定）
    style_map = {
        2019: {'color': 'rgb(0, 176, 80)',    'width': 1.75}, # 绿
        2020: {'color': 'rgb(255, 192, 0)',   'width': 1.75}, # 黄
        2021: {'color': 'rgb(122, 48, 160)',  'width': 1.75}, # 紫
        2022: {'color': 'rgb(0, 176, 240)',   'width': 1.75}, # 蓝
        2023: {'color': 'rgb(255, 0, 0)',     'width': 1.75}, # 红
        2024: {'color': 'rgb(128, 128, 128)', 'width': 1.75}, # 灰
        2025: {'color': 'rgb(0, 0, 0)',       'width': 2.6},  # 黑 (加粗)
        2026: {'color': 'rgb(0, 234, 255)',   'width': 3.0}   # 亮蓝 (最粗)
    }

    # 3. 计算子图布局 (按每排3个的网格布局)
    num_charts = len(value_cols)
    cols = 3
    rows = math.ceil(num_charts / cols)

    # 创建子图画布
    fig = make_subplots(
        rows=rows, 
        cols=cols, 
        subplot_titles=value_cols, 
        horizontal_spacing=0.05,
        vertical_spacing=0.1 if rows > 1 else 0.2
    )

    # 4. 循环遍历每一个需要绘制的指标列
    for i, col_name in enumerate(value_cols):
        row = (i // cols) + 1
        col = (i % cols) + 1
        
        # 遍历每一年（2019 - 2026）
        for year in range(2019, 2027):
            # 提取当年数据并按时间排序
            year_data = df[df['Year'] == year].sort_values('Dummy_Date')
            
            if year_data.empty or year_data[col_name].isna().all():
                continue

            # 添加折线
            fig.add_trace(
                go.Scatter(
                    x=year_data['Dummy_Date'],
                    y=year_data[col_name],
                    mode='lines',
                    name=str(year),
                    line=style_map.get(year, {'color': 'grey', 'width': 1}), 
                    connectgaps=True, # 空值插值连线
                    showlegend=True if i == 0 else False # 只在第一个图显示图例
                ),
                row=row, col=col
            )

    # 5. 全局样式设置
    fig.update_layout(
        height=max(400, 350 * rows),  # 根据行数动态调整高度，防止图表被压缩
        width=1200,         
        title_text="全部列季节性规律对比图",
        plot_bgcolor='white',
        font=dict(family="KaiTi, 楷体", size=16, color="black"),
        hovermode="x unified" 
    )

    # 将X轴设置为只显示月份
    fig.update_xaxes(
        tickformat="%m月",
        dtick="M1", 
        showgrid=True,
        gridcolor='lightgrey',
        zeroline=False
    )
    
    fig.update_yaxes(showgrid=True, gridcolor='lightgrey')

# 6. 渲染显示与保存
    # 尝试强制调用系统默认浏览器打开
    import plotly.io as pio
    pio.renderers.default = "browser"
    try:
        fig.show()
    except Exception as e:
        print(f"⚠️ 自动弹出浏览器失败: {e}")
        
    # 【最核心的保底方案】：自动在当前文件夹生成一个离线的 HTML 文件
    output_filename = "季节性图表结果.html"
    fig.write_html(output_filename)
    print(f"✅ 搞定！图表已成功保存为：{output_filename}")
    print(f"👉 请在你的代码文件夹中找到【{output_filename}】文件，双击用浏览器（Chrome/Edge等）打开即可查看！")

# ==========================================
# 主程序执行入口
# ==========================================
if __name__ == "__main__":
    
    # 1. 设置相对路径和读取文件
    file_path = '画图模板-日报.xlsm'  # 确保文件名和后缀对应
    
    print(f"正在读取文件: {file_path} ...")
    try:
        df = pd.read_excel(file_path, sheet_name='日数据', header=1)
    except FileNotFoundError:
        print(f"❌ 找不到文件 '{file_path}'，请确保文件和脚本在同一个文件夹！")
        exit()

    # 清理列名中的多余空格
    df.columns = df.columns.str.strip()
    
    # 2. 指定日期列名称
    date_col_name = '日期' # 如果你的Excel里这一列不叫"日期"，请改成实际的名字
    
    if date_col_name not in df.columns:
        print(f"❌ 在数据表中找不到名为 '{date_col_name}' 的列，请检查表头！")
        exit()

    df[date_col_name] = pd.to_datetime(df[date_col_name], errors='coerce')
    df = df.dropna(subset=[date_col_name])

    # 3. 自动获取所有需要画图的列（排除日期列）
    all_columns = [col for col in df.columns if col != date_col_name]
    
    columns_to_plot = []
    
    # 遍历所有列，强制转换为数字类型
    for col in all_columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
        # 如果转换后，这一列不全是空值（意味着有有效数据），就加入画图列表
        if not df[col].isna().all():
            columns_to_plot.append(col)

    print(f"自动识别到 {len(columns_to_plot)} 个数据列，正在生成图表...")
    
    # 4. 调用函数生成图表
    if columns_to_plot:
        generate_seasonal_charts(df, date_col=date_col_name, value_cols=columns_to_plot)
        print("✅ 图表生成完毕！")
    else:
        print("❌ 没有找到可以绘制的数字数据列。")