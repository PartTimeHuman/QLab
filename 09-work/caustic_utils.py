"""
caustic_utils.py
烧碱数据库 - 共享工具函数
--------------------------------------------------
将 CSV 数据库加载为干净的 DataFrame，并提供列分组配置。
"""

import pandas as pd
import numpy as np
from pathlib import Path

# ─────────────────────────────────────────────────────────
# 列分组配置（按PPT章节组织）
# ─────────────────────────────────────────────────────────
COLUMN_GROUPS = {
    "现货价格": [
        "山东32交割库最低价", "山东50交割库最低价", "江苏32交割库最低价",
        "魏桥采购价", "东营32价格", "西北片碱均价", "华南50", "出口FOB",
        "液氯价格-山东",
    ],
    "基差": [
        "山东32最低交割基差", "山东32最低出厂基差", "江苏32最低交割基差",
        "浙江32最低交割基差", "魏桥32基差", "鲁泰32基差", "金岭32基差",
        "金桥32基差", "山东50最低交割基差", "华南50基差", "西北05基差",
        "西北01基差", "西北片碱基差",
    ],
    "月差": [
        "1-5月差", "5-9月差", "9-1月差", "M0-M1", "M1-M2", "M1-主力",
    ],
    "比价": [
        "50-32价差：华泰", "50-32价差：金岭", "50-32价差：江苏",
        "50-32价差：海力", "32碱: 江苏-山东", "32碱: 浙江-山东",
        "50碱: 华南-山东", "片碱-32碱：西北", "东营-魏桥价差",
    ],
    "产量": [
        "烧碱周产量(wt)", "烧碱月产量(wt)", "华北烧碱月产量(wt)",
        "华北烧碱周产(wt)", "华东烧碱周产(wt)", "西北烧碱周产(wt)",
        "山东32碱产量", "山东50产量", "山东日产量",
    ],
    "开工率": [
        "月度开工率", "周度开工率", "月度装置损失率", "周度装置损失率",
        "山东液碱开工率",
    ],
    "库存": [
        "烧碱总库存(wt)", "液碱总库存(wt)", "固碱总库存(wt)",
        "液碱厂总库存推算(wt)", "固碱厂总库存推算(wt)", "烧碱厂总库存推算(wt)",
        "山东32库存-湿吨", "山东50库存湿吨", "山东总库存-湿吨",
        "魏桥库存-湿吨", "烧碱仓单",
    ],
    "出口": [
        "烧碱总出口(wt)", "烧碱出口单价(USD/t)", "净出口(wt)",
        "山东出口利润", "江苏出口利润", "05出口利润", "09出口利润",
        "烧碱出口：澳大利亚", "烧碱出口：印尼", "烧碱出口：台湾",
        "烧碱出口：越南", "烧碱出口：日本", "烧碱出口：韩国",
        "烧碱出口：印度",
    ],
    "利润成本": [
        "山东氯碱利润-完全成本", "山东氯碱利润-现金流成本",
        "鲁泰氯碱利润（推算电价）", "轻碱*1.35-烧碱",
        "魏桥价格边际利润", "江苏最低厂库边际利润", "盘面边际利润",
        "01盘面利润", "05盘面利润", "09盘面利润",
        "PVC现货利润含烧碱", "V01盘面利润含烧碱", "V05盘面利润含烧碱",
    ],
    "下游-氧化铝": [
        "铝土矿产量", "铝土矿港口库存", "氧化铝产量", "氧化铝产能利用率",
        "氧化铝利润", "氧化铝净进口", "氧化铝总库存",
    ],
    "下游-印染化纤": [
        "浙江印染开机", "江苏印染开机", "华东印染开机",
        "粘胶短纤产量(周)(万吨)", "粘胶短纤产能利用率(周)",
        "粘胶短纤产量(月)(万吨)", "粘胶短纤产能利用率(月)",
    ],
    "下游-MDI": [
        "聚合MDI产量(万吨)", "聚合MDI产能利用率",
        "纯MDI产量(万吨)", "纯MDI产能利用率", "MDI产量(万吨)",
    ],
    "下游-环氧丙烷": [
        "环氧丙烷利润-外采氯", "环氧丙烷利润-自用氯",
        "环氧丙烷完全利润", "环氧丙烷：产能利用率：中国（周）",
        "环氧丙烷：产量：中国（周）",
    ],
    "平衡表": [
        "产量-平衡表", "总需求-平衡表", "库存-平衡表", "出口-平衡表",
        "总需求-平衡表（实产法）", "库存-平衡表（实产法）",
    ],
    "波动率": [
        "日波动率", "年化波动率",
    ],
}


# ─────────────────────────────────────────────────────────
# 数据库加载
# ─────────────────────────────────────────────────────────
def load_database(filepath: str) -> pd.DataFrame:
    """
    加载烧碱CSV数据库并返回整洁的 DataFrame。

    数据结构说明:
    - 原始CSV第0行是指标名称（真实列头）
    - 第0列（序号）是日期
    - 返回: index=日期(DatetimeIndex), columns=指标名称
    """
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(f"找不到文件: {filepath}")

    # 读取原始数据（CSV的第1行是序号/1/2/3...，第2行才是真实指标名）
    df_raw = pd.read_csv(filepath, low_memory=False)

    # DataFrame行0 = 真实指标名称（日期, 烧碱周产量, ...）
    real_cols = [str(c).strip() for c in df_raw.iloc[0].tolist()]
    real_cols[0] = "日期"

    # 去重列名（有时有重名）
    seen = {}
    unique_cols = []
    for c in real_cols:
        if c in seen:
            seen[c] += 1
            unique_cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            unique_cols.append(c)

    df_raw.columns = unique_cols

    # 去掉第0行（已用作列名），保留其余数据行
    df = df_raw.iloc[1:].copy()
    df = df.set_index("日期")

    # 解析日期
    df.index = pd.to_datetime(df.index, errors="coerce")
    df = df[df.index.notna()]
    df = df.sort_index()

    # 转为数值
    df = df.apply(pd.to_numeric, errors="coerce")

    # 去掉全空列
    df = df.dropna(axis=1, how="all")

    return df


def get_latest_values(df: pd.DataFrame, n_cols: int = 30) -> pd.DataFrame:
    """
    返回每个指标的最新非空值及其日期。
    """
    records = []
    for col in df.columns:
        series = df[col].dropna()
        if len(series) == 0:
            continue
        latest_date = series.index[-1]
        latest_val = series.iloc[-1]
        # 与上周比较
        prev_val = series.iloc[-2] if len(series) > 1 else np.nan
        pct_chg = (latest_val - prev_val) / abs(prev_val) * 100 if prev_val != 0 else np.nan
        records.append({
            "指标": col,
            "最新日期": latest_date.strftime("%Y-%m-%d"),
            "最新值": round(latest_val, 4),
            "上期值": round(prev_val, 4) if not np.isnan(prev_val) else "-",
            "环比变化%": round(pct_chg, 2) if not np.isnan(pct_chg) else "-",
            "数据量": len(series),
        })
    result = pd.DataFrame(records)
    return result


def get_column_group(col_name: str) -> str:
    """返回指标所属分组名称"""
    for group, cols in COLUMN_GROUPS.items():
        if col_name in cols:
            return group
    return "其他"


if __name__ == "__main__":
    HERE = Path(__file__).parent.resolve()
    df = load_database(HERE / "工作簿2.csv")
    print(f"数据库加载成功: {df.shape[0]} 行 × {df.shape[1]} 列")
    print(f"日期范围: {df.index[0].date()} → {df.index[-1].date()}")
    print("\n前5个指标:")
    print(df.iloc[:3, :5])
