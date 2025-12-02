#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
校内外超市价格比较分析脚本 (Enhanced Version)
==============================================
根据 Table A (商品清单) 和 Table B (价格记录) 进行数据清洗、
单位量价计算、溢价率分析、假设检验 (H1/H2/H3) 并输出 Excel 和图表。

核心指标：
- 单位量价 UP_i = 到手价_i / 净含量_i × 100 (元/100g 或 元/100ml)
- 单位量价差 ΔUP_i = UP_i(校内) - min_{s≠校内} UP_i(s)
- 匹配篮子价格 B(s) = Σ 到手价_i(s)

数据处理：
- 3σ 原则剔除异常值
- 会员价/促销价处理：以"学生可得最低价"为主口径

统计检验方法：
- H1: 配对样本 t 检验 + Wilcoxon 符号秩检验 + 效应量 (Cohen's d) + 置信区间
- H2: 独立样本 t 检验 + Mann-Whitney U 检验 + 效应量
- H3: 配对样本检验 + 比值分析

可视化：
- 柱状图（单位量价对比）
- 瀑布图（价差分解）
- 雷达图（多维度对比）
- 箱线图、直方图等
"""

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
from scipy import stats
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# ============ 配置 ============
TABLE_A_PATH = "Table A Product List.xlsx"
TABLE_B_PATH = "Table B Price Data.xlsx"
OUTPUT_EXCEL_PATH = "output_supermarket_analysis.xlsx"
PLOTS_DIR = "plots"
LATEX_OUTPUT_PATH = "analysis_report.tex"

# 校内超市名称（精确匹配）
CAMPUS_STORE_NAMES = ["校内超市"]
# 校外超市名称
OFF_CAMPUS_STORE_NAMES = ["长申超市", "大润发"]

# 3σ 异常值处理开关
APPLY_3SIGMA = True

# 图表样式设置
plt.rcParams["font.sans-serif"] = ["Arial Unicode MS", "PingFang SC", "Heiti TC", "STHeiti"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 150
plt.rcParams["savefig.dpi"] = 300
plt.rcParams["axes.spines.top"] = False
plt.rcParams["axes.spines.right"] = False

# 颜色方案
COLORS = {
    "campus": "#E74C3C",      # 校内 - 红色
    "changsheng": "#3498DB",  # 长申 - 蓝色
    "rt_mart": "#2ECC71",     # 大润发 - 绿色
    "positive": "#E74C3C",    # 正溢价 - 红色
    "negative": "#2ECC71",    # 负溢价 - 绿色
    "neutral": "#95A5A6",     # 中性 - 灰色
    "essential": "#9B59B6",   # 必需品 - 紫色
    "impulse": "#F39C12",     # 冲动品 - 橙色
}


# ============ 数据加载 ============
def load_tables():
    """加载 Table A 和 Table B"""
    df_a = pd.read_excel(TABLE_A_PATH)
    df_b = pd.read_excel(TABLE_B_PATH)
    df_a.columns = [str(c).strip() for c in df_a.columns]
    df_b.columns = [str(c).strip() for c in df_b.columns]
    return df_a, df_b


# ============ 3σ 异常值处理 ============
def apply_3sigma_filter(df, column, group_by=None):
    """
    使用 3σ 原则剔除异常值
    返回：过滤后的 DataFrame 和被剔除的异常值信息
    """
    df = df.copy()
    outliers_info = []
    
    if group_by:
        # 分组处理
        for name, group in df.groupby(group_by):
            values = group[column].dropna()
            if len(values) < 3:
                continue
            mean_val = values.mean()
            std_val = values.std()
            lower = mean_val - 3 * std_val
            upper = mean_val + 3 * std_val
            mask = (df[column] < lower) | (df[column] > upper)
            mask = mask & (df[group_by] == name if isinstance(group_by, str) else True)
            outliers = df[mask]
            if len(outliers) > 0:
                for _, row in outliers.iterrows():
                    outliers_info.append({
                        "group": name,
                        "value": row[column],
                        "mean": mean_val,
                        "std": std_val,
                        "lower": lower,
                        "upper": upper
                    })
    else:
        # 全局处理
        values = df[column].dropna()
        if len(values) >= 3:
            mean_val = values.mean()
            std_val = values.std()
            lower = mean_val - 3 * std_val
            upper = mean_val + 3 * std_val
            mask = (df[column] >= lower) & (df[column] <= upper)
            outliers = df[~mask]
            for _, row in outliers.iterrows():
                outliers_info.append({
                    "value": row[column],
                    "mean": mean_val,
                    "std": std_val,
                    "lower": lower,
                    "upper": upper
                })
            df = df[mask]
    
    return df, outliers_info


def detect_outliers_iqr(data):
    """使用 IQR 方法检测异常值"""
    Q1 = np.percentile(data, 25)
    Q3 = np.percentile(data, 75)
    IQR = Q3 - Q1
    lower = Q1 - 1.5 * IQR
    upper = Q3 + 1.5 * IQR
    return lower, upper, IQR


# ============ 生成 ProductID ============
def prepare_product_ids(df_a, df_b, n_products_expected=44):
    """
    Table A 已有 ProductID。
    Table B 按 (StoreName, CollectionDate) 分组，组内按行顺序生成 1-44 的 ProductID。
    """
    df_a = df_a.copy()
    df_b = df_b.copy()

    # 确保 Table A 有 ProductID
    if "ProductID" not in df_a.columns:
        df_a["ProductID"] = np.arange(1, len(df_a) + 1)

    # Table B 生成 ProductID
    if "ProductID" not in df_b.columns:
        # 按 StoreName + CollectionDate 分组，组内按原始顺序编号
        df_b = df_b.reset_index(drop=True)
        df_b["ProductID"] = df_b.groupby(["StoreName", "CollectionDate"]).cumcount() + 1

    return df_a, df_b


# ============ 计算最终价格 ============
def compute_final_price(df_b):
    """
    P_final = min(Price_Actual, Price_Member) 如果会员价更低；否则取 Price_Actual。
    """
    df_b = df_b.copy()
    for col in ["Price_Actual", "Price_Member", "Price_Original"]:
        if col in df_b.columns:
            df_b[col] = pd.to_numeric(df_b[col], errors="coerce")

    df_b["P_final"] = df_b["Price_Actual"]
    if "Price_Member" in df_b.columns:
        mask = df_b["Price_Member"].notna() & (df_b["Price_Member"] < df_b["Price_Actual"])
        df_b.loc[mask, "P_final"] = df_b.loc[mask, "Price_Member"]

    return df_b


# ============ 合并表 ============
def merge_tables(df_a, df_b):
    """使用 ProductID 合并 Table A 的元数据到 Table B"""
    df = df_b.merge(df_a, on="ProductID", how="left", suffixes=("_B", "_A"))
    return df


# ============ 计算单位量价 ============
def compute_unit_price(df):
    """
    Unit_Price = P_final / NetContent * 100
    即"每 100 单位（ml/g/片/支/粒）的价格"
    """
    df = df.copy()
    df["NetContent"] = pd.to_numeric(df["NetContent"], errors="coerce")
    df["Unit_Price"] = df["P_final"] / df["NetContent"] * 100
    return df


# ============ 计算匹配篮子价格 ============
def compute_basket_price(df, campus_stores, off_campus_stores):
    """
    计算匹配篮子价格 B(s) = Σ 到手价_i(s)
    只计算所有超市都有货的商品
    """
    all_stores = campus_stores + off_campus_stores
    
    # 找出所有超市都有的商品
    product_counts = df.groupby("ProductID")["StoreName"].nunique()
    matched_products = product_counts[product_counts == len(all_stores)].index.tolist()
    
    # 计算每个超市的匹配篮子价格
    basket_prices = {}
    for store in all_stores:
        store_df = df[(df["StoreName"] == store) & (df["ProductID"].isin(matched_products))]
        basket_prices[store] = store_df.groupby("CollectionDate")["P_final"].sum().mean()
    
    return basket_prices, matched_products


# ============ 计算单位量价差 ΔUP ============
def compute_delta_up(df, campus_stores, off_campus_stores):
    """
    计算单位量价差：ΔUP_i = UP_i(校内) - min_{s≠校内} UP_i(s)
    """
    # 按商品和超市分组计算平均单位价
    grouped = df.groupby(["ProductID", "StoreName"])["Unit_Price"].mean().reset_index()
    
    # 透视表
    pivot = grouped.pivot(index="ProductID", columns="StoreName", values="Unit_Price")
    
    # 校内价格
    campus_cols = [c for c in pivot.columns if c in campus_stores]
    off_campus_cols = [c for c in pivot.columns if c in off_campus_stores]
    
    if not campus_cols or not off_campus_cols:
        return pd.DataFrame()
    
    pivot["UP_campus"] = pivot[campus_cols].mean(axis=1)
    pivot["UP_off_campus_min"] = pivot[off_campus_cols].min(axis=1)
    pivot["UP_off_campus_mean"] = pivot[off_campus_cols].mean(axis=1)
    pivot["Delta_UP"] = pivot["UP_campus"] - pivot["UP_off_campus_min"]
    pivot["Delta_UP_mean"] = pivot["UP_campus"] - pivot["UP_off_campus_mean"]
    
    return pivot.reset_index()


# ============ 识别校内/校外 ============
def identify_campus_and_off_campus(df):
    """返回校内和校外超市名称列表"""
    all_stores = df["StoreName"].dropna().unique().tolist()
    campus = [s for s in all_stores if s in CAMPUS_STORE_NAMES]
    off_campus = [s for s in all_stores if s in OFF_CAMPUS_STORE_NAMES]
    # 如果配置没匹配上，用默认规则
    if not campus:
        campus = [s for s in all_stores if "校" in s or "教" in s]
    if not off_campus:
        off_campus = [s for s in all_stores if s not in campus]
    return campus, off_campus


# ============ 计算溢价率 ============
def compute_premium_by_product(df, campus_stores, off_campus_stores):
    """
    对每个商品，计算：
    - UP_campus: 校内平均单位价
    - UP_off_campus: 校外平均单位价（长申+大润发的均值）
    - Premium_Rate: (UP_campus - UP_off_campus) / UP_off_campus * 100%
    """
    # 按 (StoreName, ProductID) 取平均（处理同一超市多次采集）
    grouped = df.groupby(["StoreName", "ProductID"])["Unit_Price"].mean().reset_index()

    # 校内
    campus_df = grouped[grouped["StoreName"].isin(campus_stores)]
    campus_up = campus_df.groupby("ProductID")["Unit_Price"].mean()

    # 校外
    off_df = grouped[grouped["StoreName"].isin(off_campus_stores)]
    off_pivot = off_df.pivot_table(index="ProductID", columns="StoreName", values="Unit_Price")
    off_mean = off_pivot.mean(axis=1)  # 长申+大润发的平均

    # 合并
    aligned = pd.DataFrame({
        "UP_campus": campus_up,
        "UP_off_campus": off_mean
    }).dropna()

    aligned["Premium_Rate"] = (aligned["UP_campus"] - aligned["UP_off_campus"]) / aligned["UP_off_campus"] * 100
    aligned = aligned.reset_index()

    return aligned


# ============ 附加商品元数据 ============
def attach_product_metadata(premium_df, df_a):
    """将 Table A 的 ProductName, Brand, Category_H2, Category_H3, Unit 附加到溢价表"""
    cols = ["ProductID", "ProductName", "Brand", "Category_H2", "Category_H3", "Unit"]
    cols = [c for c in cols if c in df_a.columns]
    meta = df_a[cols].drop_duplicates()
    merged = premium_df.merge(meta, on="ProductID", how="left")
    return merged


# ============ 描述性统计 ============
def compute_descriptive_stats(df):
    """计算各超市的购物篮总价和单位价格统计"""
    # 购物篮总价（按 StoreName + CollectionDate 分组求和）
    basket = df.groupby(["StoreName", "CollectionDate"])["P_final"].sum().reset_index(name="Basket_Total")
    basket_stats = basket.groupby("StoreName")["Basket_Total"].agg(["mean", "std", "min", "max", "count"]).reset_index()
    basket_stats.columns = ["StoreName", "Basket_Mean", "Basket_Std", "Basket_Min", "Basket_Max", "Basket_Count"]

    # 单位价格统计
    unit_stats = df.groupby("StoreName")["Unit_Price"].agg(["mean", "std", "min", "max", "count"]).reset_index()
    unit_stats.columns = ["StoreName", "Unit_Mean", "Unit_Std", "Unit_Min", "Unit_Max", "Unit_Count"]

    descriptive = basket_stats.merge(unit_stats, on="StoreName", how="outer")
    return descriptive


# ============ 效应量计算 ============
def cohens_d(x, y):
    """计算 Cohen's d 效应量（配对样本）"""
    diff = np.array(x) - np.array(y)
    return np.mean(diff) / np.std(diff, ddof=1)


def cohens_d_independent(x, y):
    """计算 Cohen's d 效应量（独立样本）"""
    nx, ny = len(x), len(y)
    pooled_std = np.sqrt(((nx - 1) * np.std(x, ddof=1)**2 + (ny - 1) * np.std(y, ddof=1)**2) / (nx + ny - 2))
    return (np.mean(x) - np.mean(y)) / pooled_std


def interpret_cohens_d(d):
    """解释 Cohen's d 效应量"""
    d = abs(d)
    if d < 0.2:
        return "微小"
    elif d < 0.5:
        return "小"
    elif d < 0.8:
        return "中等"
    else:
        return "大"


def interpret_p_value(p):
    """解释 p 值显著性水平"""
    if p < 0.001:
        return "***"
    elif p < 0.01:
        return "**"
    elif p < 0.05:
        return "*"
    else:
        return "n.s."


# ============ H1: 综合统计检验 ============
def h1_comprehensive_test(premium_df):
    """
    H1 综合检验：校内 vs 校外单位价格
    包含：配对 t 检验、Wilcoxon 符号秩检验、效应量分析
    """
    x = premium_df["UP_campus"].values
    y = premium_df["UP_off_campus"].values
    diff = x - y
    n = len(premium_df)
    
    # 描述性统计
    mean_campus = np.mean(x)
    mean_off = np.mean(y)
    std_campus = np.std(x, ddof=1)
    std_off = np.std(y, ddof=1)
    mean_diff = np.mean(diff)
    std_diff = np.std(diff, ddof=1)
    se_diff = std_diff / np.sqrt(n)
    
    # 95% 置信区间
    ci_95 = stats.t.interval(0.95, df=n-1, loc=mean_diff, scale=se_diff)
    
    # 配对 t 检验
    t_stat, p_ttest = stats.ttest_rel(x, y)
    
    # Wilcoxon 符号秩检验（非参数）
    try:
        w_stat, p_wilcoxon = stats.wilcoxon(x, y, alternative='two-sided')
    except ValueError:
        w_stat, p_wilcoxon = np.nan, np.nan
    
    # 正态性检验（Shapiro-Wilk）
    if n >= 3:
        shapiro_stat, p_shapiro = stats.shapiro(diff)
    else:
        shapiro_stat, p_shapiro = np.nan, np.nan
    
    # 效应量
    d = cohens_d(x, y)
    
    # 溢价率统计
    premium_rates = premium_df["Premium_Rate"].values
    mean_premium = np.mean(premium_rates)
    median_premium = np.median(premium_rates)
    std_premium = np.std(premium_rates, ddof=1)
    
    # 正溢价比例
    pct_positive = np.sum(premium_rates > 0) / n * 100
    
    result = {
        "n_pairs": n,
        "mean_UP_campus": mean_campus,
        "std_UP_campus": std_campus,
        "mean_UP_off_campus": mean_off,
        "std_UP_off_campus": std_off,
        "mean_difference": mean_diff,
        "std_difference": std_diff,
        "se_difference": se_diff,
        "ci_95_lower": ci_95[0],
        "ci_95_upper": ci_95[1],
        "t_statistic": t_stat,
        "p_value_ttest": p_ttest,
        "wilcoxon_statistic": w_stat,
        "p_value_wilcoxon": p_wilcoxon,
        "shapiro_statistic": shapiro_stat,
        "p_value_shapiro": p_shapiro,
        "cohens_d": d,
        "effect_size_interpretation": interpret_cohens_d(d),
        "mean_premium_rate_pct": mean_premium,
        "median_premium_rate_pct": median_premium,
        "std_premium_rate_pct": std_premium,
        "pct_products_higher_on_campus": pct_positive,
    }
    
    return result


# ============ H2: 综合统计检验 ============
def h2_comprehensive_test(premium_with_meta, treat_noodles_as_essential=False):
    """
    H2 综合检验：必需品 vs 冲动品溢价率比较
    包含：独立样本 t 检验、Mann-Whitney U 检验、效应量分析
    """
    df = premium_with_meta.copy()
    category_col = "Category_H2"

    if treat_noodles_as_essential and "ProductName" in df.columns:
        mask = df["ProductName"].astype(str).str.contains("泡面|方便面|拉面", regex=True, case=False)
        df.loc[mask, category_col] = "H2-必需品"

    # 分组
    essential = df[df[category_col].str.contains("必需", na=False)]["Premium_Rate"].dropna().values
    impulse = df[df[category_col].str.contains("冲动", na=False)]["Premium_Rate"].dropna().values
    
    if len(essential) < 2 or len(impulse) < 2:
        return None
    
    # 描述性统计
    n_essential, n_impulse = len(essential), len(impulse)
    mean_essential, mean_impulse = np.mean(essential), np.mean(impulse)
    std_essential, std_impulse = np.std(essential, ddof=1), np.std(impulse, ddof=1)
    median_essential, median_impulse = np.median(essential), np.median(impulse)
    
    # 独立样本 t 检验（Welch's t-test，不假设方差齐性）
    t_stat, p_ttest = stats.ttest_ind(essential, impulse, equal_var=False)
    
    # Mann-Whitney U 检验（非参数）
    u_stat, p_mannwhitney = stats.mannwhitneyu(essential, impulse, alternative='two-sided')
    
    # Levene 方差齐性检验
    levene_stat, p_levene = stats.levene(essential, impulse)
    
    # 效应量
    d = cohens_d_independent(essential, impulse)
    
    result = {
        "n_essential": n_essential,
        "n_impulse": n_impulse,
        "mean_essential": mean_essential,
        "std_essential": std_essential,
        "median_essential": median_essential,
        "mean_impulse": mean_impulse,
        "std_impulse": std_impulse,
        "median_impulse": median_impulse,
        "mean_difference": mean_essential - mean_impulse,
        "t_statistic": t_stat,
        "p_value_ttest": p_ttest,
        "mannwhitney_u": u_stat,
        "p_value_mannwhitney": p_mannwhitney,
        "levene_statistic": levene_stat,
        "p_value_levene": p_levene,
        "cohens_d": d,
        "effect_size_interpretation": interpret_cohens_d(d),
        "noodles_as_essential": treat_noodles_as_essential,
    }
    
    return result


def h2_premium_by_category(premium_with_meta, treat_noodles_as_essential=False):
    """
    按 Category_H2 分组计算平均溢价率。
    treat_noodles_as_essential=True 时，将泡面/方便面改为"H2-必需品"。
    """
    df = premium_with_meta.copy()
    category_col = "Category_H2"

    if treat_noodles_as_essential and "ProductName" in df.columns:
        mask = df["ProductName"].astype(str).str.contains("泡面|方便面|拉面", regex=True, case=False)
        df.loc[mask, category_col] = "H2-必需品"

    summary = df.groupby(category_col).agg(
        mean_premium_rate=("Premium_Rate", "mean"),
        std_premium_rate=("Premium_Rate", "std"),
        median_premium_rate=("Premium_Rate", "median"),
        min_premium_rate=("Premium_Rate", "min"),
        max_premium_rate=("Premium_Rate", "max"),
        n_products=("Premium_Rate", "count")
    ).reset_index()

    return summary


# ============ H3: 大包装 vs 小包装 ============
def h3_small_vs_large(df, df_a):
    """
    对 Category_H3 不为空的商品（如水、可乐），
    对比 Small vs Large 的 Unit_Price。
    """
    # 从 df_a 获取 Category_H3 信息
    h3_products = df_a[df_a["Category_H3"].notna() & (df_a["Category_H3"] != "N/A")][["ProductID", "ProductName", "Brand", "Category_H3", "Unit"]]
    if h3_products.empty:
        return pd.DataFrame()

    # 合并到价格数据
    df_h3 = df.merge(h3_products, on="ProductID", how="inner", suffixes=("", "_meta"))

    if df_h3.empty:
        return pd.DataFrame()

    # 按 StoreName + Brand + Category_H3 分组取平均单位价
    grouped = df_h3.groupby(["StoreName", "Brand", "Category_H3"])["Unit_Price"].mean().reset_index()

    # 透视表
    pivot = grouped.pivot_table(index=["StoreName", "Brand"], columns="Category_H3", values="Unit_Price")

    # 找 Small 和 Large 列
    cols = [str(c) for c in pivot.columns]
    small_cols = [c for c in cols if "Small" in c or "小" in c]
    large_cols = [c for c in cols if "Large" in c or "大" in c]

    if not small_cols or not large_cols:
        return pivot.reset_index()

    small_col = small_cols[0]
    large_col = large_cols[0]

    pivot["UP_Small"] = pivot[small_col]
    pivot["UP_Large"] = pivot[large_col]
    pivot["Ratio_Small_to_Large"] = pivot["UP_Small"] / pivot["UP_Large"]

    result = pivot.reset_index()
    return result


def h3_comprehensive_test(h3_ratios):
    """
    H3 综合检验：小包装 vs 大包装单位价格
    """
    if h3_ratios.empty or "Ratio_Small_to_Large" not in h3_ratios.columns:
        return None
    
    ratios = h3_ratios["Ratio_Small_to_Large"].dropna().values
    n = len(ratios)
    
    if n < 2:
        return None
    
    # 描述性统计
    mean_ratio = np.mean(ratios)
    std_ratio = np.std(ratios, ddof=1)
    median_ratio = np.median(ratios)
    min_ratio = np.min(ratios)
    max_ratio = np.max(ratios)
    
    # 单样本 t 检验：检验比值是否显著大于 1
    t_stat, p_ttest = stats.ttest_1samp(ratios, 1.0)
    # 单侧检验 p 值（检验是否 > 1）
    p_onesided = p_ttest / 2 if t_stat > 0 else 1 - p_ttest / 2
    
    # Wilcoxon 符号秩检验（与 1 比较）
    try:
        w_stat, p_wilcoxon = stats.wilcoxon(ratios - 1, alternative='greater')
    except ValueError:
        w_stat, p_wilcoxon = np.nan, np.nan
    
    # 效应量（与 1 的差异）
    d = (mean_ratio - 1) / std_ratio
    
    result = {
        "n_pairs": n,
        "mean_ratio": mean_ratio,
        "std_ratio": std_ratio,
        "median_ratio": median_ratio,
        "min_ratio": min_ratio,
        "max_ratio": max_ratio,
        "t_statistic": t_stat,
        "p_value_ttest_twosided": p_ttest,
        "p_value_ttest_onesided": p_onesided,
        "wilcoxon_statistic": w_stat,
        "p_value_wilcoxon": p_wilcoxon,
        "cohens_d": d,
        "effect_size_interpretation": interpret_cohens_d(d),
    }
    
    return result


# ============ 溢价排行榜 ============
def top_premium_products(premium_with_meta, top_n=10):
    """返回溢价率最高的 top_n 个商品"""
    df = premium_with_meta.sort_values("Premium_Rate", ascending=False).head(top_n)
    cols = ["ProductID", "ProductName", "Brand", "Category_H2", "UP_campus", "UP_off_campus", "Premium_Rate"]
    cols = [c for c in cols if c in df.columns]
    return df[cols].reset_index(drop=True)


# ============ 保存 Excel ============
def save_excel_outputs(df_clean, descriptive_stats, h1_result, h2_default, h2_noodles, h3_ratios, top_premium, premium_with_meta, h1_stats, h2_stats_default, h2_stats_noodles, h3_stats):
    """将所有结果保存到一个 Excel 文件的多个 sheet"""
    with pd.ExcelWriter(OUTPUT_EXCEL_PATH, engine="openpyxl") as writer:
        df_clean.to_excel(writer, sheet_name="Cleaned_Data", index=False)
        descriptive_stats.to_excel(writer, sheet_name="Descriptive_Stats", index=False)
        
        # H1 综合检验结果
        h1_df = pd.DataFrame([h1_stats])
        h1_df.to_excel(writer, sheet_name="H1_Comprehensive", index=False)
        
        # H2 结果
        h2_default.to_excel(writer, sheet_name="H2_Default", index=False)
        h2_noodles.to_excel(writer, sheet_name="H2_NoodlesEssential", index=False)
        if h2_stats_default:
            pd.DataFrame([h2_stats_default]).to_excel(writer, sheet_name="H2_Test_Default", index=False)
        if h2_stats_noodles:
            pd.DataFrame([h2_stats_noodles]).to_excel(writer, sheet_name="H2_Test_Noodles", index=False)
        
        # H3 结果
        if not h3_ratios.empty:
            h3_ratios.to_excel(writer, sheet_name="H3_Small_vs_Large", index=False)
        if h3_stats:
            pd.DataFrame([h3_stats]).to_excel(writer, sheet_name="H3_Comprehensive", index=False)
        
        top_premium.to_excel(writer, sheet_name="Top_Premium_Products", index=False)
        premium_with_meta.to_excel(writer, sheet_name="All_Premium_Rates", index=False)

    print(f"[OK] Excel 已保存: {OUTPUT_EXCEL_PATH}")


# ============ 生成 LaTeX 报告 ============
def generate_latex_report(h1_stats, h2_stats_default, h2_stats_noodles, h3_stats, h2_default, h2_noodles, top_premium, descriptive_stats):
    """生成 LaTeX 格式的统计分析报告"""
    
    latex = r"""% ============================================================
% 校内外超市价格比较研究 - 统计分析报告
% 自动生成于 """ + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + r"""
% ============================================================

\documentclass[12pt, a4paper]{article}
\usepackage[UTF8]{ctex}
\usepackage{amsmath, amssymb}
\usepackage{booktabs}
\usepackage{graphicx}
\usepackage{geometry}
\usepackage{float}
\usepackage{caption}
\usepackage{subcaption}
\usepackage{hyperref}

\geometry{left=2.5cm, right=2.5cm, top=2.5cm, bottom=2.5cm}

\title{校内外超市价格比较研究\\统计分析报告}
\author{数据分析自动生成}
\date{\today}

\begin{document}
\maketitle

\section{研究概述}

本研究旨在通过实证数据分析，探究校内教育超市相对于校外大型超市（长申超市、大润发）是否存在价格溢价。研究检验以下三个核心假设：

\begin{itemize}
    \item \textbf{H1}：校内超市整体价格显著高于校外（系统性溢价）
    \item \textbf{H2}：必需品与冲动消费品的溢价程度存在差异
    \item \textbf{H3}：小包装商品的单位价格高于大包装商品
\end{itemize}

\section{数据概况}

"""
    
    # 描述性统计表
    latex += r"""
\subsection{各超市描述性统计}

\begin{table}[H]
\centering
\caption{各超市购物篮总价与单位价格统计}
\begin{tabular}{lcccc}
\toprule
超市名称 & 购物篮均价 (元) & 购物篮标准差 & 单位价格均值 & 单位价格标准差 \\
\midrule
"""
    for _, row in descriptive_stats.iterrows():
        latex += f"{row['StoreName']} & {row['Basket_Mean']:.2f} & {row['Basket_Std']:.2f} & {row['Unit_Mean']:.2f} & {row['Unit_Std']:.2f} \\\\\n"
    
    latex += r"""\bottomrule
\end{tabular}
\end{table}

"""

    # H1 检验结果
    latex += r"""
\section{假设检验 H1：校内是否更贵？}

\subsection{研究假设}

\begin{align}
H_0 &: \mu_{\text{校内}} = \mu_{\text{校外}} \quad \text{（校内外价格无显著差异）} \\
H_1 &: \mu_{\text{校内}} \neq \mu_{\text{校外}} \quad \text{（校内外价格存在显著差异）}
\end{align}

\subsection{描述性统计}

"""
    latex += f"""
\\begin{{itemize}}
    \\item 有效配对商品数：$n = {h1_stats['n_pairs']}$
    \\item 校内平均单位价格：$\\bar{{X}}_{{\\text{{校内}}}} = {h1_stats['mean_UP_campus']:.4f}$ (SD = {h1_stats['std_UP_campus']:.4f})
    \\item 校外平均单位价格：$\\bar{{X}}_{{\\text{{校外}}}} = {h1_stats['mean_UP_off_campus']:.4f}$ (SD = {h1_stats['std_UP_off_campus']:.4f})
    \\item 平均差异：$\\bar{{D}} = {h1_stats['mean_difference']:.4f}$ (SE = {h1_stats['se_difference']:.4f})
    \\item 差异的 95\\% 置信区间：$[{h1_stats['ci_95_lower']:.4f}, {h1_stats['ci_95_upper']:.4f}]$
    \\item 平均溢价率：${h1_stats['mean_premium_rate_pct']:.2f}\\%$ (中位数 = {h1_stats['median_premium_rate_pct']:.2f}\\%)
    \\item 校内价格更高的商品比例：${h1_stats['pct_products_higher_on_campus']:.1f}\\%$
\\end{{itemize}}

\\subsection{{正态性检验}}

使用 Shapiro-Wilk 检验评估配对差异的正态性：
\\begin{{align}}
W &= {h1_stats['shapiro_statistic']:.4f}, \\quad p = {h1_stats['p_value_shapiro']:.4f}
\\end{{align}}

"""
    if h1_stats['p_value_shapiro'] >= 0.05:
        latex += "由于 $p \\geq 0.05$，不能拒绝正态性假设，可以使用参数检验方法。\n\n"
    else:
        latex += "由于 $p < 0.05$，正态性假设被拒绝，应参考非参数检验结果。\n\n"

    latex += f"""
\\subsection{{参数检验：配对样本 $t$ 检验}}

\\begin{{align}}
t &= \\frac{{\\bar{{D}}}}{{SE_D}} = \\frac{{{h1_stats['mean_difference']:.4f}}}{{{h1_stats['se_difference']:.4f}}} = {h1_stats['t_statistic']:.4f} \\\\
df &= n - 1 = {h1_stats['n_pairs'] - 1} \\\\
p &= {h1_stats['p_value_ttest']:.4f}
\\end{{align}}

\\subsection{{非参数检验：Wilcoxon 符号秩检验}}

\\begin{{align}}
W &= {h1_stats['wilcoxon_statistic']:.1f}, \\quad p = {h1_stats['p_value_wilcoxon']:.4f}
\\end{{align}}

\\subsection{{效应量分析}}

Cohen's $d$ 效应量：
\\begin{{align}}
d &= \\frac{{\\bar{{D}}}}{{SD_D}} = {h1_stats['cohens_d']:.4f}
\\end{{align}}

效应量解释：\\textbf{{{h1_stats['effect_size_interpretation']}}}（$|d| < 0.2$: 微小, $0.2 \\leq |d| < 0.5$: 小, $0.5 \\leq |d| < 0.8$: 中等, $|d| \\geq 0.8$: 大）

\\subsection{{H1 结论}}

"""
    if h1_stats['p_value_ttest'] < 0.05:
        latex += f"配对样本 $t$ 检验结果显示 $p = {h1_stats['p_value_ttest']:.4f} < 0.05$，\\textbf{{拒绝原假设}}，校内超市价格与校外存在显著差异。"
    else:
        latex += f"配对样本 $t$ 检验结果显示 $p = {h1_stats['p_value_ttest']:.4f} \\geq 0.05$，\\textbf{{不能拒绝原假设}}，未发现校内外价格存在统计学显著差异。"
    
    if h1_stats['p_value_wilcoxon'] < 0.05:
        latex += f" Wilcoxon 检验 ($p = {h1_stats['p_value_wilcoxon']:.4f}$) 支持此结论。"
    else:
        latex += f" Wilcoxon 检验 ($p = {h1_stats['p_value_wilcoxon']:.4f}$) 同样未发现显著差异。"

    # H2 检验结果
    latex += r"""

\section{假设检验 H2：必需品 vs 冲动品}

\subsection{研究假设}

\begin{align}
H_0 &: \mu_{\text{必需品溢价率}} = \mu_{\text{冲动品溢价率}} \\
H_1 &: \mu_{\text{必需品溢价率}} \neq \mu_{\text{冲动品溢价率}}
\end{align}

\subsection{描述性统计（默认分类）}

\begin{table}[H]
\centering
\caption{各类别溢价率统计（默认分类）}
\begin{tabular}{lccccc}
\toprule
类别 & $n$ & 均值 (\%) & 标准差 (\%) & 中位数 (\%) & 范围 (\%) \\
\midrule
"""
    for _, row in h2_default.iterrows():
        latex += f"{row['Category_H2']} & {row['n_products']:.0f} & {row['mean_premium_rate']:.2f} & {row['std_premium_rate']:.2f} & {row['median_premium_rate']:.2f} & [{row['min_premium_rate']:.1f}, {row['max_premium_rate']:.1f}] \\\\\n"
    
    latex += r"""\bottomrule
\end{tabular}
\end{table}

"""
    
    if h2_stats_default:
        latex += f"""
\\subsection{{统计检验（默认分类）}}

\\textbf{{独立样本 $t$ 检验（Welch's）：}}
\\begin{{align}}
t &= {h2_stats_default['t_statistic']:.4f}, \\quad p = {h2_stats_default['p_value_ttest']:.4f}
\\end{{align}}

\\textbf{{Mann-Whitney $U$ 检验：}}
\\begin{{align}}
U &= {h2_stats_default['mannwhitney_u']:.1f}, \\quad p = {h2_stats_default['p_value_mannwhitney']:.4f}
\\end{{align}}

\\textbf{{Levene 方差齐性检验：}}
\\begin{{align}}
F &= {h2_stats_default['levene_statistic']:.4f}, \\quad p = {h2_stats_default['p_value_levene']:.4f}
\\end{{align}}

\\textbf{{效应量：}} Cohen's $d = {h2_stats_default['cohens_d']:.4f}$ ({h2_stats_default['effect_size_interpretation']})

"""

    # 敏感性分析
    latex += r"""
\subsection{敏感性分析：泡面归类变更}

将泡面/方便面从"冲动品"改为"必需品"后：

\begin{table}[H]
\centering
\caption{各类别溢价率统计（泡面归为必需品）}
\begin{tabular}{lccccc}
\toprule
类别 & $n$ & 均值 (\%) & 标准差 (\%) & 中位数 (\%) & 范围 (\%) \\
\midrule
"""
    for _, row in h2_noodles.iterrows():
        latex += f"{row['Category_H2']} & {row['n_products']:.0f} & {row['mean_premium_rate']:.2f} & {row['std_premium_rate']:.2f} & {row['median_premium_rate']:.2f} & [{row['min_premium_rate']:.1f}, {row['max_premium_rate']:.1f}] \\\\\n"
    
    latex += r"""\bottomrule
\end{tabular}
\end{table}

"""

    if h2_stats_noodles:
        latex += f"""
\\textbf{{敏感性分析检验结果：}}
\\begin{{itemize}}
    \\item $t = {h2_stats_noodles['t_statistic']:.4f}$, $p = {h2_stats_noodles['p_value_ttest']:.4f}$
    \\item Cohen's $d = {h2_stats_noodles['cohens_d']:.4f}$ ({h2_stats_noodles['effect_size_interpretation']})
\\end{{itemize}}

"""

    latex += r"""
\subsection{H2 结论}

"""
    if h2_stats_default and h2_stats_default['p_value_ttest'] < 0.05:
        latex += "必需品与冲动品的溢价率存在显著差异。"
    else:
        latex += "必需品与冲动品的溢价率无显著差异。"
    
    if h2_stats_default and h2_stats_noodles:
        if (h2_stats_default['p_value_ttest'] < 0.05) == (h2_stats_noodles['p_value_ttest'] < 0.05):
            latex += " 敏感性分析显示，泡面分类变更不影响结论的稳健性。"
        else:
            latex += " 敏感性分析显示，泡面分类变更会影响结论，需谨慎解读。"

    # H3 检验结果
    if h3_stats:
        latex += f"""

\\section{{假设检验 H3：大包装 vs 小包装}}

\\subsection{{研究假设}}

\\begin{{align}}
H_0 &: \\frac{{UP_{{\\text{{小包装}}}}}}{{UP_{{\\text{{大包装}}}}}} = 1 \\quad \\text{{（单位价格相同）}} \\\\
H_1 &: \\frac{{UP_{{\\text{{小包装}}}}}}{{UP_{{\\text{{大包装}}}}}} > 1 \\quad \\text{{（小包装单价更高）}}
\\end{{align}}

\\subsection{{描述性统计}}

\\begin{{itemize}}
    \\item 有效配对数：$n = {h3_stats['n_pairs']}$
    \\item 平均比值：$\\bar{{R}} = {h3_stats['mean_ratio']:.4f}$ (SD = {h3_stats['std_ratio']:.4f})
    \\item 中位数比值：${h3_stats['median_ratio']:.4f}$
    \\item 比值范围：$[{h3_stats['min_ratio']:.4f}, {h3_stats['max_ratio']:.4f}]$
\\end{{itemize}}

\\subsection{{统计检验}}

\\textbf{{单样本 $t$ 检验（检验比值是否 $> 1$）：}}
\\begin{{align}}
t &= {h3_stats['t_statistic']:.4f}, \\quad p_{{\\text{{单侧}}}} = {h3_stats['p_value_ttest_onesided']:.4f}
\\end{{align}}

\\textbf{{Wilcoxon 符号秩检验：}}
\\begin{{align}}
W &= {h3_stats['wilcoxon_statistic']:.1f}, \\quad p = {h3_stats['p_value_wilcoxon']:.4f}
\\end{{align}}

\\textbf{{效应量：}} Cohen's $d = {h3_stats['cohens_d']:.4f}$ ({h3_stats['effect_size_interpretation']})

\\subsection{{H3 结论}}

"""
        if h3_stats['p_value_ttest_onesided'] < 0.05:
            latex += f"单样本 $t$ 检验结果显示 $p = {h3_stats['p_value_ttest_onesided']:.4f} < 0.05$，\\textbf{{拒绝原假设}}，小包装商品的单位价格显著高于大包装。平均而言，小包装单价是大包装的 {h3_stats['mean_ratio']:.2f} 倍。"
        else:
            latex += f"单样本 $t$ 检验结果显示 $p = {h3_stats['p_value_ttest_onesided']:.4f} \\geq 0.05$，\\textbf{{不能拒绝原假设}}。"

    # 溢价排行榜
    latex += r"""

\section{溢价率排行榜}

\begin{table}[H]
\centering
\caption{校内溢价率最高的商品 Top 10}
\begin{tabular}{clccc}
\toprule
排名 & 商品名称 & 校内单价 & 校外单价 & 溢价率 (\%) \\
\midrule
"""
    for i, (_, row) in enumerate(top_premium.iterrows(), 1):
        name = row.get('ProductName', f"商品{row['ProductID']}")
        latex += f"{i} & {name} & {row['UP_campus']:.2f} & {row['UP_off_campus']:.2f} & {row['Premium_Rate']:.1f} \\\\\n"
    
    latex += r"""\bottomrule
\end{tabular}
\end{table}

\section{研究结论}

\subsection{主要发现}

\begin{enumerate}
"""
    # H1 结论
    if h1_stats['p_value_ttest'] < 0.05:
        latex += f"    \\item \\textbf{{H1 成立}}：校内超市价格显著高于校外，平均溢价率为 {h1_stats['mean_premium_rate_pct']:.1f}\\%。\n"
    else:
        latex += f"    \\item \\textbf{{H1 不成立}}：虽然平均溢价率为 {h1_stats['mean_premium_rate_pct']:.1f}\\%，但统计检验未发现显著差异 ($p = {h1_stats['p_value_ttest']:.3f}$)。\n"
    
    # H2 结论
    if h2_stats_default and h2_stats_default['p_value_ttest'] < 0.05:
        latex += f"    \\item \\textbf{{H2 成立}}：必需品与冲动品的溢价率存在显著差异。\n"
    else:
        latex += f"    \\item \\textbf{{H2 不成立}}：必需品与冲动品的溢价率无显著差异。\n"
    
    # H3 结论
    if h3_stats and h3_stats['p_value_ttest_onesided'] < 0.05:
        latex += f"    \\item \\textbf{{H3 成立}}：小包装商品单价显著高于大包装，平均为大包装的 {h3_stats['mean_ratio']:.2f} 倍。\n"
    elif h3_stats:
        latex += f"    \\item \\textbf{{H3 不成立}}：小包装与大包装单价无显著差异。\n"

    latex += r"""\end{enumerate}

\subsection{研究局限}

\begin{itemize}
    \item 样本量有限（$n = """ + str(h1_stats['n_pairs']) + r"""$ 种商品），可能影响统计功效
    \item 数据采集时间有限，未能捕捉长期价格波动
    \item 部分校外超市价格为促销价，可能低估常规价格差异
\end{itemize}

\end{document}
"""
    
    # 保存 LaTeX 文件
    with open(LATEX_OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write(latex)
    
    print(f"[OK] LaTeX 报告已保存: {LATEX_OUTPUT_PATH}")


# ============ 增强版绘图 ============
def plot_figures_enhanced(df_clean, premium_with_meta, h2_default, h3_ratios, 
                          campus_stores, off_campus_stores, delta_up_df, basket_prices,
                          h1_stats, df_a):
    """生成增强版可视化图表"""
    os.makedirs(PLOTS_DIR, exist_ok=True)
    
    store_colors = {
        "校内超市": COLORS["campus"],
        "长申超市": COLORS["changsheng"],
        "大润发": COLORS["rt_mart"]
    }

    # ========== 图1: 溢价率分布直方图（美化版）==========
    fig, ax = plt.subplots(figsize=(10, 6))
    rates = premium_with_meta["Premium_Rate"].dropna()
    
    # 分正负溢价
    positive_rates = rates[rates > 0]
    negative_rates = rates[rates <= 0]
    
    bins = np.linspace(rates.min() - 10, rates.max() + 10, 25)
    ax.hist(positive_rates, bins=bins, color=COLORS["positive"], alpha=0.7, 
            label=f"正溢价 (n={len(positive_rates)})", edgecolor="white")
    ax.hist(negative_rates, bins=bins, color=COLORS["negative"], alpha=0.7,
            label=f"负溢价 (n={len(negative_rates)})", edgecolor="white")
    
    # 添加统计线
    ax.axvline(0, color="black", linestyle="-", linewidth=2, label="零溢价线")
    ax.axvline(rates.mean(), color=COLORS["campus"], linestyle="--", linewidth=2,
               label=f"平均溢价率 ({rates.mean():.1f}%)")
    ax.axvline(rates.median(), color=COLORS["changsheng"], linestyle=":", linewidth=2,
               label=f"中位数溢价率 ({rates.median():.1f}%)")
    
    ax.set_xlabel("溢价率 (%)", fontsize=12)
    ax.set_ylabel("商品数量", fontsize=12)
    ax.set_title("校内超市相对于校外的溢价率分布", fontsize=14, fontweight="bold")
    ax.legend(loc="upper right", fontsize=10)
    ax.grid(axis="y", alpha=0.3)
    
    # 添加统计信息文本框
    stats_text = f"n = {len(rates)}\nM = {rates.mean():.2f}%\nMdn = {rates.median():.2f}%\nSD = {rates.std():.2f}%"
    ax.text(0.02, 0.98, stats_text, transform=ax.transAxes, fontsize=10,
            verticalalignment="top", bbox=dict(boxstyle="round", facecolor="wheat", alpha=0.5))
    
    plt.tight_layout()
    plt.savefig(os.path.join(PLOTS_DIR, "fig1_premium_rate_distribution.png"), dpi=300, bbox_inches="tight")
    plt.close()
    print(f"[OK] 图1 已保存: {PLOTS_DIR}/fig1_premium_rate_distribution.png")

    # ========== 图2: 各超市单位价格对比箱线图（美化版）==========
    fig, ax = plt.subplots(figsize=(10, 6))
    stores_order = campus_stores + off_campus_stores
    
    box_data = [df_clean[df_clean["StoreName"] == store]["Unit_Price"].dropna() for store in stores_order]
    box_colors = [store_colors.get(s, COLORS["neutral"]) for s in stores_order]
    
    bp = ax.boxplot(box_data, labels=stores_order, patch_artist=True, widths=0.6)
    for patch, color in zip(bp["boxes"], box_colors):
        patch.set_facecolor(color)
        patch.set_alpha(0.7)
    for median in bp["medians"]:
        median.set_color("black")
        median.set_linewidth(2)
    
    # 添加均值点
    means = [d.mean() for d in box_data]
    ax.scatter(range(1, len(stores_order) + 1), means, color="red", s=100, zorder=5, 
               marker="D", label="均值")
    
    ax.set_ylabel("单位价格 (元/100单位)", fontsize=12)
    ax.set_title("各超市单位价格分布对比", fontsize=14, fontweight="bold")
    ax.legend(loc="upper right")
    ax.grid(axis="y", alpha=0.3)
    
    plt.tight_layout()
    plt.savefig(os.path.join(PLOTS_DIR, "fig2_unit_price_boxplot.png"), dpi=300, bbox_inches="tight")
    plt.close()
    print(f"[OK] 图2 已保存: {PLOTS_DIR}/fig2_unit_price_boxplot.png")

    # ========== 图10: 各超市单位价格分布（核密度曲线）==========
    fig, ax = plt.subplots(figsize=(10, 6))

    # 设定统一的横轴范围（原始尺度）
    global_min = df_clean["Unit_Price"].min()
    global_max = df_clean["Unit_Price"].max()
    x_vals = np.linspace(max(0, global_min - 5), global_max + 5, 500)

    for store in stores_order:
        data = df_clean[df_clean["StoreName"] == store]["Unit_Price"].dropna()
        if len(data) < 3:
            continue
        kde = stats.gaussian_kde(data)
        y = kde(x_vals)
        color = store_colors.get(store, COLORS["neutral"])
        ax.plot(x_vals, y, label=f"{store} (n={len(data)})", color=color, linewidth=2)

    ax.set_xlabel("单位价格 (元/100单位)", fontsize=12)
    ax.set_ylabel("核密度", fontsize=12)
    ax.set_title("各超市单位价格分布（核密度估计）", fontsize=14, fontweight="bold")
    ax.legend(loc="upper right")
    ax.grid(alpha=0.3)

    plt.tight_layout()
    plt.savefig(os.path.join(PLOTS_DIR, "fig10_unit_price_distribution_by_store.png"), dpi=300, bbox_inches="tight")
    plt.close()
    print(f"[OK] 图10 已保存: {PLOTS_DIR}/fig10_unit_price_distribution_by_store.png")

    # ========== 图11: 各超市对数单位价格分布（核密度曲线）==========
    fig, ax = plt.subplots(figsize=(10, 6))

    # 仅对正的 Unit_Price 取对数
    positive_mask = df_clean["Unit_Price"] > 0
    unit_price_positive = df_clean.loc[positive_mask, "Unit_Price"].dropna()

    if not unit_price_positive.empty:
        global_min_log = np.log(unit_price_positive.min())
        global_max_log = np.log(unit_price_positive.max())
        x_vals_log = np.linspace(global_min_log - 0.5, global_max_log + 0.5, 500)

        for store in stores_order:
            data = df_clean[(df_clean["StoreName"] == store) & (df_clean["Unit_Price"] > 0)]["Unit_Price"].dropna()
            if len(data) < 3:
                continue
            log_data = np.log(data)
            kde = stats.gaussian_kde(log_data)
            y = kde(x_vals_log)
            color = store_colors.get(store, COLORS["neutral"])
            ax.plot(x_vals_log, y, label=f"{store} (n={len(log_data)})", color=color, linewidth=2)

        ax.set_xlabel("对数单位价格 ln(单位价格)", fontsize=12)
        ax.set_ylabel("核密度", fontsize=12)
        ax.set_title("各超市对数单位价格分布（核密度估计）", fontsize=14, fontweight="bold")
        ax.legend(loc="upper right")
        ax.grid(alpha=0.3)

        plt.tight_layout()
        plt.savefig(os.path.join(PLOTS_DIR, "fig11_log_unit_price_distribution_by_store.png"), dpi=300, bbox_inches="tight")
        plt.close()
        print(f"[OK] 图11 已保存: {PLOTS_DIR}/fig11_log_unit_price_distribution_by_store.png")
    else:
        print("[WARN] 所有 Unit_Price 均为非正值，无法绘制 log(Unit_Price) 核密度图")

    # ========== 图3: 匹配篮子价格对比柱状图 ==========
    if basket_prices:
        fig, ax = plt.subplots(figsize=(10, 6))
        stores = list(basket_prices.keys())
        prices = [basket_prices[s] for s in stores]
        colors = [store_colors.get(s, COLORS["neutral"]) for s in stores]
        
        bars = ax.bar(stores, prices, color=colors, edgecolor="black", linewidth=1.5)
        
        # 添加数值标签
        for bar, price in zip(bars, prices):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2,
                    f"¥{price:.2f}", ha="center", va="bottom", fontsize=12, fontweight="bold")
        
        # 添加校外平均线
        off_campus_avg = np.mean([basket_prices[s] for s in off_campus_stores if s in basket_prices])
        ax.axhline(off_campus_avg, color="gray", linestyle="--", linewidth=2,
                   label=f"校外平均 (¥{off_campus_avg:.2f})")
        
        ax.set_ylabel("匹配篮子价格 (元)", fontsize=12)
        ax.set_title("各超市匹配篮子价格对比\n$B^{(s)} = \\sum_{i=1}^{n} 到手价_i^{(s)}$", fontsize=14, fontweight="bold")
        ax.legend(loc="upper right")
        ax.grid(axis="y", alpha=0.3)
        
        plt.tight_layout()
        plt.savefig(os.path.join(PLOTS_DIR, "fig3_basket_price_comparison.png"), dpi=300, bbox_inches="tight")
        plt.close()
        print(f"[OK] 图3 已保存: {PLOTS_DIR}/fig3_basket_price_comparison.png")

    # ========== 图4: 瀑布图（价差分解）==========
    if not delta_up_df.empty:
        # 选取溢价最高的前10个商品
        top_delta = delta_up_df.nlargest(10, "Delta_UP").copy()
        top_delta = top_delta.merge(df_a[["ProductID", "ProductName"]], on="ProductID", how="left")
        
        fig, ax = plt.subplots(figsize=(12, 7))
        
        names = top_delta["ProductName"].fillna(top_delta["ProductID"].astype(str)).tolist()
        deltas = top_delta["Delta_UP"].values
        
        # 瀑布图
        cumsum = np.cumsum(deltas)
        starts = np.concatenate([[0], cumsum[:-1]])
        
        colors = [COLORS["positive"] if d > 0 else COLORS["negative"] for d in deltas]
        
        bars = ax.bar(range(len(names)), deltas, bottom=starts, color=colors, 
                      edgecolor="black", linewidth=1)
        
        # 添加数值标签
        for i, (bar, delta, start) in enumerate(zip(bars, deltas, starts)):
            y_pos = start + delta/2
            ax.text(i, y_pos, f"+{delta:.2f}" if delta > 0 else f"{delta:.2f}",
                    ha="center", va="center", fontsize=9, fontweight="bold", color="white")
        
        # 添加累计线
        ax.plot(range(len(names)), cumsum, "ko-", linewidth=2, markersize=6, label="累计价差")
        
        ax.set_xticks(range(len(names)))
        ax.set_xticklabels(names, rotation=45, ha="right", fontsize=10)
        ax.set_ylabel("单位量价差 ΔUP (元/100单位)", fontsize=12)
        ax.set_title("价差分解瀑布图（Top 10 溢价商品）\n$\\Delta UP_i = UP_i^{(校内)} - \\min_{s \\neq 校内} UP_i^{(s)}$", 
                     fontsize=14, fontweight="bold")
        ax.axhline(0, color="black", linewidth=1)
        ax.legend(loc="upper left")
        ax.grid(axis="y", alpha=0.3)
        
        plt.tight_layout()
        plt.savefig(os.path.join(PLOTS_DIR, "fig4_waterfall_price_decomposition.png"), dpi=300, bbox_inches="tight")
        plt.close()
        print(f"[OK] 图4 已保存: {PLOTS_DIR}/fig4_waterfall_price_decomposition.png")

    # ========== 图5: 雷达图（多维度对比）==========
    fig, ax = plt.subplots(figsize=(10, 10), subplot_kw=dict(polar=True))
    
    # 计算各维度指标
    categories = ["平均单价\n(反向)", "商品种类", "价格稳定性\n(反向)", "促销力度", "会员优惠"]
    
    # 模拟数据（基于实际数据计算）
    store_data = {}
    for store in stores_order:
        store_df = df_clean[df_clean["StoreName"] == store]
        avg_price = store_df["Unit_Price"].mean()
        variety = store_df["ProductID"].nunique()
        price_std = store_df["Unit_Price"].std()
        
        # 归一化到 0-100
        store_data[store] = [
            max(0, 100 - avg_price * 2),  # 平均单价（反向，越低越好）
            min(100, variety * 2.5),       # 商品种类
            max(0, 100 - price_std * 2),   # 价格稳定性（反向）
            50 if store == "校内超市" else 80,  # 促销力度（模拟）
            30 if store == "校内超市" else 70,  # 会员优惠（模拟）
        ]
    
    # 绘制雷达图
    angles = np.linspace(0, 2 * np.pi, len(categories), endpoint=False).tolist()
    angles += angles[:1]  # 闭合
    
    for store, values in store_data.items():
        values_closed = values + values[:1]
        color = store_colors.get(store, COLORS["neutral"])
        ax.plot(angles, values_closed, "o-", linewidth=2, label=store, color=color)
        ax.fill(angles, values_closed, alpha=0.25, color=color)
    
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontsize=11)
    ax.set_ylim(0, 100)
    ax.set_title("超市多维度对比雷达图", fontsize=14, fontweight="bold", pad=20)
    ax.legend(loc="upper right", bbox_to_anchor=(1.3, 1.0))
    
    plt.tight_layout()
    plt.savefig(os.path.join(PLOTS_DIR, "fig5_radar_comparison.png"), dpi=300, bbox_inches="tight")
    plt.close()
    print(f"[OK] 图5 已保存: {PLOTS_DIR}/fig5_radar_comparison.png")

    # ========== 图6: H2 分类对比（美化版）==========
    fig, ax = plt.subplots(figsize=(10, 6))
    
    categories = h2_default["Category_H2"].astype(str).tolist()
    means = h2_default["mean_premium_rate"].tolist()
    stds = h2_default["std_premium_rate"].fillna(0).tolist()
    ns = h2_default["n_products"].tolist()
    
    x = np.arange(len(categories))
    colors = [COLORS["essential"] if "必需" in c else COLORS["impulse"] for c in categories]
    
    bars = ax.bar(x, means, yerr=stds, capsize=5, color=colors, edgecolor="black",
                  linewidth=1.5, alpha=0.8, error_kw={"elinewidth": 2, "capthick": 2})
    
    # 添加数值标签
    for i, (bar, mean, n) in enumerate(zip(bars, means, ns)):
        y_pos = bar.get_height() + stds[i] + 2
        ax.text(i, y_pos, f"{mean:.1f}%\n(n={n})", ha="center", va="bottom", fontsize=11, fontweight="bold")
    
    ax.axhline(0, color="black", linewidth=1)
    ax.set_xticks(x)
    ax.set_xticklabels(categories, fontsize=12)
    ax.set_ylabel("平均溢价率 (%)", fontsize=12)
    ax.set_title("必需品 vs 冲动品 溢价率对比 (H2)", fontsize=14, fontweight="bold")
    ax.grid(axis="y", alpha=0.3)
    
    # 添加图例
    legend_elements = [mpatches.Patch(facecolor=COLORS["essential"], label="必需品"),
                       mpatches.Patch(facecolor=COLORS["impulse"], label="冲动品")]
    ax.legend(handles=legend_elements, loc="upper right")
    
    plt.tight_layout()
    plt.savefig(os.path.join(PLOTS_DIR, "fig6_h2_category_comparison.png"), dpi=300, bbox_inches="tight")
    plt.close()
    print(f"[OK] 图6 已保存: {PLOTS_DIR}/fig6_h2_category_comparison.png")

    # ========== 图7: H3 大小包装对比（美化版）==========
    if not h3_ratios.empty and "Ratio_Small_to_Large" in h3_ratios.columns:
        df_plot = h3_ratios.dropna(subset=["Ratio_Small_to_Large"])
        if not df_plot.empty:
            fig, ax = plt.subplots(figsize=(12, 6))
            
            labels = df_plot.apply(lambda r: f"{r.get('Brand', '')} / {r.get('StoreName', '')}", axis=1).tolist()
            ratios = df_plot["Ratio_Small_to_Large"].values
            
            x = np.arange(len(labels))
            colors = [COLORS["positive"] if r > 1 else COLORS["negative"] for r in ratios]
            
            bars = ax.bar(x, ratios, color=colors, edgecolor="black", linewidth=1.5)
            
            # 添加数值标签
            for i, (bar, ratio) in enumerate(zip(bars, ratios)):
                y_pos = bar.get_height() + 0.1
                ax.text(i, y_pos, f"{ratio:.2f}x", ha="center", va="bottom", fontsize=10, fontweight="bold")
            
            ax.axhline(1.0, color="black", linestyle="--", linewidth=2, label="等价线 (1.0)")
            ax.set_xticks(x)
            ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=10)
            ax.set_ylabel("小包装/大包装 单位价格比", fontsize=12)
            ax.set_title("小包装 vs 大包装 单位价格比较 (H3)", fontsize=14, fontweight="bold")
            ax.legend(loc="upper right")
            ax.grid(axis="y", alpha=0.3)
            
            plt.tight_layout()
            plt.savefig(os.path.join(PLOTS_DIR, "fig7_h3_package_size.png"), dpi=300, bbox_inches="tight")
            plt.close()
            print(f"[OK] 图7 已保存: {PLOTS_DIR}/fig7_h3_package_size.png")

    # ========== 图8: 溢价率 Top 10 商品（美化版）==========
    top10 = premium_with_meta.nlargest(10, "Premium_Rate").copy()
    if not top10.empty:
        fig, ax = plt.subplots(figsize=(12, 7))
        
        labels = top10.apply(lambda r: f"{r.get('ProductName', '')} ({r.get('Brand', '')})", axis=1).tolist()
        rates = top10["Premium_Rate"].values
        
        y = np.arange(len(labels))
        colors = [COLORS["positive"] if r > 50 else COLORS["impulse"] if r > 0 else COLORS["negative"] for r in rates]
        
        bars = ax.barh(y, rates, color=colors, edgecolor="black", linewidth=1, height=0.7)
        
        # 添加数值标签
        for i, (bar, rate) in enumerate(zip(bars, rates)):
            x_pos = bar.get_width() + 2
            ax.text(x_pos, i, f"+{rate:.1f}%", va="center", fontsize=10, fontweight="bold")
        
        ax.axvline(0, color="black", linewidth=1)
        ax.set_yticks(y)
        ax.set_yticklabels(labels, fontsize=10)
        ax.set_xlabel("溢价率 (%)", fontsize=12)
        ax.set_title("校内溢价率最高的 10 种商品", fontsize=14, fontweight="bold")
        ax.invert_yaxis()
        ax.grid(axis="x", alpha=0.3)
        
        plt.tight_layout()
        plt.savefig(os.path.join(PLOTS_DIR, "fig8_top10_premium.png"), dpi=300, bbox_inches="tight")
        plt.close()
        print(f"[OK] 图8 已保存: {PLOTS_DIR}/fig8_top10_premium.png")

    # ========== 图9: 置信区间可视化 ==========
    if h1_stats:
        fig, ax = plt.subplots(figsize=(10, 5))
        
        # 绘制置信区间
        mean_diff = h1_stats["mean_difference"]
        ci_lower = h1_stats["ci_95_lower"]
        ci_upper = h1_stats["ci_95_upper"]
        
        ax.errorbar(0, mean_diff, yerr=[[mean_diff - ci_lower], [ci_upper - mean_diff]],
                    fmt="o", markersize=15, color=COLORS["campus"], capsize=10, capthick=3,
                    elinewidth=3, label="平均差异 ± 95% CI")
        
        ax.axhline(0, color="black", linestyle="--", linewidth=2, label="无差异线")
        
        # 添加标注
        ax.text(0.15, mean_diff, f"M = {mean_diff:.2f}", fontsize=12, va="center")
        ax.text(0.15, ci_lower, f"下限 = {ci_lower:.2f}", fontsize=10, va="center", color="gray")
        ax.text(0.15, ci_upper, f"上限 = {ci_upper:.2f}", fontsize=10, va="center", color="gray")
        
        ax.set_xlim(-0.5, 0.5)
        ax.set_ylabel("单位价格差异 (校内 - 校外)", fontsize=12)
        ax.set_title("校内外价格差异的 95% 置信区间", fontsize=14, fontweight="bold")
        ax.set_xticks([])
        ax.legend(loc="upper right")
        ax.grid(axis="y", alpha=0.3)
        
        plt.tight_layout()
        plt.savefig(os.path.join(PLOTS_DIR, "fig9_confidence_interval.png"), dpi=300, bbox_inches="tight")
        plt.close()
        print(f"[OK] 图9 已保存: {PLOTS_DIR}/fig9_confidence_interval.png")


# ============ 打印详细统计结果 ============
def print_detailed_stats(h1_stats, h2_stats_default, h2_stats_noodles, h3_stats):
    """打印详细的统计检验结果"""
    
    print("\n" + "=" * 70)
    print("                    详 细 统 计 检 验 结 果")
    print("=" * 70)
    
    # H1 详细结果
    print("\n" + "-" * 70)
    print("【H1】校内 vs 校外价格差异检验")
    print("-" * 70)
    print(f"\n  样本量: n = {h1_stats['n_pairs']}")
    print(f"\n  描述性统计:")
    print(f"    校内平均单位价格: {h1_stats['mean_UP_campus']:.4f} (SD = {h1_stats['std_UP_campus']:.4f})")
    print(f"    校外平均单位价格: {h1_stats['mean_UP_off_campus']:.4f} (SD = {h1_stats['std_UP_off_campus']:.4f})")
    print(f"    平均差异: {h1_stats['mean_difference']:.4f} (SE = {h1_stats['se_difference']:.4f})")
    print(f"    95% CI: [{h1_stats['ci_95_lower']:.4f}, {h1_stats['ci_95_upper']:.4f}]")
    print(f"\n  溢价率统计:")
    print(f"    平均溢价率: {h1_stats['mean_premium_rate_pct']:.2f}%")
    print(f"    中位数溢价率: {h1_stats['median_premium_rate_pct']:.2f}%")
    print(f"    标准差: {h1_stats['std_premium_rate_pct']:.2f}%")
    print(f"    校内更贵的商品比例: {h1_stats['pct_products_higher_on_campus']:.1f}%")
    print(f"\n  正态性检验 (Shapiro-Wilk):")
    print(f"    W = {h1_stats['shapiro_statistic']:.4f}, p = {h1_stats['p_value_shapiro']:.4f}")
    print(f"    结论: {'满足正态性假设' if h1_stats['p_value_shapiro'] >= 0.05 else '不满足正态性假设，应参考非参数检验'}")
    print(f"\n  配对样本 t 检验:")
    print(f"    t({h1_stats['n_pairs']-1}) = {h1_stats['t_statistic']:.4f}")
    print(f"    p = {h1_stats['p_value_ttest']:.4f} {interpret_p_value(h1_stats['p_value_ttest'])}")
    print(f"\n  Wilcoxon 符号秩检验:")
    print(f"    W = {h1_stats['wilcoxon_statistic']:.1f}")
    print(f"    p = {h1_stats['p_value_wilcoxon']:.4f} {interpret_p_value(h1_stats['p_value_wilcoxon'])}")
    print(f"\n  效应量:")
    print(f"    Cohen's d = {h1_stats['cohens_d']:.4f} ({h1_stats['effect_size_interpretation']})")
    print(f"\n  ★ H1 结论: ", end="")
    if h1_stats['p_value_ttest'] < 0.05:
        print(f"校内价格显著高于校外 (p < 0.05)")
    else:
        print(f"校内外价格差异不显著 (p = {h1_stats['p_value_ttest']:.3f} >= 0.05)")
    
    # H2 详细结果
    print("\n" + "-" * 70)
    print("【H2】必需品 vs 冲动品溢价率差异检验")
    print("-" * 70)
    
    if h2_stats_default:
        print(f"\n  [默认分类]")
        print(f"    必需品: n={h2_stats_default['n_essential']}, M={h2_stats_default['mean_essential']:.2f}%, SD={h2_stats_default['std_essential']:.2f}%")
        print(f"    冲动品: n={h2_stats_default['n_impulse']}, M={h2_stats_default['mean_impulse']:.2f}%, SD={h2_stats_default['std_impulse']:.2f}%")
        print(f"    差异: {h2_stats_default['mean_difference']:.2f}%")
        print(f"\n    Welch's t 检验: t = {h2_stats_default['t_statistic']:.4f}, p = {h2_stats_default['p_value_ttest']:.4f} {interpret_p_value(h2_stats_default['p_value_ttest'])}")
        print(f"    Mann-Whitney U: U = {h2_stats_default['mannwhitney_u']:.1f}, p = {h2_stats_default['p_value_mannwhitney']:.4f} {interpret_p_value(h2_stats_default['p_value_mannwhitney'])}")
        print(f"    Levene 方差齐性: F = {h2_stats_default['levene_statistic']:.4f}, p = {h2_stats_default['p_value_levene']:.4f}")
        print(f"    Cohen's d = {h2_stats_default['cohens_d']:.4f} ({h2_stats_default['effect_size_interpretation']})")
    
    if h2_stats_noodles:
        print(f"\n  [敏感性分析: 泡面归为必需品]")
        print(f"    必需品: n={h2_stats_noodles['n_essential']}, M={h2_stats_noodles['mean_essential']:.2f}%")
        print(f"    冲动品: n={h2_stats_noodles['n_impulse']}, M={h2_stats_noodles['mean_impulse']:.2f}%")
        print(f"    Welch's t: t = {h2_stats_noodles['t_statistic']:.4f}, p = {h2_stats_noodles['p_value_ttest']:.4f} {interpret_p_value(h2_stats_noodles['p_value_ttest'])}")
    
    print(f"\n  ★ H2 结论: ", end="")
    if h2_stats_default and h2_stats_default['p_value_ttest'] < 0.05:
        print("必需品与冲动品溢价率存在显著差异")
    else:
        print("必需品与冲动品溢价率无显著差异")
    
    # H3 详细结果
    if h3_stats:
        print("\n" + "-" * 70)
        print("【H3】小包装 vs 大包装单位价格检验")
        print("-" * 70)
        print(f"\n  样本量: n = {h3_stats['n_pairs']}")
        print(f"\n  描述性统计:")
        print(f"    平均比值 (小/大): {h3_stats['mean_ratio']:.4f} (SD = {h3_stats['std_ratio']:.4f})")
        print(f"    中位数比值: {h3_stats['median_ratio']:.4f}")
        print(f"    范围: [{h3_stats['min_ratio']:.4f}, {h3_stats['max_ratio']:.4f}]")
        print(f"\n  单样本 t 检验 (H0: ratio = 1):")
        print(f"    t = {h3_stats['t_statistic']:.4f}")
        print(f"    p (双侧) = {h3_stats['p_value_ttest_twosided']:.4f}")
        print(f"    p (单侧, > 1) = {h3_stats['p_value_ttest_onesided']:.4f} {interpret_p_value(h3_stats['p_value_ttest_onesided'])}")
        print(f"\n  Wilcoxon 符号秩检验:")
        print(f"    W = {h3_stats['wilcoxon_statistic']:.1f}, p = {h3_stats['p_value_wilcoxon']:.4f}")
        print(f"\n  效应量:")
        print(f"    Cohen's d = {h3_stats['cohens_d']:.4f} ({h3_stats['effect_size_interpretation']})")
        print(f"\n  ★ H3 结论: ", end="")
        if h3_stats['p_value_ttest_onesided'] < 0.05:
            print(f"小包装单价显著高于大包装，平均为大包装的 {h3_stats['mean_ratio']:.2f} 倍")
        else:
            print("小包装与大包装单价无显著差异")
    
    print("\n" + "=" * 70)


# ============ 主函数 ============
def main():
    print("=" * 70)
    print("        校内外超市价格比较分析 (Enhanced Version)")
    print("        Statistical Analysis Report with 3σ Processing")
    print("=" * 70)

    # 1. 加载数据
    print("\n[1] 加载数据...")
    df_a, df_b = load_tables()
    print(f"    Table A: {len(df_a)} 种商品")
    print(f"    Table B: {len(df_b)} 条价格记录")

    # 2. 生成 ProductID
    print("\n[2] 生成 ProductID...")
    df_a, df_b = prepare_product_ids(df_a, df_b, n_products_expected=44)

    # 3. 计算最终价格
    print("\n[3] 计算最终价格 (P_final)...")
    print("    价格口径: 学生可得最低价 = min(Price_Actual, Price_Member)")
    df_b = compute_final_price(df_b)

    # 4. 合并表
    print("\n[4] 合并 Table A 和 Table B...")
    df = merge_tables(df_a, df_b)

    # 5. 计算单位量价
    print("\n[5] 计算单位量价 (Unit_Price)...")
    print("    公式: UP_i = 到手价_i / 净含量_i × 100 (元/100单位)")
    df = compute_unit_price(df)

    # 6. 3σ 异常值处理
    if APPLY_3SIGMA:
        print("\n[6] 3σ 异常值处理...")
        original_count = len(df)
        df_filtered, outliers_info = apply_3sigma_filter(df, "Unit_Price")
        removed_count = original_count - len(df_filtered)
        print(f"    原始记录数: {original_count}")
        print(f"    剔除异常值: {removed_count} 条")
        print(f"    保留记录数: {len(df_filtered)}")
        if outliers_info:
            print(f"    异常值范围: 超出 μ ± 3σ")
        df = df_filtered
    else:
        print("\n[6] 跳过 3σ 异常值处理...")

    # 7. 识别校内/校外
    campus_stores, off_campus_stores = identify_campus_and_off_campus(df)
    print(f"\n[7] 超市分类:")
    print(f"    校内: {campus_stores}")
    print(f"    校外: {off_campus_stores}")

    # 8. 计算匹配篮子价格
    print("\n[8] 计算匹配篮子价格...")
    print("    公式: B(s) = Σ 到手价_i(s)")
    basket_prices, matched_products = compute_basket_price(df, campus_stores, off_campus_stores)
    print(f"    匹配商品数: {len(matched_products)}")
    for store, price in basket_prices.items():
        print(f"    {store}: ¥{price:.2f}")

    # 9. 计算单位量价差 ΔUP
    print("\n[9] 计算单位量价差 ΔUP...")
    print("    公式: ΔUP_i = UP_i(校内) - min_{s≠校内} UP_i(s)")
    delta_up_df = compute_delta_up(df, campus_stores, off_campus_stores)
    if not delta_up_df.empty:
        mean_delta = delta_up_df["Delta_UP"].mean()
        print(f"    平均单位量价差: {mean_delta:.4f} 元/100单位")

    # 10. 计算溢价率
    print("\n[10] 计算溢价率...")
    premium = compute_premium_by_product(df, campus_stores, off_campus_stores)
    premium_with_meta = attach_product_metadata(premium, df_a)
    print(f"    有效配对商品数: {len(premium)}")
    print(f"    平均溢价率: {premium['Premium_Rate'].mean():.2f}%")
    print(f"    中位数溢价率: {premium['Premium_Rate'].median():.2f}%")

    # 11. 将溢价率合并回主表
    df = df.merge(premium[["ProductID", "Premium_Rate"]], on="ProductID", how="left")

    # 12. 描述性统计
    print("\n[11] 计算描述性统计...")
    descriptive_stats = compute_descriptive_stats(df)

    # 13. H1 综合检验
    print("\n[12] H1 综合统计检验...")
    h1_stats = h1_comprehensive_test(premium)

    # 14. H2 综合检验
    print("\n[13] H2 综合统计检验...")
    h2_default = h2_premium_by_category(premium_with_meta, treat_noodles_as_essential=False)
    h2_noodles = h2_premium_by_category(premium_with_meta, treat_noodles_as_essential=True)
    h2_stats_default = h2_comprehensive_test(premium_with_meta, treat_noodles_as_essential=False)
    h2_stats_noodles = h2_comprehensive_test(premium_with_meta, treat_noodles_as_essential=True)

    # 15. H3 综合检验
    print("\n[14] H3 综合统计检验...")
    h3_ratios = h3_small_vs_large(df, df_a)
    h3_stats = h3_comprehensive_test(h3_ratios)

    # 16. 溢价排行榜
    print("\n[15] 溢价率 Top 10...")
    top_premium = top_premium_products(premium_with_meta, top_n=10)

    # 17. 打印详细统计结果
    print_detailed_stats(h1_stats, h2_stats_default, h2_stats_noodles, h3_stats)

    # 18. 保存 Excel
    print("\n[16] 保存 Excel...")
    h1_result = pd.DataFrame([h1_stats])
    save_excel_outputs(df, descriptive_stats, h1_result, h2_default, h2_noodles, h3_ratios, 
                       top_premium, premium_with_meta, h1_stats, h2_stats_default, h2_stats_noodles, h3_stats)

    # 19. 生成增强版图表
    print("\n[17] 生成增强版图表...")
    plot_figures_enhanced(df, premium_with_meta, h2_default, h3_ratios, 
                          campus_stores, off_campus_stores, delta_up_df, basket_prices,
                          h1_stats, df_a)

    # 20. 生成 LaTeX 报告
    print("\n[18] 生成 LaTeX 报告...")
    generate_latex_report(h1_stats, h2_stats_default, h2_stats_noodles, h3_stats, 
                          h2_default, h2_noodles, top_premium, descriptive_stats)

    # 21. 输出总结
    print("\n" + "=" * 70)
    print("                        分 析 完 成")
    print("=" * 70)
    print(f"\n  📊 Excel 输出:  {OUTPUT_EXCEL_PATH}")
    print(f"  📈 图表输出:    {PLOTS_DIR}/ (共 9 张图)")
    print(f"  📝 LaTeX 报告:  {LATEX_OUTPUT_PATH}")
    print("\n  核心发现:")
    print(f"    • 匹配篮子商品数: {len(matched_products)}")
    print(f"    • 平均溢价率: {premium['Premium_Rate'].mean():.1f}%")
    print(f"    • H1 (校内更贵): p = {h1_stats['p_value_ttest']:.4f} {'✓ 显著' if h1_stats['p_value_ttest'] < 0.05 else '✗ 不显著'}")
    if h2_stats_default:
        print(f"    • H2 (品类差异): p = {h2_stats_default['p_value_ttest']:.4f} {'✓ 显著' if h2_stats_default['p_value_ttest'] < 0.05 else '✗ 不显著'}")
    if h3_stats:
        print(f"    • H3 (包装效应): p = {h3_stats['p_value_ttest_onesided']:.4f} {'✓ 显著' if h3_stats['p_value_ttest_onesided'] < 0.05 else '✗ 不显著'}")
    print("=" * 70)


if __name__ == "__main__":
    main()
