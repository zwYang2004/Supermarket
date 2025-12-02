# Supermarket
# 校园封闭市场下的价格差异研究 – 代码与数据仓库

本仓库用于复现课程作业 **《校内外超市价格比较研究》** 的全部数据处理与统计分析过程。  
研究场景为扬州大学瘦西湖校区校园教育超市及其周边两家社会超市（长申超市、大润发），核心问题是：**在校园这种相对封闭的市场环境下，校内超市是否比校外超市“更贵”？**

---

## 1. 研究问题与假设

根据课堂作业要求，重点检验以下三个假设：

- **H1 系统性溢价**：校内超市的整体价格显著高于校外超市。
- **H2 品类差异**：必需品（如牛奶、卫生巾）的溢价程度高于冲动消费品（如零食）。
- **H3 规格效应**：小包装商品的单位价格显著高于大包装。

所有价格比较都基于**单位量价（Unit Price）** 或 **严格匹配的商品篮子（Matched Basket）**，以保证统计上的可比性。

---

## 2. 数据来源

本项目使用的数据来自实地手动采集，并整理为以下两个核心表格：

- `Table A Product List.xlsx`  
  商品清单与元数据（`ProductID`、品牌、品类、单位、净含量等），用于：
  - 定义研究的商品集合（44 个商品对，29 个完全匹配商品）；
  - 支持 H2（必需品 vs 冲动品）和 H3（大小规格）的分类；
  - 提供计算单位量价所需的净含量信息。

- `Table B Price Data.xlsx`  
  不同超市、不同采集日期下的价格记录（门店名称、采集日期、实际价格、会员价、原价等），用于：
  - 计算学生“实际支付价格”（到手价 `P_final`）；
  - 构造匹配篮子总价与单位量价。

**公开数据链接**

为方便复现与二次分析，原始价格数据和整理后的分析数据也同步发布在 Google Sheets：

> https://docs.google.com/spreadsheets/d/19ySf4G_klmk4n2IzGUdF69TiCI0pBh4BRroPfq1tn1Y/edit?usp=sharing

你可以通过上述链接在线查看或下载为 Excel 文件，然后放到本仓库根目录（或 `data/` 目录）中使用。

---

## 3. 代码结构

- [supermarket_price_analysis.py](cci:7://file:///Users/yangzw/Documents/%E6%89%AC%E5%B7%9E%E5%A4%A7%E5%AD%A6/%E6%95%B0%E5%AD%A6%E4%B8%93%E4%B8%9A%E8%AF%BE/%E6%95%99%E8%82%B2%E7%BB%9F%E8%AE%A1%E5%AD%A6/supermarket_price_analysis.py:0:0-0:0)  
  主分析脚本，完成：
  - 读取 Table A / Table B；
  - 生成或校正 `ProductID`；
  - 计算到手价 `P_final = min(Price_Actual, Price_Member)`；
  - 合并商品元数据，计算单位量价 `Unit_Price`；
  - 使用 3σ 原则剔除异常值（可通过 `APPLY_3SIGMA` 开关控制）；
  - 识别校内 vs 校外超市；
  - 构造匹配篮子总价并进行宏观对比；
  - 计算每个商品的溢价率（Premium Rate）；
  - 对 H1/H2/H3 进行统计检验（t 检验、Wilcoxon、Mann-Whitney、效应量等）；
  - 输出汇总 Excel：`output_supermarket_analysis.xlsx`；
  - 生成多张图表（价格分布、溢价 Top10、规格效应雷达图等）到 `plots/` 目录；


- [advanced_analysis.py](cci:7://file:///Users/yangzw/Documents/%E6%89%AC%E5%B7%9E%E5%A4%A7%E5%AD%A6/%E6%95%B0%E5%AD%A6%E4%B8%93%E4%B8%9A%E8%AF%BE/%E6%95%99%E8%82%B2%E7%BB%9F%E8%AE%A1%E5%AD%A6/advanced_analysis.py:0:0-0:0)  
  辅助分析脚本，基于 `output_supermarket_analysis.xlsx`：
  - 计算校内/校外单位量价的商品均值；
  - 绘制“校内 vs 校外单位价格”散点图（`fig10_price_scatter.png`）并画出 `y=x` 参考线；
  - 构建对数线性回归模型（OLS），将回归结果输出到 `ols_results.txt`，并给出核心系数（校内效应、小包装效应）的解读。

- [analysis_report.tex](cci:7://file:///Users/yangzw/Documents/%E6%89%AC%E5%B7%9E%E5%A4%A7%E5%AD%A6/%E6%95%B0%E5%AD%A6%E4%B8%93%E4%B8%9A%E8%AF%BE/%E6%95%99%E8%82%B2%E7%BB%9F%E8%AE%A1%E5%AD%A6/analysis_report.tex:0:0-0:0)（可选放入仓库）  
  课程论文的 LaTeX 源文件，系统呈现研究背景、理论框架、数据处理流程与统计结果。

- `ref.bib`
  LaTeX 论文使用的参考文献 BibTeX 数据库。

---

## 4. 环境依赖

建议使用 Python 3.9 及以上版本。

需要安装的第三方包（见 `requirements.txt`）：

```txt
pandas
numpy
scipy
statsmodels
matplotlib
seaborn
openpyxl
