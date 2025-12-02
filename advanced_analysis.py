import pandas as pd
import numpy as np
import statsmodels.api as sm
import matplotlib.pyplot as plt
import seaborn as sns

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS'] 
plt.rcParams['axes.unicode_minus'] = False

def run_analysis():
    try:
        # 读取数据
        df = pd.read_excel('output_supermarket_analysis.xlsx')
        print(f"数据加载成功，共 {len(df)} 条记录")
        
        # --- 1. 准备散点图数据 ---
        # 校内：按 ProductID 分组取均值（处理多次采集）
        df_campus = df[df['StoreName'] == '校内超市'].groupby('ProductID')['Unit_Price'].mean()
        # 校外：按 ProductID 分组取均值
        df_off = df[df['StoreName'] != '校内超市'].groupby('ProductID')['Unit_Price'].mean()
        
        # 合并为 DataFrame
        merged = pd.DataFrame({
            'Campus': df_campus,
            'OffCampus': df_off
        }).dropna()
        
        print(f"匹配成功的商品数: {len(merged)}")
        
        # --- 2. 绘制散点图 ---
        plt.figure(figsize=(8, 8))
        plt.scatter(merged['OffCampus'], merged['Campus'], alpha=0.6, c='blue', edgecolors='w', s=80)
        
        # 画 y=x 参考线
        max_val = max(merged['Campus'].max(), merged['OffCampus'].max()) * 1.1
        plt.plot([0, max_val], [0, max_val], 'r--', linewidth=2, label='y=x (价格持平线)')
        
        plt.xlabel('校外平均单位价格 (元/100单位)', fontsize=12)
        plt.ylabel('校内单位价格 (元/100单位)', fontsize=12)
        plt.title('校内 vs 校外单位价格对比散点图', fontsize=14)
        plt.legend(fontsize=11)
        plt.grid(True, alpha=0.3)
        
        plt.tight_layout()
        plt.savefig('fig10_price_scatter.png', dpi=300, bbox_inches='tight')
        plt.close()
        print("✓ 散点图已保存为 fig10_price_scatter.png")
        
        # --- 3. 运行 OLS 回归 ---
        reg_df = df.copy()
        
        # 剔除 Unit_Price <= 0 的异常值
        reg_df = reg_df[reg_df['Unit_Price'] > 0]
        
        reg_df['log_UP'] = np.log(reg_df['Unit_Price'])
        reg_df['Is_Campus'] = (reg_df['StoreName'] == '校内超市').astype(int)
        
        # 判断小包装：NetContent < 150
        reg_df['Is_Small_Pack'] = (reg_df['NetContent'] < 150).astype(int)
        
        # 哑变量处理: Category_H2
        category_dummies = pd.get_dummies(reg_df['Category_H2'], prefix='Cat', drop_first=True)
        category_dummies = category_dummies.astype(int)
        
        # 合并所有自变量
        X = pd.concat([reg_df[['Is_Campus', 'Is_Small_Pack']], category_dummies], axis=1)
        X = sm.add_constant(X)
        y = reg_df['log_UP']
        
        model = sm.OLS(y, X).fit()
        
        print("\n" + "="*60)
        print("OLS 回归结果")
        print("="*60)
        print(model.summary())
        
        # 保存回归结果
        with open('ols_results.txt', 'w') as f:
            f.write(model.summary().as_text())
        print("\n✓ 回归结果已保存为 ols_results.txt")
        
        # --- 4. 输出关键系数解读 ---
        print("\n" + "="*60)
        print("关键系数解读")
        print("="*60)
        
        coef_campus = model.params.get('Is_Campus', 0)
        pval_campus = model.pvalues.get('Is_Campus', 1)
        coef_small = model.params.get('Is_Small_Pack', 0)
        pval_small = model.pvalues.get('Is_Small_Pack', 1)
        
        print(f"Is_Campus 系数: {coef_campus:.4f}, p值: {pval_campus:.4f}")
        if pval_campus < 0.05:
            print(f"  → 校内效应显著，校内价格比校外高 {(np.exp(coef_campus)-1)*100:.1f}%")
        else:
            print(f"  → 校内效应不显著 (p > 0.05)")
            
        print(f"Is_Small_Pack 系数: {coef_small:.4f}, p值: {pval_small:.4f}")
        if pval_small < 0.05:
            print(f"  → 小包装效应显著，小包装单价比大包装高 {(np.exp(coef_small)-1)*100:.1f}%")
        else:
            print(f"  → 小包装效应不显著 (p > 0.05)")
            
        print("\n✓ 分析完成！")
        
    except Exception as e:
        print(f"分析脚本运行出错: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    run_analysis()
