# 定义文件路径
zhibiao = r"J:\BaiduNetdiskDownload\05马克\8亿+企业查数据\03指标计算\企业查区县统计合并_新指标和原始指标merged_data归一化_指标计算结果.xlsx"
xingzheng = r"J:\BaiduNetdiskDownload\05马克\8亿+企业查数据\01区县统计\行政边界_gjdl_县.xlsx"

# 读取表格
print("正在读取表格...")
import pandas as pd

# 读取行政边界数据
df_xingzheng = pd.read_excel(xingzheng)

# 读取指标数据
df_zhibiao = pd.read_excel(zhibiao)

# 连接行政边界和指标数据
final_df = pd.merge(df_xingzheng, df_zhibiao, how='left', left_on='gjdl_xdm', right_on='district_code')

# 导出合并后的数据到Excel
output_excel_path = r"J:\BaiduNetdiskDownload\05马克\8亿+企业查数据\03指标计算\合并数据结果.xlsx"
final_df.to_excel(output_excel_path, index=False)
print(f"合并数据已导出到: {output_excel_path}")

print("数据连接完成。")
import os

os.environ['USE_PYGEOS'] = '0'
import geopandas as gpd
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties

# 读取GIS数据
print("正在读取GIS数据...")
input_gdb_path = r"J:\BaiduNetdiskDownload\05马克\8亿+企业查数据\01区县统计\行政边界.gdb"
input_xian_layer = "国家地理_县级行政区"
gdf = gpd.read_file(input_gdb_path, layer=input_xian_layer)

print("GIS数据字段:")
for column in gdf.columns:
    print(column)

# 连接GIS数据和指标计算结果
print("正在连接GIS数据和指标计算结果...")
# 确保 'xian_code' 和 'gjdl_xdm' 列的数据类型一致
gdf['xian_code'] = gdf['xian_code'].astype(str)
final_df['gjdl_xdm'] = final_df['gjdl_xdm'].astype(str)
merged = gdf.merge(final_df, left_on="xian_code", right_on="gjdl_xdm")

# 设置中文字体
font = FontProperties(fname=r"C:\Windows\Fonts\simhei.ttf", size=10)

# 定义指标列表
indicators = ['高科技', '高效能', '高质量', '综合指数']

# 创建输出目录
output_dir = r'J:\BaiduNetdiskDownload\05马克\8亿+企业查数据\03指标计算\全国输出图表'
os.makedirs(output_dir, exist_ok=True)

# 创建Excel写入器
excel_path = os.path.join(output_dir, '全国指标百强.xlsx')
with pd.ExcelWriter(excel_path) as writer:
    # 为每个指标创建全国分布图和百强列表
    for indicator in indicators:
        # 将缺失值填充为0
        merged[indicator] = merged[indicator].fillna(0)

        # 创建图表
        fig, ax = plt.subplots(figsize=(20, 16))
        merged.plot(column=indicator, cmap='Blues', linewidth=0, edgecolor='none', ax=ax, legend=True, legend_kwds={'orientation': 'horizontal'})

        # 设置标题和轴标签
        plt.title(f'全国 - {indicator}分布图', fontproperties=font, fontsize=20)
        plt.axis('off')

        # 获取百强数据
        top_100 = merged.sort_values(by=indicator, ascending=False).head(100)
        
        # 打印前10强县级名称
        print(f"\n{indicator}指标前10强县级名称:")
        for i, (index, row) in enumerate(top_100.head(10).iterrows(), 1):
            print(f"{i}. {row['name_xian_x']}")

        # 在图上标注百强的名称，避免标注互相重叠
        from adjustText import adjust_text
        
        texts = []
        # for idx, row in top_100.iterrows():
        #     # 创建注释对象（县级名称）
        #     text = plt.text(row.geometry.centroid.x, row.geometry.centroid.y, row['name_xian_x'],
        #                     fontproperties=font, fontsize=8, ha='center', va='center')
        #     texts.append(text)

        # 标出省边界
        provinces = merged['name_sheng_x'].unique()
        for province in provinces:
            province_data = merged[merged['name_sheng_x'] == province]
            province_boundary = province_data.dissolve(by='name_sheng_x')
            province_boundary.boundary.plot(ax=ax, color='gray', linewidth=1)
            # 计算省份边界的中心点
            bounds = province_boundary.total_bounds
            center_x = (bounds[0] + bounds[2]) / 2
            center_y = (bounds[1] + bounds[3]) / 2
            # text = plt.text(center_x, center_y, province,
            #                 fontproperties=font, fontsize=12, ha='center', va='center', fontweight='bold')
            # texts.append(text)

        # 使用adjust_text函数调整文本位置，避免重叠
        adjust_text(texts, arrowprops=dict(arrowstyle='-', color='lightgray', lw=0.5))

        # 保存图表
        plt.savefig(os.path.join(output_dir, f'{indicator}_全国分布图.png'), dpi=300, bbox_inches='tight')
        plt.close()

        # 导出百强到Excel的不同sheet
        top_100[['name_sheng_x', 'name_shi_x', 'name_xian_x', indicator]].to_excel(writer, sheet_name=indicator, index=False)
        print(f"已将 {indicator} 百强数据导出到Excel的 {indicator} 工作表")

        print(f"已生成 {indicator} 全国分布图和百强Excel工作表。")

print(f"所有图表已生成完毕，百强数据已导出到 {excel_path}")

# 导出全国gpkg
gpkg_path = os.path.join(output_dir, '全国指标数据.gpkg')
merged.to_file(gpkg_path, driver='GPKG')
print(f"全国指标数据已导出到GeoPackage文件: {gpkg_path}")
