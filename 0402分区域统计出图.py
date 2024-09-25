# 定义文件路径
zhibiao = r"J:\BaiduNetdiskDownload\05马克\8亿+企业查数据\03指标计算\企业查区县统计合并_新指标和原始指标merged_data归一化_指标计算结果.xlsx"#J:\BaiduNetdiskDownload\05马克\8亿+企业查数据
xingzheng = r"J:\BaiduNetdiskDownload\05马克\8亿+企业查数据\03指标计算\行政边界_gjdl_县.xlsx"
chengshiqun = r"J:\BaiduNetdiskDownload\05马克\8亿+企业查数据\03指标计算\城市群.xlsx"

# 读取表格
print("正在读取表格...")
import pandas as pd

# 读取行政边界和城市群数据
df_xingzheng = pd.read_excel(xingzheng)
df_chengshiqun = pd.read_excel(chengshiqun)

# 连接行政边界和城市群数据
df_merged = pd.merge(df_xingzheng, df_chengshiqun, how='left', left_on='name_shi', right_on='城市')

# 读取指标数据
df_zhibiao = pd.read_excel(zhibiao)

# 连接合并后的数据和指标数据
final_df = pd.merge(df_merged, df_zhibiao, how='left', left_on='gjdl_xdm', right_on='district_code')
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

# 定义指标列表和对应的颜色映射
indicators = ['高科技', '高效能', '高质量', '综合指数']
color_maps = {
    '高科技': 'Blues',
    '高效能': 'Oranges',
    '高质量': 'Greens',
    '综合指数': 'Purples'
}

# 计算每个城市群的平均指数
city_group_averages = merged.groupby('城市群')[indicators].mean().reset_index()

# 计算城市群每个县的平均指数
county_averages = merged.groupby(['城市群', 'name_xian_x'])[indicators].mean().reset_index()

# 输出城市群每个县的平均指数到Excel
output_dir = r'J:\BaiduNetdiskDownload\05马克\8亿+企业查数据\03指标计算\区域输出图表'
os.makedirs(output_dir, exist_ok=True)
average_output_path = os.path.join(output_dir, '城市群县级平均指数.xlsx')
county_averages.to_excel(average_output_path, index=False)
print(f"已将城市群每个县的平均指数导出到: {average_output_path}")

# 计算并输出每个城市群的总体平均指数
city_group_averages = merged.groupby('城市群')[indicators].mean().reset_index()
city_group_average_path = os.path.join(output_dir, '城市群总体平均指数.xlsx')
city_group_averages.to_excel(city_group_average_path, index=False)
print(f"已将城市群总体平均指数导出到: {city_group_average_path}")

# 为每个城市群创建单独的GeoDataFrame并导出为Shapefile
for city_group in merged['城市群'].unique():
    if pd.isna(city_group):
        continue

    city_data = merged[merged['城市群'] == city_group].copy()
    
    # 导出为GeoPackage
    try:
        gpkg_path = os.path.join(output_dir, f'{city_group}.gpkg')
        city_data.to_file(gpkg_path, driver='GPKG', encoding='utf-8')
        print(f"已将 {city_group} 的数据导出为GeoPackage: {gpkg_path}")
    except Exception as e:
        print(f"导出 {city_group} 的数据时出错: {str(e)}")
        print("请检查数据兼容性和文件权限。")

    for indicator in indicators:
        # 将缺失值填充为0
        city_data[indicator] = city_data[indicator].fillna(0)

        # 创建图表
        fig, ax = plt.subplots(figsize=(12, 8))
        city_data.plot(column=indicator, cmap=color_maps[indicator], linewidth=0.5, edgecolor='0.8', ax=ax)

        # 注释掉设置标题和轴标签的代码
        # plt.title(f'{city_group} - {indicator}分布图', fontproperties=font)
        plt.axis('off')

        # 获取前10名数据
        top_10 = city_data.sort_values(by=indicator, ascending=False).head(10)

        # 在图上标注前10名的名称和省份名称，避免标注互相重叠
        from adjustText import adjust_text
        
        texts = []
        for idx, row in top_10.iterrows():
            # 创建注释对象（县级名称）
            text = plt.text(row.geometry.centroid.x, row.geometry.centroid.y, row['name_xian_x'],
                            fontproperties=font, fontsize=8, ha='center', va='center')
            texts.append(text)

        # 标出县、市、省边界
        xians = city_data['name_xian_x'].unique()
        for xian in xians:
            xian_data = city_data[city_data['name_xian_x'] == xian]
            xian_boundary = xian_data.dissolve(by='name_xian_x')
            xian_boundary.boundary.plot(ax=ax, color='lightgray', linewidth=0.5)

        # cities = city_data['name_shi_x'].unique()
        # for city in cities:
        #     city_data_subset = city_data[city_data['name_shi_x'] == city]
        #     city_boundary = city_data_subset.dissolve(by='name_shi_x')
        #     try:
        #         city_boundary.boundary.plot(ax=ax, color='darkgray', linewidth=0.8)
        #     except ValueError as e:
        #         print(f"绘制城市边界时出错: {str(e)}")
        #         print("跳过当前城市边界的绘制")

        provinces = city_data['name_sheng_x'].unique()
        for province in provinces:
            province_data = city_data[city_data['name_sheng_x'] == province]
            province_boundary = province_data.dissolve(by='name_sheng_x')
            province_boundary.boundary.plot(ax=ax, color='gray', linewidth=1)
            # 计算省份边界的中心点
            bounds = province_boundary.total_bounds
            center_x = (bounds[0] + bounds[2]) / 2
            center_y = (bounds[1] + bounds[3]) / 2
            text = plt.text(center_x, center_y, province,
                            fontproperties=font, fontsize=10, ha='center', va='center', fontweight='bold')
            texts.append(text)
        # 使用adjust_text函数调整文本位置，避免重叠
        adjust_text(texts, arrowprops=dict(arrowstyle='-', color='lightgray', lw=0.5))

        # 保存图表
        plt.savefig(os.path.join(output_dir, f'{city_group}_{indicator}_分布图.png'), dpi=300, bbox_inches='tight')
        plt.close()

        # 导出前10名到Excel的不同sheet
        sheet_name = f'{indicator}_前10名'
        excel_path = os.path.join(output_dir, f'{city_group}_数据.xlsx')
        try:
            # 检查文件是否存在
            if os.path.exists(excel_path):
                # 如果文件存在，使用openpyxl引擎以追加模式打开
                with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a') as writer:
                    top_10[['name_shi_x', 'name_xian_x', indicator]].to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                # 如果文件不存在，创建新文件
                with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                    top_10[['name_shi_x', 'name_xian_x', indicator]].to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"已将 {city_group} 的 {indicator} 前10名数据导出到 {excel_path} 的 {sheet_name} 表格中。")
        except Exception as e:
            print(f"导出 {city_group} 的 {indicator} 数据时出错: {str(e)}")

        print(f"已生成 {city_group} 的 {indicator} 分布图和前10名Excel文件。")

print("所有图表、Excel文件和Shapefile已生成完毕。")
