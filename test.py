import wikipedia
import pandas as pd

excel_file = "./data/input/E.xlsx" # Excel 文件名

# 读取 Excel 文件
df = pd.read_excel(excel_file)
query_column = '装备名称'  # 装备名称列

if query_column not in df.columns:
    raise ValueError(f"Excel 文件中未找到列 '{query_column}'，请检查列名是否正确！")

# 提取装备名称
equipment_names = df[query_column]


for index, equipment in enumerate(equipment_names, start=1):
    if not isinstance(equipment, str) or equipment.strip() == "":
        print(f"跳过空装备名称，行号：{index}")
        continue
    print(f"正在处理装备：{equipment}")
    wikipedia.set_lang("zh")

    try:
        page = wikipedia.page(equipment)

        for image in page.images:
            # if image.endswith(".jpg"):
                print(image)
                # break
        
    except wikipedia.exceptions.PageError:
        print(f"未找到 '{equipment}' 的页面，跳过。")