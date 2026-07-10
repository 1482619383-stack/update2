import pandas as pd
from phone import Phone

# ==========================================
# 1. 【设置路径并读取数据】
file_path = r'C:\Users\EDY\Desktop\update1\订单_完整处理流程_20260526_102200.xlsx'
df = pd.read_excel(file_path, sheet_name='高中待上传')
df_chu = pd.read_excel(file_path, sheet_name='初中待上传')

# 一键清理所有表头中的隐藏空格，防止 KeyError
df.columns = df.columns.str.strip()
df_chu.columns = df_chu.columns.str.strip()

# ==========================================
# 2. 【清洗步骤】：白名单提取法（规范化“新商品规格”字段）
valid_gao_specs = ['9月升高一（现初三）', '9月升高二（现高一）', '9月升高三（现高二）']
valid_chu_specs = [
    '9月升初一(现六年级)', '9月升初一（现六年级）', 
    '9月升初二（现初一）', '9月升初三（现初二）', 
    '在读初三', '9月升高一（现初三）'
]

def extract_pure_spec(text, valid_specs):
    text_str = str(text)
    for spec in valid_specs:
        if spec in text_str:
            return spec
    return text_str

df['新商品规格'] = df['新商品规格'].apply(lambda x: extract_pure_spec(x, valid_gao_specs))
df_chu['新商品规格'] = df_chu['新商品规格'].apply(lambda x: extract_pure_spec(x, valid_chu_specs))


# ==========================================
# 3. 【映射更新“年级”列】 (针对高中与初中表)
mapping_dict = {
    '9月升高一（现初三）': '高三',
    '9月升高二（现高一）': '高一',
    '9月升高三（现高二）': '高二'
}
df['年级'] = df['新商品规格'].map(mapping_dict).fillna(df['年级'])

mapping_dict_chu = {
    '9月升初一(现六年级)': '六年级',
    '9月升初一（现六年级）': '六年级',
    '9月升初二（现初一）': '初一',
    '9月升初三（现初二）': '初二',
    '在读初三': '初三',
    '9月升高一（现初三）': '高三'
}
df_chu['年级'] = df_chu['新商品规格'].map(mapping_dict_chu).fillna(df_chu['年级'])


# ==========================================
# 4. 【处理“一级渠道”空白值】 (初中、高中同步兜底)
df['一级渠道'] = df['一级渠道'].fillna('个人渠道')
df_chu['一级渠道'] = df_chu['一级渠道'].fillna('个人渠道')


# ==========================================
# 5. 【二级渠道和账号ID同时为空时，修改为“店铺”】 (初中、高中同步兜底)
condition_both_empty_gao = (df['二级渠道'].isna() | (df['二级渠道'].astype(str).str.strip() == '')) & \
                           (df['账号ID'].isna() | (df['账号ID'].astype(str).str.strip() == ''))
df.loc[condition_both_empty_gao, '二级渠道'] = '店铺'
df.loc[condition_both_empty_gao, '账号ID'] = '店铺'

condition_both_empty_chu = (df_chu['二级渠道'].isna() | (df_chu['二级渠道'].astype(str).str.strip() == '')) & \
                           (df_chu['账号ID'].isna() | (df_chu['账号ID'].astype(str).str.strip() == ''))
df_chu.loc[condition_both_empty_chu, '二级渠道'] = '店铺'
df_chu.loc[condition_both_empty_chu, '账号ID'] = '店铺'


# ==========================================
# 5.5 【跨表转移“高三”数据】 (剪切初中高三数据 -> 修改名称 -> 粘贴给高中)
condition_transfer = df_chu['新商品规格'] == '9月升高一（现初三）'
df_to_transfer = df_chu[condition_transfer].copy()

if not df_to_transfer.empty:
    df_to_transfer['推广计划名称'] = '高中-商务-达播-5元'
    df_chu = df_chu[~condition_transfer] # 从初中表剔除
    df = pd.concat([df, df_to_transfer], ignore_index=True) # 拼接到高中表


# ==========================================
# 6. 【处理筛选 (仅在高中原表修改，不再提取新表)】
df['一级渠道'] = df['一级渠道'].astype(str).str.strip()
condition_jidi = df['一级渠道'].str.contains('基地', na=False)

# 直接在原表中将符合基地条件的数据修改推广计划名称
df.loc[condition_jidi, '推广计划名称'] = '高中-商务-达播基地-5元'


# ==========================================
# 6.3 【针对高中非基地部分，单独处理“高三”】
condition_not_jidi = ~condition_jidi
condition_gaosan = condition_not_jidi & (df['年级'] == '高三')
df.loc[condition_gaosan, '推广计划名称'] = '高中-商务-达播-5元'


# ==========================================
# 6.5 【查询高中剩余数据的手机归属地】
print("\n>>> 正在进行本地高速手机号归属地查询，请稍候...")
phone_checker = Phone()

def get_province(phone_num):
    if pd.isna(phone_num):
        return '未知'
    phone_str = str(phone_num).strip().split('.')[0]
    if len(phone_str) == 11 and phone_str.isdigit():
        try:
            info = phone_checker.find(phone_str)
            if info and 'province' in info:
                return info['province']
        except:
            pass 
    return '未知/格式错误'

condition_to_localize = condition_not_jidi & (df['年级'] != '高三')

if condition_to_localize.sum() > 0:
    df.loc[condition_to_localize, '收货手机号所属地区'] = df.loc[condition_to_localize, '收货手机号'].apply(get_province)

    # 6.6 【根据省份修改高中剩余数据的推广计划名称】
    condition_shaanxi = condition_to_localize & (df['收货手机号所属地区'].str.contains('陕西', na=False))
    df.loc[condition_shaanxi, '推广计划名称'] = '高中-商务-达播西安-5元'

    condition_henan = condition_to_localize & (df['收货手机号所属地区'].str.contains('河南', na=False))
    df.loc[condition_henan, '推广计划名称'] = '高中-商务-达播郑州-5元'

    condition_hebei = condition_to_localize & (df['收货手机号所属地区'].str.contains('河北', na=False))
    df.loc[condition_hebei, '推广计划名称'] = '高中-商务-达播石家庄-5元'


# ==========================================
# 6.6.5 【处理初中待上传 sheet 的基地推广计划】
df_chu['一级渠道'] = df_chu['一级渠道'].astype(str).str.strip()

# 1. 特例：一级渠道是“福哥-郑州基地”
condition_chu_zhengzhou = df_chu['一级渠道'] == '福哥-郑州基地'
df_chu.loc[condition_chu_zhengzhou, '推广计划名称'] = '初中-商务-达播-5元-基地-郑州'

# 2. 排除特例，筛选剩余包含“基地”的数据
condition_chu_jidi = (~condition_chu_zhengzhou) & (df_chu['一级渠道'].str.contains('基地', na=False))
df_chu.loc[condition_chu_jidi, '推广计划名称'] = '初中-商务-达播-5元-基地'




# ==========================================
# 7. 【数据保存 (仅写回原 Excel 文件)】
with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='高中待上传', index=False)
    df_chu.to_excel(writer, sheet_name='初中待上传', index=False)
    
print("\n🎉 所有自动化流程已执行完毕！")
print("1. [优化] 移除了独立新表的导出，所有修改已100%直接并入原 Excel 表中。")
print(f"2. [高中] 成功在原表中将 {condition_jidi.sum()} 行包含‘基地’的数据修改为对应计划名称。")
print("- 表格已全部更新并安全保存，随时可以用于系统上传！")
