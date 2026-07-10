import streamlit as st
import pandas as pd
from phone import Phone
import io

# 网页标题和说明
st.set_page_config(page_title="自动化订单清洗系统", page_icon="📊")
st.title("📊 自动化订单清洗与映射系统")
st.write("请上传从系统导出的订单 Excel 表格（包含初中待上传、高中待上传 sheet），系统将自动完成清洗、映射和归属地查询。")

# 1. 【网页上传组件】
uploaded_file = st.file_uploader("点击或拖拽上传 Excel 文件 (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    # 加一个运行按钮
    if st.button("🚀 开始执行全自动清洗 SOP"):
        # 显示加载动画
        with st.spinner("正在拼命处理中，请稍候... (手机号本地查询可能需要几秒钟)"):
            try:
                # ==========================================
                # 读取上传的文件（直接在内存中读取，不需要存到桌面）
                df = pd.read_excel(uploaded_file, sheet_name='高中待上传')
                df_chu = pd.read_excel(uploaded_file, sheet_name='初中待上传')
                df_xiao = pd.read_excel(uploaded_file, sheet_name='小学待上传')

                df.columns = df.columns.str.strip()
                df_chu.columns = df_chu.columns.str.strip()
                df_xiao.columns = df_xiao.columns.str.strip()

                # ==========================================
                # 2. 【清洗步骤】：只保留 9月升xx 或 在读xx，删除多余描述
                valid_gao_specs = ['9月升高一', '9月升高二', '9月升高三']
                valid_chu_specs = ['9月升初一', '9月升初二', '9月升初三', '在读初三', '9月升高一',]
                valid_xiao_specs = ['二阶段', '三阶段', '四阶段', '五阶段', '六阶段',
                                    '二年级', '三年级', '四年级', '五年级', '六年级']

                def extract_pure_spec(text, valid_specs):
                    text_str = str(text)
                    for spec in valid_specs:
                        if spec in text_str:
                            return spec
                    return text_str

                df['新商品规格'] = df['新商品规格'].apply(lambda x: extract_pure_spec(x, valid_gao_specs))
                df_chu['新商品规格'] = df_chu['新商品规格'].apply(lambda x: extract_pure_spec(x, valid_chu_specs))

                # ==========================================
                # 3. 【映射更新“年级”列】
                mapping_dict = {
                    '9月升高一': '高一',
                    '9月升高二': '高二',
                    '9月升高三': '高三',
                    '9月升高一': '高一', 
                }
                df['年级'] = df['新商品规格'].map(mapping_dict).fillna(df['年级'])

                mapping_dict_chu = {
                    '9月升初一': '初一',
                    '9月升初二': '初二',
                    '9月升初三': '初三',
                    '在读初三': '高一',
                    '9月升高一': '高一'
                }
                df_chu['年级'] = df_chu['新商品规格'].map(mapping_dict_chu).fillna(df_chu['年级'])

                mapping_dict_xiao = {
                    '二阶段': '二年级',
                    '三阶段': '三年级',
                    '四阶段': '四年级',
                    '五阶段': '五年级',
                    '六阶段': '六年级',
                    '二年级': '二年级',
                    '三年级': '三年级',
                    '四年级': '四年级',
                    '五年级': '五年级',
                    '六年级': '六年级'
                }
                df_xiao['年级'] = df_xiao['新商品规格'].map(mapping_dict_xiao).fillna(df_xiao['年级'])

                # ==========================================
                # 4. 【处理“一级渠道”空白值】
                df['一级渠道'] = df['一级渠道'].fillna('个人渠道')
                df_chu['一级渠道'] = df_chu['一级渠道'].fillna('个人渠道')
                df_xiao['一级渠道'] = df_xiao['一级渠道'].fillna('个人渠道')

                # ==========================================
                # 5. 【二级渠道和账号ID为空 -> 店铺】
                condition_both_empty_gao = (df['二级渠道'].isna() | (df['二级渠道'].astype(str).str.strip() == '')) & \
                                           (df['账号ID'].isna() | (df['账号ID'].astype(str).str.strip() == ''))
                df.loc[condition_both_empty_gao, '二级渠道'] = '店铺'
                df.loc[condition_both_empty_gao, '账号ID'] = '店铺'

                condition_both_empty_chu = (df_chu['二级渠道'].isna() | (df_chu['二级渠道'].astype(str).str.strip() == '')) & \
                                            (df_chu['账号ID'].isna() | (df_chu['账号ID'].astype(str).str.strip() == ''))
                df_chu.loc[condition_both_empty_chu, '二级渠道'] = '店铺'
                df_chu.loc[condition_both_empty_chu, '账号ID'] = '店铺'

                condition_both_empty_xiao = (df_xiao['二级渠道'].isna() | (df_xiao['二级渠道'].astype(str).str.strip() == '')) & \
                                             (df_xiao['账号ID'].isna() | (df_xiao['账号ID'].astype(str).str.strip() == ''))
                df_xiao.loc[condition_both_empty_xiao, '二级渠道'] = '店铺'
                df_xiao.loc[condition_both_empty_xiao, '账号ID'] = '店铺'

                # ==========================================
                # 5.5 【跨表转移“九月升高一”数据】
                def transfer_jiuyue_gaoyi(df, df_chu):
                    condition_transfer = df_chu['新商品规格'].astype(str).str.strip() == '9月升高一'
                    df_to_transfer = df_chu.loc[condition_transfer].copy()

                    if not df_to_transfer.empty:
                        df_to_transfer['推广计划名称'] = '高中-商务-达播-5元'
                        df_chu = df_chu.loc[~condition_transfer].copy()
                        df = pd.concat([df, df_to_transfer], ignore_index=True)
                    return df, df_chu

                df, df_chu = transfer_jiuyue_gaoyi(df, df_chu)
                # ==========================================
                # 6. 【高中基地处理】
                df['一级渠道'] = df['一级渠道'].astype(str).str.strip()
                condition_jidi = df['一级渠道'].str.contains('基地', na=False)
                df.loc[condition_jidi, '推广计划名称'] = '高中-商务-达播基地-5元'

                # ==========================================
                # 6.3 【高中非基地-高三处理】
                condition_not_jidi = ~condition_jidi
                condition_gaosan = condition_not_jidi & df['年级'].isin(['高二', '高三'])
                df.loc[condition_gaosan, '推广计划名称'] = '高中-商务-达播-5元'

                # ==========================================
                # 6.5 【高中查省份】
                phone_checker = Phone()
                def get_province(phone_num):
                    if pd.isna(phone_num): return '未知'
                    phone_str = str(phone_num).strip().split('.')[0]
                    if len(phone_str) == 11 and phone_str.isdigit():
                        try:
                            info = phone_checker.find(phone_str)
                            if info and 'province' in info: return info['province']
                        except: pass 
                    return '未知/格式错误'

                condition_to_localize = condition_not_jidi
                if condition_to_localize.sum() > 0:
                    df.loc[condition_to_localize, '收货手机号所属地区'] = df.loc[condition_to_localize, '收货手机号'].apply(get_province)
                    
                    # 6.6 【根据省份改名称】
                    condition_shaanxi = condition_to_localize & (df['收货手机号所属地区'].str.contains('陕西', na=False))
                    df.loc[condition_shaanxi, '推广计划名称'] = '高中-商务-达播西安-5元'
                    condition_henan = condition_to_localize & (df['收货手机号所属地区'].str.contains('河南', na=False))
                    df.loc[condition_henan, '推广计划名称'] = '高中-商务-达播郑州-5元'
                    condition_hebei = condition_to_localize & (df['收货手机号所属地区'].str.contains('河北', na=False))
                    df.loc[condition_hebei, '推广计划名称'] = '高中-商务-达播石家庄-5元'

                    # 6.7 【将剩余推广计划名称为 '高中-商务-达播-5元' 且年级为 '高一' 的改名为 '新高一-商务-达播-5元'】
                    condition_remaining_newgaoyi = (df['推广计划名称'] == '高中-商务-达播-5元') & (df['年级'] == '高一')
                    df.loc[condition_remaining_newgaoyi, '推广计划名称'] = '新高一-商务-达播-5元'

                # ==========================================
                # 6.6.5 【初中基地处理 - 按年级细化】
                df_chu['一级渠道'] = df_chu['一级渠道'].astype(str).str.strip()
                condition_chu_zhengzhou = df_chu['一级渠道'] == '福哥-郑州基地'
                # 福哥-郑州基地：高一 -> 后缀-新高一；初一初二初三 -> 保持基地名称
                df_chu.loc[condition_chu_zhengzhou, '推广计划名称'] = '初中-商务-达播-5元-基地-郑州'

                condition_chu_jidi = (df_chu['一级渠道'].str.contains('基地', na=False)) & (df_chu['一级渠道'] != '福哥-郑州基地')
                # 其他包含'基地'的渠道：高一 -> 后缀-新高一；初一初二初三 -> 普通基地名称
                df_chu.loc[condition_chu_jidi, '推广计划名称'] = '初中-商务-达播-5元-基地'

                # 其他剩余记录：如果年级为高一，改为常规5元-新高一；初一初二初三 -> 5元2
                condition_other = ~(condition_chu_zhengzhou | df_chu['一级渠道'].str.contains('基地', na=False))
                df_chu.loc[condition_other & (df_chu['年级'].isin(['初一','初二', '初三'])), '推广计划名称'] = '初中-商务-达播-5元2'
                
                # ==========================================
                # 6.6.6 【小学处理】
                df_xiao['一级渠道'] = df_xiao['一级渠道'].astype(str).str.strip()
                df_xiao['推广计划名称'] = '小学-商务-达播-5元'

               # ==========================================
                # 6.7 【删除过桥列】
                columns_to_drop = ['新商品规格', '新学部', '收货手机号所属地区']
                df_gaoyi = df[df['年级'] == '高一'].copy()
                df = df[df['年级'] != '高一'].copy()
                df = df.drop(columns=columns_to_drop, errors='ignore')
                df_gaoyi = df_gaoyi.drop(columns=columns_to_drop, errors='ignore')
                df_chu = df_chu.drop(columns=columns_to_drop, errors='ignore')
                df_xiao = df_xiao.drop(columns=columns_to_drop, errors='ignore')

                # ==========================================
                # 7. 【拆分高中高一数据并写入输出文件】
                # df_gaoyi 已经在上面拆分完成

                # ==========================================
                # 7.5 【在年级后添加教材版本和学科列，并重命名辅导姓名列】
                for df_obj in [df, df_gaoyi, df_chu, df_xiao]:
                    # 重命名输出列
                    if '线下分配的辅导姓名' in df_obj.columns:
                        df_obj.rename(columns={'线下分配的辅导姓名': '销售邮箱前缀'}, inplace=True)

                    # 获取年级列的位置，并在其后插入教材版本和学科列
                    if '年级' in df_obj.columns:
                        year_idx = df_obj.columns.get_loc('年级')
                        if '教材版本' not in df_obj.columns:
                            df_obj.insert(year_idx + 1, '教材版本', '通用版')
                        if '学科' not in df_obj.columns:
                            insert_idx = df_obj.columns.get_loc('年级') + 2
                            df_obj.insert(insert_idx, '学科', '综合课')

                def format_order_id_as_text(df_obj, writer_obj, sheet_name):
                    if '订单编号' not in df_obj.columns:
                        return
                    df_obj['订单编号'] = df_obj['订单编号'].astype(str).replace('nan', '')
                    ws = writer_obj.sheets[sheet_name]
                    col_idx = df_obj.columns.get_loc('订单编号') + 1
                    for row in range(2, len(df_obj) + 2):
                        cell = ws.cell(row=row, column=col_idx)
                        if cell.value is not None:
                            cell.value = str(cell.value)
                        cell.number_format = '@'

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='高中待上传', index=False)
                    df_gaoyi.to_excel(writer, sheet_name='高中新高一待上传', index=False)
                    df_chu.to_excel(writer, sheet_name='初中待上传', index=False)
                    df_xiao.to_excel(writer, sheet_name='小学待上传', index=False)

                    format_order_id_as_text(df, writer, '高中待上传')
                    format_order_id_as_text(df_gaoyi, writer, '高中新高一待上传')
                    format_order_id_as_text(df_chu, writer, '初中待上传')
                    format_order_id_as_text(df_xiao, writer, '小学待上传')
                
                processed_data = output.getvalue()

                st.success("🎉 处理完毕！请点击下方按钮获取最新表格。")
                
                # 网页下载组件
                st.download_button(
                    label="📥 下载处理完成的 Excel 文件",
                    data=processed_data,
                    file_name="清洗完成_待上传订单.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"处理过程中出现错误，请检查上传的表格格式是否正确。错误信息：{e}")
