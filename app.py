import streamlit as st
import pandas as pd
import io

st.title("📊 账单自动透视工具")
st.write("上传账单 Excel，自动生成透视表供下载。")

# 1. 上传文件组件
uploaded_file = st.file_uploader("请上传 '账单.xlsx'", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 读取上传的文件
        sheet_pos_name = '分账明细-正向-团购'
        sheet_neg_name = '分账明细-退款-团购'
        
        st.info("正在读取数据...")
        
        # 注意：这里直接从内存读取
        df_pos = pd.read_excel(uploaded_file, sheet_name=sheet_pos_name)
        df_neg = pd.read_excel(uploaded_file, sheet_name=sheet_neg_name)

        # 清洗列名
        df_pos.columns = [c.strip() for c in df_pos.columns]
        df_neg.columns = [c.strip() for c in df_neg.columns]

        group_cols = ['核销门店', '商品类型']
        sum_col = '商家应得'

        # 计算
        pivot_pos = df_pos.groupby(group_cols)[sum_col].sum()
        pivot_neg = df_neg.groupby(group_cols)[sum_col].sum()
        
        # 合并
        total_series = pivot_pos.add(pivot_neg, fill_value=0)
        
        # 透视
        final_pivot_view = total_series.unstack(level='商品类型', fill_value=0)
        final_pivot_view['总计'] = final_pivot_view.sum(axis=1)

        st.success("计算完成！预览如下：")
        st.dataframe(final_pivot_view) # 在网页上展示预览

        # 2. 导出下载组件
        # 将结果写入内存中的 Excel 文件
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_pivot_view.to_excel(writer, sheet_name='门店商品透视汇总')
        
        # 提供下载按钮
        st.download_button(
            label="📥 点击下载处理后的 Excel",
            data=output.getvalue(),
            file_name="处理结果_透视表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"发生错误：{e}")
        st.warning("请检查上传的 Excel 是否包含指定的 Sheet 名称。")