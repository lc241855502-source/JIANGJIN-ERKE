import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font
from calculator import run_calculation, generate_config_template

st.set_page_config(page_title="门店提成一键计算工具", page_icon="📊", layout="centered")

# ========== 页面标题 ==========
st.title("📊 耳科门店提成一键计算工具")
st.caption("上传业务数据 + 人员配置表，自动计算绩效与提成，支持结果导出")

st.divider()

# ========== 第一步：上传文件 ==========
st.subheader("第一步：上传两个输入文件")

col1, col2 = st.columns(2)
with col1:
    business_file = st.file_uploader(
        "① 业务数据文件",
        type=["xlsx"],
        help="包含完成情况、销售明细、备用金三个Sheet，备用金Sheet名带“备用金”即可自动识别"
    )

with col2:
    config_file = st.file_uploader(
        "② 人员基础配置表",
        type=["xlsx"],
        help="按模板填写的人员信息、绩效、调差数据"
    )

# 模板下载
st.markdown("还没有配置表？")
template_df = generate_config_template()
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    template_df.to_excel(writer, index=False, sheet_name="人员基础配置")
st.download_button(
    label="📥 下载人员配置表模板",
    data=buffer.getvalue(),
    file_name="人员基础配置表-模板.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.divider()

# ========== 第二步：计算按钮 ==========
can_calc = business_file is not None and config_file is not None

if st.button("🚀 开始计算", type="primary", use_container_width=True, disabled=not can_calc):
    try:
        with st.spinner("正在计算，请稍候..."):
            df_result, df_store_summary, df_sales_result = run_calculation(
                business_file,
                config_file,
                store_sheet="26.05完成情况",
                sales_sheet="26.05销售明细"
            )

        st.success("✅ 计算完成！")

        # ========== 结果概览 ==========
        st.subheader("计算结果概览")
        c1, c2, c3 = st.columns(3)
        c1.metric("核算人数", f"{len(df_result)} 人")
        c2.metric("绩效小计总额", f"¥ {df_result['绩效小计'].sum():,.2f}")
        c3.metric("提成总计总额", f"¥ {df_result['提成总计'].sum():,.2f}")

        # ========== 结果明细表 ==========
        st.subheader("人员薪酬明细表")
        st.dataframe(df_result, use_container_width=True, hide_index=True)

        # ========== 下载结果（修复姓名标蓝逻辑） ==========
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            df_result.to_excel(writer, index=False, sheet_name="人员提成结果")
            df_store_summary.to_excel(writer, index=False, sheet_name="门店汇总")
            df_sales_result.to_excel(writer, index=False, sheet_name="销售明细(含计算)")

            # 动态定位备注列，姓名标蓝加粗
            workbook = writer.book
            worksheet = writer.sheets["人员提成结果"]
            blue_font = Font(color="0000FF", bold=True)

            # 获取表头，找到备注列的位置
            headers = [cell.value for cell in worksheet[1]]
            remark_col_idx = None
            for idx, h in enumerate(headers):
                if h == "备注":
                    remark_col_idx = idx + 1  # openpyxl列从1开始
                    break

            if remark_col_idx:
                for row_idx in range(2, len(df_result) + 2):
                    remark_cell = worksheet.cell(row=row_idx, column=remark_col_idx)
                    if remark_cell.value and str(remark_cell.value).strip():
                        name_cell = worksheet.cell(row=row_idx, column=1)  # 第1列永远是姓名
                        name_cell.font = blue_font

        st.divider()
        st.download_button(
            label="📥 下载完整计算结果Excel",
            data=output_buffer.getvalue(),
            file_name="门店人员提成计算结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"❌ 计算失败：{str(e)}")
        st.info("请检查：1. 文件格式是否正确 2. Sheet名称是否匹配 3. 必填列是否完整")

# ========== 使用说明 ==========
st.divider()
with st.expander("使用说明"):
    st.markdown("""
    1. 业务数据文件需包含：完成情况表、销售明细表、备用金表（Sheet名带“备用金”即可自动识别）
    2. 人员配置表按模板填写，部门代码必须与业务数据中的门店代码一致
    3. 产品类型自动根据品牌匹配，无需手动填写
    4. 行为绩效支持百分制（100）和百分比小数（1.0）两种格式，自动识别
    5. 输出结果对应原表列：姓名(A)、部门名称(C)、绩效小计(Z)、提成总计(AD)、调差总计(AE)
    6. 有特殊备注的人员，导出Excel中姓名会自动标蓝，方便核对
    """)