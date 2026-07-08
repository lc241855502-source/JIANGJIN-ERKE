import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font
from calculator import run_calculation, generate_config_template

st.set_page_config(page_title="门店提成绩效计算工具", layout="centered")
st.title("📊 门店提成一键计算系统")
st.caption("自动匹配品牌、多备用金合并LS、对齐Excel原表X/Y/Z公式")
st.divider()

# 文件上传
c1,c2 = st.columns(2)
with c1:
    business_upload = st.file_uploader("① 业务总xlsx文件", type=["xlsx"], help="包含26.05完成情况、销售明细、各类备用金Sheet")
with c2:
    staff_upload = st.file_uploader("② 人员配置表xlsx", type=["xlsx"])

# 模板下载（修复变量名bug）
st.subheader("人员配置模板下载")
temp_df = generate_config_template()
buf_temp = BytesIO()
# 修复：把temp → temp_df
with pd.ExcelWriter(buf_temp, engine="openpyxl") as writer:
    temp_df.to_excel(writer, index=False, sheet_name="配置模板")

st.download_button(
    label="下载人员配置模板",
    data=buf_temp.getvalue(),
    file_name="人员配置模板.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
st.divider()

# 计算按钮
can_run = business_upload is not None and staff_upload is not None
if st.button("🚀 开始计算", type="primary", disabled=not can_run):
    try:
        with st.spinner("正在计算，请等待..."):
            df_res, df_store_sum, df_sales_detail = run_calculation(
                business_upload,
                staff_upload,
                store_sheet="26.05完成情况",
                sales_sheet="26.05销售明细"
            )
        st.success("✅ 计算完成")
        # 统计
        t1,t2,t3 = st.columns(3)
        t1.metric("核算人数", len(df_res))
        t2.metric("绩效合计", f"¥{df_res['绩效小计'].sum():,.2f}")
        t3.metric("提成合计", f"¥{df_res['提成总计'].sum():,.2f}")
        st.dataframe(df_res, hide_index=True, use_container_width=True)

        # 导出带蓝色姓名
        output_buf = BytesIO()
        with pd.ExcelWriter(output_buf, engine="openpyxl") as writer:
            df_res.to_excel(writer, index=False, sheet_name="人员提成结果")
            df_store_sum.to_excel(writer, index=False, sheet_name="门店基数汇总")
            df_sales_detail.to_excel(writer, index=False, sheet_name="销售明细(含Z列)")
            wb = writer.book
            ws = w.sheets["人员提成结果"]
            blue_font = Font(color="#0000FF", bold=True)
            headers = [cell.value for cell in ws[1]]
            remark_idx = None
            for i, h in enumerate(headers):
                if h == "备注":
                    remark_idx = i+1
                    break
            if remark_idx:
                for row_num in range(2, len(df_res)+2):
                    remark_cell = ws.cell(row=row_num, column=remark_idx)
                    if remark_cell.value and str(remark_cell).strip():
                        ws.cell(row=row_num, column=1).font = blue_font
        st.download_button(
            "📥 下载完整结果Excel",
            data=output_buf.getvalue(),
            file_name="门店提成计算结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception as err:
        st.error(f"计算失败：{str(err)}")
        st.info("检查Sheet名称、列名、Excel文件完整性")

# 使用说明
with st.expander("使用说明"):
    st.markdown("""
1. 业务文件自动识别全部带「备用金」的Sheet，合并统计LS费用
2. 销售表产品类型自动根据J列品牌填充，无需手动填写
3. Y、Z列逻辑1:1复刻原校验Excel所有特殊订单规则
4. 人员备注含病假/全月缺勤自动清零绩效提成，导出Excel姓名标蓝
5. 配置表新增「个人提成调整」列，用来填张子夜、罗梦婷这类人工固定差额
6. 导出文件包含三张工作表：人员薪酬、门店汇总、原始销售明细（可核对Z列）
""")
