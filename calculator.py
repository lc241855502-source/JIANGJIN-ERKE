import pandas as pd
import numpy as np

# ===================== 常量配置 =====================
# 品牌-产品类型映射（对应导出计数_品牌表）
BRAND_PRODUCT_MAP = {
    "奥迪康": "助听器",
    "峰力": "助听器",
    "瑞声达": "助听器",
    "斯达克": "助听器",
    "科林": "助听器",
    "西嘉": "助听器",
    "瑞思迈": "呼吸机"
}

# 提成系数
HEARING_AID_RATE = {
    "店长": 0.05,
    "大店长": 0.05,
    "店长（双店）": 0.05,
    "店员": 0.02
}
VENTILATOR_RATE = {
    "店长": 0.20,
    "大店长": 0.20,
    "店长（双店）": 0.20,
    "店员": 0.08
}

# 绩效权重配置：{门店等级: {职位: (业绩权重, 行为权重)}}
PERFORMANCE_WEIGHT = {
    "A类": {"店长": (0.6, 0.4), "店员": (0.4, 0.6)},
    "B类": {"店长": (0.6, 0.4), "店员": (0.4, 0.6)},
    "C类": {"店长": (0.6, 0.4), "店员": (0.4, 0.6)},
    "C+类": {"店长": (0.6, 0.4), "店员": (0.4, 0.6)},
    "D类": {"店长": (0.7, 0.3), "店员": (0.3, 0.7)}
}

# 订单无提成关键词
NO_COMMISSION_KEYWORDS = ["无提成", "提成0%", "#N/A", "提成100元/台", "提成88元/台", "提成100元每台", "提成88元每台"]
# 人员无提成/清零关键词
STAFF_NO_COMMISSION_KEYWORDS = ["无提成", "病假", "产假", "离职", "缺勤整月", "无绩效", "全月缺勤", "整月病假"]


# ===================== 工具函数：智能读取带标题行的Sheet =====================
def smart_read_excel(file, sheet_name, required_keywords):
    df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    header_row = 0
    for i in range(min(10, len(df_raw))):
        row_text = " ".join([str(x) for x in df_raw.iloc[i].values if pd.notna(x)])
        match_count = sum([1 for kw in required_keywords if kw in row_text])
        if match_count >= len(required_keywords) * 0.6:
            header_row = i
            break
    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
    df.columns = [str(col).strip() for col in df.columns]
    df = df.dropna(axis=1, how="all")
    return df


# ===================== 工具函数：模糊匹配列名 =====================
def find_col(df, keywords):
    """在DataFrame列中查找包含任一关键词的列名，返回第一个匹配的列名"""
    for col in df.columns:
        for kw in keywords:
            if kw in col:
                return col
    return None


# ===================== 单订单最终提成基数计算 =====================
def calc_final_base(row):
    product_type = str(row.get("产品类型", "")).strip()
    brand = str(row.get("品牌", "")).strip()
    discount = float(row["成交折扣"]) if pd.notna(row.get("成交折扣")) else 0
    actual_perf = float(row["实际计提绩效"]) if pd.notna(row.get("实际计提绩效")) else 0
    remark = str(row["备注"]) if pd.notna(row.get("备注")) else ""
    retail_total = float(row["零售总价"]) if pd.notna(row.get("零售总价")) else 0
    deal_amount = float(row["成交金额"]) if pd.notna(row.get("成交金额")) else 0

    # 斯达克助听器1.2倍计提
    if product_type == "助听器" and "斯达克" in brand:
        actual_perf = actual_perf * 1.2

    # 呼吸机规则
    if product_type == "呼吸机":
        if discount < 0.7:
            return 0.0
        return round(deal_amount - retail_total * 0.7, 2)

    # 助听器规则
    elif product_type == "助听器":
        for kw in NO_COMMISSION_KEYWORDS:
            if kw in remark:
                return 0.0
        if discount >= 0.5:
            return round(actual_perf, 2)
        else:
            return round(actual_perf * 0.5, 2)

    return 0.0


# ===================== 单人绩效计算（严格对齐人工计算逻辑） =====================
def calc_person_performance(staff_row, store_info):
    store_level = store_info.get("门店类别", "C+类")
    completion_rate = store_info.get("绩效完成率", 0.0)
    position = str(staff_row["职位"]).strip()
    perf_base = float(staff_row["绩效设定"])  # W列：绩效基数
    behavior_raw = float(staff_row["行为绩效"])

    # 自动识别行为绩效格式：>1 按百分制，≤1 按百分比小数
    if behavior_raw > 1:
        behavior_score = behavior_raw / 100
    else:
        behavior_score = behavior_raw

    # 1. 计算业绩系数：≥70%生效，封顶100%，保留3位小数（百分比1位精度）
    if completion_rate >= 0.7:
        perf_ratio = round(min(completion_rate, 1.0), 3)
    else:
        perf_ratio = 0.0

    # 2. 获取权重
    level_config = PERFORMANCE_WEIGHT.get(store_level, PERFORMANCE_WEIGHT["C+类"])
    perf_weight, behavior_weight = level_config.get(position, (0.4, 0.6))

    # 3. 分步计算X、Y、Z，每步四舍五入，完全对齐Excel
    x_perf = round(perf_base * perf_weight * perf_ratio, 2)  # X列：业绩绩效
    y_behavior = round(perf_base * behavior_weight * behavior_score, 2)  # Y列：行为绩效
    z_total = round(x_perf + y_behavior, 2)  # Z列：绩效小计

    return z_total


# ===================== 主计算流程 =====================
def run_calculation(business_file, config_file,
                    store_sheet="26.05完成情况",
                    sales_sheet="26.05销售明细"):
    # 1. 智能读取核心业务Sheet
    df_store = smart_read_excel(
        business_file, store_sheet,
        required_keywords=["区域", "店名", "任务额", "计提绩效"]
    )
    df_sales = smart_read_excel(
        business_file, sales_sheet,
        required_keywords=["门店代码", "品牌", "成交金额", "实际计提绩效"]
    )
    df_staff = pd.read_excel(config_file)
    df_staff.columns = [str(col).strip() for col in df_staff.columns]

    # 2. 自动识别备用金Sheet，精准读取LS费用
    ls_summary = {}
    try:
        xls = pd.ExcelFile(business_file)
        all_sheets = xls.sheet_names
        ls_sheet_name = None
        for sheet in all_sheets:
            if "备用金" in sheet:
                ls_sheet_name = sheet
                break
        
        if ls_sheet_name:
            df_ls = smart_read_excel(
                business_file, ls_sheet_name,
                required_keywords=["门店", "费用类型", "金额"]
            )
            # 精准匹配：只在费用类型列中查找包含LS的项
            type_col = find_col(df_ls, ["费用类型", "项目", "类型", "费用项目"])
            store_col = find_col(df_ls, ["门店代码", "门店", "部门代码", "部门"])
            amount_col = find_col(df_ls, ["金额", "费用金额", "发生额"])

            if type_col and store_col and amount_col:
                ls_mask = df_ls[type_col].astype(str).str.contains("LS", na=False)
                ls_summary = df_ls[ls_mask].groupby(store_col)[amount_col].sum().to_dict()
    except Exception:
        pass

    # 3. 列名兼容映射
    store_col_map = {
        "部门": "部门", "部门代码": "部门", "门店代码": "部门",
        "门店类别": "门店类别", "任务额": "任务额", "计提绩效": "计提绩效"
    }
    sales_col_map = {
        "门店代码": "门店代码", "品牌": "品牌", "产品类型": "产品类型",
        "成交折扣": "成交折扣", "实际计提绩效": "实际计提绩效",
        "零售总价": "零售总价", "成交金额": "成交金额", "备注": "备注"
    }

    df_store = df_store.rename(columns={k: v for k, v in store_col_map.items() if k in df_store.columns})
    df_sales = df_sales.rename(columns={k: v for k, v in sales_col_map.items() if k in df_sales.columns})

    # 兼容：产品类型列不存在则自动创建空列，后续靠品牌匹配填充
    if "产品类型" not in df_sales.columns:
        df_sales["产品类型"] = ""

    # 兼容：人员备注列模糊匹配
    staff_remark_col = find_col(df_staff, ["备注", "说明", "备注说明", "人员备注"])
    if not staff_remark_col:
        df_staff["备注"] = ""
        staff_remark_col = "备注"

    # 4. 必填列校验
    required_store = ["部门", "门店类别", "任务额", "计提绩效"]
    required_sales = ["门店代码", "品牌", "成交折扣", "实际计提绩效", "零售总价", "成交金额", "备注"]
    required_staff = ["姓名", "部门名称", "部门代码", "职位", "是否有提成资格", "绩效设定", "行为绩效", "库存机奖励", "异常补差", "转介绍"]

    missing_store = [c for c in required_store if c not in df_store.columns]
    missing_sales = [c for c in required_sales if c not in df_sales.columns]
    missing_staff = [c for c in required_staff if c not in df_staff.columns]

    if missing_store:
        raise ValueError(f"门店完成情况表缺少必填列：{', '.join(missing_store)}。当前列：{', '.join(df_store.columns)}")
    if missing_sales:
        raise ValueError(f"销售明细表缺少必填列：{', '.join(missing_sales)}。当前列：{', '.join(df_sales.columns)}")
    if missing_staff:
        raise ValueError(f"人员配置表缺少必填列：{', '.join(missing_staff)}。当前列：{', '.join(df_staff.columns)}")

    # 5. 数据清洗
    df_store = df_store.dropna(subset=["部门"])
    df_sales = df_sales.dropna(subset=["门店代码"])
    for col in ["任务额", "计提绩效"]:
        df_store[col] = pd.to_numeric(df_store[col], errors="coerce").fillna(0)
    for col in ["成交折扣", "实际计提绩效", "零售总价", "成交金额"]:
        df_sales[col] = pd.to_numeric(df_sales[col], errors="coerce").fillna(0)

    # 6. 品牌自动匹配产品类型（X列填充）
    def match_product_type(brand_name):
        brand_name = str(brand_name).strip()
        for brand, ptype in BRAND_PRODUCT_MAP.items():
            if brand in brand_name:
                return ptype
        return "助听器"

    df_sales["产品类型"] = df_sales.apply(
        lambda x: str(x["产品类型"]).strip() if str(x["产品类型"]).strip() in ["助听器", "呼吸机"] else match_product_type(x["品牌"]),
        axis=1
    )

    # 7. 计算单订单最终提成基数
    df_sales["最终提成基数"] = df_sales.apply(calc_final_base, axis=1)

    # 8. 按门店+产品类型汇总基数，保留2位小数
    store_summary = df_sales.groupby(["门店代码", "产品类型"])["最终提成基数"].sum().unstack(fill_value=0)
    for col in ["助听器", "呼吸机"]:
        if col not in store_summary.columns:
            store_summary[col] = 0.0
    store_summary = store_summary.round(2)

    # 9. 计算门店绩效完成率
    df_store["绩效完成率"] = df_store["计提绩效"] / df_store["任务额"].replace(0, np.nan)
    df_store["绩效完成率"] = df_store["绩效完成率"].fillna(0)
    store_info_dict = df_store.set_index("部门")[["门店类别", "绩效完成率"]].to_dict("index")

    # 10. 逐人计算薪酬
    result_rows = []
    for _, staff in df_staff.iterrows():
        store_code = str(staff["部门代码"]).strip()
        position = str(staff["职位"]).strip()
        has_commission = str(staff["是否有提成资格"]).strip() == "是"
        staff_remark = str(staff.get(staff_remark_col, "")).strip()

        # 人员特殊备注：全月缺勤/病假等，全部清零
        is_exempt = any(kw in staff_remark for kw in STAFF_NO_COMMISSION_KEYWORDS)
        if is_exempt:
            result_rows.append({
                "姓名": staff["姓名"],
                "部门名称": staff["部门名称"],
                "绩效小计": 0.0,
                "提成总计": 0.0,
                "调差总计": 0.0,
                "备注": staff_remark
            })
            continue

        # 门店基数
        if store_code in store_summary.index:
            ha_base = float(store_summary.loc[store_code, "助听器"])
            ven_base = float(store_summary.loc[store_code, "呼吸机"])
        else:
            ha_base = 0.0
            ven_base = 0.0

        # 扣除门店LS费用
        store_ls = float(ls_summary.get(store_code, 0.0))
        ha_base_after_ls = round(max(0.0, ha_base - store_ls), 2)

        store_info = store_info_dict.get(store_code, {"门店类别": "C+类", "绩效完成率": 0.0})

        # 提成计算，每步保留2位小数
        if has_commission and position in HEARING_AID_RATE:
            ha_commission = round(ha_base_after_ls * HEARING_AID_RATE[position], 2)
            ven_commission = round(ven_base * VENTILATOR_RATE[position], 2)
        else:
            ha_commission = 0.0
            ven_commission = 0.0

        total_commission = round(ha_commission + ven_commission, 2)

        # 调差总计
        adjust_total = round(
            float(staff.get("库存机奖励", 0)) +
            float(staff.get("异常补差", 0)) +
            float(staff.get("转介绍", 0)), 2
        )

        # 绩效小计（严格对齐Z列逻辑）
        perf_total = calc_person_performance(staff, store_info)

        result_rows.append({
            "姓名": staff["姓名"],
            "部门名称": staff["部门名称"],
            "绩效小计": perf_total,
            "提成总计": total_commission,
            "调差总计": adjust_total,
            "备注": staff_remark
        })

    df_result = pd.DataFrame(result_rows)
    return df_result, store_summary.reset_index(), df_sales


# ===================== 生成配置模板 =====================
def generate_config_template():
    template = pd.DataFrame(columns=[
        "姓名", "工号", "部门名称", "部门代码", "职位", "备注",
        "是否有提成资格", "绩效设定", "行为绩效",
        "库存机奖励", "异常补差", "转介绍"
    ])
    template.loc[0] = ["示例-店长", "C0001", "北京六壹广场店", "B-BJ04", "店长", "", "是", 1600, 100, 0, 0, 0]
    template.loc[1] = ["示例-店员", "C0002", "北京六壹广场店", "B-BJ04", "店员", "", "是", 1000, 100, 0, 0, 0]
    template.loc[2] = ["示例-病假", "C0003", "北京劲松店", "B-BJ19", "店长", "5月全月病假，无提成", "是", 1400, 0, 0, 0, 0]
    return template