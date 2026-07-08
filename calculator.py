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

PERFORMANCE_WEIGHT = {
    "A类": {"店长": (0.6, 0.4), "店员": (0.4, 0.6)},
    "B类": {"店长": (0.6, 0.4), "店员": (0.4, 0.6)},
    "C类": {"店长": (0.6, 0.4), "店员": (0.4, 0.6)},
    "C+类": {"店长": (0.6, 0.4), "店员": (0.4, 0.6)},
    "D类": {"店长": (0.7, 0.3), "店员": (0.3, 0.7)}
}

NO_COMMISSION_KEYWORDS = ["无提成", "提成0%", "#N/A", "提成100元/台", "提成88元/台", "提成100元每台", "提成88元每台"]
STAFF_NO_COMMISSION_KEYWORDS = ["无提成", "病假", "产假", "离职", "缺勤整月", "无绩效"]


# ===================== 工具函数：智能读取带标题行的Sheet =====================
def smart_read_excel(file, sheet_name, required_keywords):
    """
    自动跳过前面的标题行，找到真实表头
    required_keywords: 必须包含的列名关键词，用于定位表头行
    """
    # 先读取全部内容，不设表头
    df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    
    # 逐行查找包含关键词的行，作为表头
    header_row = 0
    for i in range(min(10, len(df_raw))):  # 最多检查前10行
        row_text = " ".join([str(x) for x in df_raw.iloc[i].values if pd.notna(x)])
        match_count = sum([1 for kw in required_keywords if kw in row_text])
        if match_count >= len(required_keywords) * 0.6:  # 匹配60%以上关键词即判定为表头行
            header_row = i
            break
    
    # 从表头行开始读取
    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
    # 去除列名前后空格
    df.columns = [str(col).strip() for col in df.columns]
    # 去除全空列
    df = df.dropna(axis=1, how="all")
    return df


# ===================== 单订单计算 =====================
def calc_final_base(row):
    """计算单条销售明细的最终提成基数"""
    product_type = str(row.get("产品类型", "")).strip()
    discount = float(row["成交折扣"]) if pd.notna(row.get("成交折扣")) else 0
    actual_perf = float(row["实际计提绩效"]) if pd.notna(row.get("实际计提绩效")) else 0
    remark = str(row["备注"]) if pd.notna(row.get("备注")) else ""
    retail_total = float(row["零售总价"]) if pd.notna(row.get("零售总价")) else 0
    deal_amount = float(row["成交金额"]) if pd.notna(row.get("成交金额")) else 0

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


# ===================== 单人绩效计算 =====================
def calc_person_performance(staff_row, store_info):
    """计算单人绩效小计，精度与原表对齐（业绩绩效保留1位小数）"""
    store_level = store_info.get("门店类别", "C+类")
    completion_rate = store_info.get("绩效完成率", 0.0)
    position = staff_row["职位"]
    perf_setting = float(staff_row["绩效设定"])
    behavior_score = float(staff_row["行为绩效"]) / 100

    # 业绩绩效：≥70%生效，封顶100%，保留3位小数（对应百分比1位小数）
    if completion_rate >= 0.7:
        performance = round(min(completion_rate, 1.0), 3)
    else:
        performance = 0.0

    # 获取权重
    level_config = PERFORMANCE_WEIGHT.get(store_level, PERFORMANCE_WEIGHT["C+类"])
    perf_weight, behavior_weight = level_config.get(position, (0.4, 0.6))

    total = (performance * perf_weight + behavior_score * behavior_weight) * perf_setting
    return round(total, 2)


# ===================== 主计算流程 =====================
def run_calculation(business_file, config_file,
                    store_sheet="26.05完成情况",
                    sales_sheet="26.05销售明细"):
    """
    执行完整计算
    返回：(结果DataFrame, 门店汇总DataFrame, 销售明细带计算结果DataFrame)
    """
    # 1. 智能读取数据（自动跳过标题行）
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

    # 列名兼容映射
    store_col_map = {
        "部门": "部门",
        "部门代码": "部门",
        "门店代码": "部门",
        "门店类别": "门店类别",
        "任务额": "任务额",
        "计提绩效": "计提绩效"
    }
    sales_col_map = {
        "门店代码": "门店代码",
        "品牌": "品牌",
        "产品类型": "产品类型",
        "成交折扣": "成交折扣",
        "实际计提绩效": "实际计提绩效",
        "零售总价": "零售总价",
        "成交金额": "成交金额",
        "备注": "备注",
        "LS": "LS费用",
        "LS费用": "LS费用"
    }

    # 重命名列，统一字段名
    df_store = df_store.rename(columns={k: v for k, v in store_col_map.items() if k in df_store.columns})
    df_sales = df_sales.rename(columns={k: v for k, v in sales_col_map.items() if k in df_sales.columns})

    # 列名校验
    required_store_cols = ["部门", "门店类别", "任务额", "计提绩效"]
    required_sales_cols = ["门店代码", "品牌", "成交折扣", "实际计提绩效", "零售总价", "成交金额", "备注"]
    required_staff_cols = ["姓名", "部门名称", "部门代码", "职位", "是否有提成资格", "绩效设定", "行为绩效", "库存机奖励", "异常补差", "转介绍"]

    missing_store = [col for col in required_store_cols if col not in df_store.columns]
    missing_sales = [col for col in required_sales_cols if col not in df_sales.columns]
    missing_staff = [col for col in required_staff_cols if col not in df_staff.columns]

    if missing_store:
        raise ValueError(f"门店完成情况表缺少必填列：{', '.join(missing_store)}。当前识别到的列：{', '.join(df_store.columns)}")
    if missing_sales:
        raise ValueError(f"销售明细表缺少必填列：{', '.join(missing_sales)}。当前识别到的列：{', '.join(df_sales.columns)}")
    if missing_staff:
        raise ValueError(f"人员配置表缺少必填列：{', '.join(missing_staff)}。当前识别到的列：{', '.join(df_staff.columns)}")

    # 清洗数据：去除空行
    df_store = df_store.dropna(subset=["部门"])
    df_sales = df_sales.dropna(subset=["门店代码"])

    # 数值列转数字
    for col in ["任务额", "计提绩效"]:
        df_store[col] = pd.to_numeric(df_store[col], errors="coerce").fillna(0)
    for col in ["成交折扣", "实际计提绩效", "零售总价", "成交金额", "LS费用"]:
        if col in df_sales.columns:
            df_sales[col] = pd.to_numeric(df_sales[col], errors="coerce").fillna(0)
        else:
            df_sales[col] = 0

    # ========== 新增：自动匹配产品类型（对应导出计数_品牌规则） ==========
    def match_product_type(brand_name):
        brand_name = str(brand_name).strip()
        for brand, ptype in BRAND_PRODUCT_MAP.items():
            if brand in brand_name:
                return ptype
        return "助听器"  # 默认按助听器处理

    # 优先用已有产品类型，为空则用品牌匹配
    df_sales["产品类型"] = df_sales.apply(
        lambda x: str(x["产品类型"]).strip() if str(x["产品类型"]).strip() in ["助听器", "呼吸机"] else match_product_type(x["品牌"]),
        axis=1
    )

    # 2. 计算销售明细最终提成基数
    df_sales["最终提成基数"] = df_sales.apply(calc_final_base, axis=1)

    # 3. 按门店+产品类型汇总提成基数
    store_summary = df_sales.groupby(["门店代码", "产品类型"])["最终提成基数"].sum().unstack(fill_value=0)
    for col in ["助听器", "呼吸机"]:
        if col not in store_summary.columns:
            store_summary[col] = 0.0

    # ========== 新增：按门店汇总LS费用 ==========
    ls_summary = df_sales.groupby("门店代码")["LS费用"].sum().to_dict()

    # 4. 计算门店绩效完成率
    df_store["绩效完成率"] = df_store["计提绩效"] / df_store["任务额"].replace(0, np.nan)
    df_store["绩效完成率"] = df_store["绩效完成率"].fillna(0)
    store_info_dict = df_store.set_index("部门")[["门店类别", "绩效完成率"]].to_dict("index")

    # 5. 逐人计算
    result_rows = []
    for _, staff in df_staff.iterrows():
        store_code = str(staff["部门代码"]).strip()
        position = str(staff["职位"]).strip()
        has_commission = str(staff["是否有提成资格"]).strip() == "是"
        staff_remark = str(staff.get("备注", "")).strip()

        # ========== 新增：人员特殊备注判定（病假/无提成等直接全置0） ==========
        is_exempt = any(kw in staff_remark for kw in STAFF_NO_COMMISSION_KEYWORDS)
        if is_exempt:
            result_rows.append({
                "姓名": staff["姓名"],
                "部门名称": staff["部门名称"],
                "绩效小计": 0.0,
                "提成总计": 0.0,
                "调差总计": 0.0
            })
            continue

        # 门店汇总数据
        if store_code in store_summary.index:
            ha_base = store_summary.loc[store_code, "助听器"]
            ven_base = store_summary.loc[store_code, "呼吸机"]
        else:
            ha_base = 0.0
            ven_base = 0.0

        # ========== 新增：扣除门店LS费用 ==========
        store_ls = ls_summary.get(store_code, 0.0)
        ha_base_after_ls = max(0.0, ha_base - store_ls)

        store_info = store_info_dict.get(store_code, {"门店类别": "C+类", "绩效完成率": 0.0})

        # 提成计算
        if has_commission and position in HEARING_AID_RATE:
            ha_commission = ha_base_after_ls * HEARING_AID_RATE[position]
            ven_commission = ven_base * VENTILATOR_RATE[position]
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

        # 绩效小计
        perf_total = calc_person_performance(staff, store_info)

        result_rows.append({
            "姓名": staff["姓名"],
            "部门名称": staff["部门名称"],
            "绩效小计": perf_total,
            "提成总计": total_commission,
            "调差总计": adjust_total
        })

    df_result = pd.DataFrame(result_rows)
    return df_result, store_summary.reset_index(), df_sales


# ===================== 生成空白配置模板 =====================
def generate_config_template():
    """生成人员基础配置表模板"""
    template = pd.DataFrame(columns=[
        "姓名", "工号", "部门名称", "部门代码", "职位", "备注",
        "是否有提成资格", "绩效设定", "行为绩效",
        "库存机奖励", "异常补差", "转介绍"
    ])
    # 示例行
    template.loc[0] = ["示例-店长", "C0001", "北京六壹广场店", "B-BJ04", "店长", "", "是", 1600, 100, 0, 0, 0]
    template.loc[1] = ["示例-店员", "C0002", "北京六壹广场店", "B-BJ04", "店员", "", "是", 1000, 100, 0, 0, 0]
    template.loc[2] = ["示例-病假", "C0003", "北京劲松店", "B-BJ19", "店长", "5月全月病假，无提成", "是", 1400, 0, 0, 0, 0]
    return template