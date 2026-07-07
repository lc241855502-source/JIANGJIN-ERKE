import pandas as pd
import numpy as np

# ===================== 常量配置 =====================
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


# ===================== 单订单计算 =====================
def calc_final_base(row):
    """计算单条销售明细的最终提成基数"""
    product_type = str(row.get("产品类型", ""))
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
    """计算单人绩效小计"""
    store_level = store_info.get("门店类别", "C+类")
    completion_rate = store_info.get("绩效完成率", 0.0)
    position = staff_row["职位"]
    perf_setting = float(staff_row["绩效设定"])
    behavior_score = float(staff_row["行为绩效"]) / 100

    # 业绩绩效：≥70%生效，封顶100%
    if completion_rate >= 0.7:
        performance = min(completion_rate, 1.0)
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
    # 1. 读取数据
    df_store = pd.read_excel(business_file, sheet_name=store_sheet)
    df_sales = pd.read_excel(business_file, sheet_name=sales_sheet)
    df_staff = pd.read_excel(config_file)

    # 列名校验
    required_store_cols = ["部门", "门店类别", "任务额", "计提绩效"]
    required_sales_cols = ["门店代码", "产品类型", "成交折扣", "实际计提绩效", "零售总价", "成交金额", "备注"]
    required_staff_cols = ["姓名", "部门名称", "部门代码", "职位", "是否有提成资格", "绩效设定", "行为绩效", "库存机奖励", "异常补差", "转介绍"]

    for col in required_store_cols:
        if col not in df_store.columns:
            raise ValueError(f"门店完成情况表缺少必填列：{col}")
    for col in required_sales_cols:
        if col not in df_sales.columns:
            raise ValueError(f"销售明细表缺少必填列：{col}")
    for col in required_staff_cols:
        if col not in df_staff.columns:
            raise ValueError(f"人员配置表缺少必填列：{col}")

    # 2. 计算销售明细最终提成基数
    df_sales["最终提成基数"] = df_sales.apply(calc_final_base, axis=1)

    # 3. 按门店+产品类型汇总
    store_summary = df_sales.groupby(["门店代码", "产品类型"])["最终提成基数"].sum().unstack(fill_value=0)
    for col in ["助听器", "呼吸机"]:
        if col not in store_summary.columns:
            store_summary[col] = 0.0

    # 4. 计算门店绩效完成率
    df_store["绩效完成率"] = df_store["计提绩效"] / df_store["任务额"].replace(0, np.nan)
    df_store["绩效完成率"] = df_store["绩效完成率"].fillna(0)
    store_info_dict = df_store.set_index("部门")[["门店类别", "绩效完成率"]].to_dict("index")

    # 5. 逐人计算
    result_rows = []
    for _, staff in df_staff.iterrows():
        store_code = staff["部门代码"]
        position = staff["职位"]
        has_commission = str(staff["是否有提成资格"]).strip() == "是"

        # 门店汇总数据
        if store_code in store_summary.index:
            ha_base = store_summary.loc[store_code, "助听器"]
            ven_base = store_summary.loc[store_code, "呼吸机"]
        else:
            ha_base = 0.0
            ven_base = 0.0

        store_info = store_info_dict.get(store_code, {"门店类别": "C+类", "绩效完成率": 0.0})

        # 提成计算
        if has_commission and position in HEARING_AID_RATE:
            ha_commission = ha_base * HEARING_AID_RATE[position]
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
        "姓名", "工号", "部门名称", "部门代码", "职位",
        "是否有提成资格", "绩效设定", "行为绩效",
        "库存机奖励", "异常补差", "转介绍"
    ])
    # 示例行
    template.loc[0] = ["示例-店长", "C0001", "北京六壹广场店", "B-BJ04", "店长", "是", 1600, 100, 0, 0, 0]
    template.loc[1] = ["示例-店员", "C0002", "北京六壹广场店", "B-BJ04", "店员", "是", 1000, 100, 0, 0, 0]
    return template