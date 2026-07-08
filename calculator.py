import pandas as pd
import numpy as np

# 品牌映射 对应导出计数_品牌
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
HEARING_AID_RATE = {"店长":0.05, "大店长":0.05, "店长（双店）":0.05, "店员":0.02}
VENTILATOR_RATE = {"店长":0.20, "大店长":0.20, "店长（双店）":0.20, "店员":0.08}

# 绩效权重
PERFORMANCE_WEIGHT = {
    "A类":{"店长":(0.6,0.4),"店员":(0.4,0.6)},
    "B类":{"店长":(0.6,0.4),"店员":(0.4,0.6)},
    "C类":{"店长":(0.6,0.4),"店员":(0.4,0.6)},
    "C+类":{"店长":(0.6,0.4),"店员":(0.4,0.6)},
    "D类":{"店长":(0.7,0.3),"店员":(0.3,0.7)}
}

# 关键词定义
NO_COMMISSION_ORDER = ["无提成","提成0%","#N/A","提成100元/台","提成88元/台","提成100元每台","提成88元每台"]
STAFF_CLEAR_KEY = ["无提成","病假","产假","离职","缺勤整月","无绩效","全月缺勤","整月病假"]

def smart_read_excel(file, sheet_name, required_keywords):
    df_raw = pd.read_excel(file, sheet_name, header=None)
    header_row = 0
    for i in range(min(10, len(df_raw))):
        row_text = " ".join([str(x) for x in df_raw.iloc[i].values if pd.notna(x)])
        hit = sum([1 for k in required_keywords if k in row_text])
        if hit >= len(required_keywords)*0.6:
            header_row = i
            break
    df = pd.read_excel(file, sheet_name, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(axis=1, how="all")
    return df

def find_col(df, priority_keywords):
    for kw in priority_keywords:
        for col in df.columns:
            if kw in col:
                return col
    return None

def get_product_type(brand):
    brand = str(brand).strip()
    for b, t in BRAND_PRODUCT_MAP.items():
        if b in brand:
            return t
    return "助听器"

def calc_z_row(row):
    # 全部变量强制赋值兜底，无空值
    ptype = str(row.get("产品类型", "")).strip()
    brand = str(row.get("品牌", "")).strip()
    discount = float(row.get("成交折扣", 0.0)) if pd.notna(row.get("成交折扣")) else 0.0
    u = float(row.get("实际计提绩效", 0.0)) if pd.notna(row.get("实际计提绩效")) else 0.0
    retail_total = float(row.get("零售总价", 0.0)) if pd.notna(row.get("零售总价")) else 0.0
    deal_amt = float(row.get("成交金额", 0.0)) if pd.notna(row.get("成交金额")) else 0.0
    ls_ratio = float(row.get("LS分摊比例", 0.0)) if pd.notna(row.get("LS分摊比例")) else 0.0
    remark = str(row.get("备注", "")).strip()

    y = 0.0
    z = 0.0

    if ptype == "呼吸机":
        if discount < 0.7:
            y = 0.0
        else:
            y = round(deal_amt - retail_total * 0.7, 2)
    else:
        if any(k in remark for k in NO_COMMISSION_ORDER):
            y = 0.0
        elif "斯达克全品类1.2倍提成" in remark:
            y = round(u * 1.2, 2)
        elif "即墨医院订单" in remark:
            if deal_amt <= 0:
                ratio = 0.0
            else:
                ratio = (retail_total * 0.95) / deal_amt
            if ratio >= 0.5:
                y = round(retail_total, 2)
            elif ratio >= 0.4:
                y = round(retail_total * 0.5, 2)
            else:
                y = 0.0
        elif "武汉中南订单" in remark or "武汉人民订单" in remark:
            if deal_amt <= 0:
                ratio = 0.0
            else:
                ratio = retail_total / deal_amt
            if ratio >= 0.5:
                y = round(retail_total, 2)
            elif ratio >= 0.4:
                y = round(retail_total * 0.5, 2)
            else:
                y = 0.0
        elif "转门店" in remark:
            y = round(u * (1 - ls_ratio), 2)
        else:
            y = round(u, 2)
    # Z值计算
    if 0.4 < discount < 0.5:
        z = round(y / 2, 2)
    else:
        z = round(y, 2)
    return z, y

def calc_staff_perf(staff_data, store_data):
    completion = store_data.get("绩效完成率", 0.0)
    level = store_data.get("门店类别", "C+类")
    pos = str(staff_data["职位"]).strip()
    base_perf = float(staff_data["绩效设定"]) if pd.notna(staff_data["绩效设定"]) else 0.0
    beh_raw = float(staff_data["行为绩效"]) if pd.notna(staff_data["行为绩效"]) else 0.0

    if beh_raw > 2:
        beh_score = beh_raw / 100
    else:
        beh_score = beh_raw

    if completion >= 0.7:
        perf_rate = round(min(completion, 1.0), 3)
    else:
        perf_rate = 0.0

    wgt_perf, wgt_beh = PERFORMANCE_WEIGHT.get(level, {"店长":(0.6,0.4),"店员":(0.4,0.6)}).get(pos, (0.4, 0.6))
    x = round(base_perf * wgt_perf * perf_rate, 2)
    y_perf = round(base_perf * wgt_beh * beh_score, 2)
    z_total = round(x + y_perf, 2)
    return z_total

def run_calculation(business_file, config_file, store_sheet="26.05完成情况", sales_sheet="26.05销售明细"):
    # 读取三大基础表
    df_store = smart_read_excel(business_file, store_sheet, ["区域","店名","任务额","计提绩效"])
    df_sales = smart_read_excel(business_file, sales_sheet, ["门店代码","品牌","成交金额","实际计提绩效"])
    df_staff = pd.read_excel(config_file)
    df_staff.columns = [str(c).strip() for c in df_staff.columns]

    # 多备用金Sheet 汇总LS
    ls_dict = {}
    try:
        xls_all = pd.ExcelFile(business_file)
        all_sheets = xls_all.sheet_names
        ls_sheets = [s for s in all_sheets if "备用金" in s]
        for sh in ls_sheets:
            df_ls = smart_read_excel(business_file, sh, ["门店代码","LS费用","金额"])
            col_store = find_col(df_ls, ["门店代码","部门代码","门店"])
            col_amt = find_col(df_ls, ["LS费用","金额","费用金额"])
            if col_store and col_amt:
                df_ls[col_amt] = pd.to_numeric(df_ls, errors="coerce").fillna(0.0)
                group = df_ls.groupby(col_store)[col_amt].sum().to_dict()
                for k, v in group.items():
                    ls_dict[str(k)] = ls_dict.get(str(k), 0.0) + float(v)
    except Exception:
        pass

    # 列兼容映射
    store_map = {"部门":"部门","部门代码":"部门","门店代码":"部门","门店类别":"门店类别","任务额":"任务额","计提绩效":"计提"}
    sales_map = {"门店代码":"门店代码","品牌":"品牌","产品类型":"产品类型","成交折扣":"成交折扣","实际计提绩效":"零售总价","成交金额":"成交金额","备注":"备注","LS分摊比例":"LS分摊比例"}
    df_store.rename(columns={k:v for k,v in store_map if k in df_store.columns}, inplace=True)
    df_sales.rename(columns={k:v for k, v in sales_map.items() if k in df_sales.columns}, inplace=True)
    if "产品类型" not in df_sales.columns:
        df_sales["产品类型"] = ""

    # 必填校验
    req_store = ["部门","门店类别","任务额","计提"]
    req_sales = ["门店代码","品牌","成交折扣","实际计提绩效","零售总价","成交金额","备注"]
    req_staff = ["姓名","部门名称","部门代码","职位","是否有提成资格","绩效设定","行为绩效","库存机奖励","异常补差","转介绍","个人提成调整"]

    miss_store = [c for c in req_store if c not in df_store.columns]
    miss_sales = [c for c in req_sales if c not in df_sales.columns]
    miss_staff = [c for c in req_staff if c not in df_staff.columns]

    if miss_store:
        raise ValueError(f"门店完成表缺失必填列：{miss_store}")
    if miss_sales:
        raise ValueError(f"销售明细表缺失必填列：{miss_sales}")
    if miss_staff:
        raise ValueError(f"人员配置表缺失必填列：{miss_staff}")

    # 清洗数值，全部填充0.0
    df_store = df_store.dropna(subset=["部门"])
    for c in ["任务额","计提"]:
        df_store[c] = pd.to_numeric(df_store[c], errors="coerce").fillna(0.0)

    df_sales = df_sales.dropna(subset=["门店代码"])
    for c in ["成交折扣","实际计提绩效","零售总价","成交金额"]:
        df_sales[c] = pd.to_numeric(df_sales[c], errors="coerce").fillna(0.0)

    # 填充产品类型
    df_sales["产品类型"] = df_sales.apply(lambda r: get_product_type(r["品牌"]) if str(r["产品类型"]).strip() not in ["助听器","呼吸机"] else r["产品类型"], axis=1)
    # 逐行计算Y Z
    z_list = []
    y_list = []
    for _, row in df_sales.iterrows():
        z,y = calc_z_row(row)
        z_list.append(z)
        y_list.append(y)
    df_sales["Y计提基数"] = y_list
    df_sales["最终提成基数Z"] = z_list

    # 门店汇总
    store_group = df_sales.groupby(["门店代码","产品类型"])["最终提成基数Z"].sum().unstack(fill_value=0.0)
    for t in ["助听器","呼吸机"]:
        if t not in store_group.columns:
            store_group[t] = 0.0
    store_group = store_group.round(2)

    # 门店完成率字典
    df_store["绩效完成率"] = df_store["计提"] / df_store["任务额"].replace(0, np.nan)
    df_store["绩效完成率"] = df_store["绩效完成率"].fillna(0.0)
    store_info = df_store.set_index("部门")[["门店类别","绩效完成率"]].to_dict("index")

    # 处理备注列
    staff_remark_col = find_col(df_staff, ["备注","说明","人员备注"])
    if not staff_remark_col:
        df_staff["备注"] = ""
        staff_remark_col = "备注"

    res_rows = []
    for _, s in df_staff.iterrows():
        # 备注强制转字符串，空值替换为空，解决NaN
        remark_raw = s.get(staff_remark_col, "")
        remark = str(remark_raw).strip() if pd.notna(remark_raw) else ""
        is_clear = any(k in remark for k in STAFF_CLEAR_KEY)
        if is_clear:
            res_rows.append({
                "姓名": str(s.get("姓名", "")),
                "部门名称": str(s.get("部门名称", "")),
                "绩效小计": 0.0,
                "提成总计": 0.0,
                "调差总计": 0.0,
                "备注": remark
            })
            continue

        dept_code = str(s.get("部门代码", "")).strip()
        pos = str(s.get("职位", "")).strip()
        has_comm = str(s.get("是否有提成资格", "")).strip() == "是"

        # 门店基数 不存在直接返回0.0
        ha_total = float(store_group.loc[dept_code, "助听器"]) if dept_code in store_group.index else 0.0
        ven_total = float(store_group.loc[dept_code, "呼吸机"]) if dept_code in store_group.index else 0.0
        dept_ls = ls_dict.get(dept_code, 0.0)
        ha_after_ls = round(max(0.0, ha_total - dept_ls), 2)

        # 提成计算 全部兜底0
        ha_comm = 0.0
        ven_comm = 0.0
        if has_comm and pos in HEARING_AID_RATE:
            ha_comm = round(float(ha_after_ls) * HEARING_AID_RATE[pos], 2)
            ven_comm = round(float(ven_total) * VENTILATOR_RATE[pos], 2)
        base_total = round(ha_comm + ven_comm, 2)

        # 调差全部强制读取兜底0，解决None
        inv = float(s.get("库存机奖励", 0.0)) if pd.notna(s.get("库存机奖励")) else 0.0
        ab = float(s.get("异常补差", 0.0)) if pd.notna(s.get("转介绍")) else 0.0
        tr = float(s.get("转介绍", 0.0)) if pd.notna(s.get("转介绍")) else 0.0
        adj_person = float(s.get("个人提成调整", 0.0)) if pd.notna(s.get("个人提成调整")) else 0.0
        adjust = round(inv + ab + tr + adj_person, 2)
        total_comm = round(base_total + adjust, 2)

        # 绩效
        store_cfg = store_info.get(dept_code, {"门店类别":"C+类","绩效完成率":0.0})
        perf_sum = calc_staff_perf(s, store_cfg)

        res_rows.append({
            "姓名": str(s.get("姓名", "")),
            "部门名称": str(s.get("部门名称", "")),
            "绩效小计": perf_sum,
            "提成总计": total_comm,
            "调差总计": adjust,
            "备注": remark
        })
    df_result = pd.DataFrame(res_rows)
    return df_result, store_group.reset_index(), df_sales

def generate_config_template():
    cols = [
        "姓名","工号","部门名称","部门代码","职位","备注","是否有提成资格",
        "绩效设定","行为绩效","库存机奖励","异常补差","转介绍","个人提成调整"
    ]
    temp = pd.DataFrame(columns=cols)
    temp.loc[0] = ["示例店长","C001","北京六壹广场店","B-BJ04","店长","","是",1600,100,0,0,0,0]
    temp.loc[1] = ["示例店员","C002","北京六壹广场店","B-BJ04","店员","","是",100,100,0,0,0,0]
    temp.loc[2] = ["病假店长","C003","北京劲松店","B-BJ04","店长","5月全月病假，无提成","是",1400,0,0,0,0,0]
    return temp
