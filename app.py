# ============================================================
# 亚马逊FBA补货计算器（AI库存管家版）
# 功能：补货计算 + 成本估算 + 断货损失 + 亚马逊对比 + 超储预警
# ============================================================

import os
import uuid
import io
import csv
import re
import streamlit as st
import requests
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="FBA补货计算器", layout="centered")
st.title("📦 亚马逊FBA补货计算器")
st.write("上传库存报表，3秒算出补什么、补多少、要花多少钱、不补会亏多少。")

# ==================== 密钥读取 ====================
try:
    API_KEY = st.secrets["DASHSCOPE_API_KEY"]
except Exception:
    API_KEY = os.getenv("DASHSCOPE_API_KEY", "")

if not API_KEY:
    st.error("未找到 API Key，请检查配置")
    st.stop()

# ==================== 密码 ====================
CUSTOMER_PASSWORDS = {"fba2026": "试用客户"}

if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())

password = st.text_input("请输入访问密码", type="password")
if password not in CUSTOMER_PASSWORDS:
    st.warning("密码错误，请联系管理员")
    st.stop()
else:
    st.success(f"欢迎，{CUSTOMER_PASSWORDS[password]}")

# ==================== 调用AI ====================
def call_llm(system_prompt, user_message):
    url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": "qwen-plus",
        "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": user_message}]
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]

# ==================== 万能文本拆行拆列 ====================
def smart_split(text):
    if text is None:
        return []
    text = str(text).strip()
    if not text:
        return []
    lines = [ln.strip() for ln in re.split(r"[\r\n]+", text) if ln.strip()]
    rows = []
    for ln in lines:
        if "\t" in ln:
            parts = ln.split("\t")
        elif "," in ln:
            parts = ln.split(",")
        else:
            parts = re.split(r"\s{2,}", ln)
            if len(parts) == 1:
                parts = ln.split(" ")
        parts = [p.strip() for p in parts if p.strip() != ""]
        if parts:
            rows.append(parts)
    return rows

# ==================== 模糊匹配列名 ====================
def find_column(headers, keywords):
    for i, h in enumerate(headers):
        h_lower = str(h).lower().replace(" ", "").replace("_", "")
        for kw in keywords:
            if kw.lower() in h_lower:
                return i
    return None

# ==================== 安全转数字 ====================
def to_number(val):
    if val is None:
        return None
    s = str(val).strip().replace(",", "").replace("$", "").replace("%", "").replace("￥", "")
    if s in ["", "-", "n/a", "na", "null", "none", "N/A"]:
        return None
    try:
        return float(s)
    except Exception:
        return None

# ==================== 参数设置 ====================
with st.expander("⚙️ 参数设置（默认即可，不需要改）"):
    safety_days = st.number_input("安全库存天数", min_value=0, max_value=90, value=14)
    default_lead = st.number_input("默认采购提前期（表格没填时使用）", min_value=1, max_value=90, value=7)
    default_ship = st.number_input("默认物流时效（表格没填时使用）", min_value=1, max_value=120, value=20)

    st.write("---")
    st.write("**成本估算参数（表格没填时使用）**")
    default_cost = st.number_input("默认单件采购成本（元）", min_value=0.0, max_value=10000.0, value=0.0, step=1.0)
    default_price = st.number_input("默认单件售价（元）", min_value=0.0, max_value=10000.0, value=0.0, step=1.0)

# ==================== 上传文件 ====================
uploaded_file = st.file_uploader("上传库存报表（Excel 或 CSV，格式乱也没事）", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    if "last_file" not in st.session_state or st.session_state.last_file != uploaded_file.name:
        st.session_state.session_id = str(uuid.uuid4())
        st.session_state.last_file = uploaded_file.name

    temp_path = f"temp_{st.session_state.session_id}_{uploaded_file.name}"
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    with st.status("📋 正在处理...", expanded=True) as status:
        st.write("1. 读取表格...")
        raw_rows = []

        if uploaded_file.name.lower().endswith(".csv"):
            with open(temp_path, "r", encoding="utf-8-sig", errors="ignore") as f:
                reader = csv.reader(f)
                for row in reader:
                    if row and any(str(c).strip() for c in row):
                        raw_rows.append(tuple(str(c) for c in row))
        else:
            wb = load_workbook(temp_path, data_only=True)
            ws = wb.active
            raw_rows = [tuple(str(c) if c is not None else "" for c in r) for r in ws.iter_rows(values_only=True)]

        # ========== 万能识别 ==========
        st.write("2. 解析表格结构...")

        def count_non_empty(row):
            return len([c for c in row if c is not None and str(c).strip() != ""])

        max_cols = max([count_non_empty(r) for r in raw_rows]) if raw_rows else 0

        if max_cols <= 1:
            st.write("🔧 检测到数据被挤在一列里，正在自动拆分...")
            full_text = "\n".join([str(r[0]) for r in raw_rows if r and r[0] is not None])
            raw_rows = smart_split(full_text)

        if not raw_rows or len(raw_rows) < 2:
            st.error(f"❌ 表格数据不足，只读到了 {len(raw_rows)} 行。")
            st.stop()

        clean_rows = []
        for r in raw_rows:
            if r and any(str(c).strip() for c in r):
                clean_rows.append(tuple(str(c).strip() for c in r))
        raw_rows = clean_rows

        headers = list(raw_rows[0])
        data_rows = raw_rows[1:]

        # ========== 匹配列名 ==========
        st.write("3. 识别列名...")
        col_sku = find_column(headers, ["sku", "产品", "编号", "货号", "asin", "fnsku", "msku"])
        col_stock = find_column(headers, ["库存", "stock", "available", "可售", "在库", "qty", "quantity"])
        col_sales = find_column(headers, ["销量", "sales", "日均", "velocity", "sold"])
        col_lead = find_column(headers, ["提前", "采购", "lead", "交期"])
        col_ship = find_column(headers, ["时效", "物流", "shipping", "运输", "transit"])
        col_transit = find_column(headers, ["在途", "inbound", "transit", "在运", "pending"])
        col_cost = find_column(headers, ["采购价", "成本", "cost", "单价", "进价"])
        col_price = find_column(headers, ["售价", "价格", "price", "销售价"])
        col_amazon = find_column(headers, ["亚马逊建议", "amazon", "推荐补货", "官方建议"])

        missing = []
        if col_sku is None: missing.append("SKU/产品编号")
        if col_stock is None: missing.append("库存")
        if col_sales is None: missing.append("日均销量")

        if missing:
            st.error(f"❌ 没找到这些列：{'、'.join(missing)}")
            st.info(f"我读到的表头是：{headers}")
            st.stop()

        st.write(f"✅ 识别成功：SKU列='{headers[col_sku]}'，库存列='{headers[col_stock]}'，销量列='{headers[col_sales]}'")

        # ========== 逐行计算 ==========
        st.write("4. 计算补货量...")
        results = []
        skipped = 0

        for row in data_rows:
            if not row:
                continue

            def get_cell(idx):
                if idx is None or idx >= len(row):
                    return None
                return row[idx]

            sku = get_cell(col_sku)
            if sku is None or str(sku).strip() == "":
                skipped += 1
                continue

            stock = to_number(get_cell(col_stock))
            daily_sales = to_number(get_cell(col_sales))
            lead_time = to_number(get_cell(col_lead))
            shipping_time = to_number(get_cell(col_ship))
            in_transit = to_number(get_cell(col_transit))
            cost = to_number(get_cell(col_cost))
            price = to_number(get_cell(col_price))
            amazon_qty = to_number(get_cell(col_amazon))

            if stock is None and daily_sales is None:
                skipped += 1
                continue

            stock = stock if stock is not None else 0
            daily_sales = daily_sales if daily_sales is not None else 0
            lead_time = lead_time if lead_time is not None else default_lead
            shipping_time = shipping_time if shipping_time is not None else default_ship
            in_transit = in_transit if in_transit is not None else 0

            if daily_sales == 0:
                skipped += 1
                continue

            total_days = lead_time + shipping_time + safety_days
            needed = daily_sales * total_days
            restock = max(0, int(needed - stock - in_transit))

            days_left = stock / daily_sales
            if days_left < lead_time + 7:
                risk = "🔴 高（即将断货）"
            elif days_left < lead_time + 14:
                risk = "🟡 中（需要关注）"
            else:
                risk = "🟢 低（库存健康）"

            days_after = (stock + restock) / daily_sales if daily_sales > 0 else 0
            if days_after > 120:
                overstock = "⚠️ 补后超储（可能被收长期仓储费）"
            else:
                overstock = "正常"

            # ========== 新增：补货成本估算 ==========
            unit_cost = cost if cost is not None else default_cost
            restock_cost = restock * unit_cost if unit_cost > 0 else 0
            cost_display = f"¥{restock_cost:,.0f}" if restock_cost > 0 else "-"

            # ========== 新增：断货损失估算 ==========
            unit_price = price if price is not None else default_price
            daily_loss = daily_sales * unit_price if unit_price > 0 else 0
            loss_display = f"¥{daily_loss:,.0f}/天" if daily_loss > 0 else "-"

            # ========== 新增：亚马逊建议对比 ==========
            if amazon_qty is not None and amazon_qty > 0:
                diff = restock - amazon_qty
                if abs(diff) < 5:
                    compare = "✅ 与亚马逊一致"
                elif diff > 0:
                    compare = f"⬆️ 比亚马逊多补 {int(diff)} 个"
                else:
                    compare = f"⬇️ 比亚马逊少补 {int(abs(diff))} 个"
            else:
                compare = "-"

            results.append({
                "SKU": str(sku),
                "当前库存": int(stock),
                "日均销量": round(daily_sales, 1),
                "建议补货量": restock,
                "本次补货成本": cost_display,
                "断货一天损失": loss_display,
                "断货风险": risk,
                "距离断货天数": int(days_left),
                "补后可用天数": int(days_after),
                "亚马逊建议": int(amazon_qty) if amazon_qty is not None else "-",
                "对比": compare,
                "计算过程": f"{daily_sales}×({lead_time}+{shipping_time}+{safety_days}) - {stock} - {in_transit} = {restock}",
                "备注": overstock
            })

        # 清理临时文件
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception:
            pass

        if not results:
            st.error(f"❌ 没有可计算的数据。跳过了 {skipped} 行无效数据。")
            st.stop()

        status.update(label=f"✅ 计算完成！共 {len(results)} 个SKU", state="complete", expanded=False)

    # ==================== 展示结果 ====================
    st.success(f"共计算 {len(results)} 个SKU，跳过 {skipped} 行无效数据")

    risk_order = {"🔴 高（即将断货）": 0, "🟡 中（需要关注）": 1, "🟢 低（库存健康）": 2}
    results_sorted = sorted(results, key=lambda x: risk_order.get(x["断货风险"], 3))

    st.dataframe(results_sorted, use_container_width=True)

    # ========== 汇总卡 ==========
    st.write("---")
    st.subheader("📊 本次补货汇总")

    total_restock = sum([r["建议补货量"] for r in results_sorted if isinstance(r["建议补货量"], (int, float))])
    high_risk_count = len([r for r in results_sorted if "🔴" in r["断货风险"]])

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("需要补货的SKU", f"{len([r for r in results_sorted if r['建议补货量'] > 0])} 个")
    with col2:
        st.metric("总建议补货量", f"{total_restock:,} 个")
    with col3:
        st.metric("高危断货SKU", f"{high_risk_count} 个")

    # ==================== 一键下载 ====================
    st.write("---")
    st.subheader("📥 下载补货报告")
    st.write("把这份表格发给采购，直接照着下单就行。")

    output = io.BytesIO()
    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "补货建议"
    headers_out = list(results_sorted[0].keys())
    ws_out.append(headers_out)
    for r in results_sorted:
        ws_out.append([r[h] for h in headers_out])
    for i, h in enumerate(headers_out, 1):
        ws_out.column_dimensions[get_column_letter(i)].width = max(len(str(h)) * 2, 12)
    wb_out.save(output)
    output.seek(0)

    st.download_button(
        label="📥 下载补货建议表（Excel）",
        data=output,
        file_name=f"补货建议_{st.session_state.session_id[:8]}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ==================== AI解读 ====================
    with st.expander("💬 让AI帮你解读这份补货建议"):
        if st.button("生成解读"):
            with st.spinner("AI正在分析数据..."):
                summary_prompt = f"""你是一名资深的亚马逊运营专家。
下面是刚刚计算好的补货结果（已按风险排序）：
{results_sorted[:10]}

请用5-8句话，给卖家解读：
1. 哪些SKU最紧急，需要立刻下单？
2. 整体库存健康度如何？
3. 本次补货大概要花多少钱？
4. 如果不补货，一天大概损失多少钱？
5. 有没有超储风险，需要注意仓储费？
6. 有没有什么坑需要注意？

用老板能听懂的大白话，不要重复表格数据。"""
                try:
                    st.markdown(call_llm(summary_prompt, "请解读"))
                except Exception as e:
                    st.error(f"AI解读失败：{e}")

else:
    st.info("👆 请上传库存报表，支持 Excel 和 CSV，格式乱也没关系。")