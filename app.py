# ============================================================
# 亚马逊FBA智能补货系统（作战指挥版）
# 特点：深色商务风 + 复杂表头兼容 + 资金透视 + 72小时急救警报 + AI运营总监解读
# 本地/云端通用，支持 Excel 和 CSV
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

# ==================== 1. 页面基础配置 ====================
st.set_page_config(
    page_title="FBA智能补货系统",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==================== 2. 商务深色风 CSS ====================
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    .stApp {
        background: #0A1120;
        color: #E2E8F0;
    }

    .main-title {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        font-size: 2.5rem;
        font-weight: 700;
        text-align: center;
        color: #F8FAFC;
        letter-spacing: -0.02em;
        margin-bottom: 0.3rem;
    }

    .sub-title {
        text-align: center;
        color: #94A3B8;
        font-size: 1rem;
        margin-bottom: 2.5rem;
    }

    div[data-testid="stVerticalBlock"] > div {
        background: #131E35;
        border-radius: 8px;
        padding: 1.2rem;
        border: 1px solid #1E293B;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.3);
    }

    .stButton > button {
        width: 100%;
        border-radius: 6px;
        background: #10B981;
        color: #FFFFFF;
        font-weight: 600;
        border: none;
        transition: all 0.2s ease;
    }
    .stButton > button:hover {
        background: #059669;
    }

    .stDataFrame {
        border-radius: 6px;
        overflow: hidden;
        border: 1px solid #1E293B;
    }

    div[data-testid="stMetric"] {
        background: #131E35;
        border: 1px solid #1E293B;
        border-radius: 8px;
        padding: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

# ==================== 3. 标题 ====================
st.markdown('<div class="main-title">FBA 智能补货系统</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">上传库存报表 · 精准计算补货量 · 透明展示每一步</div>', unsafe_allow_html=True)


# ==================== 4. 安全读取 API Key ====================
def get_api_key():
    try:
        return st.secrets["DASHSCOPE_API_KEY"]
    except Exception:
        return os.getenv("DASHSCOPE_API_KEY", "")


API_KEY = get_api_key()
if not API_KEY:
    st.error("⚠️ 未检测到 API Key，请检查环境变量或 Streamlit Secrets 配置。")
    st.stop()

# ==================== 5. 密码管理 ====================
CUSTOMER_PASSWORDS = {
    "fba2026": "试用客户",
    # 收到钱后，在这里加一行，比如：
    # "fbaVIP001": "客户张三",
}

if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())

with st.sidebar:
    st.markdown("### 🔐 用户登录")
    password = st.text_input("请输入访问密码", type="password", label_visibility="collapsed")
    if password not in CUSTOMER_PASSWORDS:
        st.warning("密码错误，请联系管理员")
        st.stop()
    else:
        st.success(f"欢迎，{CUSTOMER_PASSWORDS[password]}")


# ==================== 6. 调用 AI ====================
def call_llm(system_prompt, user_message):
    url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": "qwen-plus",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message}
        ]
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]


# ==================== 7. 工具函数 ====================
def smart_split(text):
    if text is None: return []
    text = str(text).strip()
    if not text: return []
    lines = [ln.strip() for ln in re.split(r"[\r\n]+", text) if ln.strip()]
    rows = []
    for ln in lines:
        if "\t" in ln:
            parts = ln.split("\t")
        elif "," in ln:
            parts = ln.split(",")
        else:
            parts = re.split(r"\s{2,}", ln)
            if len(parts) == 1: parts = ln.split(" ")
        parts = [p.strip() for p in parts if p.strip() != ""]
        if parts: rows.append(parts)
    return rows


def find_column(headers, keywords):
    for i, h in enumerate(headers):
        h_lower = str(h).lower().replace(" ", "").replace("_", "")
        for kw in keywords:
            if kw.lower() in h_lower:
                return i
    return None


def to_number(val):
    if val is None: return None
    s = str(val).strip().replace(",", "").replace("$", "").replace("%", "").replace("￥", "")
    if s in ["", "-", "n/a", "na", "null", "none", "N/A"]: return None
    try:
        return float(s)
    except:
        return None


# ==================== 8. 参数设置 ====================
with st.expander("⚙️ 参数设置（默认即可，按需调整）", expanded=False):
    col1, col2, col3 = st.columns(3)
    with col1:
        safety_days = st.number_input("安全库存天数", min_value=0, max_value=90, value=14)
    with col2:
        default_lead = st.number_input("默认采购提前期（表格没填时使用）", min_value=1, max_value=90, value=30)
    with col3:
        default_ship = st.number_input("默认物流时效（表格没填时使用）", min_value=1, max_value=120, value=30)

    exclude_skus_raw = st.text_input(
        "排除指定 SKU（可选，多个用逗号分隔）",
        placeholder="例如：A009, B010"
    )

# ==================== 9. 上传与计算 ====================
uploaded_file = st.file_uploader("📂 上传库存报表（支持 Excel 或 CSV，格式乱也没事）", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    temp_path = f"temp_{st.session_state.session_id}_{uploaded_file.name}"
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    with st.status("🤖 正在分析数据...", expanded=True) as status:
        st.write("读取表格...")
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

        st.write("智能解析结构...")
        max_cols = max(
            [len([c for c in r if c is not None and str(c).strip() != ""]) for r in raw_rows]) if raw_rows else 0
        if max_cols <= 1:
            full_text = "\n".join([str(r[0]) for r in raw_rows if r and r[0] is not None])
            raw_rows = smart_split(full_text)

        if not raw_rows or len(raw_rows) < 2:
            st.error("❌ 表格数据不足。")
            st.stop()

        clean_rows = [tuple(str(c).strip() for c in r) for r in raw_rows if r and any(str(c).strip() for c in r)]
        raw_rows = clean_rows
        headers = list(raw_rows[0])
        data_rows = raw_rows[1:]

        st.write("匹配关键列...")
        # ========== 扩展后的列匹配规则 ==========
        col_sku = find_column(headers, ["sku", "产品", "编号", "货号", "asin", "fnsku", "msku", "品名"])

        # 库存识别细分
        col_stock = find_column(headers, ["库存", "stock", "available", "可售", "在库", "qty", "quantity", "总库存"])
        col_fba_stock = find_column(headers, ["fba可售", "fba库存", "fba available", "亚马逊可售", "fba"])
        col_overseas_stock = find_column(headers, ["海外仓", "overseas", "海外库存"])

        col_sales = find_column(headers,
                                ["销量", "sales", "日均", "velocity", "sold", "7天销量", "30天销量", "平均日销"])
        col_lead = find_column(headers, ["提前", "采购", "lead", "交期", "采购周期", "质检", "上架"])
        col_ship = find_column(headers, ["时效", "物流", "shipping", "运输", "transit", "运输周期"])
        col_transit = find_column(headers, ["在途", "inbound", "transit", "在运", "pending", "已下单"])
        col_cost = find_column(headers, ["采购价", "成本", "cost", "单价", "进价", "拿货价"])
        col_price = find_column(headers, ["售价", "价格", "price", "销售价"])

        if col_sku is None or col_stock is None or col_sales is None:
            st.error("❌ 未找到必备列：SKU、库存、日均销量。请检查表头。")
            st.stop()

        exclude_skus = set()
        if exclude_skus_raw:
            exclude_skus = {s.strip().upper() for s in exclude_skus_raw.split(",") if s.strip()}

        st.write("精准计算补货量...")
        results = []
        for row in data_rows:
            if not row: continue


            def get_cell(idx):
                return row[idx] if idx is not None and idx < len(row) else None


            sku = get_cell(col_sku)
            if sku is None or str(sku).strip() == "": continue
            if str(sku).strip().upper() in exclude_skus: continue

            daily_sales = to_number(get_cell(col_sales)) or 0
            if daily_sales == 0: continue

            # 库存精细拆分逻辑
            total_stock = to_number(get_cell(col_stock)) or 0
            fba_stock = to_number(get_cell(col_fba_stock)) if col_fba_stock is not None else total_stock
            overseas_stock = to_number(get_cell(col_overseas_stock)) or 0

            lead_time = to_number(get_cell(col_lead)) or default_lead
            shipping_time = to_number(get_cell(col_ship)) or default_ship
            in_transit = to_number(get_cell(col_transit)) or 0
            cost = to_number(get_cell(col_cost)) or 0.0
            price = to_number(get_cell(col_price)) or 0.0

            total_days = lead_time + shipping_time + safety_days
            needed = daily_sales * total_days
            restock = max(0, int(needed - total_stock - in_transit))
            days_left = total_stock / daily_sales if daily_sales > 0 else 0
            fba_days_left = fba_stock / daily_sales if daily_sales > 0 else 0

            # 紧急调拨建议
            transfer_action = "-"
            if fba_days_left < 7 and overseas_stock > 0 and needed > fba_stock:
                transfer_amount = min(int(needed - fba_stock), int(overseas_stock))
                if transfer_amount > 0:
                    transfer_action = f"🚨 紧急调拨 {transfer_amount} 个从海外仓到FBA"

            risk = "🔴 高（即将断货）" if days_left < lead_time + 7 else (
                "🟡 中（需要关注）" if days_left < lead_time + 14 else "🟢 低（库存健康）")
            days_after = (total_stock + restock) / daily_sales if daily_sales > 0 else 0
            overstock = "⚠️ 补后超储" if days_after > 120 else "正常"
            transit_flag = f"⚠️ 已有 {int(in_transit)} 个在途" if in_transit > 0 else "-"

            results.append({
                "SKU": str(sku),
                "当前总库存": int(total_stock),
                "FBA可售": int(fba_stock),
                "海外仓备胎": int(overseas_stock),
                "日均销量": round(daily_sales, 1),
                "建议补货量": restock,
                "补货成本(元)": f"¥{restock * cost:,.0f}" if cost > 0 else "-",
                "断货损失(元/天)": f"¥{daily_sales * price:,.0f}" if price > 0 else "-",
                "断货风险": risk,
                "距离断货(天)": int(days_left),
                "调拨建议": transfer_action,
                "补后可用(天)": int(days_after),
                "在途库存提示": transit_flag,
                "计算过程": f"{daily_sales}×({lead_time}+{shipping_time}+{safety_days}) - {total_stock} - {in_transit} = {restock}",
                "备注": overstock
            })

        try:
            if os.path.exists(temp_path): os.remove(temp_path)
        except:
            pass

        if not results:
            st.error("❌ 没有可计算的有效数据。")
            st.stop()
        status.update(label="✅ 计算完成！", state="complete", expanded=False)

    # ==================== 10. 结果展示（增加资金透视与急救线） ====================
    st.success(f"✅ 共计算 {len(results)} 个 SKU，按风险从高到低排序")

    risk_order = {"🔴 高（即将断货）": 0, "🟡 中（需要关注）": 1, "🟢 低（库存健康）": 2}
    results_sorted = sorted(results, key=lambda x: risk_order.get(x["断货风险"], 3))

    high_risk = len([r for r in results_sorted if "🔴" in r["断货风险"]])
    total_restock = sum([r["建议补货量"] for r in results_sorted])

    # 计算总资金与总损失（去除格式化符号）
    total_cost = 0
    total_loss = 0
    for r in results_sorted:
        if r["补货成本(元)"] != "-":
            total_cost += int(r["补货成本(元)"].replace('¥', '').replace(',', ''))
        if r["断货损失(元/天)"] != "-":
            total_loss += int(r["断货损失(元/天)"].replace('¥', '').replace(',', ''))

    colA, colB, colC, colD = st.columns(4)
    with colA:
        st.metric("🚨 高危断货 SKU", f"{high_risk} 个", delta="需立即处理", delta_color="inverse")
    with colB:
        st.metric("📦 总建议补货量", f"{total_restock:,} 个")
    with colC:
        st.metric("💰 补货需要资金", f"¥{total_cost:,}", help="需要拿出的现金流")
    with colD:
        st.metric("🔥 断货每日损失", f"¥{total_loss:,}", delta="高危风险", delta_color="inverse",
                  help="不补货每天流失的钱")

    # 急救警报
    if high_risk > 0:
        top_risk_sku = results_sorted[0]
        st.error(
            f"🚨 **【72小时急救警报】** 最危险的 SKU 是 {top_risk_sku['SKU']}，距离断货仅剩 {top_risk_sku['距离断货(天)']} 天！{top_risk_sku['调拨建议']}。若不处理，预计每日损失 {top_risk_sku['断货损失(元/天)']}。建议本周优先处理前 {high_risk} 个高危项。")

    st.dataframe(results_sorted, use_container_width=True)

    # ==================== 11. 下载与AI解读 ====================
    st.write("---")
    colX, colY = st.columns(2)

    with colX:
        st.subheader("📥 下载补货报告")
        output = io.BytesIO()
        wb_out = Workbook()
        ws_out = wb_out.active
        headers_out = list(results_sorted[0].keys())
        ws_out.append(headers_out)
        for r in results_sorted:
            ws_out.append([r[h] for h in headers_out])
        for i, h in enumerate(headers_out, 1):
            ws_out.column_dimensions[get_column_letter(i)].width = max(len(str(h)) * 2, 12)
        wb_out.save(output)
        output.seek(0)
        st.download_button(
            "📥 下载 Excel 补货建议表",
            data=output,
            file_name="补货建议.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with colY:
        st.subheader("💬 AI 运营总监解读")
        if st.button("生成解读报告", use_container_width=True):
            with st.spinner("AI正在分析数据..."):
                # 构建更高级的提示词
                system_prompt = "你是一个拥有10年实战经验的亚马逊供应链运营总监。你的语气冷静、果断、直击要害。你不喜欢废话，只给老板下达具体的行动指令。"
                user_prompt = f"""
                以下是我的FBA库存补货计算结果（已按风险降序排列）：
                {results_sorted[:8]}

                请按以下结构输出行动简报：
                **1. 致命风险（一句话点名）**
                直接指出本周如果不处理，一定会断货的SKU是哪个，每天损失多少。

                **2. 本周必做动作（分点列出）**
                - 【紧急补货】：需要补哪些SKU，总量多少。
                - 【海外仓调拨】：如果有调拨建议，明确写出调拨指令。
                - 【资金占用建议】：对比补货成本和断货损失，告诉老板这笔钱花得值不值。

                **3. 供应链健康度点评**
                用一句话评价当前整体库存的容错率。
                """
                try:
                    st.info(call_llm(system_prompt, user_prompt))
                except Exception as e:
                    st.error(f"解读失败：{e}")

else:
    st.info("👆 请上传你的库存报表（支持 Excel 和 CSV），系统将自动分析并生成补货建议。")
