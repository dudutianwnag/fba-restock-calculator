# ============================================================
# 亚马逊FBA智能补货系统（经营指挥中心 - 极致对比看板版）
# ============================================================

import os
import uuid
import io
import csv
import re
from datetime import datetime
import streamlit as st
import streamlit.components.v1 as components
import requests
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

# ==================== 1. 页面基础配置 ====================
st.set_page_config(
    page_title="FBA库存经营指挥中心",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 2. 商务深色风全局 UI ====================
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stApp { background: #0A1120; color: #E2E8F0; }

    div[data-testid="stMetric"] {
        background: #131E35; border: 1px solid #1E293B; border-radius: 10px;
        padding: 15px 20px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.5);
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] { font-size: 2rem; font-weight: bold; }
    div[data-testid="stMetric"] div[data-testid="stMetricDelta"] { font-size: 1rem; }

    .stTabs [data-baseweb="tab-list"] { gap: 24px; background-color: #0A1120; }
    .stTabs [data-baseweb="tab"] {
        height: 50px; white-space: pre-wrap; background-color: transparent;
        border-radius: 4px 4px 0px 0px; gap: 1px; padding: 10px; color: #94A3B8;
    }
    .stTabs [aria-selected="true"] { color: #10B981 !important; border-bottom: 2px solid #10B981 !important; }

    .stButton > button {
        width: 100%; border-radius: 6px; background: #10B981; color: #FFFFFF;
        font-weight: 600; border: none; transition: all 0.2s ease;
    }
    .stButton > button:hover { background: #059669; }

    .rank-item {
        padding: 10px; border-bottom: 1px solid #1E293B;
        display: flex; justify-content: space-between;
    }
    .rank-item:last-child { border-bottom: none; }
    .rank-highlight { color: #EF4444; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)


# ==================== 3. 安全读取 API Key ====================
def get_api_key():
    try:
        return st.secrets["DASHSCOPE_API_KEY"]
    except Exception:
        return os.getenv("DASHSCOPE_API_KEY", "")


API_KEY = get_api_key()

# ==================== 4. 状态初始化 ====================
CUSTOMER_PASSWORDS = {
    "fba2026": "试用客户",
}

if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())
if "is_logged_in" not in st.session_state:
    st.session_state.is_logged_in = False
if "history_records" not in st.session_state:
    st.session_state.history_records = []
if "last_results" not in st.session_state:
    st.session_state.last_results = None

with st.sidebar:
    st.markdown("### 🔐 库存指挥中心")
    password = st.text_input("请输入访问密码", type="password", label_visibility="collapsed")
    if password in CUSTOMER_PASSWORDS:
        st.session_state.is_logged_in = True
        st.success(f"欢迎，{CUSTOMER_PASSWORDS[password]}")
    elif password:
        st.warning("密码错误，请联系管理员")


# ==================== 5. 调用大模型 ====================
def call_llm(system_prompt, user_message):
    url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    payload = {"model": "qwen-plus",
               "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": user_message}]}
    resp = requests.post(url, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]


# ==================== 6. 工具函数 ====================
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
        h_lower = str(h).lower().replace(" ", "").replace("_", "").replace("-", "").replace("（", "").replace("）",
                                                                                                             "").replace(
            "(", "").replace(")", "")
        for kw in keywords:
            if kw.lower() in h_lower: return i
    return None


def to_number(val):
    if val is None: return None
    s = str(val).strip().replace(",", "").replace("$", "").replace("%", "").replace("￥", "").replace("¥", "")
    if s in ["", "-", "n/a", "na", "null", "none", "N/A"]: return None
    try:
        return float(s)
    except:
        return None


# ==================== 7. 极简粒子特效首页 ====================
def render_particle_home():
    # (为了节省空间，此处保留你之前的粒子特效代码，可直接复用)
    particle_html = """
    <!DOCTYPE html>
    <html><head><style>body { margin: 0; overflow: hidden; background-color: transparent; display: flex; justify-content: center; align-items: center; height: 100%; font-family: sans-serif; }canvas { display: block; }</style></head>
    <body><canvas id="particleCanvas"></canvas><script>
        const canvas = document.getElementById('particleCanvas');const ctx = canvas.getContext('2d');let width = canvas.width = window.innerWidth;let height = canvas.height = window.innerHeight;
        let particles = [];const text = "FBA 库存中枢";
        function initText() {ctx.clearRect(0, 0, width, height);ctx.font = `900 80px "Microsoft YaHei", Arial`;ctx.fillStyle = "#10B981";ctx.textAlign = "center";ctx.textBaseline = "middle";ctx.fillText(text, width / 2, height / 2);
            const imageData = ctx.getImageData(0, 0, width, height);const data = imageData.data;particles = [];
            for (let y = 0; y < height; y += 3) {for (let x = 0; x < width; x += 3) {const index = (y * width + x) * 4;const alpha = data[index + 3];if (alpha > 128) {particles.push({x: x + (Math.random() - 0.5) * 200,y: y + (Math.random() - 0.5) * 200,targetX: x,targetY: y,vx: (Math.random() - 0.5) * 5,vy: (Math.random() - 0.5) * 5,size: 1.5,color: `rgba(16, 185, 129, ${Math.random() * 0.5 + 0.5})`});}}}ctx.clearRect(0, 0, width, height);}
        let mouse = { x: 0, y: 0 };let isMouseOverText = false;
        function animate() {ctx.clearRect(0, 0, width, height);
            if (Math.abs(mouse.x - width/2) < 250 && Math.abs(mouse.y - height/2) < 80) {isMouseOverText = true;} else {isMouseOverText = false;}
            particles.forEach(p => {if (isMouseOverText) {const dx = p.x - mouse.x;const dy = p.y - mouse.y;const dist = Math.sqrt(dx * dx + dy * dy);if (dist < 150) {p.vx += (dx / dist) * 2;p.vy += (dy / dist) * 2;}}
                p.vx += (p.targetX - p.x) * 0.03;p.vy += (p.targetY - p.y) * 0.03;p.vx *= 0.85;p.vy *= 0.85;p.x += p.vx;p.y += p.vy;
                const floatY = Math.sin(Date.now() * 0.002 + p.targetX * 0.01) * 5;
                ctx.beginPath();ctx.arc(p.x, p.y + floatY, p.size, 0, Math.PI * 2);ctx.fillStyle = p.color;ctx.fill();});
            requestAnimationFrame(animate);}
        window.addEventListener('mousemove', (e) => {const rect = canvas.getBoundingClientRect();mouse.x = e.clientX - rect.left;mouse.y = e.clientY - rect.top;});
        window.addEventListener('resize', () => {width = canvas.width = window.innerWidth;height = canvas.height = window.innerHeight;initText();});
        initText();animate();</script></body></html>
    """
    components.html(particle_html, height=500)


# ==================== 8. 页面布局与逻辑 ====================
if not st.session_state.is_logged_in:
    render_particle_home()
    st.markdown(
        "<h3 style='text-align: center; color: #94A3B8; margin-top: -50px;'>👈 请在左侧输入访问密码解锁系统</h3>",
        unsafe_allow_html=True)
    st.stop()

st.markdown('<h1 style="text-align: center; color: #F8FAFC; margin-bottom: 0;">📦 FBA 库存经营指挥中心</h1>',
            unsafe_allow_html=True)
st.markdown(
    '<p style="text-align: center; color: #94A3B8; margin-top: 5px;">上传一次Excel，每天监控你的断货、滞销与资金风险</p>',
    unsafe_allow_html=True)
st.write("")

with st.expander("⚙️ 核心参数设置（默认即可，可按需调整）", expanded=False):
    col1, col2, col3 = st.columns(3)
    with col1: safety_days = st.number_input("安全库存天数", min_value=0, max_value=90, value=14)
    with col2: default_lead = st.number_input("默认采购提前期（表格没填时使用）", min_value=1, max_value=90, value=30)
    with col3: default_ship = st.number_input("默认物流时效（表格没填时使用）", min_value=1, max_value=120, value=30)
    exclude_skus_raw = st.text_input("排除指定 SKU（可选，多个用逗号分隔）", placeholder="例如：A009, B010")

st.write("---")
uploaded_file = st.file_uploader("📂 上传你的库存报表（支持 Excel 或 CSV，表头多乱都行）", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    temp_path = f"temp_{st.session_state.session_id}_{uploaded_file.name}"
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    with st.status("🤖 正在解析数据并计算...", expanded=True) as status:
        # --- 数据读取逻辑 (压缩展示，与之前一致) ---
        st.write("读取原始数据...")
        raw_rows = []
        if uploaded_file.name.lower().endswith(".csv"):
            with open(temp_path, "r", encoding="utf-8-sig", errors="ignore") as f:
                reader = csv.reader(f)
                for row in reader:
                    if row and any(str(c).strip() for c in row): raw_rows.append(tuple(str(c) for c in row))
        else:
            wb = load_workbook(temp_path, data_only=True)
            ws = wb.active
            raw_rows = [tuple(str(c) if c is not None else "" for c in r) for r in ws.iter_rows(values_only=True)]

        st.write("智能清洗并映射列名...")
        max_cols = max(
            [len([c for c in r if c is not None and str(c).strip() != ""]) for r in raw_rows]) if raw_rows else 0
        if max_cols <= 1:
            full_text = "\n".join([str(r[0]) for r in raw_rows if r and r[0] is not None])
            raw_rows = smart_split(full_text)

        if not raw_rows or len(raw_rows) < 2: st.error("❌ 表格数据不足。"); st.stop()

        clean_rows = [tuple(str(c).strip() for c in r) for r in raw_rows if r and any(str(c).strip() for c in r)]
        headers = list(clean_rows[0]);
        data_rows = clean_rows[1:]

        # --- 表头字典 ---
        col_sku = find_column(headers,
                              ["sku", "产品", "编号", "货号", "asin", "fnsku", "msku", "品名", "商品编码", "产品编号",
                               "产品代码", "item", "product", "产品id", "商品id", "款号", "型号", "product-id",
                               "seller-sku"])
        col_stock = find_column(headers,
                                ["库存", "stock", "available", "可售", "在库", "qty", "quantity", "总库存", "可售库存",
                                 "可用库存", "现有库存", "当前库存", "在售", "总数量", "可用数", "可售数量", "库存数量",
                                 "fba库存", "afn-total-quantity", "afn-warehouse-quantity", "afn-fulfillable-quantity"])
        col_fba_stock = find_column(headers, ["fba可售", "fba库存", "fba available", "亚马逊可售", "fba", "fba可售库存",
                                              "fba在库", "afn-fulfillable-quantity", "afn-warehouse-quantity",
                                              "fba-available"])
        col_overseas_stock = find_column(headers,
                                         ["海外仓", "overseas", "海外库存", "海外仓库存", "海外仓可售", "第三方仓",
                                          "海外仓库", "overseas-stock", "海外仓储"])
        col_awd_stock = find_column(headers,
                                    ["awd", "awd库存", "awd仓", "awd-warehouse", "amazon-warehouse-distribution"])
        col_domestic_stock = find_column(headers, ["国内仓", "国内库存", "自建仓", "国内仓库", "本地仓", "国内仓库存"])
        col_sales = find_column(headers,
                                ["销量", "sales", "日均", "velocity", "sold", "7天销量", "30天销量", "平均日销",
                                 "日均销量", "日销量", "日销", "平均销量", "7日销量", "30日销量", "周销量", "月销量",
                                 "销售速度", "daily-sales", "avg-sales", "units-sold", "销量(30天)"])
        col_lead = find_column(headers,
                               ["提前", "采购", "lead", "交期", "采购周期", "质检", "上架", "采购提前期", "备货周期",
                                "采购天数", "订货周期", "生产周期", "采购交期", "lead-time", "采购提前", "交货周期"])
        col_ship = find_column(headers,
                               ["时效", "物流", "shipping", "运输", "transit", "运输周期", "物流时效", "运输天数",
                                "头程", "头程时效", "物流天数", "运输时间", "ship-time", "shipping-time", "配送周期"])
        col_transit = find_column(headers,
                                  ["在途", "inbound", "transit", "在运", "pending", "已下单", "在途库存", "在途数量",
                                   "已发未到", "运输中", "途中", "in-transit", "在途数", "发货中", "已发货"])
        col_cost = find_column(headers,
                               ["采购价", "成本", "cost", "单价", "进价", "拿货价", "采购成本", "采购单价", "进货价",
                                "产品成本", "采购价格", "成本价", "unit-cost", "purchase-price", "buy-price"])
        col_price = find_column(headers,
                                ["售价", "价格", "price", "销售价", "零售价", "listing价", "销售价格", "商品售价",
                                 "your-price", "sell-price", "售价(USD)"])

        if col_sku is None or col_stock is None or col_sales is None:
            st.error("❌ 未找到必备列：SKU、库存、日均销量。请检查表头。");
            st.stop()

        exclude_skus = {s.strip().upper() for s in exclude_skus_raw.split(",") if
                        s.strip()} if exclude_skus_raw else set()

        st.write("正在计算风险与资金占用...")
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

            total_stock = to_number(get_cell(col_stock)) or 0
            fba_stock = to_number(get_cell(col_fba_stock)) if col_fba_stock is not None else total_stock
            overseas_stock = to_number(get_cell(col_overseas_stock)) or 0
            awd_stock = to_number(get_cell(col_awd_stock)) or 0
            domestic_stock = to_number(get_cell(col_domestic_stock)) or 0
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

            transfer_action = "-"
            if fba_days_left < 7 and overseas_stock > 0 and needed > fba_stock:
                transfer_amount = min(int(needed - fba_stock), int(overseas_stock))
                if transfer_amount > 0: transfer_action = f"🚨 紧急调拨 {transfer_amount} 个从海外仓到FBA"

            risk = "🔴 高（即将断货）" if days_left < lead_time + 7 else (
                "🟡 中（需要关注）" if days_left < lead_time + 14 else "🟢 低（库存健康）")
            days_after = (total_stock + restock) / daily_sales if daily_sales > 0 else 0
            overstock = "⚠️ 补后超储" if days_after > 120 else "正常"
            transit_flag = f"⚠️ 已有 {int(in_transit)} 个在途" if in_transit > 0 else "-"
            is_slow_moving = 1 if days_left > 90 and daily_sales > 0 else 0

            results.append({
                "SKU": str(sku), "当前总库存": int(total_stock), "FBA可售": int(fba_stock),
                "海外仓备胎": int(overseas_stock), "AWD库存": int(awd_stock),
                "国内仓": int(domestic_stock), "日均销量": round(daily_sales, 1),
                "建议补货量": restock, "补货成本(元)": f"¥{restock * cost:,.0f}" if cost > 0 else "-",
                "库存占用资金(元)": f"¥{total_stock * cost:,.0f}" if cost > 0 else "-",
                "断货损失(元/天)": f"¥{daily_sales * price:,.0f}" if price > 0 else "-",
                "断货风险": risk, "距离断货(天)": int(days_left), "调拨建议": transfer_action,
                "补后可用(天)": int(days_after), "在途库存提示": transit_flag,
                "计算过程": f"{daily_sales}×({lead_time}+{shipping_time}+{safety_days}) - {total_stock} - {in_transit} = {restock}",
                "滞销预警": "⚠️ 滞销积压" if is_slow_moving else "正常", "备注": overstock
            })

        try:
            if os.path.exists(temp_path): os.remove(temp_path)
        except:
            pass

        if not results: st.error("❌ 没有可计算的有效数据。"); st.stop()
        status.update(label="✅ 计算完成！", state="complete", expanded=False)

    # ==================== 9. 仪表盘展示区 ====================
    st.write("")
    risk_order = {"🔴 高（即将断货）": 0, "🟡 中（需要关注）": 1, "🟢 低（库存健康）": 2}
    results_sorted = sorted(results, key=lambda x: risk_order.get(x["断货风险"], 3))

    high_risk = len([r for r in results_sorted if "🔴" in r["断货风险"]])
    slow_moving = len([r for r in results_sorted if "⚠️" in r.get("滞销预警", "")])
    total_capital = sum([int(r["库存占用资金(元)"].replace('¥', '').replace(',', '')) for r in results_sorted if
                         r["库存占用资金(元)"] != "-"])
    health_score = max(0, 100 - high_risk * 10 - slow_moving * 5)

    # 记录本次数据
    now_str = datetime.now().strftime("%m-%d %H:%M:%S")
    current_record = {
        "上传时间": now_str,
        "高危SKU数": high_risk,
        "滞销SKU数": slow_moving,
        "资金占用(元)": total_capital,
        "健康分": health_score,
        "full_data": results_sorted  # 保存全量数据用于SKU级对比
    }
    st.session_state.history_records.append(current_record)

    # 顶部四个大卡片
    colA, colB, colC, colD = st.columns(4)
    if len(st.session_state.history_records) >= 2:
        last = st.session_state.history_records[-2]
        with colA:
            st.metric("🚨 今日即将断货 SKU", f"{high_risk} 个", delta=f"{high_risk - last['高危SKU数']:+d} 个 vs 上次",
                      delta_color="inverse")
        with colB:
            st.metric("📉 持续滞销 SKU", f"{slow_moving} 个", delta=f"{slow_moving - last['滞销SKU数']:+d} 个 vs 上次",
                      delta_color="inverse")
        with colC:
            st.metric("💰 当前占用资金", f"¥{total_capital:,}",
                      delta=f"{total_capital - last['资金占用(元)']:+,} 元 vs 上次", delta_color="inverse")
        with colD:
            st.metric("💚 库存健康评分", f"{health_score} 分", delta=f"{health_score - last['健康分']:+d} 分 vs 上次",
                      delta_color="inverse")
    else:
        with colA:
            st.metric("🚨 今日即将断货 SKU", f"{high_risk} 个", "首次分析")
        with colB:
            st.metric("📉 持续滞销 SKU", f"{slow_moving} 个", "首次分析")
        with colC:
            st.metric("💰 当前占用资金", f"¥{total_capital:,}", "首次分析")
        with colD:
            st.metric("💚 库存健康评分", f"{health_score} 分", "首次分析")

    # --- 极致对比看板 ---
    if len(st.session_state.history_records) > 1:
        with st.expander("📊 查看极致库存对比看板（点击展开）", expanded=True):
            last_record = st.session_state.history_records[-2]

            # 1. 并排柱状图对比
            st.markdown("### 📈 核心指标并排对比")
            col_chart1, col_chart2, col_chart3 = st.columns(3)

            # 构建对比数据
            compare_df = pd.DataFrame({
                '指标': ['断货风险 (个)', '断货风险 (个)', '滞销风险 (个)', '滞销风险 (个)', '资金占用 (元)',
                         '资金占用 (元)'],
                '周期': ['上次', '本次', '上次', '本次', '上次', '本次'],
                '数值': [last_record['高危SKU数'], high_risk, last_record['滞销SKU数'], slow_moving,
                         last_record['资金占用(元)'], total_capital]
            })

            with col_chart1:
                df_risk = compare_df[compare_df['指标'] == '断货风险 (个)']
                fig_risk = px.bar(df_risk, x='周期', y='数值', color='周期', title="断货风险对比",
                                  color_discrete_map={'上次': '#F59E0B', '本次': '#EF4444'}, text_auto=True)
                fig_risk.update_layout(template='plotly_dark', showlegend=False, height=300,
                                       margin=dict(l=10, r=10, t=40, b=10))
                st.plotly_chart(fig_risk, use_container_width=True)

            with col_chart2:
                df_slow = compare_df[compare_df['指标'] == '滞销风险 (个)']
                fig_slow = px.bar(df_slow, x='周期', y='数值', color='周期', title="滞销风险对比",
                                  color_discrete_map={'上次': '#F59E0B', '本次': '#EF4444'}, text_auto=True)
                fig_slow.update_layout(template='plotly_dark', showlegend=False, height=300,
                                       margin=dict(l=10, r=10, t=40, b=10))
                st.plotly_chart(fig_slow, use_container_width=True)

            with col_chart3:
                df_cap = compare_df[compare_df['指标'] == '资金占用 (元)']
                fig_cap = px.bar(df_cap, x='周期', y='数值', color='周期', title="资金占用对比",
                                 color_discrete_map={'上次': '#10B981', '本次': '#F59E0B'}, text_auto='.2s')
                fig_cap.update_layout(template='plotly_dark', showlegend=False, height=300,
                                      margin=dict(l=10, r=10, t=40, b=10))
                st.plotly_chart(fig_cap, use_container_width=True)

            # 2. SKU级全量对比明细表
            st.markdown("### 🔍 问题SKU级变化明细（红=恶化，绿=好转）")
            current_map = {r["SKU"]: r for r in results_sorted}
            previous_map = {r["SKU"]: r for r in last_record['full_data']}
            all_skus = set(current_map.keys()) | set(previous_map.keys())

            sku_deltas = []
            for sku in all_skus:
                prev = previous_map.get(sku, {})
                curr = current_map.get(sku, {})

                prev_risk = "🔴" in prev.get("断货风险", "")
                curr_risk = "🔴" in curr.get("断货风险", "")

                prev_cap = to_number(str(prev.get("库存占用资金(元)", "0")).replace('¥', '').replace(',', '')) or 0
                curr_cap = to_number(str(curr.get("库存占用资金(元)", "0")).replace('¥', '').replace(',', '')) or 0
                delta_cap = curr_cap - prev_cap

                if prev_risk != curr_risk or abs(delta_cap) > 500:
                    status = "🔴 恶化（新入高危）" if curr_risk and not prev_risk else (
                        "🟢 好转（解除高危）" if not curr_risk and prev_risk else "➖ 持续高危" if curr_risk else "正常")
                    sku_deltas.append({
                        "SKU": sku,
                        "状态": status,
                        "上次断货风险": "高危" if prev_risk else "正常",
                        "本次断货风险": "高危" if curr_risk else "正常",
                        "上次库存": prev.get("当前总库存", "-"),
                        "本次库存": curr.get("当前总库存", "-"),
                        "库存变化": curr.get("当前总库存", 0) - prev.get("当前总库存", 0),
                        "资金变化(元)": f"{delta_cap:+,.0f}",
                        "备注": "⚠️ 库存大幅波动" if abs(delta_cap) > 1000 else ""
                    })

            if sku_deltas:
                st.dataframe(pd.DataFrame(sku_deltas), use_container_width=True)
            else:
                st.success("✅ 本次与上次对比，未发现显著恶化的SKU。")

            # 3. 核心洞察（老板体）
            st.markdown("### 💡 核心洞察（老板必看）")
            delta_cap_total = total_capital - last_record['资金占用(元)']
            delta_risk_total = high_risk - last_record['高危SKU数']

            insights = []
            if delta_risk_total > 0:
                insights.append(
                    f"⚠️ **紧急警告**：断货风险比上次增加了 **{delta_risk_total}** 个！请立刻排查上面的红色SKU，优先补货。")
            elif delta_risk_total < 0:
                insights.append(f"✅ **风险缓解**：断货风险比上次减少了 **{abs(delta_risk_total)}** 个，干得漂亮。")

            if delta_cap_total > 0:
                insights.append(
                    f"🔴 **资金预警**：资金占用比上次多压了 **¥{abs(delta_cap_total):,}**。如果这些钱压在滞销货上，请尽快清仓。")
            elif delta_cap_total < 0:
                insights.append(
                    f"💚 **资金效率提升**：资金占用比上次减少了 **¥{abs(delta_cap_total):,}**，库存周转在变快。")

            if not insights:
                insights.append("✅ 整体库存状况稳定，请继续保持警惕。")

            st.info(" | ".join(insights))

    else:
        st.info("💡 提示：再次上传一张新报表，这里就会出现专业的库存健康趋势对比看板。")

    # 今日补货行动清单
    st.write("---")
    st.markdown("### 📋 【今日补货行动清单】")
    urgent_skus = [r for r in results_sorted if "🔴" in r["断货风险"]][:5]
    if urgent_skus:
        for r in urgent_skus:
            col_cb, col_info = st.columns([1, 10])
            with col_cb: st.checkbox("", key=f"cb_{r['SKU']}_{now_str}", label_visibility="collapsed")
            with col_info:
                st.markdown(
                    f"🔴 **{r['SKU']}**：建议补货 **{r['建议补货量']}** 个（当前库存 {r['当前总库存']}，距离断货 {r['距离断货(天)']} 天）")
    else:
        st.success("✅ 当前没有紧急补货项，库存健康。")

    # ==================== 10. 选项卡详细展示区 ====================
    st.write("---")
    tab1, tab2, tab3 = st.tabs(["📊 全量风险明细表", "💸 资金与滞销监控", "💬 AI运营总监解读"])
    with tab1:
        st.dataframe(results_sorted, use_container_width=True)

    with tab2:
        col_rank1, col_rank2 = st.columns(2)
        with col_rank1:
            st.markdown("### 💸 压钱最多的 SKU TOP 10")
            capital_sorted = sorted(results_sorted, key=lambda x: to_number(
                x["库存占用资金(元)"].replace('¥', '').replace(',', '')) or 0, reverse=True)[:10]
            for rank, r in enumerate(capital_sorted, 1):
                amount = r["库存占用资金(元)"] if r["库存占用资金(元)"] != "-" else "¥0"
                st.markdown(
                    f'<div class="rank-item"><span>Top {rank}. {r["SKU"]} (库存 {r["当前总库存"]} 个)</span><span class="rank-highlight">{amount}</span></div>',
                    unsafe_allow_html=True)
        with col_rank2:
            st.markdown("### 📉 滞销积压清理 TOP 10")
            slow_sorted = [r for r in results_sorted if r.get("滞销预警") == "⚠️ 滞销积压"]
            if slow_sorted:
                for rank, r in enumerate(slow_sorted[:10], 1):
                    st.markdown(
                        f'<div class="rank-item"><span>Top {rank}. {r["SKU"]} (库存 {r["当前总库存"]} 个)</span><span style="color: #F59E0B;">建议清仓/调拨</span></div>',
                        unsafe_allow_html=True)
            else:
                st.info("✅ 当前没有滞销积压的 SKU。")

    with tab3:
        if st.button("生成 AI 运营总监行动指令", use_container_width=True):
            with st.spinner("AI正在分析资金与风险数据..."):
                system_prompt = "你是一个拥有10年实战经验的亚马逊供应链运营总监。语气冷静果断，只给老板下达具体行动指令。"
                user_prompt = f"以下是我的FBA库存数据计算结果（含资金与滞销风险）：\n{results_sorted[:8]}\n请输出行动简报：\n1. 致命风险（一句话点名）\n2. 本周必做动作（紧急补货、海外仓调拨、滞销清理）\n3. 整体资金健康度点评"
                try:
                    st.info(call_llm(system_prompt, user_prompt))
                except Exception as e:
                    st.error(f"解读失败：{e}")

    st.write("---")
    output = io.BytesIO()
    wb_out = Workbook();
    ws_out = wb_out.active
    headers_out = list(results_sorted[0].keys());
    ws_out.append(headers_out)
    for r in results_sorted: ws_out.append([r[h] for h in headers_out])
    for i, h in enumerate(headers_out, 1): ws_out.column_dimensions[get_column_letter(i)].width = max(len(str(h)) * 2,
                                                                                                      12)
    wb_out.save(output);
    output.seek(0)
    st.download_button("📥 下载完整 Excel 智能诊断报告", data=output, file_name="FBA库存诊断报告.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)

else:
    render_particle_home()
    st.markdown(
        "<p style='text-align: center; color: #475569; font-size: 1rem; margin-top: -50px;'>↑ 在上方上传你的库存报表，立刻体验每日经营指挥中心 ↑</p>",
        unsafe_allow_html=True)
