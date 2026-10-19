# ============================================================
# 白粥 · FBA库存指挥中心（高级 SaaS 落地页版）
# ============================================================

import os
import uuid
import io
import csv
import re
import hashlib
from datetime import datetime
import streamlit as st
import streamlit.components.v1 as components
import requests
import pandas as pd
import plotly.express as px
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

# ==================== 1. 页面基础配置 ====================
st.set_page_config(
    page_title="白粥 · FBA库存指挥中心",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# 访问密码（按需求先写死为 666）
ACCESS_PASSWORD = "666"

# ==================== 2. 全局高级 SaaS 深色 UI ====================
st.markdown("""
<style>
/* ---------- 隐藏 Streamlit 默认元素：菜单栏 / 页脚 / 顶部红线 ---------- */
#MainMenu { visibility: hidden; height: 0; }
footer { visibility: hidden; }
header[data-testid="stHeader"] { display: none !important; }
[data-testid="stToolbar"] { display: none !important; }
[data-testid="stDecoration"] { display: none !important; }   /* 顶部红线 */
[data-testid="stStatusWidget"] { display: none !important; }
[data-testid="stSidebarCollapsedControl"] { display: none !important; }

/* ---------- 全局暗色底 (#0A0E17) + 浅色文字 ---------- */
html, body { background: #0A0E17 !important; }
.stApp,
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
section.main, .main, .block-container, [data-testid="stHeader"] {
    background: transparent !important;
}
.stApp, .stApp p, .stApp span, .stApp li, .stApp label,
[data-testid="stMarkdownContainer"] { color: #CBD5E1; }
h1, h2, h3, h4, h5, h6 { color: #F8FAFC !important; }

/* ---------- 全屏粒子背景 iframe：固定铺满、置于内容之下 ---------- */
/* 这里用 id 精确命中背景组件：它的 HTML 内部脚本会把自己所在的宿主 iframe 标记为 #bg_particles。
   不能写成全局 iframe 选择器——否则登录页的「粒子散开文字」组件会被一并钉成全屏，
   并被 pointer-events:none 屏蔽鼠标事件，导致鼠标滑过散开的效果完全失效
   （两个组件 iframe 的 title/testid/class 完全一致，无法用属性区分）。*/
iframe#bg_particles {
    position: fixed !important;
    top: 0 !important;
    left: 0 !important;
    width: 100vw !important;
    height: 100vh !important;
    z-index: -1 !important;
    border: none !important;
    pointer-events: none !important;
}

/* ---------- 玻璃质感卡片 ---------- */
.glass-card {
    background: rgba(255, 255, 255, 0.035);
    backdrop-filter: blur(14px);
    -webkit-backdrop-filter: blur(14px);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 18px;
    padding: 26px 22px;
    text-align: center;
    transition: all 0.3s ease;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.45);
    min-height: 190px;
}
.glass-card:hover {
    transform: translateY(-6px);
    border-color: rgba(16, 185, 129, 0.55);
    box-shadow: 0 14px 44px rgba(16, 185, 129, 0.22);
}
.glass-icon { font-size: 2.2rem; margin-bottom: 12px; }
.glass-title { font-size: 1.15rem; font-weight: 700; color: #F8FAFC; margin-bottom: 10px; }
.glass-desc { font-size: 0.9rem; color: #94A3B8; line-height: 1.65; }

/* ---------- 指标卡 ---------- */
div[data-testid="stMetric"] {
    background: rgba(255, 255, 255, 0.03);
    border: 1px solid #1E293B;
    border-radius: 12px;
    padding: 15px 20px;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.5);
}
div[data-testid="stMetric"] div[data-testid="stMetricValue"] { font-size: 2rem; font-weight: bold; color: #F8FAFC; }
div[data-testid="stMetric"] div[data-testid="stMetricDelta"] { font-size: 1rem; }

/* ---------- Tabs ---------- */
.stTabs [data-baseweb="tab-list"] { gap: 24px; background-color: transparent; border-bottom: 1px solid #1E293B; }
.stTabs [data-baseweb="tab"] {
    height: 52px; white-space: pre-wrap; background-color: transparent;
    border-radius: 6px 6px 0 0; padding: 10px; color: #94A3B8; font-weight: 600;
}
.stTabs [data-baseweb="tab"]:hover { color: #E2E8F0; }
.stTabs [aria-selected="true"] { color: #10B981 !important; border-bottom: 2px solid #10B981 !important; }

/* ---------- 按钮：悬停变色 ---------- */
.stButton > button, .stFormSubmitButton > button, .stDownloadButton > button {
    width: 100%; border-radius: 8px; background: #10B981; color: #FFFFFF !important;
    font-weight: 600; border: none; transition: all 0.22s ease; padding: 0.6rem 1rem;
}
.stButton > button:hover, .stFormSubmitButton > button:hover, .stDownloadButton > button:hover {
    background: #34D399; color: #06281F !important;
    box-shadow: 0 0 22px rgba(16, 185, 129, 0.55); transform: translateY(-2px);
}

/* ---------- 输入控件暗色化 ---------- */
.stTextInput input, .stNumberInput input,
div[data-baseweb="input"], div[data-baseweb="base-input"] {
    background: rgba(15, 23, 42, 0.85) !important; color: #E2E8F0 !important;
    border: 1px solid #1E293B !important; border-radius: 8px !important;
}
.stTextInput input:focus, .stNumberInput input:focus { border-color: #10B981 !important; }

/* ---------- 上传区 ---------- */
[data-testid="stFileUploader"] section,
[data-testid="stFileUploaderDropzone"] {
    background: rgba(255, 255, 255, 0.03) !important;
    border: 1px dashed #334155 !important; border-radius: 12px !important;
}
[data-testid="stFileUploader"] section:hover,
[data-testid="stFileUploaderDropzone"]:hover { border-color: #10B981 !important; }

/* ---------- 榜单条目 ---------- */
.rank-item {
    padding: 10px; border-bottom: 1px solid #1E293B;
    display: flex; justify-content: space-between;
}
.rank-item:last-child { border-bottom: none; }
.rank-highlight { color: #EF4444; font-weight: bold; }

/* ---------- 高危 SKU 行动卡片 ---------- */
.sku-card {
    background: rgba(255, 255, 255, 0.035);
    border: 1px solid #1E293B;
    border-left: 4px solid #EF4444;
    border-radius: 12px;
    padding: 16px 18px;
    margin-bottom: 12px;
    box-shadow: 0 6px 20px rgba(0, 0, 0, 0.35);
    transition: all 0.25s ease;
}
.sku-card:hover { border-color: #334155; border-left-color: #F43F5E; transform: translateX(3px); }
.sku-card.fatal { background: rgba(244, 63, 94, 0.08); border-left-color: #F43F5E; }
.sku-card-title { font-size: 1.05rem; font-weight: 800; color: #F8FAFC; margin-bottom: 6px; }
.sku-card-line { color: #94A3B8; font-size: 0.88rem; line-height: 1.7; }
.sku-card-line b { color: #E2E8F0; }
.fatal-warn {
    color: #F43F5E; font-weight: 800; font-size: 0.95rem;
    margin-top: 8px; letter-spacing: 0.4px;
    animation: fatalPulse 1.6s ease-in-out infinite;
}
@keyframes fatalPulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.55; } }
</style>
""", unsafe_allow_html=True)


# ==================== 3. 全屏粒子连线背景 ====================
PARTICLE_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<style>
  html, body { margin: 0; padding: 0; width: 100%; height: 100%; overflow: hidden; background: #0A0E17; }
  #tsparticles { position: fixed; inset: 0; width: 100%; height: 100%; }
</style>
</head>
<body>
<div id="tsparticles"></div>
<script>
  /* 把自己所在的宿主 iframe 打上唯一 id，供父页面样式精确命中本组件。
     components.html 使用 srcdoc，与父页面同源，故 window.frameElement 可访问。
     这样父页面就不必用全局 iframe 选择器，也就不会误伤登录页的「粒子散开文字」组件。*/
  try { if (window.frameElement) window.frameElement.id = 'bg_particles'; } catch (e) {}
</script>

<!-- 需求指定的主 CDN -->
<script src="https://cdn.jsdelivr.net/npm/@tsparticles/engine@3.0.0/tsparticles.engine.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@tsparticles/all@3.0.0/tsparticles.all.bundle.min.js"></script>

<script>
  var OPTIONS = {
    fullScreen: { enable: false },
    background: { color: { value: "#0A0E17" } },
    fpsLimit: 60,
    detectRetina: true,
    particles: {
      number: { value: 80, density: { enable: true, width: 1920, height: 1080 } },
      color: { value: ["#10B981", "#38BDF8", "#64748B", "#334155"] },
      opacity: { value: { min: 0.15, max: 0.6 } },
      size: { value: { min: 1, max: 3 } },
      links: { enable: true, distance: 150, color: "#1E3A8A", opacity: 0.4, width: 1 },
      move: {
        enable: true, speed: 1.2, direction: "none",
        random: true, straight: false, outModes: { default: "out" }
      }
    },
    interactivity: {
      events: { onHover: { enable: false }, onClick: { enable: false }, resize: true }
    }
  };

  // 主 CDN 不可用时（部分网络会拦截 cdn.jsdelivr.net）自动切换同源镜像
  var MIRRORS = [
    ["https://fastly.jsdelivr.net/npm/@tsparticles/engine@3.0.0/tsparticles.engine.min.js",
     "https://fastly.jsdelivr.net/npm/@tsparticles/all@3.0.0/tsparticles.all.bundle.min.js"],
    ["https://unpkg.com/@tsparticles/engine@3.0.0/tsparticles.engine.min.js",
     "https://unpkg.com/@tsparticles/all@3.0.0/tsparticles.all.bundle.min.js"]
  ];

  function loadScript(src) {
    return new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = src; s.async = true;
      s.onload = resolve; s.onerror = reject;
      document.head.appendChild(s);
    });
  }

  function ready() {
    return !!(window.tsParticles && typeof window.tsParticles.load === "function"
              && typeof window.loadAll === "function");
  }

  async function ensureEngine() {
    if (ready()) return true;
    for (var i = 0; i < MIRRORS.length; i++) {
      try {
        await loadScript(MIRRORS[i][0]);
        await loadScript(MIRRORS[i][1]);
        if (ready()) return true;
      } catch (e) { /* 换下一个镜像继续尝试 */ }
    }
    return false;
  }

  function start() {
    if (!window.tsParticles || typeof window.tsParticles.load !== "function") return;
    window.tsParticles.load({ id: "tsparticles", options: OPTIONS });
  }

  (async function () {
    if (!(await ensureEngine())) return;
    try { await loadAll(window.tsParticles); } catch (e) { /* 插件已内置则忽略 */ }
    start();
    setTimeout(start, 500);
  })();
</script>
</body>
</html>
"""


def render_particle_background():
    """全屏暗色动态粒子连线背景（通过 CSS 固定铺满并置于内容之下）。

    PARTICLE_HTML 内的脚本会把自己所在的宿主 iframe 标记为 id="bg_particles"，
    页面样式据此精确命中本组件，避免误伤登录页同类型的「粒子散开文字」组件。
    """
    components.html(PARTICLE_HTML, height=1, scrolling=False)


render_particle_background()


# ==================== 4. 安全读取 API Key ====================
def get_api_key():
    try:
        return st.secrets["DASHSCOPE_API_KEY"]
    except Exception:
        return os.getenv("DASHSCOPE_API_KEY", "")


API_KEY = get_api_key()

# ==================== 5. 状态初始化 ====================
if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())
if "is_logged_in" not in st.session_state:
    st.session_state.is_logged_in = False
if "last_results" not in st.session_state:
    st.session_state.last_results = None

# --- 模块 2：情景模拟控制台默认值（先写入 session_state，供控件直接读取）---
if "scenario" not in st.session_state:
    st.session_state.scenario = "日常平销"
if "promo_multiplier" not in st.session_state:
    st.session_state.promo_multiplier = 1.0
if "delay_days" not in st.session_state:
    st.session_state.delay_days = 0

# --- 模块 4：上一次「上传」时的基线快照，用于计算 delta 趋势 ---
if "upload_token" not in st.session_state:
    st.session_state.upload_token = None
if "prev_metrics" not in st.session_state:
    st.session_state.prev_metrics = None
if "prev_results" not in st.session_state:
    st.session_state.prev_results = None


# ==================== 6. 调用大模型 ====================
def call_llm(system_prompt, user_message):
    url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    payload = {"model": "qwen-plus",
               "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": user_message}]}
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
        h_lower = str(h).lower().replace(" ", "").replace("_", "").replace("-", "").replace("（", "").replace("）",
                                                                                                             "").replace(
            "(", "").replace(")", "")
        for kw in keywords:
            if kw.lower() in h_lower: return i
    return None


def find_history_columns(headers, exclude_idx=None):
    """模块 4：找出可用于计算销量标准差的『历史销量』列（至少 2 列才有统计意义）。"""
    kws = ["7天", "30天", "近7", "近30", "上周", "本周", "本月", "上月", "周销量", "月销量", "近一周", "近一月",
           "第1周", "第2周", "第3周", "第4周", "week", "month", "day7", "day30"]
    idxs = []
    for i, h in enumerate(headers):
        if i == exclude_idx: continue
        h_lower = str(h).lower().replace(" ", "").replace("_", "").replace("-", "")
        for kw in kws:
            if kw.lower() in h_lower:
                idxs.append(i)
                break
    return idxs


def parse_history_period(header):
    """模块 4：解析历史销量列的时间跨度（天），用于把「周期总量」归一化成「日均」。

    返回 (天数, 类型)：
      - ('daily')  该列本身就是日均口径，天数=1，无需换算
      - ('period') 该列是周期总量，已识别出跨度，需 ÷ 天数
      - ('assumed')无法识别跨度，按日均处理（会在界面上标注假设）
    """
    h = str(header).lower()
    for ch in [" ", "_", "-", "（", "）", "(", ")", "：", ":"]:
        h = h.replace(ch, "")

    # 1) 本身就是日均口径：7日均销 / 30天平均 / daily sales ...
    if any(k in h for k in ["日均", "平均", "avg", "average", "perday", "daily"]):
        return 1, "daily"
    # 2) 明确天数：7天销量 / 近30日 / 30days
    m = re.search(r"近?(\d+)\s*(?:天|日)", h)
    if m: return max(1, int(m.group(1))), "period"
    m = re.search(r"(\d+)\s*days?", h)
    if m: return max(1, int(m.group(1))), "period"
    m = re.search(r"day(\d+)", h)
    if m: return max(1, int(m.group(1))), "period"
    m = re.search(r"(\d+)d", h)          # sales_30d
    if m: return max(1, int(m.group(1))), "period"
    m = re.search(r"近(\d+)", h)          # 近7 / 近30（未带单位）
    if m: return max(1, int(m.group(1))), "period"
    # 3) 周：第1周/2周/1week => 7 天倍数
    m = re.search(r"第?(\d+)\s*(?:周|weeks?|wk)", h)
    if m: return max(1, int(m.group(1))) * 7, "period"
    if "周" in h or "week" in h:
        return 7, "period"
    # 4) 月：3个月=90天；裸「3月」按单月 30 天；本月/上月=30
    m = re.search(r"(\d+)\s*个月", h)
    if m: return max(1, int(m.group(1))) * 30, "period"
    if re.search(r"\d+月", h):
        return 30, "period"
    if "月" in h or "month" in h:
        return 30, "period"
    return 1, "assumed"


def to_number(val):
    if val is None: return None
    s = str(val).strip().replace(",", "").replace("$", "").replace("%", "").replace("￥", "").replace("¥", "")
    if s in ["", "-", "n/a", "na", "null", "none", "N/A"]: return None
    try:
        return float(s)
    except:
        return None


# ==================== 7.1 模块 2：情景模拟预设 ====================
SCENARIO_PRESETS = {
    "日常平销": 1.0,
    "季节性旺季(如夏季/冬季)": 1.6,
    "大促爆发(如黑五/网一/Prime Day)": 3.0,
    "清仓淡季": 0.6,
}


def apply_scenario_preset():
    """选择预设场景时，自动把促销倍数调到对应档位（延误天数保持不变）。"""
    st.session_state.promo_multiplier = SCENARIO_PRESETS.get(st.session_state.scenario, 1.0)


# ==================== 8. 密码拦截：高级 SaaS 落地页 ====================
if not st.session_state.is_logged_in:
    # ---------- 主页粒子散开文字特效（鼠标滑过散开、移开聚合） ----------
    particle_html = """
<!DOCTYPE html>
<html>
<head>
<style>
  html, body { margin: 0; padding: 0; height: 100%; background: transparent; overflow: hidden; }
  canvas { display: block; width: 100%; height: 170px; }
</style>
</head>
<body>
<canvas id="canvas"></canvas>
<script>
(function(){
  const canvas = document.getElementById('canvas');
  const ctx = canvas.getContext('2d');
  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  const text = "白粥 · FBA库存中枢";
  // tx/ty 为鼠标目标位置，x/y 为平滑后的实际位置，避免快速移动时粒子受力跳变
  const mouse = { x: -9999, y: -9999, tx: -9999, ty: -9999, radius: 100 };
  let cssW = 800, cssH = 170;
  let particles = [];
  let textCanvas = null;
  let holeCanvas = null;     // 预渲染的"柔洞"精灵，避免每帧创建大量 gradient
  const HOLE_R = 6;           // 柔洞半径

  function resize(){
    cssW = canvas.clientWidth || 800;
    cssH = 170;
    canvas.width = Math.floor(cssW * dpr);
    canvas.height = Math.floor(cssH * dpr);
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  }

  function buildText(){
    textCanvas = document.createElement('canvas');
    textCanvas.width = Math.floor(cssW);
    textCanvas.height = Math.floor(cssH);
    const t = textCanvas.getContext('2d');
    const fontSize = Math.max(46, Math.min(104, Math.floor(cssW / 9.2)));
    t.fillStyle = '#00C853';
    t.font = 'bold ' + fontSize + 'px "Microsoft YaHei","PingFang SC","Hiragino Sans GB","Heiti SC",sans-serif';
    t.textAlign = 'center';
    t.textBaseline = 'middle';
    t.fillText(text, cssW / 2, cssH / 2 + 2);
  }

  function buildHole(){
    holeCanvas = document.createElement('canvas');
    const s = HOLE_R * 2;
    holeCanvas.width = s; holeCanvas.height = s;
    const hc = holeCanvas.getContext('2d');
    const g = hc.createRadialGradient(s/2, s/2, 0, s/2, s/2, s/2);
    g.addColorStop(0, 'rgba(0,0,0,1)');
    g.addColorStop(1, 'rgba(0,0,0,0)');
    hc.fillStyle = g;
    hc.fillRect(0, 0, s, s);
  }

  function buildParticles(){
    if(!textCanvas) buildText();
    const data = textCanvas.getContext('2d').getImageData(0, 0, textCanvas.width, textCanvas.height).data;
    particles = [];
    const step = 2;
    for(let y = 0; y < textCanvas.height; y += step){
      for(let x = 0; x < textCanvas.width; x += step){
        if(data[(y * textCanvas.width + x) * 4 + 3] > 128){
          particles.push({
            x: x, y: y,
            baseX: x, baseY: y,
            size: Math.random() * 0.9 + 1.5,
            color: '#00E676',
            vx: 0, vy: 0
          });
        }
      }
    }
  }

  function rebuild(){ resize(); buildText(); buildHole(); buildParticles(); }

  canvas.addEventListener('mousemove', (e) => {
    const rect = canvas.getBoundingClientRect();
    mouse.tx = (e.clientX - rect.left) * (cssW / rect.width);
    mouse.ty = (e.clientY - rect.top) * (cssH / rect.height);
  });
  canvas.addEventListener('mouseleave', () => { mouse.tx = -9999; mouse.ty = -9999; });
  window.addEventListener('resize', rebuild);

  function animate(){
    // 鼠标位置平滑过渡（lerp），让斥力场跟随鼠标移动而非瞬移
    mouse.x += (mouse.tx - mouse.x) * 0.28;
    mouse.y += (mouse.ty - mouse.y) * 0.28;

    ctx.clearRect(0, 0, cssW, cssH);

    // 1) 先画整块实心文字（清晰大字）
    ctx.drawImage(textCanvas, 0, 0);

    // 2) 在已经"散开"的粒子的原位置打一个柔化的洞（destination-out），
    //    让实心文字在该处"溶解"为粒子。复用预渲染精灵，性能远高于逐帧建 gradient。
    ctx.globalCompositeOperation = 'destination-out';
    for(let i = 0; i < particles.length; i++){
      const p = particles[i];
      const ox = p.x - p.baseX, oy = p.y - p.baseY;
      if(ox * ox + oy * oy > 1.2){
        ctx.drawImage(holeCanvas, p.baseX - HOLE_R, p.baseY - HOLE_R);
      }
    }
    ctx.globalCompositeOperation = 'source-over';

    // 3) 更新粒子物理 + 绘制处于散开状态的粒子
    for(let i = 0; i < particles.length; i++){
      const p = particles[i];
      const dx = mouse.x - p.x, dy = mouse.y - p.y;
      const dist = Math.sqrt(dx * dx + dy * dy);
      if(dist < mouse.radius && dist > 0.01){
        const force = (mouse.radius - dist) / mouse.radius;
        p.vx -= (dx / dist) * force * 3.3;
        p.vy -= (dy / dist) * force * 3.3;
      }
      // 弹簧回归 + 阻尼（降低刚度、减小阻尼，运动更柔和飘逸）
      p.vx += (p.baseX - p.x) * 0.05;
      p.vy += (p.baseY - p.y) * 0.05;
      p.vx *= 0.90;
      p.vy *= 0.90;
      p.x += p.vx;
      p.y += p.vy;

      // 贴近基点时吸附归零，确保文字完全还原为实心
      const ox = p.x - p.baseX, oy = p.y - p.baseY;
      if(ox * ox + oy * oy < 0.5){
        p.x = p.baseX; p.y = p.baseY; p.vx = 0; p.vy = 0;
      }
      if(ox * ox + oy * oy > 0.6){
        ctx.fillStyle = p.color;
        ctx.beginPath();
        ctx.arc(p.x, p.y, p.size, 0, Math.PI * 2);
        ctx.fill();
      }
    }
    requestAnimationFrame(animate);
  }

  rebuild();
  animate();
})();
</script>
</body>
</html>
"""
    components.html(particle_html, height=180)

    st.markdown("""
    <div style="text-align:center; padding: 0 0 6px 0;">
        <p style="font-size: 1.22rem; color: #94A3B8; margin-top: 2px; font-weight: 300;">
            上传一次 Excel，每天监控你的断货、滞销与资金风险。
        </p>
    </div>
    """, unsafe_allow_html=True)

    st.write("")
    st.write("")

    card1, card2, card3 = st.columns(3)
    with card1:
        st.markdown("""
        <div class="glass-card">
            <div class="glass-icon">🔒</div>
            <div class="glass-title">免授权安全解析</div>
            <div class="glass-desc">无需绑定店铺、无需 API 授权。Excel 本地解析，数据不出你的电脑。</div>
        </div>
        """, unsafe_allow_html=True)
    with card2:
        st.markdown("""
        <div class="glass-card">
            <div class="glass-icon">🧮</div>
            <div class="glass-title">AI CFO 算账</div>
            <div class="glass-desc">自动测算补货量、库存占用资金与断货损失，像 CFO 一样替你算清账。</div>
        </div>
        """, unsafe_allow_html=True)
    with card3:
        st.markdown("""
        <div class="glass-card">
            <div class="glass-icon">🎯</div>
            <div class="glass-title">明确行动指令</div>
            <div class="glass-desc">不堆图表，只给结论：今天先补哪个 SKU、该清哪批库存，一清二楚。</div>
        </div>
        """, unsafe_allow_html=True)

    st.write("")
    st.write("")

    _, mid_col, _ = st.columns([1, 1.15, 1])
    with mid_col:
        with st.form("login_form", clear_on_submit=False):
            pwd = st.text_input("访问密码", type="password", placeholder="请输入访问密码",
                                label_visibility="collapsed")
            submitted = st.form_submit_button("进入系统", use_container_width=True)
        if submitted:
            if pwd == ACCESS_PASSWORD:
                st.session_state.is_logged_in = True
                st.rerun()
            else:
                st.error("❌ 密码错误，请重新输入")

    st.stop()


# ==================== 9. 登录后的业务区 ====================
st.markdown('<h1 style="text-align: center; color: #F8FAFC; margin-bottom: 0;">📦 FBA 库存经营指挥中心</h1>',
            unsafe_allow_html=True)
st.markdown(
    '<p style="text-align: center; color: #94A3B8; margin-top: 5px;">上传一次 Excel，每天监控你的断货、滞销与资金风险</p>',
    unsafe_allow_html=True)
st.write("")

# ---------- 9.1 模块 2：极端情况压力测试控制台 ----------
with st.expander("🛠 极端情况压力测试（情景模拟）", expanded=True):
    sc1, sc2, sc3 = st.columns(3)
    with sc1:
        st.selectbox("预设场景", list(SCENARIO_PRESETS.keys()), key="scenario",
                     on_change=apply_scenario_preset)
    with sc2:
        st.slider("促销销量预期倍数", min_value=0.5, max_value=5.0, step=0.1,
                  key="promo_multiplier", help="预测日均销量 = 原日均销量 × 该倍数")
    with sc3:
        st.slider("极端交期延误天数", min_value=0, max_value=30, step=1,
                  key="delay_days", help="动态总交期 = 原交期 + 该延误天数")
    st.caption(
        f"当前情景 **{st.session_state.scenario}** ｜ 预测日销 = 原日销 × "
        f"**{st.session_state.promo_multiplier:.1f}** ｜ 动态总交期 = 原交期 + "
        f"**{st.session_state.delay_days}** 天。下方全部风险、补货量与致命警告均已按此实时重算。"
    )

with st.expander("⚙️ 核心参数设置（默认即可，可按需调整）", expanded=False):
    col1, col2, col3 = st.columns(3)
    with col1: safety_days = st.number_input("安全库存天数（动态统计不可用时的经验值）", min_value=0, max_value=90, value=14)
    with col2: default_lead = st.number_input("默认采购提前期（表格没填时使用）", min_value=1, max_value=90, value=30)
    with col3: default_ship = st.number_input("默认物流时效（表格没填时使用）", min_value=1, max_value=120, value=30)
    use_dynamic_ss = st.checkbox(
        "启用动态安全库存（1.65 × 历史销量标准差σ × √补货周期）", value=True,
        help="表格里若有多列历史销量（如 7天销量/30天销量/上周/本月），用统计学公式算安全库存；"
             "只有单一日均销量时数据不足，会自动降级为上面的经验天数。")
    exclude_skus_raw = st.text_input("排除指定 SKU（可选，多个用逗号分隔）", placeholder="例如：A009, B010")

st.write("---")
uploaded_file = st.file_uploader("📂 上传你的库存报表（支持 Excel 或 CSV，表头多乱都行）", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    # 纯内存解析，不再往磁盘写临时文件（避免目录只读导致的 PermissionError）
    file_bytes = bytes(uploaded_file.getbuffer())
    # 历史销量列口径校准控件的渲染位（放在计算进度框之外，计算完成后也保持可见）
    hist_cal_placeholder = st.empty()

    with st.status("🤖 正在解析数据并计算...", expanded=True) as status:
        # --- 数据读取逻辑 ---
        st.write("读取原始数据...")
        raw_rows = []
        if uploaded_file.name.lower().endswith(".csv"):
            text = file_bytes.decode("utf-8-sig", errors="ignore")
            reader = csv.reader(io.StringIO(text))
            for row in reader:
                if row and any(str(c).strip() for c in row): raw_rows.append(tuple(str(c) for c in row))
        else:
            wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
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

        # 模块 4：历史销量列（用于动态安全库存的 σ 计算）
        history_cols = find_history_columns(headers, exclude_idx=col_sales)
        # 每列的时间跨度：把「周期总量」列归一化到「日均」后再算 σ
        history_col_meta = {ci: parse_history_period(headers[ci]) for ci in history_cols}

        # 模块 4：历史销量列口径手动校准（默认「自动识别」→ 不改变任何现有结果）
        history_choice, effective_meta = {}, {}
        with hist_cal_placeholder.container():
            if history_cols:
                with st.expander(
                        f"🧮 历史销量列口径校准（检测到 {len(history_cols)} 列；自动识别不准时可在此手动改）",
                        expanded=False):
                    st.caption("自动识别结果会显示在计算结果上方的『安全库存 σ 口径』说明里。"
                               "没把握就保持「自动识别」；识别错了可手动指定，或直接忽略该列。")
                    _ncol = min(3, len(history_cols))
                    cal_cols = st.columns(_ncol)
                    for n, ci in enumerate(history_cols):
                        with cal_cols[n % _ncol]:
                            history_choice[ci] = st.selectbox(
                                f"`{headers[ci]}`", ["自动识别", "按日均", "按7天总量", "按30天总量", "忽略此列"],
                                index=0, key=f"histcal_{ci}_{headers[ci]}")
                    for ci in history_cols:
                        choice = history_choice.get(ci, "自动识别")
                        if choice == "忽略此列":
                            continue                       # 该列不参与 σ 计算
                        if choice == "按日均":
                            effective_meta[ci] = (1, "manual_daily")
                        elif "天总量" in choice:
                            effective_meta[ci] = (int(re.search(r"(\d+)", choice).group(1)), "manual_period")
                        else:
                            effective_meta[ci] = history_col_meta[ci]

        exclude_skus = {s.strip().upper() for s in exclude_skus_raw.split(",") if
                        s.strip()} if exclude_skus_raw else set()

        st.write("正在计算风险与资金占用...")
        # 模块 2：读取情景模拟参数（预测日销、动态交期均由它们驱动）
        promo_multiplier = st.session_state.promo_multiplier
        delay_days = st.session_state.delay_days
        results = []
        abc_source = []   # 模块 3：收集 (SKU, 日均销量, 售价) 供 ABC 分类使用
        for row in data_rows:
            if not row: continue


            def get_cell(idx):
                return row[idx] if idx is not None and idx < len(row) else None


            sku = get_cell(col_sku)
            if sku is None or str(sku).strip() == "": continue
            if str(sku).strip().upper() in exclude_skus: continue

            base_daily_sales = to_number(get_cell(col_sales)) or 0
            if base_daily_sales == 0: continue
            # 模块 2：预测日均销量 = 原日均销量 × 促销倍数
            daily_sales = base_daily_sales * promo_multiplier

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

            # 模块 4：动态安全库存 = 1.65 × 历史销量标准差σ × √补货周期；数据不足自动降级
            replenish_cycle = lead_time + shipping_time + delay_days
            ss_units, ss_source = None, ""
            if use_dynamic_ss and effective_meta:
                norm_series, meta_used = [], []
                for ci, (days, kind) in effective_meta.items():
                    v = to_number(get_cell(ci))
                    if v is None: continue
                    norm_series.append(v / days)      # 周期总量 ÷ 天数 = 日均
                    meta_used.append((days, kind))
                if len(norm_series) >= 2:
                    sd = float(pd.Series(norm_series).std(ddof=1))
                    if sd > 0:
                        ss_units = 1.65 * sd * (replenish_cycle ** 0.5)
                        marks = []
                        if any(d > 1 for d, _ in meta_used): marks.append("已归一化为日均")
                        if any(k == "assumed" for _, k in meta_used): marks.append("含按日均假设列")
                        if any(str(k).startswith("manual") for _, k in meta_used): marks.append("含手工校准")
                        note = ("，" + "、".join(marks)) if marks else ""
                        ss_source = f"动态统计(σ={sd:.2f}，{len(norm_series)}期{note})"
            if ss_units is None or ss_units <= 0:
                ss_units = daily_sales * safety_days
                ss_source = "经验值(数据不足，已降级)"

            # 模块 2：动态交期 = 采购提前期 + 物流时效 + 极端延误天数
            dynamic_lead = lead_time + delay_days
            needed = daily_sales * replenish_cycle + ss_units
            total_days = needed / daily_sales if daily_sales > 0 else 0
            restock = max(0, int(needed - total_stock - in_transit))
            days_left = total_stock / daily_sales if daily_sales > 0 else 0
            fba_days_left = fba_stock / daily_sales if daily_sales > 0 else 0

            transfer_action = "-"
            if fba_days_left < 7 and overseas_stock > 0 and needed > fba_stock:
                transfer_amount = min(int(needed - fba_stock), int(overseas_stock))
                if transfer_amount > 0: transfer_action = f"🚨 紧急调拨 {transfer_amount} 个从海外仓到FBA"

            risk = "🔴 高（即将断货）" if days_left < dynamic_lead + 7 else (
                "🟡 中（需要关注）" if days_left < dynamic_lead + 14 else "🟢 低（库存健康）")
            days_after = (total_stock + restock) / daily_sales if daily_sales > 0 else 0
            overstock = "⚠️ 补后超储" if days_after > 120 else "正常"
            transit_flag = f"⚠️ 已有 {int(in_transit)} 个在途" if in_transit > 0 else "-"
            is_slow_moving = 1 if days_left > 90 and daily_sales > 0 else 0

            # 模块 2：按极端情景推算，库存撑不到货到 → 致命警告（红字）
            extreme_on = (promo_multiplier != 1.0) or (delay_days > 0)
            if days_left < total_days:
                fatal_msg = (f"【致命：叠加促销与延误，{int(days_left)}天后断货】" if extreme_on
                             else f"【致命：预计 {int(days_left)} 天后断货】")
            else:
                fatal_msg = ""

            abc_source.append((str(sku), base_daily_sales, price))   # 模块 3

            results.append({
                "SKU": str(sku), "当前总库存": int(total_stock), "FBA可售": int(fba_stock),
                "海外仓备胎": int(overseas_stock), "AWD库存": int(awd_stock),
                "国内仓": int(domestic_stock), "日均销量": round(base_daily_sales, 1),
                "情景预测日销": round(daily_sales, 1),
                "安全库存(个)": int(ss_units), "安全库存来源": ss_source,
                "建议补货量": restock, "补货成本(元)": f"¥{restock * cost:,.0f}" if cost > 0 else "-",
                "库存占用资金(元)": f"¥{total_stock * cost:,.0f}" if cost > 0 else "-",
                "断货损失(元/天)": f"¥{daily_sales * price:,.0f}" if price > 0 else "-",
                "断货风险": risk, "距离断货(天)": int(days_left), "调拨建议": transfer_action,
                "补后可用(天)": int(days_after), "在途库存提示": transit_flag,
                "情景致命警告": fatal_msg,
                "计算过程": f"{daily_sales}×{replenish_cycle} + 安全库存{int(ss_units)} - 库存{total_stock} - 在途{in_transit} = {restock}",
                "滞销预警": "⚠️ 滞销积压" if is_slow_moving else "正常", "备注": overstock
            })

        if not results: st.error("❌ 没有可计算的有效数据。"); st.stop()
        status.update(label="✅ 计算完成！", state="complete", expanded=False)

    # ==================== 9.9 模块 3：ABC 分类（销售额贡献占比）====================
    abc_df = pd.DataFrame(abc_source, columns=["SKU", "daily", "price"])
    # 有售价用「销量×售价」算销售额；缺售价时退化为用销量加权
    abc_df["weight"] = [d * p if p > 0 else d for d, p in zip(abc_df["daily"], abc_df["price"])]
    total_weight = abc_df["weight"].sum()
    if total_weight <= 0:
        abc_df["weight"] = 1.0          # 完全无量价数据时退化为等权，避免除零
        total_weight = abc_df["weight"].sum()
    abc_df = abc_df.sort_values("weight", ascending=False).reset_index(drop=True)
    # 累计占比（不含本行）→ 跨过 80%/95% 门槛的那个 SKU 仍归入上一档
    abc_df["cum_before"] = abc_df["weight"].cumsum().shift(1, fill_value=0) / total_weight
    abc_df["share"] = abc_df["weight"] / total_weight
    abc_df["ABC"] = ["A" if c < 0.80 else ("B" if c < 0.95 else "C") for c in abc_df["cum_before"]]
    abc_map = {row.SKU: (row.ABC, row.share) for row in abc_df.itertuples()}
    for r in results:
        cls, share = abc_map.get(r["SKU"], ("C", 0.0))
        r["ABC分类"] = cls
        r["销售额贡献占比"] = f"{share:.1%}"

    # ==================== 10. 统计与「上次上传」基线 ====================
    st.write("")
    risk_order = {"🔴 高（即将断货）": 0, "🟡 中（需要关注）": 1, "🟢 低（库存健康）": 2}
    results_sorted = sorted(
        results, key=lambda x: (0 if x.get("情景致命警告") else 1, risk_order.get(x["断货风险"], 3)))

    high_risk = len([r for r in results_sorted if "🔴" in r["断货风险"]])
    slow_moving = len([r for r in results_sorted if "⚠️" in r.get("滞销预警", "")])
    fatal_cnt = len([r for r in results_sorted if r.get("情景致命警告")])
    total_capital = sum([int(r["库存占用资金(元)"].replace('¥', '').replace(',', '')) for r in results_sorted if
                         r["库存占用资金(元)"] != "-"])
    health_score = max(0, 100 - high_risk * 10 - slow_moving * 5)

    # 模块 3：ABC 分布 + 「资金占用大」的相对阈值（前 20%）
    abc_counts = {c: len([r for r in results_sorted if r.get("ABC分类") == c]) for c in ("A", "B", "C")}
    _cap_vals = [to_number(str(r["库存占用资金(元)"]).replace('¥', '').replace(',', '')) or 0 for r in results_sorted]
    cap_threshold = float(pd.Series(_cap_vals).quantile(0.8)) if len(_cap_vals) >= 2 else 0.0

    def capital_of(rec):
        return to_number(str(rec.get("库存占用资金(元)", "0")).replace('¥', '').replace(',', '')) or 0

    # 模块 4：仅在「上传了新文件」时刷新基线，避免拖动情景模拟控件污染历史趋势
    now_str = datetime.now().strftime("%m-%d %H:%M:%S")
    upload_token = hashlib.md5(bytes(uploaded_file.getbuffer())).hexdigest()[:16]
    is_new_upload = st.session_state.upload_token != upload_token
    if is_new_upload:
        st.session_state.prev_metrics = st.session_state.get("last_metrics")
        st.session_state.prev_results = st.session_state.last_results
        st.session_state.upload_token = upload_token

    st.session_state.last_metrics = {
        "上传时间": now_str,
        "高危SKU数": high_risk,
        "滞销SKU数": slow_moving,
        "资金占用(元)": total_capital,
        "健康分": health_score,
    }
    st.session_state.last_results = results_sorted
    prev = st.session_state.prev_metrics

    # ==================== 11. 分页展示区（模块 1：彻底分页，解决拥挤）====================
    st.write("")
    tab1, tab2, tab3, tab4 = st.tabs(
        ["🚨 核心行动指令", "� 资金对比与账本", "� 全景库存明细", "🧠 AI CFO 深度解读"])

    # ---------- Tab 1：核心行动指令 ----------
    with tab1:
        colA, colB, colC, colD = st.columns(4)
        with colA:
            st.metric("🚨 即将断货 SKU", f"{high_risk} 个",
                      delta=(f"{high_risk - prev['高危SKU数']:+d} 个 vs 上次" if prev else "首次分析"),
                      delta_color="inverse")
        with colB:
            st.metric("📉 持续滞销 SKU", f"{slow_moving} 个",
                      delta=(f"{slow_moving - prev['滞销SKU数']:+d} 个 vs 上次" if prev else "首次分析"),
                      delta_color="inverse")
        with colC:
            st.metric("� 当前占用资金", f"¥{total_capital:,}",
                      delta=(f"{total_capital - prev['资金占用(元)']:+,} 元 vs 上次" if prev else "首次分析"),
                      delta_color="inverse")
        with colD:
            st.metric("💚 库存健康评分", f"{health_score} 分",
                      delta=(f"{health_score - prev['健康分']:+d} 分 vs 上次" if prev else "首次分析"))

        if fatal_cnt:
            st.error(f"🚨 当前情景下有 **{fatal_cnt}** 个 SKU 会在补货到货前断货，已置顶为最高优先级。")

        # 模块 4：把历史列的口径解析结果透明展示（含手动校准），便于核对归一化是否正确
        if history_cols:
            _desc = []
            for ci in history_cols:
                if ci not in effective_meta:
                    if history_choice.get(ci) == "忽略此列":
                        _desc.append(f"`{headers[ci]}` → 已手动忽略（不参与 σ）")
                    continue
                d, k = effective_meta[ci]
                if k == "manual_daily":
                    tag = "手动指定：日均口径"
                elif k == "manual_period":
                    tag = f"手动指定：{d}天总量，已 ÷{d} 归一化"
                elif k == "daily":
                    tag = "日均口径，无需换算"
                elif k == "period":
                    tag = f"{d}天总量，已 ÷{d} 归一化"
                else:
                    tag = "未识别跨度，按日均假设"
                _desc.append(f"`{headers[ci]}` → {tag}")
            st.caption("🧮 安全库存 σ 口径：" + " ｜ ".join(_desc))

        st.write("")
        st.markdown("### 🎯 最致命的高危 SKU（最多展示 5 个）")

        # 模块 3：只让 A/B 类的「断货风险高 或 资金占用大」进入核心区；C 类长尾一律不展示
        def in_core(rec):
            if rec.get("ABC分类") not in ("A", "B"):
                return False
            return bool(rec.get("情景致命警告")) or "🔴" in rec.get("断货风险", "") or capital_of(rec) > cap_threshold

        core_skus = [r for r in results_sorted if in_core(r)]
        longtail_skus = [r for r in results_sorted if not in_core(r)]

        if not core_skus:
            st.success("✅ 当前情景下 A/B 类核心 SKU 全部健康，无需紧急动作。")

        for r in core_skus[:5]:
            fatal_cls = " fatal" if r.get("情景致命警告") else ""
            fatal_html = (f'<div class="fatal-warn">{r["情景致命警告"]}</div>'
                          if r.get("情景致命警告") else "")
            transfer_html = (f'<div class="sku-card-line">🚚 {r["调拨建议"]}</div>'
                             if r.get("调拨建议", "-") != "-" else "")
            st.markdown(
                f'<div class="sku-card{fatal_cls}">'
                f'<div class="sku-card-title">🔴 {r["SKU"]}'
                f'<span style="font-size:0.78rem;font-weight:600;color:#94A3B8;"> ｜ {r["ABC分类"]}类'
                f' ｜ 销售额占比 {r["销售额贡献占比"]}</span></div>'
                f'<div class="sku-card-line">建议补货 <b>{r["建议补货量"]}</b> 件 ｜ 当前库存 '
                f'<b>{r["当前总库存"]}</b> ｜ 预测日销 <b>{r["情景预测日销"]}</b> ｜ 距离断货 '
                f'<b>{r["距离断货(天)"]}</b> 天 ｜ 安全库存 <b>{r["安全库存(个)"]}</b></div>'
                f'{transfer_html}{fatal_html}</div>',
                unsafe_allow_html=True)

        rest_core = core_skus[5:]
        if rest_core:
            with st.expander(f"展开查看其余 {len(rest_core)} 个核心高危 SKU（已折叠，避免干扰决策）"):
                for r in rest_core:
                    warn = (f' ｜ <span style="color:#F43F5E;font-weight:800;">{r["情景致命警告"]}</span>'
                            if r.get("情景致命警告") else "")
                    st.markdown(
                        f'🔴 **{r["SKU"]}**（{r["ABC分类"]}类）：建议补货 **{r["建议补货量"]}** 件 ｜ 当前库存 '
                        f'{r["当前总库存"]} ｜ 距离断货 {r["距离断货(天)"]} 天{warn}',
                        unsafe_allow_html=True)

        transfer_list = [r for r in results_sorted if r.get("调拨建议", "-") != "-"]
        if transfer_list:
            with st.expander(f"🚚 紧急调拨建议（{len(transfer_list)} 条）"):
                for r in transfer_list:
                    st.markdown(f"- **{r['SKU']}**：{r['调拨建议']}")

        # 模块 3：C 类长尾 + 低风险 SKU 打包折叠，避免几百个 SKU 淹没决策
        with st.expander(f"🗂 长尾低风险观察区（C类 + 低风险，共 {len(longtail_skus)} 个，已自动折叠）"):
            if longtail_skus:
                st.caption("这些 SKU 属于 C 类长尾或当前无断货/资金风险，日常无需盯盘，按季度复核即可。")
                st.dataframe(
                    pd.DataFrame([{
                        "SKU": r["SKU"], "ABC分类": r["ABC分类"], "销售额贡献占比": r["销售额贡献占比"],
                        "当前总库存": r["当前总库存"], "日均销量": r["日均销量"],
                        "距离断货(天)": r["距离断货(天)"], "库存占用资金(元)": r["库存占用资金(元)"],
                        "断货风险": r["断货风险"], "滞销预警": r["滞销预警"],
                    } for r in longtail_skus]),
                    use_container_width=True)
            else:
                st.success("✅ 没有长尾/低风险 SKU。")

    # ---------- Tab 2：资金对比与账本 ----------
    with tab2:
        if prev is not None:
            st.markdown("### 📈 上次 vs 本次 核心指标并排对比")
            compare_df = pd.DataFrame({
                '指标': ['断货风险 (个)', '断货风险 (个)', '滞销风险 (个)', '滞销风险 (个)', '资金占用 (元)',
                         '资金占用 (元)'],
                '周期': ['上次', '本次', '上次', '本次', '上次', '本次'],
                '数值': [prev['高危SKU数'], high_risk, prev['滞销SKU数'], slow_moving,
                         prev['资金占用(元)'], total_capital]
            })

            col_chart1, col_chart2, col_chart3 = st.columns(3)
            with col_chart1:
                df_risk = compare_df[compare_df['指标'] == '断货风险 (个)']
                fig_risk = px.bar(df_risk, x='周期', y='数值', color='周期', title="断货风险对比",
                                  color_discrete_map={'上次': '#F59E0B', '本次': '#EF4444'}, text_auto=True)
                fig_risk.update_layout(template='plotly_dark', showlegend=False, height=300,
                                       margin=dict(l=10, r=10, t=40, b=10), paper_bgcolor='rgba(0,0,0,0)',
                                       plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_risk, use_container_width=True)

            with col_chart2:
                df_slow = compare_df[compare_df['指标'] == '滞销风险 (个)']
                fig_slow = px.bar(df_slow, x='周期', y='数值', color='周期', title="滞销风险对比",
                                  color_discrete_map={'上次': '#F59E0B', '本次': '#EF4444'}, text_auto=True)
                fig_slow.update_layout(template='plotly_dark', showlegend=False, height=300,
                                       margin=dict(l=10, r=10, t=40, b=10), paper_bgcolor='rgba(0,0,0,0)',
                                       plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_slow, use_container_width=True)

            with col_chart3:
                df_cap = compare_df[compare_df['指标'] == '资金占用 (元)']
                fig_cap = px.bar(df_cap, x='周期', y='数值', color='周期', title="资金占用对比",
                                 color_discrete_map={'上次': '#10B981', '本次': '#F59E0B'}, text_auto='.2s')
                fig_cap.update_layout(template='plotly_dark', showlegend=False, height=300,
                                      margin=dict(l=10, r=10, t=40, b=10), paper_bgcolor='rgba(0,0,0,0)',
                                      plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_cap, use_container_width=True)

            st.markdown("### � 核心洞察（老板必看）")
            delta_cap_total = total_capital - prev['资金占用(元)']
            delta_risk_total = high_risk - prev['高危SKU数']
            insights = []
            if delta_risk_total > 0:
                insights.append(
                    f"⚠️ **紧急警告**：断货风险比上次增加了 **{delta_risk_total}** 个！请立刻排查「核心行动指令」页的红色 SKU。")
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

            prev_map = {r["SKU"]: r for r in (st.session_state.prev_results or [])}
            curr_map = {r["SKU"]: r for r in results_sorted}
            with st.expander("🔍 查看 SKU 级变化明细（红=恶化，绿=好转）"):
                sku_deltas = []
                for sku in set(curr_map) | set(prev_map):
                    p_rec = prev_map.get(sku, {})
                    c_rec = curr_map.get(sku, {})
                    p_risk = "🔴" in p_rec.get("断货风险", "")
                    c_risk = "🔴" in c_rec.get("断货风险", "")
                    p_cap = to_number(str(p_rec.get("库存占用资金(元)", "0")).replace('¥', '').replace(',', '')) or 0
                    c_cap = to_number(str(c_rec.get("库存占用资金(元)", "0")).replace('¥', '').replace(',', '')) or 0
                    d_cap = c_cap - p_cap
                    if p_risk != c_risk or abs(d_cap) > 500:
                        status = ("🔴 恶化（新入高危）" if c_risk and not p_risk else
                                  "🟢 好转（解除高危）" if not c_risk and p_risk else
                                  "➖ 持续高危" if c_risk else "正常")
                        sku_deltas.append({
                            "SKU": sku, "状态": status,
                            "上次断货风险": "高危" if p_risk else "正常",
                            "本次断货风险": "高危" if c_risk else "正常",
                            "上次库存": p_rec.get("当前总库存", "-"),
                            "本次库存": c_rec.get("当前总库存", "-"),
                            "库存变化": c_rec.get("当前总库存", 0) - p_rec.get("当前总库存", 0),
                            "资金变化(元)": f"{d_cap:+,.0f}",
                            "备注": "⚠️ 库存大幅波动" if abs(d_cap) > 1000 else ""
                        })
                if sku_deltas:
                    st.dataframe(pd.DataFrame(sku_deltas), use_container_width=True)
                else:
                    st.success("✅ 本次与上次对比，未发现显著恶化的SKU。")
        else:
            st.info(" 提示：再次上传一张新报表，这里就会出现专业的资金与风险趋势对比看板。")

        st.write("---")
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

    # ---------- Tab 3：全景库存明细 ----------
    with tab3:
        st.markdown("### 🗂 全景库存明细")
        st.caption(f"共 **{len(results_sorted)}** 个有效 SKU ｜ A类 **{abc_counts['A']}** / B类 **{abc_counts['B']}** / "
                   f"C类 **{abc_counts['C']}** ｜ 高危 **{high_risk}** 个 ｜ 滞销 **{slow_moving}** 个 ｜ "
                   f"情景致命 **{fatal_cnt}** 个 ｜ 资金占用 **¥{total_capital:,}**")
        with st.expander("查看完整数据表", expanded=True):
            st.dataframe(results_sorted, use_container_width=True)

    # ---------- Tab 4：AI CFO 解读 ----------
    with tab4:
        st.markdown("### 🧠 AI CFO 资金与风险解读")
        st.caption(
            f"已锁定当前情景：**{st.session_state.scenario}** ｜ 促销倍数 **{st.session_state.promo_multiplier:.1f}×** "
            f"｜ 交期延误 **{st.session_state.delay_days}** 天。AI 会强制结合这些参数给出确定性指令。")

        # 模块 5：按「致命优先」排序，把最危急的 SKU 连同情景参数一起喂给大模型
        ai_candidates = sorted(
            results_sorted, key=lambda x: (0 if x.get("情景致命警告") else 1, x["距离断货(天)"]))[:8]
        ai_payload = [{
            "SKU": r["SKU"], "ABC分类": r["ABC分类"], "预测日销": r["情景预测日销"],
            "当前库存": r["当前总库存"], "距离断货天数": r["距离断货(天)"],
            "建议补货量": r["建议补货量"], "库存占用资金": r["库存占用资金(元)"],
            "断货风险": r["断货风险"], "情景致命警告": r["情景致命警告"],
        } for r in ai_candidates]

        if st.button("生成 AI 运营总监行动指令", use_container_width=True):
            with st.spinner("AI CFO 正在结合情景参数下达指令..."):
                system_prompt = (
                    "你是一个拥有10年实战经验的亚马逊供应链运营总监兼CFO。语气冷静、强硬、绝对确定，"
                    "只下达可执行的行动指令，禁止模棱两可的『建议考虑』类措辞。禁止输出任何免责声明。"
                )
                user_prompt = f"""【本批次情景模拟参数（必须结合并逐条引用）】
- 预设场景：{st.session_state.scenario}
- 促销销量预期倍数：{st.session_state.promo_multiplier:.1f}×
- 极端交期延误天数：{st.session_state.delay_days} 天
- 当前全盘资金占用：¥{total_capital:,}，高危及致命 SKU 共 {high_risk} 个

【SKU 明细（已按致命程度排序，预测日销与距离断货天数均已按上述情景重算）】
{ai_payload}

【输出要求】
1. 针对每一个「情景致命警告」非空的 SKU，严格按以下句式输出一条强制指令：
「鉴于您开启了 {st.session_state.promo_multiplier:.1f} 倍销量预期，且交期延误 {st.session_state.delay_days} 天，SKU-XXX 将在 N 天后断货。系统强制指令：立即对 SKU-XXX 进行空运补货 M 件！其余走海运。」
（N 与 M 必须直接取自明细中的真实数值，不得编造）
2. 给出资金侧结论：本次情景下被占用的资金是否合理，哪些 SKU 应立即停止采购。
3. 最后输出一行『今日必做』：用祈使句列出不超过 3 条动作。
"""
                try:
                    st.info(call_llm(system_prompt, user_prompt))
                except Exception as e:
                    st.error(f"解读失败：{e}")

    # ==================== 12. 报告下载 ====================
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
    st.markdown("""
    <div style="text-align:center; padding: 36px 0;">
        <p style="color: #475569; font-size: 1rem;">↑ 在上方上传你的库存报表，立刻体验每日经营指挥中心 ↑</p>
    </div>
    """, unsafe_allow_html=True)
