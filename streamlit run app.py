"""
APS Planning Intelligence — v4
══════════════════════════════════════════════════════════════
✅ نفس بنية البيانات والـ BOM Explosion من الكود القديم (بدون تعديل)
✅ مقارنة خطتين: V1 (قديمة) و V2 (جديدة) — نفس الملف القديم + ملف جديد
✅ Dashboard Executive + Tabs + تنبيه حرج + تلوين + تصدير ذكي
══════════════════════════════════════════════════════════════
مدخلات البيانات:
  - ملف V1: plan + Component + (MRP Controller اختياري)  ← نفس الكود القديم
  - ملف V2: plan فقط (نفس تنسيق sheet plan)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from collections import defaultdict
import datetime

# ══════════════════════════════════════════════════════════════
# ⚙️ إعداد الصفحة
# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="APS Intelligence v4",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .critical-banner {
        background: linear-gradient(135deg, #ff4444, #cc0000);
        color: white; padding: 14px 20px; border-radius: 10px;
        margin: 8px 0; font-size: 15px; font-weight: bold;
        text-align: center; animation: pulse 2s infinite;
    }
    @keyframes pulse { 0%,100%{ opacity:1 } 50%{ opacity:.82 } }
    .section-title {
        font-size: 19px; font-weight: 700;
        padding-bottom: 4px; border-bottom: 2px solid #4f8ef7;
        margin: 8px 0 6px 0;
    }
    div[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# 1.  COLUMN MAP  — من الكود القديم بدون تعديل
# ══════════════════════════════════════════════════════════════
COLUMN_NAMES = {
    "material":             ["Material", "Item", "code", "Code", "المادة", "Product"],
    "material_desc":        ["Material Description", "Description", "وصف"],
    "order_type":           ["Order Type", "OT", "نوع الطلب", "Sales Org."],
    "component":            ["Component", "Comp", "المكون"],
    "component_desc":       ["Component Description", "Comp Desc", " المسمى", "وصف المكون"],
    "component_uom":        ["Component UoM", "UoM", "الوحدة"],
    "component_qty":        ["Component Quantity", "Qty", "كمية المكون"],
    "base_qty":             ["Base Quantity", "Base Qty", "الكمية الأساسية"],
    "mrp_controller":       ["MRP Controller", "مسؤول MRP"],
    "current_stock":        ["Current Stock", "Stock", "المخزون الحالي", "Unrestricted"],
    "component_order_type": ["Component Order Type", "Order Category", "نوع أمر المكون", "Procurement Type"],
    "hierarchy_level":      ["Hierarchy Level", "Level", "المستوى الهرمي"],
    "parent_material":      ["Parent Material", "Direct Parent", "الأب المباشر"],
}

def col(name_key):
    return COLUMN_NAMES[name_key][0]

def normalize_columns(df, column_map):
    rename_dict = {}
    for key, aliases in column_map.items():
        for alias in aliases:
            if alias in df.columns and alias != aliases[0]:
                rename_dict[alias] = aliases[0]
    return df.rename(columns=rename_dict)


# ══════════════════════════════════════════════════════════════
# 2.  LOAD & VALIDATE — من الكود القديم بدون تعديل
# ══════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def load_and_validate_data(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
        required_sheets = ["plan", "Component"]
        missing_sheets = [s for s in required_sheets if s not in xls.sheet_names]
        if missing_sheets:
            st.error(f"❌ الملف لا يحتوي على الأوراق المطلوبة: {', '.join(missing_sheets)}")
            st.stop()

        plan_df      = normalize_columns(xls.parse("plan"),      COLUMN_NAMES)
        component_df = normalize_columns(xls.parse("Component"), COLUMN_NAMES)
        mrp_df = (
            normalize_columns(xls.parse("MRP Controller"), COLUMN_NAMES)
            if "MRP Controller" in xls.sheet_names else pd.DataFrame()
        )

        known_cols = [aliases[0] for aliases in COLUMN_NAMES.values()]
        extra_cols = [c for c in component_df.columns if c not in known_cols]
        if extra_cols:
            component_df.drop(columns=extra_cols, inplace=True)

        required_plan_cols = [col("material"), col("material_desc"), col("order_type")]
        if not all(c in plan_df.columns for c in required_plan_cols):
            st.error(f"❌ جدول الخطة ناقص أعمدة: {required_plan_cols}")
            st.stop()

        required_comp_cols = [col("material"), col("component"), col("component_qty")]
        if not all(c in component_df.columns for c in required_comp_cols):
            st.error(f"❌ جدول المكونات ناقص أعمدة: {required_comp_cols}")
            st.stop()

        comp_qty_col = col("component_qty")
        base_qty_col = col("base_qty")
        component_df[comp_qty_col] = pd.to_numeric(component_df[comp_qty_col], errors='coerce').fillna(0)

        if base_qty_col in component_df.columns:
            component_df[base_qty_col] = (
                pd.to_numeric(component_df[base_qty_col], errors='coerce').fillna(1).replace(0, 1)
            )
            zero_base = (
                pd.to_numeric(
                    xls.parse("Component").get(base_qty_col, pd.Series(dtype=float)),
                    errors='coerce'
                ) == 0
            ).sum() if base_qty_col in xls.parse("Component").columns else 0
            if zero_base > 0:
                st.warning(f"⚠️ يوجد {zero_base} قيمة صفرية في Base Quantity — تم استبدالها بـ 1.")
            component_df[comp_qty_col] = component_df[comp_qty_col] / component_df[base_qty_col]
            component_df.drop(columns=[base_qty_col], inplace=True)

        if col("current_stock") not in component_df.columns:
            component_df[col("current_stock")] = 0
        else:
            component_df[col("current_stock")] = pd.to_numeric(
                component_df[col("current_stock")], errors='coerce').fillna(0)

        if col("component_order_type") not in component_df.columns:
            component_df[col("component_order_type")] = "غير محدد"
        if col("hierarchy_level") not in component_df.columns:
            component_df[col("hierarchy_level")] = 1
        else:
            component_df[col("hierarchy_level")] = pd.to_numeric(
                component_df[col("hierarchy_level")], errors='coerce').fillna(1).astype(int)
        if col("component_desc") not in component_df.columns:
            component_df[col("component_desc")] = ""
        if col("component_uom") not in component_df.columns:
            component_df[col("component_uom")] = ""
        if col("mrp_controller") not in component_df.columns:
            component_df[col("mrp_controller")] = "غير محدد"

        if col("parent_material") in component_df.columns:
            component_df[col("parent_material")] = (
                component_df[col("parent_material")].astype(str).str.strip()
            )
        else:
            component_df[col("parent_material")] = component_df[col("material")]

        # توحيد وحدات الوزن → KG
        gram_variants = {"g", "gm", "gr", "gram", "grams", "جرام", "جم"}
        uom_c = col("component_uom")
        qty_c = col("component_qty")
        stk_c = col("current_stock")
        is_gram = component_df[uom_c].astype(str).str.strip().str.lower().isin(gram_variants)
        if is_gram.any():
            component_df.loc[is_gram, qty_c] = component_df.loc[is_gram, qty_c] / 1000
            component_df.loc[is_gram, stk_c] = component_df.loc[is_gram, stk_c] / 1000
            component_df.loc[is_gram, uom_c] = "KG"

        # توحيد CM2 → M2
        cm2_variants = {"cm2", "cm^2", "cm²", "سم2", "سم²"}
        is_cm2 = component_df[uom_c].astype(str).str.strip().str.lower().isin(cm2_variants)
        if is_cm2.any():
            component_df.loc[is_cm2, qty_c] = component_df.loc[is_cm2, qty_c] / 10000
            component_df.loc[is_cm2, stk_c] = component_df.loc[is_cm2, stk_c] / 10000
            component_df.loc[is_cm2, uom_c] = "M2"

        return plan_df, component_df, mrp_df

    except Exception as e:
        st.error(f"❌ فشل تحميل الملف: {str(e)}")
        st.stop()


# ══════════════════════════════════════════════════════════════
# 3.  تحميل ملف V2 (plan فقط — نفس تنسيق sheet plan القديم)
# ══════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def load_plan_v2(uploaded_file):
    """يقرأ sheet 'plan' من ملف V2 ويطبق نفس normalize_columns"""
    try:
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
        if "plan" not in xls.sheet_names:
            st.error("❌ ملف V2 لا يحتوي على sheet باسم 'plan'.")
            st.stop()
        df = normalize_columns(xls.parse("plan"), COLUMN_NAMES)
        required = [col("material"), col("material_desc"), col("order_type")]
        if not all(c in df.columns for c in required):
            st.error(f"❌ sheet plan في V2 ناقص أعمدة: {required}")
            st.stop()
        return df
    except Exception as e:
        st.error(f"❌ فشل تحميل ملف V2: {str(e)}")
        st.stop()


# ══════════════════════════════════════════════════════════════
# 4.  BOM EXPLOSION — من الكود القديم بدون تعديل
# ══════════════════════════════════════════════════════════════
def bom_explosion(plan_melted, component_df):
    component_df = component_df.copy()
    component_df[col("component")] = component_df[col("component")].astype(str).str.strip()
    component_df[col("material")]  = component_df[col("material")].astype(str).str.strip()

    has_parent_col = col("parent_material") in component_df.columns
    parent_col = col("parent_material") if has_parent_col else col("material")
    if has_parent_col:
        component_df[parent_col] = component_df[parent_col].astype(str).str.strip()

    component_df = component_df.drop_duplicates(
        subset=[col("material"), parent_col, col("component"), col("component_qty")],
        keep="first"
    )
    bom_core = component_df.groupby(
        [col("material"), parent_col, col("component")], as_index=False
    )[col("component_qty")].sum()

    bom_dict = {}
    for mat, group in bom_core.groupby(col("material")):
        tree = defaultdict(list)
        for _, row in group.iterrows():
            tree[row[parent_col]].append(
                (row[col("component")], row[col("component_qty")])
            )
        bom_dict[mat] = tree

    comp_info = (
        component_df
        .drop_duplicates(subset=[col("component")], keep="last")
        .set_index(col("component"))[[
            col("component_desc"), col("component_uom"),
            col("mrp_controller"), col("current_stock"),
            col("component_order_type"),
        ]]
    )

    def explode(root_material, parent, qty, path, level, row_buf):
        if parent in path or level > 10:
            return
        tree = bom_dict.get(root_material, {})
        children = tree.get(parent, [])
        if not children:
            tree = bom_dict.get(parent, {})
            children = tree.get(parent, [])
        if not children:
            return
        new_path = path | {parent}
        for comp, comp_qty in children:
            needed = qty * comp_qty
            row_buf.append({
                "Parent":                      parent,
                col("component"):              comp,
                col("component_qty"):          comp_qty,
                "Required Component Quantity": needed,
                "BOM Level":                   level,
            })
            explode(root_material, comp, needed, new_path, level + 1, row_buf)

    all_rows = []
    for _, plan_row in plan_melted[plan_melted["Planned Quantity"] > 0].iterrows():
        mat  = str(plan_row[col("material")]).strip()
        qty  = plan_row["Planned Quantity"]
        ot   = plan_row[col("order_type")]
        date = plan_row["Date"]
        row_buf = []
        explode(mat, mat, qty, set(), level=1, row_buf=row_buf)
        mat_desc = str(plan_row.get(col("material_desc"), "")).strip()
        for r in row_buf:
            r[col("material")]      = mat
            r[col("material_desc")] = mat_desc
            r["Order Type"]         = ot
            r["Date"]               = date
        all_rows.extend(row_buf)

    if not all_rows:
        return pd.DataFrame()

    result = pd.DataFrame(all_rows)
    comp_info_clean = (
        comp_info.reset_index()
        .rename(columns={col("component"): "_comp_key"})
        .drop_duplicates(subset=["_comp_key"])
    )
    result = result.merge(comp_info_clean, left_on=col("component"),
                          right_on="_comp_key", how="left"
                          ).drop(columns=["_comp_key"], errors="ignore")
    return result


# ══════════════════════════════════════════════════════════════
# 5.  MELT HELPER — يحول plan_df إلى long format
# ══════════════════════════════════════════════════════════════
def melt_plan(plan_df):
    date_cols = [c for c in plan_df.columns
                 if isinstance(c, (datetime.datetime, pd.Timestamp))]
    if not date_cols:
        # fallback: أي عمود مش من الـ id_cols
        id_set = {col("material"), col("material_desc"), col("order_type")}
        date_cols = [c for c in plan_df.columns if c not in id_set]
    melted = plan_df.melt(
        id_vars=[col("material"), col("material_desc"), col("order_type")],
        value_vars=date_cols,
        var_name="Date",
        value_name="Planned Quantity"
    )
    melted["Date"] = pd.to_datetime(melted["Date"], errors='coerce')
    melted["Planned Quantity"] = pd.to_numeric(melted["Planned Quantity"], errors='coerce').fillna(0)
    return melted[(melted["Planned Quantity"] > 0) & (melted["Date"].notna())].copy()


# ══════════════════════════════════════════════════════════════
# 6.  COMPARE PLANS
# ══════════════════════════════════════════════════════════════
def compare_plans(melted_v1, melted_v2):
    """مقارنة الخطتين على مستوى (Material × Date × Order Type)"""
    m1 = melted_v1.groupby(
        [col("material"), col("material_desc"), col("order_type"), "Date"]
    )["Planned Quantity"].sum().reset_index().rename(columns={"Planned Quantity": "V1_Qty"})

    m2 = melted_v2.groupby(
        [col("material"), col("order_type"), "Date"]
    )["Planned Quantity"].sum().reset_index().rename(columns={"Planned Quantity": "V2_Qty"})

    cmp = pd.merge(m1, m2, on=[col("material"), col("order_type"), "Date"], how="outer").fillna(0)
    cmp["Delta"] = cmp["V2_Qty"] - cmp["V1_Qty"]
    cmp["Change"] = np.select([
        (cmp["V1_Qty"] == 0) & (cmp["V2_Qty"] > 0),
        (cmp["V1_Qty"] > 0)  & (cmp["V2_Qty"] == 0),
        (cmp["V1_Qty"] != cmp["V2_Qty"])
    ], ["🆕 جديد", "🗑️ محذوف", "✏️ متغير"], "✅ مطابق")
    return cmp


# ══════════════════════════════════════════════════════════════
# 7.  AGGREGATE BOM RESULTS → احتياج إجمالي لكل مكون
# ══════════════════════════════════════════════════════════════
def aggregate_requirements(result_df):
    if result_df.empty:
        return pd.DataFrame()
    agg_cols = [
        col("component"), col("component_desc"), col("component_uom"),
        col("mrp_controller"), col("current_stock"), col("component_order_type"),
        "Order Type", "Date", "BOM Level"
    ]
    agg_cols = [c for c in agg_cols if c in result_df.columns]
    return (
        result_df.groupby(agg_cols, as_index=False)["Required Component Quantity"].sum()
    )


# ══════════════════════════════════════════════════════════════
# 8.  RISK ENGINE
# ══════════════════════════════════════════════════════════════
def calc_risk(req_df, model="Days"):
    """
    req_df: نتيجة aggregate_requirements مجمّعة على مستوى Component
    يضيف: required / new_stock / shortage / risk / days
    """
    comp_total = (
        req_df.groupby([col("component"), col("component_desc"),
                        col("component_uom"), col("mrp_controller"),
                        col("current_stock"), col("component_order_type")])
        ["Required Component Quantity"].sum().reset_index()
        .rename(columns={"Required Component Quantity": "Total_Required"})
    )
    df = comp_total.copy()
    df["new_stock"]  = df[col("current_stock")] - df["Total_Required"]
    df["shortage"]   = np.where(df["new_stock"] < 0, -df["new_stock"], 0)

    if model == "Coverage %":
        ratio = np.where(df["Total_Required"] > 0,
                         df[col("current_stock")] / df["Total_Required"], np.inf)
        df["risk"] = np.select(
            [ratio < 0.3, ratio < 0.6, ratio < 0.9],
            ["🔴 حرج", "🟠 مرتفع", "🟡 متوسط"], default="🟢 آمن"
        )
        df["days"] = np.where(df["Total_Required"] > 0,
                               df["new_stock"] / (df["Total_Required"] / 30), np.inf)
    else:
        df["days"] = np.where(df["Total_Required"] > 0,
                               df["new_stock"] / (df["Total_Required"] / 30), np.inf)
        df["risk"] = np.select(
            [df["days"] < 0, df["days"] < 5, df["days"] < 10],
            ["🔴 حرج", "🟠 مرتفع", "🟡 متوسط"], default="🟢 آمن"
        )
    return df


def calc_procurement(risk_df):
    df = risk_df.copy()
    df["action"] = np.select(
        [df["shortage"] > 0, df["risk"] == "🟠 مرتفع"],
        ["🚨 Buy Now", "⚡ Expedite"], default="📅 Monitor"
    )
    return df


# ══════════════════════════════════════════════════════════════
# 9.  TABLE STYLING
# ══════════════════════════════════════════════════════════════
def style_risk(df):
    COLOR = {"🔴 حرج":   "background-color:#3d0000;color:#ff9999",
             "🟠 مرتفع": "background-color:#3d1f00;color:#ffbb66",
             "🟡 متوسط": "background-color:#2e2b00;color:#ffee55",
             "🟢 آمن":   "background-color:#002a00;color:#88ff88"}
    def row(r):
        v = str(r.get("risk", ""))
        s = COLOR.get(v, "")
        return [s] * len(r)
    return df.style.apply(row, axis=1)


def style_compare(df):
    COLOR = {"🆕 جديد":   "background-color:#002a1a;color:#55ffaa",
             "🗑️ محذوف": "background-color:#3d0000;color:#ff8888",
             "✏️ متغير":  "background-color:#1a1a3d;color:#88aaff"}
    def row(r):
        s = COLOR.get(str(r.get("Change", "")), "")
        return [s] * len(r)
    return df.style.apply(row, axis=1)


# ══════════════════════════════════════════════════════════════
# 10.  SMART EXPORT
# ══════════════════════════════════════════════════════════════
def smart_export(sheets_dict):
    out = BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        for name, df in sheets_dict.items():
            if df is not None and not df.empty:
                # تحويل التواريخ إلى نص لتجنب مشاكل Excel
                df_exp = df.copy()
                for c in df_exp.columns:
                    if pd.api.types.is_datetime64_any_dtype(df_exp[c]):
                        df_exp[c] = df_exp[c].dt.strftime("%Y-%m-%d")
                df_exp.to_excel(w, sheet_name=str(name)[:31], index=False)
    return out.getvalue()


# ══════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════
#                     MAIN APP
# ══════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════

st.markdown(
    "## 🚀 APS Intelligence Dashboard "
    "<span style='font-size:13px;color:#888'>v4 — مقارنة الخطط + BOM متعدد المستويات</span>",
    unsafe_allow_html=True
)

# ─── رفع الملفات ───
c1, c2 = st.columns(2)
with c1:
    st.markdown("**📂 ملف الخطة V1** *(plan + Component)*")
    f1 = st.file_uploader("V1", type=["xlsx"], label_visibility="collapsed")
with c2:
    st.markdown("**📂 ملف الخطة V2** *(plan فقط — نفس التنسيق)*")
    f2 = st.file_uploader("V2", type=["xlsx"], label_visibility="collapsed")

if not (f1 and f2):
    st.info("👆 ارفع ملفي الخطة V1 و V2 للبدء.")
    st.stop()

# ══════════════════════════════════════════════════════════════
# LOAD
# ══════════════════════════════════════════════════════════════
with st.spinner("⏳ جاري تحميل البيانات..."):
    plan_v1, component_df, mrp_df = load_and_validate_data(f1)
    plan_v2                        = load_plan_v2(f2)

# ── Melt ──
melted_v1 = melt_plan(plan_v1)
melted_v2 = melt_plan(plan_v2)

# ══════════════════════════════════════════════════════════════
# SIDEBAR — الفلاتر
# ══════════════════════════════════════════════════════════════
st.sidebar.header("⚙️ الإعدادات والفلاتر")

# فترة زمنية
all_dates = pd.concat([melted_v1["Date"], melted_v2["Date"]]).dropna()
min_d, max_d = all_dates.min().date(), all_dates.max().date()
st.sidebar.markdown("### 📅 الفترة الزمنية")
d_from = st.sidebar.date_input("من", value=min_d, min_value=min_d, max_value=max_d)
d_to   = st.sidebar.date_input("إلى", value=max_d, min_value=min_d, max_value=max_d)

melted_v1 = melted_v1[(melted_v1["Date"] >= pd.Timestamp(d_from)) &
                       (melted_v1["Date"] <= pd.Timestamp(d_to))]
melted_v2 = melted_v2[(melted_v2["Date"] >= pd.Timestamp(d_from)) &
                       (melted_v2["Date"] <= pd.Timestamp(d_to))]

# نموذج المخاطر
st.sidebar.markdown("### 📊 نموذج المخاطر")
risk_model = st.sidebar.selectbox("النموذج", ["Days", "Coverage %"])

# MRP Controller
st.sidebar.markdown("### 🏭 MRP Controller")
mrp_opts = sorted(component_df[col("mrp_controller")].dropna().unique())
sel_mrp  = st.sidebar.multiselect("اختر", options=mrp_opts, default=mrp_opts)

# مستوى الخطر
st.sidebar.markdown("### ⚠️ مستوى الخطر")
sel_risk = st.sidebar.multiselect(
    "إظهار",
    ["🔴 حرج", "🟠 مرتفع", "🟡 متوسط", "🟢 آمن"],
    default=["🔴 حرج", "🟠 مرتفع", "🟡 متوسط", "🟢 آمن"]
)

# نوع الأمر
ot_opts = sorted(set(melted_v1[col("order_type")].unique()) |
                 set(melted_v2[col("order_type")].unique()))
st.sidebar.markdown("### 📋 نوع الأمر")
sel_ot = st.sidebar.multiselect("Order Type", options=ot_opts, default=ot_opts)

melted_v1 = melted_v1[melted_v1[col("order_type")].isin(sel_ot)]
melted_v2 = melted_v2[melted_v2[col("order_type")].isin(sel_ot)]

# ══════════════════════════════════════════════════════════════
# BOM EXPLOSION (V1 و V2)
# ══════════════════════════════════════════════════════════════
with st.spinner("⚙️ جارٍ تفجير هيكل المنتجات (BOM)..."):
    res_v1 = bom_explosion(melted_v1, component_df)
    res_v2 = bom_explosion(melted_v2, component_df)

req_v1 = aggregate_requirements(res_v1)
req_v2 = aggregate_requirements(res_v2)

# مقارنة الخطتين
cmp_df = compare_plans(melted_v1, melted_v2)

# مقارنة الاحتياجات (V1 vs V2 لكل مكون)
def comp_requirements_summary(req_df, label):
    if req_df.empty:
        return pd.DataFrame(columns=[col("component"), f"{label}_Required"])
    agg = (req_df.groupby([col("component"), col("component_desc"),
                            col("component_uom"), col("mrp_controller"),
                            col("current_stock"), col("component_order_type")])
           ["Required Component Quantity"].sum().reset_index()
           .rename(columns={"Required Component Quantity": f"{label}_Required"}))
    return agg

req_sum_v1 = comp_requirements_summary(req_v1, "V1")
req_sum_v2 = comp_requirements_summary(req_v2, "V2")

req_compare = pd.merge(req_sum_v1, req_sum_v2,
                        on=[col("component"), col("component_desc"),
                            col("component_uom"), col("mrp_controller"),
                            col("current_stock"), col("component_order_type")],
                        how="outer").fillna(0)
req_compare["Delta"]  = req_compare["V2_Required"] - req_compare["V1_Required"]
req_compare["Delta%"] = np.where(
    req_compare["V1_Required"] > 0,
    (req_compare["Delta"] / req_compare["V1_Required"] * 100).round(1), np.nan
)
req_compare = req_compare.sort_values("Delta", ascending=False)

# Risk & Procurement على V2 (الخطة الجديدة)
risk_base = req_sum_v2.rename(columns={"V2_Required": "Total_Required"}) if not req_sum_v2.empty else pd.DataFrame()
if not risk_base.empty:
    risk_df = calc_risk(req_v2, risk_model)
    proc_df = calc_procurement(risk_df)

    # تطبيق الفلاتر
    if col("mrp_controller") in risk_df.columns:
        risk_df = risk_df[risk_df[col("mrp_controller")].isin(sel_mrp)]
        proc_df = proc_df[proc_df[col("mrp_controller")].isin(sel_mrp)]
    risk_df_f = risk_df[risk_df["risk"].isin(sel_risk)]
    proc_df_f = proc_df[proc_df["risk"].isin(sel_risk)]
else:
    risk_df = risk_df_f = proc_df = proc_df_f = pd.DataFrame()

# ══════════════════════════════════════════════════════════════
# 🚨 تنبيه فوري للبنود الحرجة
# ══════════════════════════════════════════════════════════════
if not risk_df.empty:
    critical = risk_df[risk_df["risk"] == "🔴 حرج"]
    if not critical.empty:
        st.markdown(
            f'<div class="critical-banner">🚨 تحذير: {len(critical)} مكون في مستوى الخطر الحرج '
            f'— إجمالي العجز: {critical["shortage"].sum():,.0f} وحدة — مراجعة فورية مطلوبة!</div>',
            unsafe_allow_html=True
        )
        with st.expander("📋 البنود الحرجة — عرض فوري", expanded=True):
            show_c = [col("component"), col("component_desc"), col("mrp_controller"),
                      "Total_Required", col("current_stock"), "shortage", "days", "action"]
            show_c = [c for c in show_c if c in proc_df.columns]
            st.dataframe(
                style_risk(proc_df[proc_df["risk"] == "🔴 حرج"][show_c]),
                use_container_width=True, height=230
            )

# ══════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Dashboard",
    "🔄 مقارنة الخطط",
    "📦 BOM & الاحتياجات",
    "⚠️ المخاطر",
    "🛒 خطة الشراء",
    "📥 التصدير"
])

# ══════════════════════════════════════════════════════════════
# TAB 1 — EXECUTIVE DASHBOARD
# ══════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-title">📊 لوحة الأداء التنفيذية</div>', unsafe_allow_html=True)

    # ── KPI Row 1: تغييرات الخطة ──
    new_i  = int((cmp_df["Change"] == "🆕 جديد").sum())
    del_i  = int((cmp_df["Change"] == "🗑️ محذوف").sum())
    chg_i  = int((cmp_df["Change"] == "✏️ متغير").sum())
    eql_i  = int((cmp_df["Change"] == "✅ مطابق").sum())
    inc_q  = cmp_df[cmp_df["Delta"] > 0]["Delta"].sum()
    dec_q  = abs(cmp_df[cmp_df["Delta"] < 0]["Delta"].sum())

    k = st.columns(6)
    k[0].metric("✅ مطابق",         eql_i)
    k[1].metric("🆕 جديد",          new_i)
    k[2].metric("🗑️ محذوف",        del_i)
    k[3].metric("✏️ متغير",         chg_i)
    k[4].metric("📈 زيادة الكمية",  f"{inc_q:,.0f}")
    k[5].metric("📉 نقص الكمية",    f"{dec_q:,.0f}")

    st.divider()

    # ── KPI Row 2: المخاطر ──
    if not risk_df.empty:
        crit = int((risk_df["risk"] == "🔴 حرج").sum())
        high = int((risk_df["risk"] == "🟠 مرتفع").sum())
        med  = int((risk_df["risk"] == "🟡 متوسط").sum())
        safe = int((risk_df["risk"] == "🟢 آمن").sum())
        tot_short = risk_df["shortage"].sum()
        days_c    = risk_df["days"].replace([np.inf, -np.inf], np.nan).dropna()
        avg_days  = days_c.mean() if not days_c.empty else np.nan

        r = st.columns(6)
        r[0].metric("🔴 حرج",            crit)
        r[1].metric("🟠 مرتفع",          high)
        r[2].metric("🟡 متوسط",          med)
        r[3].metric("🟢 آمن",            safe)
        r[4].metric("⚠️ إجمالي العجز",   f"{tot_short:,.0f}")
        r[5].metric("📅 متوسط أيام التغطية",
                    f"{avg_days:.1f}" if not np.isnan(avg_days) else "N/A")

        st.divider()

    # ── الرسوم البيانية التنفيذية ──
    g1, g2, g3 = st.columns(3)

    with g1:
        cc = cmp_df["Change"].value_counts().reset_index()
        cc.columns = ["Change", "count"]
        fig = px.pie(cc, values="count", names="Change", title="توزيع تغييرات الخطة")
        st.plotly_chart(fig, use_container_width=True)

    with g2:
        if not risk_df.empty:
            rc = risk_df["risk"].value_counts().reset_index()
            rc.columns = ["risk", "count"]
            cmap = {"🔴 حرج":"#cc3333","🟠 مرتفع":"#cc7700",
                    "🟡 متوسط":"#ccaa00","🟢 آمن":"#228833"}
            fig2 = px.pie(rc, values="count", names="risk", color="risk",
                          color_discrete_map=cmap, title="توزيع مستويات المخاطر (V2)")
            st.plotly_chart(fig2, use_container_width=True)

    with g3:
        if not risk_df.empty and "shortage" in risk_df.columns:
            top5 = risk_df.nlargest(5, "shortage")[[col("component"), "shortage"]]
            if not top5.empty:
                fig3 = px.bar(top5, x="shortage", y=col("component"), orientation="h",
                              title="أكبر 5 عجز في المخزون (V2)",
                              color="shortage", color_continuous_scale="Reds")
                st.plotly_chart(fig3, use_container_width=True)

    # ── Gauge: نسبة الأمان ──
    if not risk_df.empty:
        gg1, gg2 = st.columns(2)
        with gg1:
            total_items = len(risk_df)
            safe_pct = (safe / total_items * 100) if total_items > 0 else 0
            fig_g = go.Figure(go.Indicator(
                mode="gauge+number+delta",
                value=safe_pct,
                delta={"reference": 80, "increasing": {"color": "#228833"}},
                title={"text": "نسبة البنود الآمنة %"},
                gauge={"axis": {"range": [0, 100]}, "bar": {"color": "#228833"},
                       "steps": [{"range": [0, 50], "color": "#cc3333"},
                                  {"range": [50, 75], "color": "#cc7700"},
                                  {"range": [75, 100], "color": "#ccee88"}],
                       "threshold": {"line": {"color": "white", "width": 3}, "value": 80}}
            ))
            fig_g.update_layout(height=260, margin=dict(t=40, b=10, l=20, r=20))
            st.plotly_chart(fig_g, use_container_width=True)

        with gg2:
            # Trend V1 vs V2
            t1 = melted_v1.groupby("Date")["Planned Quantity"].sum().reset_index().rename(columns={"Planned Quantity": "V1"})
            t2 = melted_v2.groupby("Date")["Planned Quantity"].sum().reset_index().rename(columns={"Planned Quantity": "V2"})
            trend = pd.merge(t1, t2, on="Date", how="outer").fillna(0)
            fig_t = px.line(trend, x="Date", y=["V1", "V2"],
                            title="إجمالي الإنتاج: V1 vs V2",
                            color_discrete_map={"V1": "#4477cc", "V2": "#ff7744"})
            st.plotly_chart(fig_t, use_container_width=True)


# ══════════════════════════════════════════════════════════════
# TAB 2 — مقارنة الخطط
# ══════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">🔄 مقارنة V1 مقابل V2</div>', unsafe_allow_html=True)

    s = st.columns(4)
    s[0].metric("✅ مطابق", eql_i)
    s[1].metric("🆕 جديد",  new_i)
    s[2].metric("🗑️ محذوف", del_i)
    s[3].metric("✏️ متغير",  chg_i)

    # فلتر نوع التغيير
    fc = st.multiselect("فلتر التغيير",
                         ["🆕 جديد", "🗑️ محذوف", "✏️ متغير", "✅ مطابق"],
                         default=["🆕 جديد", "🗑️ محذوف", "✏️ متغير"])
    cmp_show = cmp_df[cmp_df["Change"].isin(fc)] if fc else cmp_df

    # رسم خطي
    t1 = melted_v1.groupby("Date")["Planned Quantity"].sum().reset_index().rename(columns={"Planned Quantity": "V1"})
    t2 = melted_v2.groupby("Date")["Planned Quantity"].sum().reset_index().rename(columns={"Planned Quantity": "V2"})
    trend = pd.merge(t1, t2, on="Date", how="outer").fillna(0)
    fig_tr = px.line(trend, x="Date", y=["V1", "V2"],
                     title="إجمالي الإنتاج اليومي: V1 vs V2",
                     color_discrete_map={"V1": "#4477cc", "V2": "#ff7744"})
    st.plotly_chart(fig_tr, use_container_width=True)

    cmp_show_disp = cmp_show.copy()
    cmp_show_disp["Date"] = cmp_show_disp["Date"].astype(str)
    st.dataframe(style_compare(cmp_show_disp.head(3000)), use_container_width=True, height=420)


# ══════════════════════════════════════════════════════════════
# TAB 3 — BOM & الاحتياجات
# ══════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-title">📦 BOM ومقارنة الاحتياجات</div>', unsafe_allow_html=True)

    sub_a, sub_b, sub_c = st.tabs([
        "📊 مقارنة الاحتياجات V1 vs V2",
        "📋 تفاصيل V1",
        "📋 تفاصيل V2"
    ])

    with sub_a:
        ca, cb = st.columns(2)
        ca.metric("⬆️ مكونات زاد احتياجها",  int((req_compare["Delta"] > 0).sum()))
        cb.metric("⬇️ مكونات نقص احتياجها", int((req_compare["Delta"] < 0).sum()))

        disp_cols = [col("component"), col("component_desc"), col("component_uom"),
                     col("mrp_controller"), "V1_Required", "V2_Required", "Delta", "Delta%"]
        disp_cols = [c for c in disp_cols if c in req_compare.columns]
        st.dataframe(req_compare[disp_cols], use_container_width=True, height=380)

        top_d = req_compare.nlargest(10, "Delta")
        if not top_d.empty:
            fig_d = px.bar(top_d, x=col("component"), y=["V1_Required", "V2_Required"],
                           barmode="group", title="أعلى 10 مكونات: V1 vs V2",
                           color_discrete_map={"V1_Required":"#4477cc","V2_Required":"#ff7744"})
            st.plotly_chart(fig_d, use_container_width=True)

    with sub_b:
        if not req_v1.empty:
            rv1_disp = req_v1.copy()
            rv1_disp["Date"] = rv1_disp["Date"].astype(str)
            st.dataframe(rv1_disp, use_container_width=True, height=450)
        else:
            st.info("لا توجد بيانات لـ V1.")

    with sub_c:
        if not req_v2.empty:
            rv2_disp = req_v2.copy()
            rv2_disp["Date"] = rv2_disp["Date"].astype(str)
            st.dataframe(rv2_disp, use_container_width=True, height=450)
        else:
            st.info("لا توجد بيانات لـ V2.")


# ══════════════════════════════════════════════════════════════
# TAB 4 — المخاطر
# ══════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-title">⚠️ تحليل المخاطر (بناءً على V2)</div>', unsafe_allow_html=True)

    if risk_df_f.empty:
        st.info("لا توجد بيانات مخاطر — تحقق من فلاتر MRP ومستوى الخطر.")
    else:
        r = st.columns(4)
        r[0].metric("🔴 حرج",   int((risk_df_f["risk"] == "🔴 حرج").sum()))
        r[1].metric("🟠 مرتفع", int((risk_df_f["risk"] == "🟠 مرتفع").sum()))
        r[2].metric("🟡 متوسط", int((risk_df_f["risk"] == "🟡 متوسط").sum()))
        r[3].metric("🟢 آمن",   int((risk_df_f["risk"] == "🟢 آمن").sum()))

        # Scatter: required vs stock
        fig_sc = px.scatter(
            risk_df_f, x="Total_Required", y=col("current_stock"),
            color="risk", hover_name=col("component"),
            title="المخزون الحالي مقابل الاحتياج الكلي",
            color_discrete_map={"🔴 حرج":"#cc3333","🟠 مرتفع":"#cc7700",
                                 "🟡 متوسط":"#ccaa00","🟢 آمن":"#228833"},
            size="shortage", size_max=40
        )
        st.plotly_chart(fig_sc, use_container_width=True)

        st.dataframe(style_risk(risk_df_f), use_container_width=True, height=440)


# ══════════════════════════════════════════════════════════════
# TAB 5 — خطة الشراء
# ══════════════════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="section-title">🛒 خطة الشراء والإجراءات</div>', unsafe_allow_html=True)

    if proc_df_f.empty:
        st.info("لا توجد بيانات — تحقق من الفلاتر.")
    else:
        p = st.columns(3)
        p[0].metric("🚨 شراء فوري", int((proc_df_f["action"] == "🚨 Buy Now").sum()))
        p[1].metric("⚡ تسريع",      int((proc_df_f["action"] == "⚡ Expedite").sum()))
        p[2].metric("📅 متابعة",    int((proc_df_f["action"] == "📅 Monitor").sum()))

        # Sunburst
        if col("mrp_controller") in proc_df_f.columns:
            sun = (proc_df_f.groupby([col("mrp_controller"), "risk", "action"])
                   .size().reset_index(name="count"))
            if not sun.empty:
                fig_sun = px.sunburst(sun, path=[col("mrp_controller"), "risk", "action"],
                                      values="count",
                                      title="توزيع الإجراءات حسب MRP Controller")
                st.plotly_chart(fig_sun, use_container_width=True)

        st.dataframe(style_risk(proc_df_f), use_container_width=True, height=440)


# ══════════════════════════════════════════════════════════════
# TAB 6 — التصدير الذكي
# ══════════════════════════════════════════════════════════════
with tab6:
    st.markdown('<div class="section-title">📥 التصدير الذكي</div>', unsafe_allow_html=True)

    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("#### 📋 اختر الـ Sheets")
        ex_cmp      = st.checkbox("🔄 مقارنة الخطتين",              value=True)
        ex_req_cmp  = st.checkbox("📊 مقارنة الاحتياجات V1 vs V2",  value=True)
        ex_req_v1   = st.checkbox("📋 الاحتياجات التفصيلية V1",     value=False)
        ex_req_v2   = st.checkbox("📋 الاحتياجات التفصيلية V2",     value=True)
        ex_risk     = st.checkbox("⚠️ تحليل المخاطر",               value=True)
        ex_proc     = st.checkbox("🛒 خطة الشراء",                  value=True)
        ex_bom_v1   = st.checkbox("🔩 نتائج BOM الخام V1",          value=False)
        ex_bom_v2   = st.checkbox("🔩 نتائج BOM الخام V2",          value=False)

    with col_b:
        st.markdown("#### ⚙️ إعدادات")
        file_name   = st.text_input("اسم الملف", value="APS_Report_v4")
        export_note = st.text_area("ملاحظات (تُضاف كـ sheet)", height=90, placeholder="اكتب ملاحظاتك...")
        ex_risk_lvl = st.multiselect(
            "تصدير مستويات خطر محددة فقط",
            ["🔴 حرج", "🟠 مرتفع", "🟡 متوسط", "🟢 آمن"],
            default=["🔴 حرج", "🟠 مرتفع", "🟡 متوسط", "🟢 آمن"]
        )

    # بناء قاموس الـ sheets
    sheets = {}
    if ex_cmp:
        tmp = cmp_df.copy(); tmp["Date"] = tmp["Date"].astype(str)
        sheets["Plan Comparison"] = tmp
    if ex_req_cmp:
        sheets["Req. Comparison"] = req_compare
    if ex_req_v1 and not req_v1.empty:
        tmp = req_v1.copy(); tmp["Date"] = tmp["Date"].astype(str)
        sheets["Requirements V1"] = tmp
    if ex_req_v2 and not req_v2.empty:
        tmp = req_v2.copy(); tmp["Date"] = tmp["Date"].astype(str)
        sheets["Requirements V2"] = tmp
    if ex_risk and not risk_df.empty:
        sheets["Risk Analysis"] = risk_df[risk_df["risk"].isin(ex_risk_lvl)]
    if ex_proc and not proc_df.empty:
        sheets["Procurement"] = proc_df[proc_df["risk"].isin(ex_risk_lvl)]
    if ex_bom_v1 and not res_v1.empty:
        tmp = res_v1.copy(); tmp["Date"] = tmp["Date"].astype(str)
        sheets["BOM Raw V1"] = tmp
    if ex_bom_v2 and not res_v2.empty:
        tmp = res_v2.copy(); tmp["Date"] = tmp["Date"].astype(str)
        sheets["BOM Raw V2"] = tmp
    if export_note.strip():
        sheets["Notes"] = pd.DataFrame({"ملاحظات": [export_note]})

    st.markdown(
        f"**📋 سيتم تصدير {len(sheets)} sheet(s):** " +
        " | ".join(f"`{k}`" for k in sheets.keys())
    )

    if st.button("🚀 توليد ملف التصدير", type="primary", use_container_width=True):
        if sheets:
            with st.spinner("جارٍ إنشاء الملف..."):
                file_data = smart_export(sheets)
            fname = (file_name.strip() or "APS_Report_v4") + ".xlsx"
            st.download_button(
                f"⬇️ تحميل: {fname}", data=file_data, file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.success(f"✅ الملف جاهز — {len(sheets)} sheets")
        else:
            st.warning("⚠️ لم تختر أي sheet!")

# ── Footer ──
st.divider()
st.markdown(
    "<div style='text-align:center;color:#666;font-size:12px'>"
    "م. رضا رشدي — APS Intelligence v4 | Streamlit + Pandas + SAP CS12"
    "</div>", unsafe_allow_html=True
)
