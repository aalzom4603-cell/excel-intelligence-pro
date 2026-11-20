import io, os, json, hashlib
from datetime import datetime
import pandas as pd
import plotly.express as px
import streamlit as st
from supabase import create_client, Client

st.set_page_config(page_title="EXCEL INTELLIGENCE PRO", page_icon="📊", layout="wide")

# ---------- Theme ----------
st.markdown("""
<style>
:root { --navy:#10233f; --blue:#2563eb; --bg:#f4f7fb; }
.stApp { background:var(--bg); }
.block-container { padding-top:1.5rem; max-width:1500px; }
.hero {background:linear-gradient(110deg,#10233f,#1e4b83);padding:26px 30px;border-radius:18px;color:white;margin-bottom:18px}
.hero h1 {font-size:30px;margin:0 0 6px 0}
.hero p {opacity:.86;margin:0}
.kpi {background:white;border:1px solid #e5eaf2;border-radius:15px;padding:18px 20px;min-height:112px;box-shadow:0 4px 14px rgba(16,35,63,.04)}
.kpi-label {color:#64748b;font-size:13px}
.kpi-value {color:#10233f;font-size:28px;font-weight:750;margin-top:7px}
.section {font-size:20px;font-weight:750;color:#10233f;margin:18px 0 10px}
.small-muted {color:#64748b;font-size:12px}
[data-testid="stSidebar"] {background:#10233f}
[data-testid="stSidebar"] * {color:#f8fafc !important}
</style>
""", unsafe_allow_html=True)

def tr(ar, en):
    return ar if st.session_state.get("lang","العربية") == "العربية" else en

def get_supabase():
    url = st.secrets.get("SUPABASE_URL", os.getenv("SUPABASE_URL", ""))
    key = st.secrets.get("SUPABASE_ANON_KEY", os.getenv("SUPABASE_ANON_KEY", ""))
    if url and key:
        return create_client(url, key)
    return None

def fmt_num(x):
    try: return f"{int(x):,}"
    except: return str(x)

def safe_df(df):
    return df.replace({pd.NA: None}).where(pd.notnull(df), None)

def login_ui(sb):
    st.markdown('<div class="hero"><h1>EXCEL INTELLIGENCE PRO</h1><p>Smart Excel Analysis • Interactive Business Dashboards</p></div>', unsafe_allow_html=True)
    st.subheader(tr("تسجيل الدخول","Sign in"))
    mode = st.radio(tr("اختر العملية","Choose action"), [tr("دخول","Sign in"), tr("إنشاء حساب","Create account")], horizontal=True)
    email = st.text_input("Email")
    password = st.text_input(tr("كلمة المرور","Password"), type="password")
    if st.button(tr("متابعة","Continue"), type="primary", use_container_width=True):
        try:
            if "إنشاء" in mode or "Create" in mode:
                res = sb.auth.sign_up({"email":email,"password":password})
                st.success(tr("تم إنشاء الحساب. راجع بريدك إذا طُلب تأكيده.","Account created. Check your email if confirmation is required."))
            else:
                res = sb.auth.sign_in_with_password({"email":email,"password":password})
                st.session_state["auth_user"] = res.user
                st.session_state["auth_session"] = res.session
                st.rerun()
        except Exception as e:
            st.error(f"{tr('تعذر إتمام العملية','Authentication failed')}: {e}")

def analyze_df(df):
    rows, cols = df.shape
    missing = int(df.isna().sum().sum())
    dup = int(df.duplicated().sum()) if rows else 0
    numeric = df.select_dtypes(include="number").columns.tolist()
    dates = []
    for c in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            dates.append(c)
        elif df[c].dtype == "object":
            sample = df[c].dropna().astype(str).head(30)
            if len(sample) and sample.str.match(r"\d{1,4}[-/]\d{1,2}[-/]\d{1,4}").mean() > .7:
                dates.append(c)
    categorical = [c for c in df.columns if c not in numeric]
    return {"rows":rows,"cols":cols,"missing":missing,"duplicates":dup,
            "numeric":numeric,"dates":dates,"categorical":categorical}

def make_insights(df, meta):
    out=[]
    out.append(tr(f"يحتوي الشيت على {meta['rows']:,} سجل و{meta['cols']} عمود.",
                  f"This sheet contains {meta['rows']:,} records and {meta['cols']} columns."))
    if meta["missing"]:
        out.append(tr(f"يوجد {meta['missing']:,} خلية فارغة؛ يُنصح بمراجعتها قبل اعتماد النتائج.",
                      f"There are {meta['missing']:,} empty cells; review them before relying on the results."))
    else:
        out.append(tr("لا توجد خلايا فارغة في البيانات المقروءة.","No empty cells were detected in the imported data."))
    if meta["duplicates"]:
        out.append(tr(f"تم اكتشاف {meta['duplicates']:,} صف مكرر بالكامل.",
                      f"{meta['duplicates']:,} fully duplicated rows were detected."))
    if meta["numeric"]:
        out.append(tr("الأعمدة الرقمية: ","Numeric columns: ") + ", ".join(map(str,meta["numeric"][:12])))
    if meta["dates"]:
        out.append(tr("أعمدة يُحتمل أنها تواريخ: ","Potential date columns: ") + ", ".join(map(str,meta["dates"][:12])))
    if meta["rows"] == 0:
        out.append(tr("الشيت فارغ أو لا يحتوي على سجلات قابلة للتحليل.","The sheet is empty or has no analyzable records."))
    return out

def upload_to_cloud(sb, user_id, name, raw):
    path = f"{user_id}/{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}_{name}"
    sb.storage.from_("excel-files").upload(path, raw, {"content-type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","upsert":"true"})
    sb.table("excel_uploads").insert({"user_id":user_id,"file_name":name,"storage_path":path,"file_size":len(raw)}).execute()

# ---------- Sidebar ----------
with st.sidebar:
    st.markdown("## 📊 EXCEL INTELLIGENCE")
    st.session_state["lang"] = st.selectbox("Language / اللغة", ["العربية","English"], index=0)
    st.markdown("---")
    page = st.radio(tr("القائمة الرئيسية","MAIN MENU"),
                    [tr("لوحة التحكم","Dashboard"),tr("تحليل ملف جديد","New Analysis"),tr("التحليل المخصص","Custom Analysis"),tr("عن النظام","About")])
    st.markdown("---")
    st.caption("EXCEL INTELLIGENCE PRO")
    sb = get_supabase()
    if sb and st.session_state.get("auth_user"):
        if st.button(tr("تسجيل الخروج","Sign out")):
            st.session_state.pop("auth_user",None); st.session_state.pop("auth_session",None); st.rerun()

# ---------- Auth gate ----------
if sb and not st.session_state.get("auth_user"):
    login_ui(sb)
    st.stop()

# ---------- Session data ----------
if "workbooks" not in st.session_state: st.session_state["workbooks"] = {}
if "active_file" not in st.session_state: st.session_state["active_file"] = None

st.markdown('<div class="hero"><h1>EXCEL INTELLIGENCE PRO</h1><p>تحويل ملفات Excel إلى معرفة وقرارات واضحة | Turn spreadsheets into clear insights and decisions</p></div>', unsafe_allow_html=True)

if not sb:
    st.info(tr("وضع المعاينة: التحليل يعمل، لكن تسجيل الدخول والحفظ السحابي يحتاجان إعداد Supabase. لا ترفع بيانات سرية في وضع المعاينة.",
               "Preview mode: analysis works, but login and cloud saving require Supabase setup. Do not upload confidential data in preview mode."))

if page in [tr("تحليل ملف جديد","New Analysis"), tr("لوحة التحكم","Dashboard")]:
    st.markdown(f'<div class="section">{tr("رفع ملف Excel","Upload an Excel file")}</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(tr("اختر ملف Excel أو CSV","Choose an Excel or CSV file"), type=["xlsx","xls","xlsm","csv"])
    if uploaded:
        raw = uploaded.getvalue()
        try:
            if uploaded.name.lower().endswith(".csv"):
                book = {"CSV": pd.read_csv(io.BytesIO(raw))}
            else:
                book = pd.read_excel(io.BytesIO(raw), sheet_name=None)
            st.session_state["workbooks"][uploaded.name] = {"sheets":book,"raw":raw}
            st.session_state["active_file"] = uploaded.name
            if sb:
                try:
                    uid = st.session_state["auth_user"].id
                    upload_to_cloud(sb, uid, uploaded.name, raw)
                    st.success(tr("تم رفع الملف وحفظ نسخة سحابية.","File uploaded and a cloud copy was saved."))
                except Exception as e:
                    st.warning(tr("تم التحليل في الجلسة، لكن تعذر الحفظ السحابي: ","Analyzed in this session, but cloud save failed: ")+str(e))
            else:
                st.success(tr("تم تحميل الملف وقراءة الشيتات بنجاح.","File loaded and sheets were read successfully."))
        except Exception as e:
            st.error(tr("تعذر قراءة الملف: ","Could not read file: ")+str(e))

    if st.session_state["workbooks"]:
        names=list(st.session_state["workbooks"].keys())
        selected=st.selectbox(tr("الملف الحالي","Current file"),names,index=names.index(st.session_state["active_file"]) if st.session_state["active_file"] in names else 0)
        st.session_state["active_file"]=selected
        sheets=st.session_state["workbooks"][selected]["sheets"]
        sheet_name=st.selectbox(tr("اختر الشيت","Select sheet"),list(sheets.keys()))
        df=sheets[sheet_name]
        meta=analyze_df(df)
        st.markdown(f'<div class="section">{tr("المؤشرات الرئيسية","Key metrics")} — {sheet_name}</div>',unsafe_allow_html=True)
        k1,k2,k3,k4=st.columns(4)
        for col,label,val in [(k1,tr("عدد السجلات","Records"),fmt_num(meta["rows"])),(k2,tr("عدد الأعمدة","Columns"),fmt_num(meta["cols"])),(k3,tr("الخلايا الفارغة","Missing cells"),fmt_num(meta["missing"])),(k4,tr("الصفوف المكررة","Duplicate rows"),fmt_num(meta["duplicates"]))]:
            col.markdown(f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div></div>',unsafe_allow_html=True)
        left,right=st.columns([1,1])
        with left:
            st.markdown(f'<div class="section">{tr("التفسير التلقائي","Automatic interpretation")}</div>',unsafe_allow_html=True)
            for item in make_insights(df,meta): st.write("• "+item)
        with right:
            st.markdown(f'<div class="section">{tr("أنواع البيانات","Data types")}</div>',unsafe_allow_html=True)
            type_df=pd.DataFrame({"Column":df.columns.astype(str),"Data type":[str(x) for x in df.dtypes],"Filled values":[int(df[c].notna().sum()) for c in df.columns]})
            st.dataframe(type_df,use_container_width=True,hide_index=True)
        st.markdown(f'<div class="section">{tr("استكشاف البيانات","Data explorer")}</div>',unsafe_allow_html=True)
        st.dataframe(df.head(500),use_container_width=True)
        if meta["numeric"]:
            st.markdown(f'<div class="section">{tr("ملخص الأعمدة الرقمية","Numeric summary")}</div>',unsafe_allow_html=True)
            st.dataframe(df[meta["numeric"]].describe().T,use_container_width=True)
        csv=df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(tr("تنزيل بيانات الشيت CSV","Download sheet as CSV"),csv,file_name=f"{sheet_name}.csv",mime="text/csv")

elif page == tr("التحليل المخصص","Custom Analysis"):
    if not st.session_state["workbooks"]:
        st.warning(tr("ارفع ملفاً أولاً من صفحة تحليل ملف جديد.","Upload a file first from New Analysis."))
    else:
        f=st.selectbox(tr("الملف","File"),list(st.session_state["workbooks"].keys()))
        sheets=st.session_state["workbooks"][f]["sheets"]
        sn=st.selectbox(tr("الشيت","Sheet"),list(sheets.keys()))
        df=sheets[sn]
        cols=list(df.columns)
        c1,c2=st.columns(2)
        with c1: x=st.selectbox(tr("عمود التصنيف / المحور الأفقي","Category / X-axis"),cols)
        with c2: y=st.selectbox(tr("عمود القيمة / المحور الرأسي (اختياري)","Value / Y-axis (optional)"),["—"]+cols)
        chart=st.selectbox(tr("نوع الرسم","Chart type"),["Bar","Pie","Line","Histogram"])
        agg=st.selectbox(tr("طريقة التجميع","Aggregation"),["count","sum","mean"])
        if st.button(tr("إنشاء التحليل","Generate analysis"),type="primary"):
            try:
                if chart=="Histogram":
                    fig=px.histogram(df,x=x)
                elif y=="—":
                    counts=df[x].astype(str).value_counts().head(30).reset_index()
                    counts.columns=[x,"Count"]
                    fig=px.pie(counts,names=x,values="Count") if chart=="Pie" else px.bar(counts,x=x,y="Count")
                else:
                    tmp=df.groupby(x,dropna=False)[y].agg(agg).reset_index()
                    if chart=="Pie": fig=px.pie(tmp,names=x,values=y)
                    elif chart=="Line": fig=px.line(tmp,x=x,y=y,markers=True)
                    else: fig=px.bar(tmp,x=x,y=y)
                st.plotly_chart(fig,use_container_width=True)
                st.success(tr("تم إنشاء التحليل بناءً على الأعمدة المحددة.","Analysis generated from the selected columns."))
            except Exception as e:
                st.error(tr("تعذر إنشاء الرسم. تأكد أن عمود القيمة رقمي عند استخدام الجمع أو المتوسط. ","Could not create chart. Ensure the value column is numeric for sum or mean. ")+str(e))

elif page == tr("عن النظام","About"):
    st.subheader(tr("عن EXCEL INTELLIGENCE PRO","About EXCEL INTELLIGENCE PRO"))
    st.write(tr("منصة لتحليل ملفات Excel، واستكشاف الشيتات، واستخراج مؤشرات الأداء، وإنشاء رسوم بيانية مخصصة.",
                "A platform for analyzing Excel workbooks, exploring sheets, extracting KPIs, and creating custom charts."))
    st.warning(tr("هذه نسخة أولية قابلة للتطوير. التحليل يصف البيانات إحصائياً ولا يفترض معنى أعمدة غير واضحة دون تعريفها.",
                  "This is an initial extensible version. It describes data statistically and does not assume the meaning of unclear columns without definitions."))
