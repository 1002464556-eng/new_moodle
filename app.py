import streamlit as st
import pandas as pd
import os
import base64

# ==================== הגדרות עמוד ועיצוב ====================
st.set_page_config(page_title="ישראל ראלית - משימות מודל", layout="wide", page_icon="🇮🇱")

# עיצוב מותאם אישית (CSS) 
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(180deg, #f0f4f8 0%, #e0e8f0 100%);
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    * { direction: rtl; text-align: right; }
    
    div[data-testid="metric-container"] {
        background-color: white;
        border: 1px solid #d1d5db;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        border-right: 5px solid #1D3557;
    }
    .math-title { color: #E63946; font-weight: bold; margin-bottom: 0px; }
    .sci-title { color: #1D3557; font-weight: bold; margin-bottom: 0px; }
    
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ==================== פונקציות עזר ====================
def safe_read_file(filepath):
    if filepath.endswith('.xlsx'):
        try: return pd.read_excel(filepath, engine='openpyxl')
        except: pass
    else:
        for enc in ['utf-8-sig', 'cp1255', 'iso-8859-8', 'utf-8']:
            try: return pd.read_csv(filepath, encoding=enc, dtype=str)
            except: continue
    return pd.DataFrame()

def get_image_base64(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except:
        return None


# ==================== טעינת ועיבוד הנתונים ====================
@st.cache_data
def load_and_process_data():
    all_files = os.listdir('.')
    
    # 1. משיכת החרגות
    exc_file = next((f for f in all_files if 'להחרגה' in f), None)
    excluded_ids = []
    if exc_file:
        df_ex = safe_read_file(exc_file)
        if not df_ex.empty:
            for col in df_ex.columns:
                extracted = df_ex[col].astype(str).str.extract(r'(\d{6})')[0].dropna().tolist()
                if extracted: excluded_ids.extend(extracted)

    # 2. עיבוד קבצי המודל (לוקח רק את הקבצים האחרונים לטובת הרמזור)
    model_frames = []
    for f in sorted([f for f in all_files if 'מודל' in f]):
        domain = 'מתמטיקה' if 'מתמטיקה' in f else ('מדעים' if 'מדעים' in f else 'כללי')
        df = safe_read_file(f)
        if df.empty: continue

        df.columns = df.columns.astype(str).str.strip()
        col_school = next((c for c in df.columns if 'מוסד' in c and 'סמל' not in c), None) or next((c for c in df.columns if 'מוסד' in c), None)
        col_dist = next((c for c in df.columns if 'מחוז' in c), None)
        col_sup = next((c for c in df.columns if 'מפקח' in c), None)
        col_avg = next((c for c in df.columns if 'ממוצע' in c), None)

        if not col_school: continue

        df['סמל מוסד'] = df[col_school].astype(str).str.extract(r'(\d{6})')[0]
        df = df.dropna(subset=['סמל מוסד'])
        df = df[~df['סמל מוסד'].isin(excluded_ids)]
        
        res = pd.DataFrame()
        res['סמל מוסד'] = df['סמל מוסד']
        res['מוסד'] = df[col_school].astype(str).str.replace(r'^\d{6}\s*-\s*', '', regex=True)
        res['מחוז תקשוב'] = df[col_dist].astype(str).str.strip() if col_dist else 'לא ידוע'
        res['שם מפקח'] = df[col_sup].astype(str).str.strip() if col_sup else 'לא ידוע'
        res['תחום'] = domain
        res['ממוצע משימות'] = pd.to_numeric(df[col_avg], errors='coerce').fillna(0).round(2) if col_avg else 0.0
        res['filename'] = f # שמירת שם הקובץ לטובת מיון
        model_frames.append(res)
            
    df_latest = pd.concat(model_frames, ignore_index=True) if model_frames else pd.DataFrame()
    if not df_latest.empty:
        # שומר רק את הנתון מהקובץ הכי מעודכן לכל בית ספר ותחום
        df_latest = df_latest.sort_values('filename').drop_duplicates(subset=['סמל מוסד', 'תחום'], keep='last')

    # 3. עיבוד קבצי התפעולי (סינון מתחת ל-50%)
    op_frames = []
    for f in sorted([f for f in all_files if 'תפעולי' in f]):
        domain = 'מתמטיקה' if 'מתמטיקה' in f else ('מדעים' if 'מדעים' in f else 'כללי')
        df_op = safe_read_file(f)
        if df_op.empty: continue
        
        col_school = next((c for c in df_op.columns if 'מוסד' in c), None)
        col_auth = next((c for c in df_op.columns if 'רשות' in c), None)
        col_dist = next((c for c in df_op.columns if 'מחוז' in c), None)
        col_sup = next((c for c in df_op.columns if 'מפקח' in c), None)
        
        # זיהוי עמודות פוטנציאל וביצוע לצורך חישוב מדויק
        col_pot = next((c for c in df_op.columns if 'פוטנציאל' in c), None)
        col_perf = next((c for c in df_op.columns if 'שביצעו' in c and 'אחוז' not in c), None)
        
        if not col_school or not col_pot or not col_perf: continue
        
        df_op['סמל מוסד'] = df_op[col_school].astype(str).str.extract(r'(\d{6})')[0]
        df_op = df_op.dropna(subset=['סמל מוסד'])
        df_op = df_op[~df_op['סמל מוסד'].isin(excluded_ids)]
        
        df_op['מוסד_נקי'] = df_op[col_school].astype(str).str.replace(r'^\d{6}\s*-\s*', '', regex=True)
        df_op['pot_num'] = pd.to_numeric(df_op[col_pot], errors='coerce').fillna(0)
        df_op['perf_num'] = pd.to_numeric(df_op[col_perf], errors='coerce').fillna(0)
        
        # חיבור כל הכיתות ברמת בית הספר
        grouped = df_op.groupby('סמל מוסד').agg({
            'מוסד_נקי': 'first',
            col_auth: 'first',
            col_dist: 'first',
            col_sup: 'first',
            'pot_num': 'sum',
            'perf_num': 'sum'
        }).reset_index()
        
        # חישוב אחוז ביצוע כולל לבית הספר
        grouped['אחוז_ביצוע'] = grouped.apply(lambda x: (x['perf_num'] / x['pot_num'] * 100) if x['pot_num'] > 0 else 100, axis=1)
        
        # סינון: רק מוסדות מתחת ל-50%
        urgent = grouped[grouped['אחוז_ביצוע'] < 50].copy()
        if not urgent.empty:
            urgent['תחום'] = domain
            urgent.rename(columns={'מוסד_נקי': 'מוסד', col_auth: 'רשות', col_dist: 'מחוז תקשוב', col_sup: 'שם מפקח'}, inplace=True)
            urgent['filename'] = f
            op_frames.append(urgent)

    df_urgent = pd.concat(op_frames, ignore_index=True) if op_frames else pd.DataFrame()
    if not df_urgent.empty:
        df_urgent = df_urgent.sort_values('filename').drop_duplicates(subset=['סמל מוסד', 'תחום'], keep='last')

    return df_latest, df_urgent

df_latest, df_urgent = load_and_process_data()

# ==================== בניית ממשק המשתמש ====================

# הטמעת הלוגו של משרד החינוך (אם התמונה קיימת בתיקייה)
logo_base64 = get_image_base64('image_5e4888.png')
if logo_base64:
    st.markdown(f'<img src="data:image/png;base64,{logo_base64}" style="max-height: 80px; float: right; margin-left: 20px;">', unsafe_allow_html=True)

# כותרת ראשית מדויקת ויעד דינמי
st.title("ישראל ראלית משימות מודל לכיתה ז")
st.markdown(f"**{get_dynamic_target()}**")
st.divider()

if df_latest.empty:
    st.error("🚨 לא נמצאו קבצי מודל. ודאי שהעלית קבצים תקינים.")
    st.stop()

valid_districts = df_latest[df_latest['מחוז תקשוב'] != 'לא ידוע']['מחוז תקשוב'].dropna().unique()
district_list = sorted([str(d) for d in valid_districts])
district = st.sidebar.selectbox("בחר/י מחוז (מומלץ: העיר ירושלים):", district_list) if district_list else ""

if not district:
    st.stop()

df_lat_dist = df_latest[df_latest['מחוז תקשוב'] == district]
df_urg_dist = df_urgent[df_urgent['מחוז תקשוב'] == district] if not df_urgent.empty else pd.DataFrame()

# --- רובריקה 1: מאקרו מחוז ---
st.header(f"תמונת מצב עדכנית - מחוז {district}")
col1, col2 = st.columns(2)
with col1:
    math_avg = df_lat_dist[df_lat_dist['תחום'] == 'מתמטיקה']['ממוצע משימות'].mean()
    st.markdown("<h3 class='math-title'>📐 מתמטיקה</h3>", unsafe_allow_html=True)
    st.metric("ממוצע משימות רשותי", f"{math_avg:.1f}" if pd.notna(math_avg) else "0.0")
with col2:
    sci_avg = df_lat_dist[df_lat_dist['תחום'] == 'מדעים']['ממוצע משימות'].mean()
    st.markdown("<h3 class='sci-title'>🔬 מדעים</h3>", unsafe_allow_html=True)
    st.metric("ממוצע משימות רשותי", f"{sci_avg:.1f}" if pd.notna(sci_avg) else "0.0")
st.divider()

# --- רובריקה 2: מפקחים ורמזור ---
st.header("ניתוח ביצועים לפי מפקח/ת")
valid_sups = df_lat_dist[df_lat_dist['שם מפקח'] != 'לא ידוע']['שם מפקח'].dropna().unique()
supervisors = sorted([str(s) for s in valid_sups])
supervisor = st.selectbox("בחר/י מפקח/ת:", supervisors) if supervisors else ""

if supervisor:
    df_lat_sup = df_lat_dist[df_lat_dist['שם מפקח'] == supervisor]
    
    st.markdown("### סטטוס עדכני - פירוט מוסדות (שיטת הרמזור)")
    def style_row(row, domain):
        val = row['ממוצע משימות']
        if pd.isna(val): return [''] * len(row)
        if domain == 'מתמטיקה': color = '#fad2e1' if val < 5 else ('#fefae0' if val < 12 else '#d8f3dc')
        else: color = '#fad2e1' if val < 2 else ('#fefae0' if val < 6 else '#d8f3dc')
        return [f'background-color: {color}; color: #333;' if col in ['מוסד', 'ממוצע משימות'] else '' for col in row.index]

    t1, t2 = st.tabs(["מתמטיקה", "מדעים"])
    with t1:
        d_m = df_lat_sup[df_lat_sup['תחום'] == 'מתמטיקה'][['סמל מוסד', 'מוסד', 'ממוצע משימות']].sort_values('ממוצע משימות', ascending=False)
        if not d_m.empty: st.dataframe(d_m.style.apply(style_row, domain='מתמטיקה', axis=1), use_container_width=True, hide_index=True)
    with t2:
        d_s = df_lat_sup[df_lat_sup['תחום'] == 'מדעים'][['סמל מוסד', 'מוסד', 'ממוצע משימות']].sort_values('ממוצע משימות', ascending=False)
        if not d_s.empty: st.dataframe(d_s.style.apply(style_row, domain='מדעים', axis=1), use_container_width=True, hide_index=True)

    st.divider()

    # --- רובריקה 3: התערבות דחופה (סינון מתחת ל-50%) ---
    st.header("🚨 מוקדי התערבות דחופים (ביצוע מתחת ל-50%)")
    
    if not df_urg_dist.empty and 'שם מפקח' in df_urg_dist.columns:
        df_urg_sup = df_urg_dist[df_urg_dist['שם מפקח'] == supervisor]
        math_no_course = df_urg_sup[df_urg_sup['תחום'] == 'מתמטיקה']
        sci_no_course = df_urg_sup[df_urg_sup['תחום'] == 'מדעים']

        col_no1, col_no2 = st.columns(2)
        with col_no1:
            with st.expander(f"מתמטיקה: לחץ לצפייה ב-{len(math_no_course)} מוסדות"):
                if not math_no_course.empty:
                    st.dataframe(math_no_course[['מוסד', 'רשות']], hide_index=True, use_container_width=True)
                else:
                    st.success("אין מוסדות הדורשים התערבות.")
        with col_no2:
            with st.expander(f"מדעים: לחץ לצפייה ב-{len(sci_no_course)} מוסדות"):
                if not sci_no_course.empty:
                    st.dataframe(sci_no_course[['מוסד', 'רשות']], hide_index=True, use_container_width=True)
                else:
                    st.success("אין מוסדות הדורשים התערבות.")
    else:
        st.info("לא נמצאו מוסדות עם פחות מ-50% ביצוע באזור זה.")
