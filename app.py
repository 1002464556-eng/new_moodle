import streamlit as st
import pandas as pd
import plotly.express as px
import os
import re
import base64

# ==================== הגדרות עמוד ועיצוב ====================
st.set_page_config(page_title="מערכת בינה עסקית (BI) - משימות מודל", layout="wide", page_icon="📊")

# עיצוב מותאם אישית (CSS) - רקע ממלכתי, יישור לימין, ועיצוב כרטיסיות
st.markdown("""
<style>
    /* רקע כללי בסגנון ישראל ראלית / משרד החינוך */
    .stApp {
        background: linear-gradient(180deg, #f0f4f8 0%, #e0e8f0 100%);
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    /* יישור לימין של כל האפליקציה */
    * { direction: rtl; text-align: right; }
    
    /* עיצוב כותרות וכרטיסיות נתונים (Metrics) */
    div[data-testid="metric-container"] {
        background-color: white;
        border: 1px solid #d1d5db;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        border-right: 5px solid #1D3557;
    }
    /* צבעים ייעודיים למתמטיקה ולמדעים */
    .math-title { color: #E63946; font-weight: bold; }
    .sci-title { color: #1D3557; font-weight: bold; }
    
    /* הסתרת אלמנטים מיותרים של Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# קבועים
COLOR_MAP = {'מתמטיקה': '#E63946', 'מדעים': '#1D3557'}
TARGETS = {'מתמטיקה': 17.0, 'מדעים': 8.0}

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
    
    # טעינת קובץ החרגות
    exc_file = next((f for f in all_files if 'להחרגה' in f), None)
    excluded_ids = []
    if exc_file:
        df_ex = safe_read_file(exc_file)
        if not df_ex.empty:
            for col in df_ex.columns:
                extracted = df_ex[col].astype(str).str.extract(r'(\d{6})')[0].dropna().tolist()
                if extracted: excluded_ids.extend(extracted)

    # 1. עיבוד קבצי המודל (היסטוריה)
    def process_model_df(filepath, domain):
        df = safe_read_file(filepath)
        if df.empty: return None

        df.columns = df.columns.astype(str).str.strip()
        col_school = next((c for c in df.columns if 'מוסד' in c and 'סמל' not in c), None) or next((c for c in df.columns if 'מוסד' in c), None)
        col_dist = next((c for c in df.columns if 'מחוז' in c), None)
        col_sup = next((c for c in df.columns if 'מפקח' in c), None)
        col_avg = next((c for c in df.columns if 'ממוצע' in c), None)

        if not col_school: return None

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
        res['אחוז ביצוע מהיעד'] = (res['ממוצע משימות'] / TARGETS[domain] * 100).round(1)
        
        filename_no_ext = os.path.splitext(os.path.basename(filepath))[0]
        # חילוץ התאריך מהשם (לדוגמה מ"מודל מתמטיקה 16.03" חלץ "16.03")
        period_match = re.search(r'(\d{1,2}\.\d{1,2})', filename_no_ext)
        res['תקופה'] = period_match.group(1) if period_match else filename_no_ext
            
        return res

    hist_frames = []
    for f in sorted([f for f in all_files if 'מודל' in f]):
        domain = 'מתמטיקה' if 'מתמטיקה' in f else ('מדעים' if 'מדעים' in f else 'כללי')
        df = process_model_df(f, domain)
        if df is not None: hist_frames.append(df)
        
    df_history = pd.concat(hist_frames, ignore_index=True) if hist_frames else pd.DataFrame()
    df_latest = pd.DataFrame()
    if not df_history.empty:
        df_latest = df_history.sort_values('תקופה').drop_duplicates(subset=['סמל מוסד', 'תחום'], keep='last')

    # 2. עיבוד קבצי התפעולי (מוקדי התערבות - ללא קורסים)
    op_frames = []
    for f in all_files:
        if 'תפעולי' in f:
            domain = 'מתמטיקה' if 'מתמטיקה' in f else ('מדעים' if 'מדעים' in f else 'כללי')
            df_op = safe_read_file(f)
            if df_op.empty: continue
            
            col_school = next((c for c in df_op.columns if 'מוסד' in c), None)
            col_dist = next((c for c in df_op.columns if 'מחוז' in c), None)
            col_sup = next((c for c in df_op.columns if 'מפקח' in c), None)
            col_courses = next((c for c in df_op.columns if 'קורסים שנפתחו' in c), None)
            
            if not col_school or not col_courses: continue
            
            df_op['סמל מוסד'] = df_op[col_school].astype(str).str.extract(r'(\d{6})')[0]
            df_op = df_op.dropna(subset=['סמל מוסד'])
            df_op = df_op[~df_op['סמל מוסד'].isin(excluded_ids)]
            df_op['מוסד_נקי'] = df_op[col_school].astype(str).str.replace(r'^\d{6}\s*-\s*', '', regex=True)
            df_op['קורסים_num'] = pd.to_numeric(df_op[col_courses], errors='coerce').fillna(0)
            
            grouped = df_op.groupby(['סמל מוסד', 'מוסד_נקי']).agg({
                'קורסים_num': 'sum',
                col_dist: 'first',
                col_sup: 'first'
            }).reset_index()
            
            zeros = grouped[grouped['קורסים_num'] == 0].copy()
            if not zeros.empty:
                zeros['תחום'] = domain
                zeros.rename(columns={col_dist: 'מחוז תקשוב', col_sup: 'שם מפקח', 'מוסד_נקי': 'מוסד'}, inplace=True)
                op_frames.append(zeros)

    df_urgent = pd.concat(op_frames, ignore_index=True) if op_frames else pd.DataFrame()
    if not df_urgent.empty:
        df_urgent = df_urgent.drop_duplicates(subset=['סמל מוסד', 'תחום'])

    return df_history, df_latest, df_urgent

df_history, df_latest, df_urgent = load_and_process_data()

# בדיקת התקדמות
def get_progress_status(df_trend):
    if df_trend['תקופה'].nunique() <= 1: return "SINGLE_POINT"
    # בודק אם סכום ההפרשים בין המקסימום למינימום לאורך התקופות הוא 0
    diff_sum = df_trend.groupby('תחום')['ממוצע משימות'].apply(lambda x: x.max() - x.min()).sum()
    return "NO_PROGRESS" if diff_sum == 0 else "PROGRESS"

# ==================== עיצוב ממשק המשתמש ====================

# הטמעת לוגו
logo_base64 = get_image_base64('image_5e4888.png')
if logo_base64:
    st.markdown(f'<img src="data:image/png;base64,{logo_base64}" style="max-height: 80px; float: right; margin-left: 20px;">', unsafe_allow_html=True)

st.title("מערכת בינה עסקית (BI) - משימות מודל")
st.markdown(f"**יעד מחוזי לחודש מרץ: 95% ביצוע | {int(TARGETS['מתמטיקה'])} משימות מתמטיקה | {int(TARGETS['מדעים'])} משימות מדעים**")
st.divider()

if df_latest.empty:
    st.error("🚨 לא נמצאו נתוני מודל תקינים להצגה.")
    st.stop()

valid_districts = df_latest[df_latest['מחוז תקשוב'] != 'לא ידוע']['מחוז תקשוב'].dropna().unique()
district_list = sorted([str(d) for d in valid_districts])
district = st.sidebar.selectbox("בחר/י מחוז (מומלץ: העיר ירושלים):", district_list) if district_list else ""

if not district:
    st.stop()

df_lat_dist = df_latest[df_latest['מחוז תקשוב'] == district]
df_hist_dist = df_history[df_history['מחוז תקשוב'] == district]
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
    df_hist_sup = df_hist_dist[df_hist_dist['שם מפקח'] == supervisor]
    
    # גרף מפקח
    if not df_hist_sup.empty:
        trend_sup = df_hist_sup.groupby(['תקופה', 'תחום']).agg({'ממוצע משימות': 'mean', 'אחוז ביצוע מהיעד': 'mean'}).reset_index()
        trend_sup = trend_sup.sort_values('תקופה')
        
        status = get_progress_status(trend_sup)
        if status == "SINGLE_POINT":
            st.info("קיימת רק נקודת זמן אחת למוסד/מפקח זה. ברגע שתעלי קבצים נוספים בעתיד, יופיע כאן קו מגמה.")
        elif status == "NO_PROGRESS":
            st.warning("לא בוצעה פעילות מה16.03") # הטקסט המדויק למפקח
        else:
            fig_sup = px.line(trend_sup, x='תקופה', y='אחוז ביצוע מהיעד', color='תחום', markers=True,
                              title=f"התקדמות יחסית ליעד (%) - מפקח/ת: {supervisor}",
                              hover_data={'ממוצע משימות': True, 'אחוז ביצוע מהיעד': ':.1f'},
                              color_discrete_map=COLOR_MAP)
            fig_sup.update_traces(mode='lines+markers', line=dict(width=3), marker=dict(size=8))
            fig_sup.update_layout(yaxis_title="אחוז ביצוע מהיעד (%)", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            fig_sup.update_yaxes(matches=None, autorange=True, gridcolor='lightgrey') 
            st.plotly_chart(fig_sup, use_container_width=True)
    
    # טבלאות רמזור
    st.markdown("### סטטוס עדכני - פירוט מוסדות (שיטת הרמזור)")
    def style_row(row, domain):
        val = row['ממוצע משימות']
        if pd.isna(val): return [''] * len(row)
        # צבעי פסטל רכים יותר למראה יוקרתי
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

    # --- רובריקה 3: התערבות דחופה ---
    st.header("מוקדי התערבות דחופים (ללא קורסים)")
    if not df_urg_dist.empty:
        df_urg_sup = df_urg_dist[df_urg_dist['שם מפקח'] == supervisor]
        if not df_urg_sup.empty:
            df_urg_unique = df_urg_sup[['סמל מוסד', 'מוסד', 'תחום']].sort_values('תחום')
            st.error(f"⚠️ שימו לב: נמצאו {len(df_urg_unique)} מוסדות באחריות מפקח/ת זה שבהם טרם נפתחו קורסי מודל כלל (על פי קבצי התפעולי).")
            st.dataframe(df_urg_unique, hide_index=True, use_container_width=True)
        else:
            st.success("לא נמצאו מוסדות ללא קורסים באחריות מפקח/ת זה.")
    else:
        st.success("לא נמצאו נתונים אודות מוסדות ללא קורסים מקבצי התפעולי באזור זה.")

    st.divider()

    # --- רובריקה 4: חקר ביצועים רמת בית ספר ---
    st.header("ניתוח עומק ברמת מוסד (Micro Analysis)")
    if not df_hist_sup.empty:
        schools = sorted(df_hist_sup['מוסד'].dropna().unique())
        selected_school = st.selectbox("בחר/י מוסד לבחינת מגמת שיפור אישית:", schools)
        
        if selected_school:
            df_school = df_hist_sup[df_hist_sup['מוסד'] == selected_school].sort_values('תקופה')
            
            status_sch = get_progress_status(df_school)
            if status_sch == "SINGLE_POINT":
                st.info("קיימת רק נקודת זמן אחת למוסד זה. ברגע שתעלי קבצים נוספים בעתיד, יופיע כאן קו מגמה.")
            elif status_sch == "NO_PROGRESS":
                st.warning("לא בוצעה שום פעילות מ16.03") # הטקסט המדויק לבית ספר
            else:
                fig_sch = px.line(df_school, x='תקופה', y='אחוז ביצוע מהיעד', color='תחום', markers=True,
                                  title=f"מעקב התקדמות יחסית - {selected_school}",
                                  hover_data={'ממוצע משימות': True, 'אחוז ביצוע מהיעד': ':.1f'},
                                  color_discrete_map=COLOR_MAP)
                fig_sch.update_traces(mode='lines+markers', line=dict(width=3), marker=dict(size=8))
                fig_sch.update_layout(yaxis_title="אחוז ביצוע מהיעד (%)", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                fig_sch.update_yaxes(matches=None, autorange=True, gridcolor='lightgrey')
                st.plotly_chart(fig_sch, use_container_width=True)
