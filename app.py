import streamlit as st
import pandas as pd
import plotly.express as px
import os
import re

st.set_page_config(page_title="ישראל ראלית - מחוז והעיר ירושלים", layout="wide", page_icon="🇮🇱")
COLOR_MAP = {'מתמטיקה': '#E63946', 'מדעים': '#1D3557'}
TARGETS = {'מתמטיקה': 17.0, 'מדעים': 8.0}

def safe_read_file(filepath):
    if filepath.endswith('.xlsx'):
        try: return pd.read_excel(filepath, engine='openpyxl')
        except: pass
    else:
        for enc in ['utf-8-sig', 'cp1255', 'iso-8859-8', 'utf-8']:
            try: return pd.read_csv(filepath, encoding=enc, dtype=str)
            except: continue
    return pd.DataFrame()

@st.cache_data
def load_and_process_data():
    all_files = os.listdir('.')
    
    exc_file = next((f for f in all_files if 'להחרגה' in f), None)
    excluded_ids = []
    if exc_file:
        df_ex = safe_read_file(exc_file)
        if not df_ex.empty:
            for col in df_ex.columns:
                extracted = df_ex[col].astype(str).str.extract(r'(\d{6})')[0].dropna().tolist()
                if extracted: excluded_ids.extend(extracted)

    # 1. עיבוד קבצי המודל (היסטוריה ותמונת מצב)
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
        
        # חישוב אחוז ביצוע ליצירת גרף יחסי
        target = TARGETS[domain]
        res['אחוז ביצוע מהיעד'] = (res['ממוצע משימות'] / target * 100).round(1)
        
        filename_no_ext = os.path.splitext(os.path.basename(filepath))[0]
        period = re.sub(r'(מודל|מתמטיקה|מדעים)', '', filename_no_ext).strip()
        res['תקופה'] = period if period else 'נוכחי'
            
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

    # 2. עיבוד קבצי התפעולי (למציאת מוסדות ללא קורסים)
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
            
            # קיבוץ לפי בית ספר - האם פתח קורסים בכלל הכיתות?
            grouped = df_op.groupby(['סמל מוסד', 'מוסד_נקי']).agg({
                'קורסים_num': 'sum',
                col_dist: 'first',
                col_sup: 'first'
            }).reset_index()
            
            # שואבים רק מוסדות שסך הקורסים שלהם הוא 0
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

# ==================== פונקציית בדיקת התקדמות ====================
def has_progress(df_trend):
    """ בודקת אם יש שינוי בנתונים בין התקופות השונות """
    if df_trend['תקופה'].nunique() <= 1: return True
    # חישוב ההפרש בין המקסימום למינימום של כל תחום
    diff_sum = df_trend.groupby('תחום')['ממוצע משימות'].apply(lambda x: x.max() - x.min()).sum()
    return diff_sum > 0

# ==================== ממשק המשתמש ====================
st.title("🇮🇱 ישראל ראלית - מחוז והעיר ירושלים")
st.markdown("#### משימות מודל - כיתה ז'")
st.divider()

if df_latest.empty:
    st.error("🚨 לא נמצאו קבצי מודל תקינים.")
    st.stop()

valid_districts = df_latest[df_latest['מחוז תקשוב'] != 'לא ידוע']['מחוז תקשוב'].dropna().unique()
district_list = sorted([str(d) for d in valid_districts])
district = st.sidebar.selectbox("בחר/י מחוז:", district_list) if district_list else ""

if not district:
    st.info("אנא בחר/י מחוז מהתפריט בצד.")
    st.stop()

df_lat_dist = df_latest[df_latest['מחוז תקשוב'] == district]
df_hist_dist = df_history[df_history['מחוז תקשוב'] == district]
df_urg_dist = df_urgent[df_urgent['מחוז תקשוב'] == district] if not df_urgent.empty else pd.DataFrame()

# --- רובריקה 1: מאקרו מחוז ---
st.header(f"📌 תמונת מצב עדכנית - מחוז {district}")
col1, col2 = st.columns(2)
with col1:
    math_avg = df_lat_dist[df_lat_dist['תחום'] == 'מתמטיקה']['ממוצע משימות'].mean()
    st.markdown(f"<h3 style='color:{COLOR_MAP['מתמטיקה']};'>📐 מתמטיקה</h3>", unsafe_allow_html=True)
    st.metric("ממוצע משימות רשותי", f"{math_avg:.1f}" if pd.notna(math_avg) else "0.0")
with col2:
    sci_avg = df_lat_dist[df_lat_dist['תחום'] == 'מדעים']['ממוצע משימות'].mean()
    st.markdown(f"<h3 style='color:{COLOR_MAP['מדעים']};'>🔬 מדעים</h3>", unsafe_allow_html=True)
    st.metric("ממוצע משימות רשותי", f"{sci_avg:.1f}" if pd.notna(sci_avg) else "0.0")
st.divider()

# --- רובריקה 2: מפקחים ורמזור ---
st.header("👥 ניתוח ביצועים לפי מפקח/ת")
valid_sups = df_lat_dist[df_lat_dist['שם מפקח'] != 'לא ידוע']['שם מפקח'].dropna().unique()
supervisors = sorted([str(s) for s in valid_sups])
supervisor = st.selectbox("בחר/י מפקח/ת:", supervisors) if supervisors else ""

if supervisor:
    df_lat_sup = df_lat_dist[df_lat_dist['שם מפקח'] == supervisor]
    df_hist_sup = df_hist_dist[df_hist_dist['שם מפקח'] == supervisor]
    
    # גרף מפקח (עם נירמול ליחסיות ובדיקת קיפאון)
    if not df_hist_sup.empty:
        trend_sup = df_hist_sup.groupby(['תקופה', 'תחום']).agg({'ממוצע משימות': 'mean', 'אחוז ביצוע מהיעד': 'mean'}).reset_index()
        trend_sup = trend_sup.sort_values('תקופה')
        
        if trend_sup['תקופה'].nunique() > 1 and not has_progress(trend_sup):
            st.warning("לא בוצעה פעילות מה16.03")
        else:
            fig_sup = px.line(trend_sup, x='תקופה', y='אחוז ביצוע מהיעד', color='תחום', markers=True,
                              title=f"📊 התקדמות יחסית ליעד (מפקח/ת: {supervisor})",
                              hover_data=['ממוצע משימות'],
                              color_discrete_map=COLOR_MAP)
            fig_sup.update_traces(mode='lines+markers')
            fig_sup.update_layout(yaxis_title="אחוז ביצוע מהיעד (%)")
            fig_sup.update_yaxes(matches=None, autorange=True) # משחרר את הציר כדי שהשיפועים יבלטו לעין
            st.plotly_chart(fig_sup, use_container_width=True)
    
    # טבלאות רמזור
    st.markdown("### 📋 סטטוס עדכני - פירוט מוסדות")
    def style_row(row, domain):
        val = row['ממוצע משימות']
        if pd.isna(val): return [''] * len(row)
        if domain == 'מתמטיקה': color = '#ffcccc' if val < 5 else ('#ffffcc' if val < 12 else '#ccffcc')
        else: color = '#ffcccc' if val < 2 else ('#ffffcc' if val < 6 else '#ccffcc')
        return [f'background-color: {color}; color: black;' if col in ['מוסד', 'ממוצע משימות'] else '' for col in row.index]

    t1, t2 = st.tabs(["📐 מתמטיקה", "🔬 מדעים"])
    with t1:
        d_m = df_lat_sup[df_lat_sup['תחום'] == 'מתמטיקה'][['סמל מוסד', 'מוסד', 'ממוצע משימות']].sort_values('ממוצע משימות')
        if not d_m.empty: st.dataframe(d_m.style.apply(style_row, domain='מתמטיקה', axis=1), use_container_width=True, hide_index=True)
    with t2:
        d_s = df_lat_sup[df_lat_sup['תחום'] == 'מדעים'][['סמל מוסד', 'מוסד', 'ממוצע משימות']].sort_values('ממוצע משימות')
        if not d_s.empty: st.dataframe(d_s.style.apply(style_row, domain='מדעים', axis=1), use_container_width=True, hide_index=True)

    st.divider()

    # --- רובריקה 3: התערבות דחופה ---
    st.header("🚨 מוקדי התערבות דחופים")
    if not df_urg_dist.empty:
        df_urg_sup = df_urg_dist[df_urg_dist['שם מפקח'] == supervisor]
        if not df_urg_sup.empty:
            df_urg_unique = df_urg_sup[['סמל מוסד', 'מוסד', 'תחום']].sort_values('תחום')
            st.warning(f"המפקח/ת {supervisor} אחראי/ת על {len(df_urg_unique)} מוסדות שטרם פתחו קורסי מודל.")
            st.dataframe(df_urg_unique, hide_index=True, use_container_width=True)
        else:
            st.success("אין מוסדות ללא קורסים באחריות מפקח זה. מצוין!")
    else:
        st.info("קובץ תפעולי לא הניב תוצאות חריגות (כל בתי הספר פתחו קורסים).")

    st.divider()

    # --- רובריקה 4: חקר ביצועים רמת בית ספר ---
    st.header("🏫 מגמות ברמת מוסד")
    if not df_hist_sup.empty:
        schools = sorted(df_hist_sup['מוסד'].dropna().unique())
        selected_school = st.selectbox("בחר/י מוסד לבחינת מגמת שיפור אישית:", schools)
        
        if selected_school:
            df_school = df_hist_sup[df_hist_sup['מוסד'] == selected_school].sort_values('תקופה')
            
            if df_school['תקופה'].nunique() > 1 and not has_progress(df_school):
                st.warning("לא בוצעה שום פעילות מ16.03")
            else:
                fig_sch = px.line(df_school, x='תקופה', y='אחוז ביצוע מהיעד', color='תחום', markers=True,
                                  title=f"📈 מעקב התקדמות יחסית - {selected_school}",
                                  hover_data=['ממוצע משימות'],
                                  color_discrete_map=COLOR_MAP)
                fig_sch.update_traces(mode='lines+markers')
                fig_sch.update_layout(yaxis_title="אחוז ביצוע מהיעד (%)")
                fig_sch.update_yaxes(matches=None, autorange=True) # משחרר את הציר כדי להבליט שינויים
                st.plotly_chart(fig_sch, use_container_width=True)
