import streamlit as st
import pandas as pd
import plotly.express as px
import os
import re

# הגדרת דף וצבעים גלובליים (אדום-מתמטיקה, כחול-מדעים)
st.set_page_config(page_title="דשבורד משימות מודל - אנליטיקס", layout="wide", page_icon="📈")
COLOR_MAP = {'מתמטיקה': '#E63946', 'מדעים': '#1D3557'}

# פונקציית קריאה חכמה ועמידה לשגיאות קידוד
def safe_read_file(filepath):
    if filepath.endswith('.xlsx'):
        try: return pd.read_excel(filepath, engine='openpyxl')
        except: pass
    else:
        for enc in ['utf-8-sig', 'cp1255', 'iso-8859-8', 'utf-8']:
            try: return pd.read_csv(filepath, encoding=enc, dtype=str) # קורא הכל כטקסט למניעת שגיאות
            except: continue
    return pd.DataFrame()

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

    # 2. מנוע עיבוד קבצים מרכזי - חסין לתקלות שורות ריקות/Totals
    def process_df(filepath, domain, file_type):
        df = safe_read_file(filepath)
        if df.empty: return None

        df.columns = df.columns.astype(str).str.strip()
        
        # זיהוי עמודות חכם
        col_school = next((c for c in df.columns if 'מוסד' in c and 'סמל' not in c), None) or next((c for c in df.columns if 'מוסד' in c), None)
        col_dist = next((c for c in df.columns if 'מחוז' in c), None)
        col_sup = next((c for c in df.columns if 'מפקח' in c), None)
        col_avg = next((c for c in df.columns if 'ממוצע' in c), None)

        if not col_school: return None

        # *** הפתרון למחיקת שורות ריקות או Totals ***
        # רק שורה שמכילה בדיוק 6 ספרות תשרוד את הסינון הזה
        df['סמל מוסד'] = df[col_school].astype(str).str.extract(r'(\d{6})')[0]
        df = df.dropna(subset=['סמל מוסד']) # מעיף את Totals או שורות ריקות לחלוטין
        df = df[~df['סמל מוסד'].isin(excluded_ids)] # סינון מוסדות מוחרגים

        df['מוסד_נקי'] = df[col_school].astype(str).str.replace(r'^\d{6}\s*-\s*', '', regex=True)

        res = pd.DataFrame()
        res['סמל מוסד'] = df['סמל מוסד']
        res['מוסד'] = df['מוסד_נקי']
        res['מחוז תקשוב'] = df[col_dist].astype(str).str.strip() if col_dist else 'לא ידוע'
        res['שם מפקח'] = df[col_sup].astype(str).str.strip() if col_sup else 'לא ידוע'
        
        # זיהוי תחום - אם נשלח מבחוץ או אם קיים בקובץ
        col_domain = next((c for c in df.columns if 'תחום' in c), None)
        if domain:
            res['תחום'] = domain
        elif col_domain:
            res['תחום'] = df[col_domain].astype(str).str.strip()
        else:
            res['תחום'] = 'כללי'
        
        if file_type == 'מודל':
            # המרה בטוחה למספרים
            res['ממוצע משימות'] = pd.to_numeric(df[col_avg], errors='coerce').fillna(0).round(2) if col_avg else 0.0
            
            # חילוץ התאריך/מספר משם הקובץ לבניית הגרף
            filename_no_ext = os.path.splitext(os.path.basename(filepath))[0]
            period = re.sub(r'(מודל|ללא קורסים|מתמטיקה|מדעים)', '', filename_no_ext).strip()
            res['תקופה'] = period if period else 'נוכחי'
            
        return res

    # 3. איסוף היסטוריה מלאה מהמודל
    hist_frames = []
    for f in sorted([f for f in all_files if 'מודל' in f and 'מתמטיקה' in f]):
        df = process_df(f, 'מתמטיקה', 'מודל')
        if df is not None: hist_frames.append(df)
        
    for f in sorted([f for f in all_files if 'מודל' in f and 'מדעים' in f]):
        df = process_df(f, 'מדעים', 'מודל')
        if df is not None: hist_frames.append(df)
        
    df_history = pd.concat(hist_frames, ignore_index=True) if hist_frames else pd.DataFrame()
    
    # 4. חילוץ תמונת המצב העדכנית ביותר מתוך ההיסטוריה
    df_latest = pd.DataFrame()
    if not df_history.empty:
        df_latest = df_history.sort_values('תקופה').drop_duplicates(subset=['סמל מוסד', 'תחום'], keep='last')

    # 5. טיפול בקבצי "ללא קורסים"
    nc_frames = []
    nc_files = [f for f in all_files if 'ללא' in f]
    
    # מיון כדי לקחת את הקובץ העדכני ביותר אם יש כמה
    for f in sorted(nc_files):
        # לא מניח מראש מה התחום, נותן לפונקציה לחפש בפנים או שם 'כללי'
        domain_guess = 'מתמטיקה' if 'מתמטיקה' in f else ('מדעים' if 'מדעים' in f else None)
        df_nc = process_df(f, domain_guess, 'ללא_קורסים')
        if df_nc is not None: nc_frames.append(df_nc)
        
    df_no_courses = pd.concat(nc_frames, ignore_index=True) if nc_frames else pd.DataFrame()
    if not df_no_courses.empty:
        df_no_courses = df_no_courses.drop_duplicates(subset=['סמל מוסד', 'תחום'], keep='last')

    return df_history, df_latest, df_no_courses

df_history, df_latest, df_no_courses = load_and_process_data()

# ==================== בניית ממשק המשתמש ====================
st.title("📈 מערכת בינה עסקית (BI) - משימות מודל")
st.markdown("### 🎯 יעד מחוזי לחודש מרץ: 95% ביצוע | 17 משימות מתמטיקה | 8 משימות מדעים")
st.divider()

if df_latest.empty:
    st.error("🚨 לא נמצאו קבצי מודל. ודאי שהעלית קבצים בשם 'מודל מתמטיקה [תאריך].csv' וכו'.")
    st.stop()

# יצירת רשימת מחוזות נקייה (ללא NaN או 'לא ידוע' אם אפשר)
valid_districts = df_latest[df_latest['מחוז תקשוב'] != 'לא ידוע']['מחוז תקשוב'].dropna().unique()
district_list = sorted([str(d) for d in valid_districts])
district = st.sidebar.selectbox("בחר/י מחוז:", district_list) if district_list else ""

if not district:
    st.info("אנא בחר/י מחוז מהתפריט בצד.")
    st.stop()

df_lat_dist = df_latest[df_latest['מחוז תקשוב'] == district]
df_hist_dist = df_history[df_history['מחוז תקשוב'] == district]
df_nc_dist = df_no_courses[df_no_courses['מחוז תקשוב'] == district] if not df_no_courses.empty else pd.DataFrame()

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
    
    # גרף מגמות - רמת מפקח
    if not df_hist_sup.empty and df_hist_sup['תקופה'].nunique() > 1:
        trend_sup = df_hist_sup.groupby(['תקופה', 'תחום'])['ממוצע משימות'].mean().reset_index()
        # סידור התקופות לפי סדר אלפביתי (שעובד מצוין עם תאריכים כמו 16.03, 18.03)
        trend_sup = trend_sup.sort_values('תקופה') 
        
        fig_sup = px.line(trend_sup, x='תקופה', y='ממוצע משימות', color='תחום', markers=True,
                          title=f"📊 מגמת התקדמות - ממוצע כלל בתי הספר (מפקח/ת: {supervisor})",
                          color_discrete_map=COLOR_MAP)
        fig_sup.update_layout(xaxis_title="תאריך/תקופה", yaxis_title="ממוצע משימות")
        st.plotly_chart(fig_sup, use_container_width=True)
    
    # טבלאות רמזור
    st.markdown("### 📋 סטטוס עדכני - פירוט מוסדות (שיטת הרמזור)")
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
    st.header("🚨 מוקדי התערבות דחופים (ללא קורסים)")
    if not df_nc_dist.empty:
        df_nc_sup = df_nc_dist[df_nc_dist['שם מפקח'] == supervisor]
        if not df_nc_sup.empty:
            df_nc_unique = df_nc_sup[['סמל מוסד', 'מוסד', 'תחום']].drop_duplicates().sort_values('תחום')
            st.warning(f"המפקח/ת {supervisor} אחראי/ת על {len(df_nc_unique)} מוסדות שטרם פתחו קורסי מודל.")
            st.dataframe(df_nc_unique, hide_index=True, use_container_width=True)
        else:
            st.success("אין מוסדות ללא קורסים באחריות מפקח זה. מצוין!")
    else:
        st.info("לא נמצאו נתונים מקובץ 'ללא קורסים'.")

    st.divider()

    # --- רובריקה 4: חקר ביצועים רמת בית ספר ---
    st.header("🏫 ניתוח עומק ברמת מוסד (Micro Analysis)")
    if not df_hist_sup.empty:
        schools = sorted(df_hist_sup['מוסד'].dropna().unique())
        selected_school = st.selectbox("בחר/י מוסד לבחינת מגמת שיפור אישית:", schools)
        
        if selected_school:
            df_school = df_hist_sup[df_hist_sup['מוסד'] == selected_school].sort_values('תקופה')
            if df_school['תקופה'].nunique() > 1:
                fig_sch = px.line(df_school, x='תקופה', y='ממוצע משימות', color='תחום', markers=True,
                                  title=f"📈 מעקב התקדמות - {selected_school}",
                                  color_discrete_map=COLOR_MAP)
                fig_sch.update_layout(xaxis_title="תאריך/תקופה", yaxis_title="ממוצע משימות")
                st.plotly_chart(fig_sch, use_container_width=True)
            else:
                st.info("💡 קיימת רק נקודת זמן אחת למוסד זה. ברגע שתעלי קבצים נוספים בעתיד, יופיע כאן קו מגמה.")
