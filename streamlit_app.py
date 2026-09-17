import streamlit as st
import pandas as pd
import plotly.express as px
import io
import os

# -----------------------------------------------------------------------------
# 1. PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="MOCO - Data Validation 2026",
    page_icon="⛏️", 
    layout="wide"
)

# Header Utama dengan Logo
col_logo, col_title = st.columns([1, 6])

with col_logo:
    if os.path.exists("BN.png"):
        st.image("BN.png", width=90)
    else:
        st.write("📌 *(Upload logo.png ke Git)*")

with col_title:
    st.title("BIMA NUSA INT MOCO - Mining Operational")
    st.caption("MOHH Anomaly & Latensi Input User (OB & COAL)")

# -----------------------------------------------------------------------------
# 2. HELPER FUNCTIONS FOR EXCEL PARSING
# -----------------------------------------------------------------------------
def process_aturan_1_summary(file):
    """
    Aturan 1: Membaca Summary Productivity
    - Filter OB & COAL
    - Hitung MOHH = EWH + STB + BD
    - Filter Anomali MOHH > 24 jam (dengan toleransi pembulatan float)
    - Ambil kolom: DATE, SITE, UNITNO, WORKGROUP, EWH, STB, BD, MOHH, USERNAMES DAY, USERNAMES N
    """
    try:
        df_raw = pd.read_excel(file, header=None)
        
        # Cari baris header utama
        header_row = 1
        for idx, row in df_raw.head(15).iterrows():
            row_str = [str(x).upper() for x in row.values if pd.notna(x)]
            if any("WORKGROUP" in item for item in row_str):
                header_row = idx
                break

        # Read dengan MultiIndex header (2 baris)
        df = pd.read_excel(file, header=[header_row, header_row + 1])
        
        # Flatten nama kolom
        cols = []
        for c in df.columns:
            top = str(c[0]).strip() if pd.notna(c[0]) and not str(c[0]).startswith("Unnamed") else ""
            bot = str(c[1]).strip() if pd.notna(c[1]) and not str(c[1]).startswith("Unnamed") else ""
            if top and bot:
                cols.append(f"{top}_{bot}".upper())
            elif top:
                cols.append(top.upper())
            elif bot:
                cols.append(bot.upper())
            else:
                cols.append(f"COL_{len(cols)}")
        df.columns = cols
        
        # Mapping nama kolom standar
        col_map = {}
        for c in df.columns:
            if "WORKGROUP" in c: col_map[c] = "WORKGROUP"
            elif "SITE" in c: col_map[c] = "SITE"
            elif "UNIT" in c or "CN" in c or "EQUIPMENT" in c: col_map[c] = "UNITNO"
            elif "DATE" in c or "TANGGAL" in c: col_map[c] = "DATE"
            elif c.endswith("_EWH") or c == "EWH": col_map[c] = "EWH"
            elif c.endswith("_STB") or c == "STB": col_map[c] = "STB"
            elif c.endswith("_BD") or c == "BD": col_map[c] = "BD"
            elif c.endswith("_MOHH") or c == "MOHH": col_map[c] = "MOHH"
            elif "USER" in c and ("DAY" in c or "_D" in c): col_map[c] = "USERNAMES DAY"
            elif "USER" in c and ("NIGHT" in c or "_N" in c): col_map[c] = "USERNAMES N"

        df = df.rename(columns=col_map)
        
        # Format kolom DATE agar tidak menampilkan jam
        if "DATE" in df.columns:
            df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce").dt.strftime("%Y-%m-%d")

        # Filter Workgroup OB & COAL
        if "WORKGROUP" in df.columns:
            df["WORKGROUP_STR"] = df["WORKGROUP"].astype(str).str.strip().str.upper()
            df = df[df["WORKGROUP_STR"].str.contains("OB|COAL|OVERBURDEN", regex=True, na=False)].copy()
            df = df.drop(columns=["WORKGROUP_STR"])
            
        # Pastikan kolom numerik untuk kalkulasi
        for num_col in ["EWH", "STB", "BD", "MOHH"]:
            if num_col in df.columns:
                df[num_col] = pd.to_numeric(df[num_col], errors="coerce").fillna(0)
            else:
                df[num_col] = 0.0

        # Hitung MOHH berdasarkan kalkulasi murni EWH + STB + BD
        df["MOHH"] = (df["EWH"] + df["STB"] + df["BD"]).round(2)
        
        # Filter Anomali: Murni di atas 24 Jam
        df_anomali = df[df["MOHH"] > 24.001].copy()
        
        # Pilih kolom sesuai Aturan 1
        target_cols = ["DATE", "SITE", "UNITNO", "WORKGROUP", "EWH", "STB", "BD", "MOHH", "USERNAMES DAY", "USERNAMES N"]
        existing_target = [c for c in target_cols if c in df_anomali.columns]
        
        return df_anomali[existing_target], df[[c for c in target_cols if c in df.columns]], None
    except Exception as e:
        return None, None, str(e)


def process_aturan_2_input_time(file):
    """
    Aturan 2: Membaca Input Time (Time Entry)
    - Menangani header bertingkat
    - Filter Workgroup OB & COAL
    - Format kolom Date agar tanpa jam
    - Menyaring data yang pada kolom Dev (Hours text) mengandung kata 'jam'
    """
    try:
        df_raw = pd.read_excel(file, header=None)
        
        # 1. Cari baris yang mengandung 'DEV (HOURS TEXT)' atau 'INPUT'
        sub_header_row = None
        for idx, row in df_raw.head(15).iterrows():
            row_str = [str(x).upper() for x in row.values if pd.notna(x)]
            if any("DEV (HOURS TEXT)" in item or "DEV (HOURS)" in item for item in row_str):
                sub_header_row = idx
                break

        if sub_header_row is not None:
            df = pd.read_excel(file, header=sub_header_row)
        else:
            df = pd.read_excel(file)

        # Bersihkan nama kolom dari whitespace
        df.columns = [str(c).strip() for c in df.columns]

        # 2. Filter Workgroup OB & COAL
        wg_col = None
        for c in df.columns:
            if "WORKGROUP" in str(c).upper():
                wg_col = c
                break
        
        if wg_col:
            mask_wg = df[wg_col].astype(str).str.strip().str.upper().str.contains("OB|COAL|OVERBURDEN", regex=True, na=False)
            df = df[mask_wg].copy()

        # 3. Format kolom Date agar bersih tanpa jam (00:00:00)
        date_col = None
        for c in df.columns:
            if "DATE" in str(c).upper() or "TANGGAL" in str(c).upper():
                date_col = c
                break
        
        if date_col:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce").dt.strftime("%Y-%m-%d")

        # 4. Cari kolom 'Dev (Hours text)'
        dev_txt_col = None
        for c in df.columns:
            c_upper = str(c).upper()
            if "DEV" in c_upper and "TEXT" in c_upper:
                dev_txt_col = c
                break

        if not dev_txt_col:
            dev_cols = [c for c in df.columns if "DEV" in str(c).upper()]
            if len(dev_cols) >= 2:
                dev_txt_col = dev_cols[1]
            elif len(dev_cols) == 1:
                dev_txt_col = dev_cols[0]

        if not dev_txt_col:
            return None, None, "Kolom 'Dev (Hours text)' tidak ditemukan di dalam file Excel."

        # 5. Filter data yang mengandung kata 'jam'
        mask_delay = df[dev_txt_col].astype(str).str.lower().str.contains("jam", na=False)
        df_delay = df[mask_delay].copy()

        return df_delay, df, None
    except Exception as e:
        return None, None, str(e)


def convert_df_to_excel(df):
    """Helper untuk fitur ekspor laporan ke format Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Audit Report')
    return output.getvalue()

# -----------------------------------------------------------------------------
# 3. SIDEBAR UPLOAD & BRANDING
# -----------------------------------------------------------------------------
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    
    st.header("📁 Submit Daily Operational Files")
    file_prod = st.file_uploader("Summary Productivity", type=["xlsx", "xls"])
    file_time = st.file_uploader("Input Time / Time Entry", type=["xlsx", "xls"])

# -----------------------------------------------------------------------------
# 4. MAIN AUDIT DISPLAY
# -----------------------------------------------------------------------------
if file_prod is not None and file_time is not None:
    df_m_anomali, df_m_all, err1 = process_aturan_1_summary(file_prod)
    df_t_delay, df_t_all, err2 = process_aturan_2_input_time(file_time)
    
    if err1:
        st.error(f"Error (Summary Productivity): {err1}")
    elif err2:
        st.error(f"Error (Input Time): {err2}")
    else:
        st.success("✅ File Berhasil Diproses! Menampilkan Hasil...")
        
        tab1, tab2 = st.tabs(["Anomali MOHH (>24 Jam)", "Keterlambatan Input User"])
        
        # TAB 1: ANOMALI MOHH
        with tab1:
            st.subheader("Tabel Anomali MOHH (> 24 Jam)")
            st.caption("Menampilkan unit kerja OB & COAL yang total MOHH-nya melebihi 24 jam dalam 1 hari operasional.")
            
            col_m1, col_m2, col_m3 = st.columns(3)
            col_m1.metric("Total Anomali MOHH Found", f"{len(df_m_anomali)} Record")
            col_m2.metric("Total Data OB & COAL Evaluated", f"{len(df_m_all)} Record")
            pct_m = (len(df_m_anomali) / len(df_m_all) * 100) if len(df_m_all) > 0 else 0
            col_m3.metric("Rasio Anomali", f"{pct_m:.1f}%")
            
            st.markdown("---")
            
            # Visualisasi Aturan 1
            if len(df_m_all) > 0:
                cg1, cg2 = st.columns(2)
                with cg1:
                    st.markdown("##### 📊 Komposisi Rata-Rata Jam Operasional")
                    avg_hours = pd.DataFrame({
                        'Status': ['EWH', 'STB', 'BD'],
                        'Rata-Rata Jam': [df_m_all['EWH'].mean(), df_m_all['STB'].mean(), df_m_all['BD'].mean()]
                    })
                    fig_pie = px.pie(avg_hours, values='Rata-Rata Jam', names='Status', hole=0.4,
                                     color_discrete_sequence=['#22c55e', '#eab308', '#ef4444'])
                    fig_pie.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=280)
                    st.plotly_chart(fig_pie, use_container_width=True)
                
                with cg2:
                    st.markdown("##### 🏗️ Anomali per Site")
                    if len(df_m_anomali) > 0 and 'SITE' in df_m_anomali.columns:
                        site_cnt = df_m_anomali['SITE'].value_counts().reset_index()
                        site_cnt.columns = ['SITE', 'Jumlah']
                        fig_bar = px.bar(site_cnt, x='SITE', y='Jumlah', color='Jumlah', color_continuous_scale='Reds')
                        fig_bar.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=280)
                        st.plotly_chart(fig_bar, use_container_width=True)
                    else:
                        st.info("Tidak ada grafik site (0 anomali).")

            st.markdown("### Detail Data Anomali")
            if len(df_m_anomali) > 0:
                st.dataframe(df_m_anomali, use_container_width=True)
                st.download_button(
                    label="📥 Unduh Data Anomali MOHH (.xlsx)",
                    data=convert_df_to_excel(df_m_anomali),
                    file_name="Audit_MOHH_Anomali.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Tidak ditemukan anomali MOHH > 24 jam pada file ini.")

        # TAB 2: KETERLAMBATAN INPUT USER
        with tab2:
            st.subheader("Tabel Input (> 1 Jam)")
            st.caption("Menampilkan log input dispatch / CCR")
            
            col_t1, col_t2, col_t3 = st.columns(3)
            col_t1.metric("Total Terlambat (>1 Jam)", f"{len(df_t_delay)} Record")
            col_t2.metric("Total Time Entry Evaluated", f"{len(df_t_all)} Record")
            pct_t = (len(df_t_delay) / len(df_t_all) * 100) if len(df_t_all) > 0 else 0
            col_t3.metric("Rasio Keterlambatan", f"{pct_t:.1f}%")

            st.markdown("---")
            
            # Visualisasi Aturan 2
            if len(df_t_delay) > 0:
                cg3, cg4 = st.columns(2)
                with cg3:
                    st.markdown("##### 👤 Top User Delay Input")
                    user_cols = [c for c in df_t_delay.columns if "USER" in str(c).upper()]
                    if user_cols:
                        top_users = df_t_delay[user_cols[0]].value_counts().head(5).reset_index()
                        top_users.columns = ['User', 'Jumlah Log']
                        fig_u = px.bar(top_users, y='User', x='Jumlah Log', orientation='h',
                                       color='Jumlah Log', color_continuous_scale='Oranges')
                        fig_u.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(t=10, b=10, l=10, r=10), height=280)
                        st.plotly_chart(fig_u, use_container_width=True)
                
                with cg4:
                    st.markdown("##### Visual per Shift")
                    shift_cols = [c for c in df_t_delay.columns if "SHIFT" in str(c).upper()]
                    if shift_cols:
                        shift_cnt = df_t_delay[shift_cols[0]].value_counts().reset_index()
                        shift_cnt.columns = ['Shift', 'Jumlah']
                        fig_s = px.pie(shift_cnt, values='Jumlah', names='Shift', color_discrete_sequence=px.colors.qualitative.Set2)
                        fig_s.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=280)
                        st.plotly_chart(fig_s, use_container_width=True)

            st.markdown("### Detail Data Delay")
            if len(df_t_delay) > 0:
                st.dataframe(df_t_delay, use_container_width=True)
                st.download_button(
                    label="📥 Unduh Data Keterlambatan Input (.xlsx)",
                    data=convert_df_to_excel(df_t_delay),
                    file_name="Audit_Keterlambatan_Input.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Tidak ditemukan keterlambatan input user > 1 jam pada file ini.")

else:
    st.info("👋 Silakan unggah **file Excel di sidebar kiri** untuk memulai.")
    
    st.markdown("---")
    col_i1, col_i2 = st.columns(2)
    with col_i1:
        st.markdown("""
        #### 🚨 Parameter MOHH
        """)
    with col_i2:
        st.markdown("""
        #### ⏱️ Parameter Latensi Input
        """)
