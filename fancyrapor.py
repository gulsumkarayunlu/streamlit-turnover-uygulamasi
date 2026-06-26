import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(
    page_title="Koton Mağazacılık - Turnover Analizi",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .stApp { background-color: #0d1b3e; }
    section[data-testid="stSidebar"] { background-color: #f0f2f6; }
    section[data-testid="stSidebar"] * { color: #1a1a2e !important; }
    div[data-testid="metric-container"] { background-color: #1a2f5e; border-radius: 8px; padding: 12px; }
    h1, h2, h3, h4, p { color: white !important; }
    .stTabs [data-baseweb="tab"] { color: white; }
    .stTabs [data-baseweb="tab-list"] { background-color: #0d1b3e; }
    .magaza-kart {
        background-color: #1a2f5e;
        border-left: 4px solid #e05c5c;
        border-radius: 8px;
        padding: 20px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def veri_yukle():
    df = pd.read_excel("Mayıs_TO.xlsx", sheet_name="TO-Yeni Format MTD-YTD", header=2)
    df.columns = [
        "Masraf_Kodu", "Yeni_Masraf_Kodu", "Magaza", "BM", "PM", "HRBP", "Segment", "Il",
        "Oca26_Cikis", "Oca26_Ort", "Oca26_TO",
        "Sub26_Cikis", "Sub26_Ort", "Sub26_TO",
        "Mar26_Cikis", "Mar26_Ort", "Mar26_TO",
        "Nis26_Cikis", "Nis26_Ort", "Nis26_TO",
        "Mayis26_Cikis", "Mayis26_Ort", "Mayis26_TO",
        "YTD26_Cikis", "YTD26_Ort", "YTD26_TO",
        "Oca25_Cikis", "Oca25_Ort", "Oca25_TO",
        "Sub25_Cikis", "Sub25_Ort", "Sub25_TO",
        "Mar25_Cikis", "Mar25_Ort", "Mar25_TO",
        "Nis25_Cikis", "Nis25_Ort", "Nis25_TO",
        "Mayis25_Cikis", "Mayis25_Ort", "Mayis25_TO",
        "YTD25_Cikis", "YTD25_Ort", "YTD25_TO"
    ]
    df = df.dropna(subset=["Masraf_Kodu"])
    df = df[df["Masraf_Kodu"].astype(str).str.strip() != "Toplam"]
    df = df.reset_index(drop=True)
    to_cols = ["Oca26_TO","Sub26_TO","Mar26_TO","Nis26_TO","Mayis26_TO","YTD26_TO",
               "Oca25_TO","Sub25_TO","Mar25_TO","Nis25_TO","Mayis25_TO","YTD25_TO"]
    for col in to_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce") * 100
    return df

df = veri_yukle()

# Ay → sütun eşleştirmesi
ay_map = {
    "Ocak":  ("Oca26_Cikis", "Oca26_Ort", "Oca26_TO", "Oca25_TO"),
    "Şubat": ("Sub26_Cikis", "Sub26_Ort", "Sub26_TO", "Sub25_TO"),
    "Mart":  ("Mar26_Cikis", "Mar26_Ort", "Mar26_TO", "Mar25_TO"),
    "Nisan": ("Nis26_Cikis", "Nis26_Ort", "Nis26_TO", "Nis25_TO"),
    "Mayıs": ("Mayis26_Cikis", "Mayis26_Ort", "Mayis26_TO", "Mayis25_TO"),
}

def genel_to(df_f, cikis_col, ort_col):
    c = pd.to_numeric(df_f[cikis_col], errors="coerce").sum()
    o = pd.to_numeric(df_f[ort_col], errors="coerce").sum()
    return round((c / o) * 100, 1) if o > 0 else 0.0

def risk_etiketi(to):
    if pd.isna(to): return "—", "⚪"
    if to > 20: return "Yüksek", "🔴"
    if to > 10: return "Normal", "🟡"
    return "Düşük", "🟢"

# ── SOL PANEL ──────────────────────────────────
with st.sidebar:
    try:
        st.image("koton_siyah.png", width=180)
    except:
            st.markdown("**KOTON**")
    st.markdown("**Karne Verisi - Turnover**")
    st.divider()
    st.markdown("📆 **2026 Yılı Ayları**")
    sec_ay = st.selectbox("Ay", ["Tümü (Nisan & Mayıs)", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs"], label_visibility="collapsed")


    bm_listesi  = ["Tümü"] + sorted(df["BM"].dropna().unique().tolist())
    pm_listesi  = ["Tümü"] + sorted(df["PM"].dropna().unique().tolist())
    hrbp_listesi = ["Tümü"] + sorted(df["HRBP"].dropna().unique().tolist())

    st.markdown("👤 **Bölge Müdürü (BM)**")
    sec_bm   = st.selectbox("BM",   bm_listesi,   label_visibility="collapsed")
    st.markdown("🏪 **Perakende Müdürü (PM)**")
    sec_pm   = st.selectbox("PM",   pm_listesi,   label_visibility="collapsed")
    st.markdown("🤝 **İK İş Ortağı (HRBP)**")
    sec_hrbp = st.selectbox("HRBP", hrbp_listesi, label_visibility="collapsed")


# ── FİLTRELEME ────────────────────────────────
df_f = df.copy()
if sec_bm   != "Tümü": df_f = df_f[df_f["BM"]   == sec_bm]
if sec_pm   != "Tümü": df_f = df_f[df_f["PM"]   == sec_pm]
if sec_hrbp != "Tümü": df_f = df_f[df_f["HRBP"] == sec_hrbp]

# Seçili aya göre hangi sütunları göstereceğimizi belirle
if sec_ay == "Tümü (Nisan & Mayıs)":
    cikis26, ort26, to26_col, to25_col = "Nis26_Cikis", "Nis26_Ort", "Nis26_TO", "Nis25_TO"
    cikis25, ort25 = "Nis25_Cikis", "Nis25_Ort"
    ana_baslik26, ana_baslik25 = "2026 Nisan TO", "2025 Nisan TO"
    ytd26_col, ytd25_col = "YTD26_TO", "YTD25_TO"
    ytd26_cikis, ytd26_ort = "YTD26_Cikis", "YTD26_Ort"
    ytd25_cikis, ytd25_ort = "YTD25_Cikis", "YTD25_Ort"
    may26_col, may25_col = "Mayis26_TO", "Mayis25_TO"
    may26_cikis, may26_ort = "Mayis26_Cikis", "Mayis26_Ort"
    may25_cikis, may25_ort = "Mayis25_Cikis", "Mayis25_Ort"
    goster_ytd = True
    goster_mayis = True
else:
    cikis26, ort26, to26_col, to25_col = ay_map[sec_ay]
    cikis25 = cikis26.replace("26", "25")
    ort25   = ort26.replace("26", "25")
    ana_baslik26 = f"2026 {sec_ay} TO"
    ana_baslik25 = f"2025 {sec_ay} TO"
    ytd26_col, ytd25_col = "YTD26_TO", "YTD25_TO"
    ytd26_cikis, ytd26_ort = "YTD26_Cikis", "YTD26_Ort"
    ytd25_cikis, ytd25_ort = "YTD25_Cikis", "YTD25_Ort"
    may26_col = may25_col = None
    goster_ytd = True
    goster_mayis = False

# KPI değerleri
kpi26     = genel_to(df_f, cikis26, ort26)
kpi25     = genel_to(df_f, cikis25, ort25)
kpi_ytd26 = genel_to(df_f, ytd26_cikis, ytd26_ort)
kpi_ytd25 = genel_to(df_f, ytd25_cikis, ytd25_ort)
kpi_may26 = genel_to(df_f, "Mayis26_Cikis", "Mayis26_Ort") if goster_mayis else None
kpi_may25 = genel_to(df_f, "Mayis25_Cikis", "Mayis25_Ort") if goster_mayis else None

filtre_adi = []
if sec_bm   != "Tümü": filtre_adi.append(f"BM: {sec_bm}")
if sec_pm   != "Tümü": filtre_adi.append(f"PM: {sec_pm}")
if sec_hrbp != "Tümü": filtre_adi.append(f"HRBP: {sec_hrbp}")
if sec_ay   != "Tümü (Nisan & Mayıs)": filtre_adi.append(f"Ay: {sec_ay}")
alt_yazi = " | ".join(filtre_adi) if filtre_adi else "Tüm Türkiye Mağazalarının Turnover Detay Verisi Aşağıda Görüntülenebilmektedir."

# ── ANA BAŞLIK ────────────────────────────────
col_logo, col_baslik = st.columns([1, 6])
with col_logo:
    try:
        st.image("koton_logo.png", width=180)
    except:
        st.markdown("**KOTON**")
with col_baslik:
    st.markdown("<h1 style='color:white; margin-bottom:0;'>Mağazacılık Turnover Analiz Paneli</h1>", unsafe_allow_html=True)
    st.markdown(f"<p style='color:#aaaacc; margin-top:4px;'>{alt_yazi}</p>", unsafe_allow_html=True)

st.divider()

# ── SEKMELER ──────────────────────────────────
sekme1, sekme2 = st.tabs(["📊 Genel Performans & Trendler", "🏪 Mağaza Detay Analiz Kartı"])

# ── SEKME 1 ───────────────────────────────────
with sekme1:
    st.markdown("### Filtrelenmiş Mağaza Performans KPI Özetleri")

    st.markdown("#### 🚀 2026 Yılı Dönemsel Performans")
    if goster_mayis:
        k1, k2, k3 = st.columns(3)
        with k1:
            delta = round(kpi26 - kpi25, 1)
            st.metric("📅 2026 Nisan TO", f"%{kpi26:.1f}", delta=f"%{delta:.1f} (vs 2025 Nisan)", delta_color="inverse")
        with k2:
            delta = round(kpi_may26 - kpi_may25, 1)
            st.metric("📆 2026 Mayıs TO MTD", f"%{kpi_may26:.1f}", delta=f"%{delta:.1f} (vs 2025 Mayıs)", delta_color="inverse")
        with k3:
            delta = round(kpi_ytd26 - kpi_ytd25, 1)
            st.metric("📊 2026 YTD Toplam TO", f"%{kpi_ytd26:.1f}", delta=f"%{delta:.1f} (vs 2025 YTD)", delta_color="inverse")
    else:
        k1, k2 = st.columns(2)
        with k1:
            delta = round(kpi26 - kpi25, 1)
            st.metric(f"📅 {ana_baslik26}", f"%{kpi26:.1f}", delta=f"%{delta:.1f} (vs 2025)", delta_color="inverse")
        with k2:
            delta = round(kpi_ytd26 - kpi_ytd25, 1)
            st.metric("📊 2026 YTD Toplam TO", f"%{kpi_ytd26:.1f}", delta=f"%{delta:.1f} (vs 2025)", delta_color="inverse")

    st.markdown("#### 🌐 2025 Yılı Dönemsel"
                " Performans")
    if goster_mayis:
        k4, k5, k6 = st.columns(3)
        with k4:
            st.metric("📅 2025 Nisan TO", f"%{kpi25:.1f}")
        with k5:
            st.metric("📆 2025 Mayıs TO MTD", f"%{kpi_may25:.1f}")
        with k6:
            st.metric("📊 2025 YTD Toplam TO", f"%{kpi_ytd25:.1f}")
    else:
        k4, k5 = st.columns(2)
        with k4:
            st.metric(f"📅 {ana_baslik25}", f"%{kpi25:.1f}")
        with k5:
            st.metric("📊 2025 YTD Toplam TO", f"%{kpi_ytd25:.1f}")

    st.divider()
    st.markdown("### İlgili Filtreye Ait Mağaza Listesi")

    col_ara, col_risk = st.columns([2, 1])
    with col_ara:
        st.markdown("🔍 **Liste İçi Arama**")
        arama = st.text_input("Ara", placeholder="Mağaza kodu veya adı...", label_visibility="collapsed")
    with col_risk:
        st.markdown("🎯 **Risk Grubu Filtresi (2026 YTD)**")
        risk_filtre = st.selectbox("Risk", ["Tümü", "Yüksek (>%30)", "Normal (%20-30)", "Düşük/Normal (<=%20)"], label_visibility="collapsed")

    df_liste = df_f.copy()
    if arama:
        df_liste = df_liste[
            df_liste["Masraf_Kodu"].astype(str).str.lower().str.contains(arama.lower()) |
            df_liste["Magaza"].astype(str).str.lower().str.contains(arama.lower())
        ]
    if risk_filtre == "Yüksek (>%30)":
        df_liste = df_liste[df_liste["YTD26_TO"] > 30]
    elif risk_filtre == "Normal (%20-30)":
        df_liste = df_liste[(df_liste["YTD26_TO"] > 20) & (df_liste["YTD26_TO"] <= 30)]
    elif risk_filtre == "Düşük/Normal (<=%20)":
        df_liste = df_liste[df_liste["YTD26_TO"] <= 10]

    st.markdown(f"**Görüntülenen Mağaza Sayısı: {len(df_liste)}**")

    # Tabloda seçili aya göre sütunlar
    if sec_ay == "Tümü (Nisan & Mayıs)":
        goster = df_liste[["Masraf_Kodu","Magaza","Il","Segment","BM","PM","HRBP",
                             "Nis25_TO","Mayis25_TO","YTD25_TO",
                             "Nis26_TO","Mayis26_TO","YTD26_TO"]].copy()
        goster.columns = ["Masraf Kodu","Mağaza Adı","İl","Segment","BM","PM","HRBP",
                           "2025 Nis%","2025 May%","2025 YTD%",
                           "2026 Nis%","2026 May%","2026 YTD%"]
        to_cols_goster = ["2025 Nis%","2025 May%","2025 YTD%","2026 Nis%","2026 May%","2026 YTD%"]
    else:
        goster = df_liste[["Masraf_Kodu","Magaza","Il","Segment","BM","PM","HRBP",
                             to25_col, to26_col, "YTD26_TO"]].copy()
        goster.columns = ["Masraf Kodu","Mağaza Adı","İl","Segment","BM","PM","HRBP",
                           f"2025 {sec_ay}%", f"2026 {sec_ay}%", "2026 YTD%"]
        to_cols_goster = [f"2025 {sec_ay}%", f"2026 {sec_ay}%", "2026 YTD%"]

    for c in to_cols_goster:
        goster[c] = goster[c].apply(lambda x: f"%{x:.1f}" if not pd.isna(x) else "—")
    st.dataframe(goster, use_container_width=True, hide_index=True)

# ── SEKME 2 ───────────────────────────────────
with sekme2:
    st.markdown("### Detaylı Mağaza Karnesi")
    st.markdown("Aşağıdaki filtreden seçim yapabilirsiniz.")

    st.markdown("🔍 **İncelenecek Mağazayı Seçin**")
    magaza_listesi = list(df_f["Masraf_Kodu"] + " — " + df_f["Magaza"])
    sec_magaza = st.selectbox("Mağaza", magaza_listesi, label_visibility="collapsed")

    if sec_magaza:
        kod = sec_magaza.split(" — ")[0].strip()
        s = df[df["Masraf_Kodu"] == kod].iloc[0]

        st.markdown(f"""
        <div class="magaza-kart">
            <p style="color:#aaaacc; font-size:12px; margin-bottom:8px;">MAĞAZA BİLGİLERİ 
            <h2 style="color:white; margin:0;">{s['Masraf_Kodu']} — {s['Magaza']}</h2>
            <br>
            <span style="color:#e0e0e0; margin-right:24px;">📍 İl: <b>{s['Il']}</b></span>
            <span style="color:#e0e0e0; margin-right:24px;">➕️ Segment: <b>{s['Segment']}</b></span>
            <br><br>
            <span style="color:#e0e0e0; margin-right:24px;">👤 BM: <b>{s['BM']}</b></span>
            <span style="color:#e0e0e0;">🏪 PM: <b>{s['PM']}</b></span>
            <span style="color:#e0e0e0;">🤝 HRBP: <b>{s['HRBP']}</b></span>
        </div>
        """, unsafe_allow_html=True)

        col_sol, col_sag = st.columns([1, 1])

        with col_sol:
            if sec_ay == "Tümü (Nisan & Mayıs)":
                st.markdown("#### 🚀 2026 Turnover Seviyeleri")
                k1, k2, k3 = st.columns(3)
                for kart, ay, to_val in zip([k1,k2,k3],
                    ["Nisan 2026","Mayıs 2026","YTD 2026"],
                    [s["Nis26_TO"], s["Mayis26_TO"], s["YTD26_TO"]]):
                    with kart:
                        etiket, emoji = risk_etiketi(to_val)
                        st.metric(ay, f"%{to_val:.1f}" if not pd.isna(to_val) else "—")
                        st.markdown(f"<p style='color:#aaaacc; font-size:12px;'>{emoji} {etiket}</p>", unsafe_allow_html=True)

                st.markdown("#### 🌐 2025 Turnover Seviyeleri")
                k4, k5, k6 = st.columns(3)
                for kart, ay, to_val in zip([k4,k5,k6],
                    ["Nisan 2025","Mayıs 2025","YTD 2025"],
                    [s["Nis25_TO"], s["Mayis25_TO"], s["YTD25_TO"]]):
                    with kart:
                        etiket, emoji = risk_etiketi(to_val)
                        st.metric(ay, f"%{to_val:.1f}" if not pd.isna(to_val) else "—")
                        st.markdown(f"<p style='color:#aaaacc; font-size:12px;'>{emoji} {etiket}</p>", unsafe_allow_html=True)
            else:
                st.markdown(f"#### 🚀 {sec_ay} 2026 Turnover")
                k1, k2 = st.columns(2)
                to26_val = s[to26_col]
                to25_val = s[to25_col]
                ytd_val  = s["YTD26_TO"]
                with k1:
                    etiket, emoji = risk_etiketi(to26_val)
                    st.metric(f"{sec_ay} 2026", f"%{to26_val:.1f}" if not pd.isna(to26_val) else "—")
                    st.markdown(f"<p style='color:#aaaacc; font-size:12px;'>{emoji} {etiket}</p>", unsafe_allow_html=True)
                with k2:
                    etiket, emoji = risk_etiketi(to25_val)
                    st.metric(f"{sec_ay} 2025", f"%{to25_val:.1f}" if not pd.isna(to25_val) else "—")
                    st.markdown(f"<p style='color:#aaaacc; font-size:12px;'>{emoji} {etiket}</p>", unsafe_allow_html=True)

                st.markdown("#### 📊 YTD 2026")
                etiket, emoji = risk_etiketi(ytd_val)
                st.metric("YTD 2026", f"%{ytd_val:.1f}" if not pd.isna(ytd_val) else "—")
                st.markdown(f"<p style='color:#aaaacc; font-size:12px;'>{emoji} {etiket}</p>", unsafe_allow_html=True)

        with col_sag:
            st.markdown("#### 📈 Yıllık Karşılaştırma Grafiği")
            aylar = ["Oca", "Şub", "Mar", "Nis", "May"]
            vals26 = [s["Oca26_TO"], s["Sub26_TO"], s["Mar26_TO"], s["Nis26_TO"], s["Mayis26_TO"]]
            vals25 = [s["Oca25_TO"], s["Sub25_TO"], s["Mar25_TO"], s["Nis25_TO"], s["Mayis25_TO"]]

            fig = go.Figure()
            fig.add_trace(go.Bar(name="2025", x=aylar, y=vals25, marker_color="#e05c5c", opacity=0.8))
            fig.add_trace(go.Bar(name="2026", x=aylar, y=vals26, marker_color="#4a6fa5", opacity=0.9))
            fig.update_layout(
                paper_bgcolor="#1a2f5e", plot_bgcolor="#1a2f5e",
                font=dict(color="white"), legend=dict(font=dict(color="white")),
                barmode="group", margin=dict(l=20, r=20, t=20, b=20), height=300
            )
            st.plotly_chart(fig, use_container_width=True)