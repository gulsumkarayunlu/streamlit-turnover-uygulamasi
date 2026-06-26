import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import re
import os

st.set_page_config(
    page_title="Koton Mağazacılık - Turnover Analizi",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── AKILLI EXCEL DOSYASI BULUCU (FILE NOT FOUND ÖNLEYİCİ) ────────
hedef_excel = None
try:
    # Klasördeki tüm dosyaları tarayıp "mayis" geçen ilk excel dosyasını buluyoruz
    for file in os.listdir("."):
        file_clean = str(file).lower().replace("ı", "i").strip()
        if "mayis" in file_clean and file.endswith(".xlsx") and not file.startswith("~$"):
            hedef_excel = file
            break

    # Eğer mayıs içeren dosya bulunamazsa, klasördeki ilk excel dosyasını yedek olarak seçelim
    if hedef_excel is None:
        for file in os.listdir("."):
            if file.endswith(".xlsx") and not file.startswith("~$"):
                hedef_excel = file
                break

    # Hiçbir şey bulunamazsa hata vermemesi için varsayılan isme dönelim
    if hedef_excel is None:
        hedef_excel = "Mayis_TO.xlsx"
except Exception:
    hedef_excel = "Mayis_TO.xlsx"

# ── AKILLI EXCEL SAYFASI BULUCU (SHEET NOT FOUND ÖNLEYİCİ) ────────
try:
    xls = pd.ExcelFile(hedef_excel)
    sayfa_isimleri = xls.sheet_names
    hedef_sayfa = None

    # Sayfa ismini Türkçe karakterlere ve boşluklara karşı duyarsız olarak arıyoruz
    for sheet in sayfa_isimleri:
        sheet_clean = str(sheet).strip().lower().replace("ı", "i")
        if "mayis" in sheet_clean and "detay" in sheet_clean:
            hedef_sayfa = sheet
            break

    # Eğer detay sayfası bulunamazsa, adında detay veya mayis geçen ilk sayfayı seçelim
    if hedef_sayfa is None:
        for sheet in sayfa_isimleri:
            sheet_clean = str(sheet).strip().lower().replace("ı", "i")
            if "detay" in sheet_clean or "mayis" in sheet_clean:
                hedef_sayfa = sheet
                break

    # Hala bulunamadıysa ilk sayfayı seçelim
    if hedef_sayfa is None:
        hedef_sayfa = sayfa_isimleri[0]

    df_detay = pd.read_excel(hedef_excel, sheet_name=hedef_sayfa, header=2)
except Exception as e:
    st.error(
        f"🚨 Excel dosyası okunamadı! Lütfen projenizde geçerli bir Excel dosyası ve içinde detay sayfası bulunduğundan emin olun. Hata: {str(e)}")
    df_detay = pd.DataFrame()


def bul_kolon(columns, anahtar_kelimeler, haric_tutulacaklar=None):
    """
    Excel başlıklarında yazım farklılıkları olsa bile
    en yakın eşleşen sütun adını güvenli şekilde bulur.
    """
    if haric_tutulacaklar is None:
        haric_tutulacaklar = []

    # 1. Aşama: Tüm kelimelerin sütun adında tam olarak geçmesi
    for col in columns:
        col_str = str(col).strip().lower()
        if all(kw.lower() in col_str for kw in anahtar_kelimeler):
            if not any(h.lower() in col_str for h in haric_tutulacaklar):
                return col

    # 2. Aşama: Kelimelerden en az birinin geçmesi (Yedek plan)
    for col in columns:
        col_str = str(col).strip().lower()
        if any(kw.lower() in col_str for kw in anahtar_kelimeler):
            if not any(h.lower() in col_str for h in haric_tutulacaklar):
                return col
    return None


def parse_percentage(val):
    """
    Gelen veriyi (ondalık sayı, yüzde formatlı metin vb.)
    güvenli bir şekilde yüzdelik sayıya dönüştürür.
    """
    if pd.isna(val) or val == "—":
        return None
    if isinstance(val, (int, float)):
        if val <= 1.0 and val > 0:
            return val * 100
        return val
    if isinstance(val, str):
        val_clean = val.replace("%", "").strip()
        try:
            v = float(val_clean)
            if v <= 1.0 and v > 0 and "%" not in val:
                return v * 100
            return v
        except ValueError:
            return None
    return None


def oran_formatla(val):
    if pd.isna(val) or val == "—":
        return "—"
    parsed = parse_percentage(val)
    if parsed is None:
        return "—"
    return f"%{parsed:.1f}"


# Detay tablo için mağaza sütununu bulup temizliyoruz
if not df_detay.empty:
    magaza_col_detay = (
            bul_kolon(df_detay.columns, ["mağaza"]) or
            bul_kolon(df_detay.columns, ["magaza"]) or
            bul_kolon(df_detay.columns, ["masraf"]) or
            df_detay.columns[2]
    )

    df_detay = df_detay.dropna(subset=[magaza_col_detay])
    df_detay = df_detay[df_detay[magaza_col_detay].astype(str).str.strip() != "Toplam"]
    df_detay = df_detay[df_detay[magaza_col_detay].astype(str).str.strip() != ""]
else:
    magaza_col_detay = None

# ── STİL DOSYALARI ─────────────────────────────
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
    .segment-kart {
        background-color: #1a2f5e;
        border-left: 4px solid #4a6fa5;
        border-radius: 8px;
        padding: 16px;
        margin: 10px 0;
    }
    /* 2x2 Grid için Lacivert Kutuların Stilleri */
    .turnover-kutu {
        background-color: #1a2f5e;
        border-radius: 8px;
        padding: 22px;
        margin-bottom: 15px;
        text-align: center;
        border-bottom: 4px solid #85B7EB;
        box-shadow: 0 4px 6px rgba(0,0,0,0.15);
    }
    .turnover-kutu.gonullu { border-bottom: 4px solid #4a6fa5; }
    .turnover-kutu.gonulsuz { border-bottom: 4px solid #e05c5c; }
    .turnover-kutu.zorunlu { border-bottom: 4px solid #ffd166; }
</style>
""", unsafe_allow_html=True)


@st.cache_data
def veri_yukle():
    df = pd.read_excel(hedef_excel, sheet_name="TO-Yeni Format MTD-YTD", header=2)
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
    to_cols = ["Oca26_TO", "Sub26_TO", "Mar26_TO", "Nis26_TO", "Mayis26_TO", "YTD26_TO",
               "Oca25_TO", "Sub25_TO", "Mar25_TO", "Nis25_TO", "Mayis25_TO", "YTD25_TO"]
    for col in to_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce") * 100
    return df


df = veri_yukle()

ay_map = {
    "Ocak": ("Oca26_Cikis", "Oca26_Ort", "Oca26_TO", "Oca25_TO"),
    "Şubat": ("Sub26_Cikis", "Sub26_Ort", "Sub26_TO", "Sub25_TO"),
    "Mart": ("Mar26_Cikis", "Mar26_Ort", "Mar26_TO", "Mar25_TO"),
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
    sec_ay = st.selectbox("Ay", ["Tümü (Nisan & Mayıs)", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs"],
                          label_visibility="collapsed")

    pm_listesi = ["Tümü"] + sorted(df["PM"].dropna().unique().tolist())
    st.markdown("🏪 **Perakende Müdürlüğü (PM)**")
    sec_pm = st.selectbox("PM", pm_listesi, label_visibility="collapsed")

    if sec_pm != "Tümü":
        df_temp_pm = df[df["PM"] == sec_pm]
    else:
        df_temp_pm = df.copy()

    bm_listesi = ["Tümü"] + sorted(df_temp_pm["BM"].dropna().unique().tolist())
    st.markdown("👤 **Bölge Müdürlüğü (BM)**")
    sec_bm = st.selectbox("BM", bm_listesi, label_visibility="collapsed")

    if sec_bm != "Tümü":
        df_temp_bm = df_temp_pm[df_temp_pm["BM"] == sec_bm]
    else:
        df_temp_bm = df_temp_pm.copy()

    hrbp_listesi = ["Tümü"] + sorted(df_temp_bm["HRBP"].dropna().unique().tolist())
    st.markdown("🤝 **İK İş Ortağı (HRBP)**")
    sec_hrbp = st.selectbox("HRBP", hrbp_listesi, label_visibility="collapsed")

    segment_sirasi = ["FS", "A++", "A+", "A", "B", "C", "D"]
    segment_listesi = ["Tümü"] + [s for s in segment_sirasi if s in df["Segment"].dropna().unique().tolist()]
    st.markdown("🏷️ **Segment**")
    sec_segment = st.selectbox("Segment", segment_listesi, label_visibility="collapsed")

# ── FİLTRELEME ────────────────────────────────
df_f = df.copy()
if sec_bm != "Tümü": df_f = df_f[df_f["BM"] == sec_bm]
if sec_pm != "Tümü": df_f = df_f[df_f["PM"] == sec_pm]
if sec_hrbp != "Tümü": df_f = df_f[df_f["HRBP"] == sec_hrbp]
if sec_segment != "Tümü": df_f = df_f[df_f["Segment"] == sec_segment]

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
    ort25 = ort26.replace("26", "25")
    ana_baslik26 = f"2026 {sec_ay} TO"
    ana_baslik25 = f"2025 {sec_ay} TO"
    ytd26_col, ytd25_col = "YTD26_TO", "YTD25_TO"
    ytd26_cikis, ytd26_ort = "YTD26_Cikis", "YTD26_Ort"
    ytd25_cikis, ytd25_ort = "YTD25_Cikis", "YTD25_Ort"
    may26_col = may25_col = None
    goster_ytd = True
    goster_mayis = False

kpi26 = genel_to(df_f, cikis26, ort26)
kpi25 = genel_to(df_f, cikis25, ort25)
kpi_ytd26 = genel_to(df_f, ytd26_cikis, ytd26_ort)
kpi_ytd25 = genel_to(df_f, ytd25_cikis, ytd25_ort)
kpi_may26 = genel_to(df_f, "Mayis26_Cikis", "Mayis26_Ort") if goster_mayis else None
kpi_may25 = genel_to(df_f, "Mayis25_Cikis", "Mayis25_Ort") if goster_mayis else None

filtre_adi = []
if sec_bm != "Tümü": filtre_adi.append(f"BM: {sec_bm}")
if sec_pm != "Tümü": filtre_adi.append(f"PM: {sec_pm}")
if sec_hrbp != "Tümü": filtre_adi.append(f"HRBP: {sec_hrbp}")
if sec_segment != "Tümü": filtre_adi.append(f"Segment: {sec_segment}")
if sec_ay != "Tümü (Nisan & Mayıs)": filtre_adi.append(f"Ay: {sec_ay}")
alt_yazi = " | ".join(
    filtre_adi) if filtre_adi else "Tüm Türkiye Mağazalarının Turnover Detay Verisi Aşağıda Görüntülenebilmektedir."

# ── ANA BAŞLIK ────────────────────────────────
col_logo, col_baslik = st.columns([1, 6])
with col_logo:
    try:
        st.image("koton_logo.png", width=180)
    except:
        st.markdown("**KOTON**")
with col_baslik:
    st.markdown("<h1 style='color:white; margin-bottom:0;'>Mağazacılık Turnover Analiz Paneli</h1>",
                unsafe_allow_html=True)
    st.markdown(f"<p style='color:#aaaacc; margin-top:4px;'>{alt_yazi}</p>", unsafe_allow_html=True)

st.divider()

# ── SEKMELER ──────────────────────────────────
sekme1, sekme2, sekme3 = st.tabs([
    "📊 Genel Performans & Trendler",
    "🏪 Mağaza Detay Analiz Kartı",
    "📅 Mayıs Detay Analizi"
])

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
            st.metric("📆 2026 Mayıs TO MTD", f"%{kpi_may26:.1f}", delta=f"%{delta:.1f} (vs 2025 Mayıs)",
                      delta_color="inverse")
        with k3:
            delta = round(kpi_ytd26 - kpi_ytd25, 1)
            st.metric("📊 2026 YTD Toplam TO", f"%{kpi_ytd26:.1f}", delta=f"%{delta:.1f} (vs 2025 YTD)",
                      delta_color="inverse")
    else:
        k1, k2 = st.columns(2)
        with k1:
            delta = round(kpi26 - kpi25, 1)
            st.metric(f"📅 {ana_baslik26}", f"%{kpi26:.1f}", delta=f"%{delta:.1f} (vs 2025)", delta_color="inverse")
        with k2:
            delta = round(kpi_ytd26 - kpi_ytd25, 1)
            st.metric("📊 2026 YTD Toplam TO", f"%{kpi_ytd26:.1f}", delta=f"%{delta:.1f} (vs 2025)",
                      delta_color="inverse")

    st.markdown("#### 🌐 2025 Yılı Dönemsel Performans")
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
        risk_filtre = st.selectbox("Risk", ["Tümü", "Yüksek (>%30)", "Normal (%20-30)", "Düşük/Normal (<=%20)"],
                                   label_visibility="collapsed")

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
        df_liste = df_liste[df_liste["YTD26_TO"] <= 20]

    st.markdown(f"**Görüntülenen Mağaza Sayısı: {len(df_liste)}**")

    if sec_ay == "Tümü (Nisan & Mayıs)":
        goster = df_liste[["Masraf_Kodu", "Magaza", "Il", "Segment", "BM", "PM", "HRBP",
                           "Nis25_TO", "Mayis25_TO", "YTD25_TO",
                           "Nis26_TO", "Mayis26_TO", "YTD26_TO"]].copy()
        goster.columns = ["Masraf Kodu", "Mağaza Adı", "İl", "Segment", "BM", "PM", "HRBP",
                          "2025 Nis%", "2025 May%", "2025 YTD%",
                          "2026 Nis%", "2026 May%", "2026 YTD%"]
        to_cols_goster = ["2025 Nis%", "2025 May%", "2025 YTD%", "2026 Nis%", "2026 May%", "2026 YTD%"]
    else:
        goster = df_liste[["Masraf_Kodu", "Magaza", "Il", "Segment", "BM", "PM", "HRBP",
                           to25_col, to26_col, "YTD26_TO"]].copy()
        goster.columns = ["Masraf Kodu", "Mağaza Adı", "İl", "Segment", "BM", "PM", "HRBP",
                          f"2025 {sec_ay}%", f"2026 {sec_ay}%", "2026 YTD%"]
        to_cols_goster = [f"2025 {sec_ay}%", f"2026 {sec_ay}%", "2026 YTD%"]

    for c in to_cols_goster:
        goster[c] = goster[c].apply(lambda x: f"%{x:.1f}" if not pd.isna(x) else "—")
    st.dataframe(goster, use_container_width=True, hide_index=True)

# ── SEKME 2 ───────────────────────────────────
with sekme2:
    st.markdown("### Detaylı Mağaza Karnesi")
    st.markdown("Aşağıdaki filtreden seçim yapabilirsiniz.")

    col_mag, col_il = st.columns([3, 1])
    with col_il:
        st.markdown("📍 **İl Filtresi**")
        il_listesi = ["Tümü"] + sorted(df_f["Il"].dropna().unique().tolist())
        sec_il = st.selectbox("İl", il_listesi, label_visibility="collapsed")

    df_magaza = df_f.copy()
    if sec_il != "Tümü":
        df_magaza = df_magaza[df_magaza["Il"] == sec_il]

    with col_mag:
        st.markdown("🔍 **İncelenecek Mağazayı Seçin**")
        magaza_listesi = list(df_magaza["Masraf_Kodu"] + " — " + df_magaza["Magaza"])
        if not magaza_listesi:
            st.warning("Seçili filtreye uygun mağaza bulunamadı.")
            st.stop()
        sec_magaza = st.selectbox("Mağaza", magaza_listesi, label_visibility="collapsed")

    if sec_magaza:
        kod = sec_magaza.split(" — ")[0].strip()
        s = df[df["Masraf_Kodu"] == kod].iloc[0]

        st.markdown(f"""
        <div class="magaza-kart">
            <p style="color:#aaaacc; font-size:12px; margin-bottom:8px;">MAĞAZA BİLGİLERİ</p>
            <h2 style="color:white; margin:0;">{s['Masraf_Kodu']} — {s['Magaza']}</h2>
            <br>
            <span style="color:#e0e0e0; margin-right:24px;">📍 İl: <b>{s['Il']}</b></span>
            <span style="color:#e0e0e0; margin-right:24px;">➕️ Segment: <b>{s['Segment']}</b></span>
            <br><br>
            <span style="color:#e0e0e0; margin-right:24px;">👤 BM: <b>{s['BM']}</b></span>
            <span style="color:#e0e0e0; margin-right:24px;">🏪 PM: <b>{s['PM']}</b></span>
            <span style="color:#e0e0e0;">🤝 HRBP: <b>{s['HRBP']}</b></span>
        </div>
        """, unsafe_allow_html=True)

        col_sol, col_sag = st.columns([1, 1])

        with col_sol:
            if sec_ay == "Tümü (Nisan & Mayıs)":
                st.markdown("#### 🚀 2026 Turnover Seviyeleri")
                k1, k2, k3 = st.columns(3)
                for kart, ay, to_val in zip([k1, k2, k3],
                                            ["Nisan 2026", "Mayıs 2026", "YTD 2026"],
                                            [s["Nis26_TO"], s["Mayis26_TO"], s["YTD26_TO"]]):
                    with kart:
                        etiket, emoji = risk_etiketi(to_val)
                        st.metric(ay, f"%{to_val:.1f}" if not pd.isna(to_val) else "—")
                        st.markdown(f"<p style='color:#aaaacc; font-size:12px;'>{emoji} {etiket}</p>",
                                    unsafe_allow_html=True)

                st.markdown("#### 🌐 2025 Turnover Seviyeleri")
                k4, k5, k6 = st.columns(3)
                for kart, ay, to_val in zip([k4, k5, k6],
                                            ["Nisan 2025", "Mayıs 2025", "YTD 2025"],
                                            [s["Nis25_TO"], s["Mayis25_TO"], s["YTD25_TO"]]):
                    with kart:
                        etiket, emoji = risk_etiketi(to_val)
                        st.metric(ay, f"%{to_val:.1f}" if not pd.isna(to_val) else "—")
                        st.markdown(f"<p style='color:#aaaacc; font-size:12px;'>{emoji} {etiket}</p>",
                                    unsafe_allow_html=True)
            else:
                st.markdown(f"#### 🚀 {sec_ay} 2026 Turnover")
                k1, k2 = st.columns(2)
                to26_val = s[to26_col]
                to25_val = s[to25_col]
                ytd_val = s["YTD26_TO"]
                with k1:
                    etiket, emoji = risk_etiketi(to26_val)
                    st.metric(f"{sec_ay} 2026", f"%{to26_val:.1f}" if not pd.isna(to26_val) else "—")
                    st.markdown(f"<p style='color:#aaaacc; font-size:12px;'>{emoji} {etiket}</p>",
                                unsafe_allow_html=True)
                with k2:
                    etiket, emoji = risk_etiketi(to25_val)
                    st.metric(f"{sec_ay} 2025", f"%{to25_val:.1f}" if not pd.isna(to25_val) else "—")
                    st.markdown(f"<p style='color:#aaaacc; font-size:12px;'>{emoji} {etiket}</p>",
                                unsafe_allow_html=True)

                st.markdown("#### 📊 YTD 2026")
                etiket, emoji = risk_etiketi(ytd_val)
                st.metric("YTD 2026", f"%{ytd_val:.1f}" if not pd.isna(ytd_val) else "—")
                st.markdown(f"<p style='color:#aaaacc; font-size:12px;'>{emoji} {etiket}</p>", unsafe_allow_html=True)

        with col_sol:
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

        st.divider()
        st.markdown(f"#### 🏷️ Segment Karşılaştırması — {s['Segment']}")
        st.markdown(
            f"<p style='color:#aaaacc;'>Bu mağazanın aynı segmentteki ({s['Segment']}) diğer mağazalarla karşılaştırması</p>",
            unsafe_allow_html=True)

        df_seg = df[df["Segment"] == s["Segment"]]
        seg_nis26 = genel_to(df_seg, "Nis26_Cikis", "Nis26_Ort")
        seg_may26 = genel_to(df_seg, "Mayis26_Cikis", "Mayis26_Ort")
        seg_ytd26 = genel_to(df_seg, "YTD26_Cikis", "YTD26_Ort")
        seg_nis25 = genel_to(df_seg, "Nis25_Cikis", "Nis25_Ort")
        seg_may25 = genel_to(df_seg, "Mayis25_Cikis", "Mayis25_Ort")
        seg_ytd25 = genel_to(df_seg, "YTD25_Cikis", "YTD25_Ort")

        st.markdown(f"""
        <div class="segment-kart">
            <p style="color:#85B7EB; font-weight:bold; margin-bottom:12px;">Segment Ortalaması ({s['Segment']}) — {len(df_seg)} Mağaza</p>
        </div>
        """, unsafe_allow_html=True)

        sc1, sc2, sc3 = st.columns(3)
        karsilastirmalar = [
            ("📅 Nisan 2026", s["Nis26_TO"], seg_nis26),
            ("📆 Mayıs 2026", s["Mayis26_TO"], seg_may26),
            ("📊 YTD 2026", s["YTD26_TO"], seg_ytd26),
        ]
        for kart, (baslik, magaza_val, seg_val) in zip([sc1, sc2, sc3], karsilastirmalar):
            with kart:
                mag_str = f"%{magaza_val:.1f}" if not pd.isna(magaza_val) else "—"
                seg_str = f"%{seg_val:.1f}"
                delta = round(magaza_val - seg_val, 1) if not pd.isna(magaza_val) else 0
                st.metric(
                    label=baslik,
                    value=mag_str,
                    delta=f"%{delta:.1f} (Seg. Ort: {seg_str})",
                    delta_color="inverse"
                )

        sc4, sc5, sc6 = st.columns(3)
        karsilastirmalar25 = [
            ("📅 Nisan 2025", s["Nis25_TO"], seg_nis25),
            ("📆 Mayıs 2025", s["Mayis25_TO"], seg_may25),
            ("📊 YTD 2025", s["YTD25_TO"], seg_ytd25),
        ]
        for kart, (baslik, magaza_val, seg_val) in zip([sc4, sc5, sc6], karsilastirmalar25):
            with kart:
                mag_str = f"%{magaza_val:.1f}" if not pd.isna(magaza_val) else "—"
                seg_str = f"%{seg_val:.1f}"
                delta = round(magaza_val - seg_val, 1) if not pd.isna(magaza_val) else 0
                st.metric(
                    label=baslik,
                    value=mag_str,
                    delta=f"%{delta:.1f} (Seg. Ort: {seg_str})",
                    delta_color="inverse"
                )

# ── SEKME 3 ───────────────────────────────────
with sekme3:
    st.markdown("### 📅 Mayıs Detay Veri Analizi")
    st.markdown(
        "Mağaza ve unvan seçimi yaparak Nisan Dönem İçi hareketlerini ve turnover oranlarını dinamik olarak inceleyebilirsiniz.")

    if df_detay.empty:
        st.warning(
            "⚠️ Excel dosyasında 'Mayıs Detay Tablo' sayfası bulunamadığı için bu sekmedeki analizler yüklenemedi. Lütfen Excel dosyasını kontrol edin.")
    else:
        # 1. Filtre: Mağaza Seçimi (En Üstte)
        st.markdown("🔍 **Mağaza Seçimi**")
        magaza_listesi_detay = sorted(df_detay[magaza_col_detay].astype(str).unique().tolist())
        sec_magaza_detay = st.selectbox("Mağaza (Detay Tablo)", magaza_listesi_detay, label_visibility="collapsed")

        if sec_magaza_detay:
            # Seçilen mağazanın satırını çekiyoruz
            s_detay = df_detay[df_detay[magaza_col_detay].astype(str) == sec_magaza_detay].iloc[0]

            # 2. Üst Sütunların (Full Time ve Part-Time) Altındaki Sütunları Bulma
            ft_tum_col = None
            ft_gonullu_col = None
            pt_tum_col = None
            pt_gonullu_col = None

            # 1. Aşama: Net Kelime Araması (Hatasız eşleşme için arada '-' işareti olmayan "Full Time" / "Part-Time" aranır)
            for col in df_detay.columns:
                col_str = str(col).strip().lower()
                if "sayısı" in col_str or "sayisi" in col_str:
                    continue

                if "full time" in col_str or "fulltime" in col_str or "full" in col_str:
                    if "tüm" in col_str or "tum" in col_str:
                        ft_tum_col = col
                    elif "gönüllü" in col_str or "gonullu" in col_str:
                        ft_gonullu_col = col
                elif "part-time" in col_str or "part time" in col_str or "parttime" in col_str or "part" in col_str:
                    if "tüm" in col_str or "tum" in col_str:
                        pt_tum_col = col
                    elif "gönüllü" in col_str or "gonullu" in col_str:
                        pt_gonullu_col = col

            # 2. Aşama: Eğer bulunamadıysa (Pandas'ın duplicate kolonlar için sonlarına .1 eklediği durum fallback'i)
            if not ft_tum_col or not pt_tum_col:
                # Unvan detaylarını ve "sayı", "sayisi", "çıkış" kelimelerini hariç tutarak sadece Tüm ve Gönüllü turnover sütunlarını filtreliyoruz
                tum_all = [c for c in df_detay.columns if ("tüm" in str(c).lower() or "tum" in str(c).lower()) and (
                            "turnover" in str(c).lower() or "to" in str(c).lower()) and not any(
                    u in str(c).lower() for u in
                    ["müdür", "vm", "sorumlu", "satış", "danışman", "mmy", "yardım", "sayısı", "sayisi", "çıkış",
                     "cikis"])]
                gonullu_all = [c for c in df_detay.columns if "gönüllü" in str(c).lower() and (
                            "turnover" in str(c).lower() or "to" in str(c).lower()) and not any(
                    u in str(c).lower() for u in
                    ["müdür", "vm", "sorumlu", "satış", "danışman", "mmy", "yardım", "sayısı", "sayisi", "çıkış",
                     "cikis"])]

                if len(tum_all) >= 2:
                    # Kullanıcının uyarısı üzerine Full Time ve Part-Time yerleşim eşleşmesini ters çevirerek (swap) düzelttik
                    ft_tum_col = tum_all[1]
                    pt_tum_col = tum_all[0]
                if len(gonullu_all) >= 2:
                    ft_gonullu_col = gonullu_all[1]
                    pt_gonullu_col = gonullu_all[0]

            # Değerleri formatlama işlemi (Düzeltilmiş ters eşleşme (swap) atamaları)
            val_ft_tum = oran_formatla(s_detay[ft_tum_col]) if ft_tum_col else "—"
            val_ft_gonullu = oran_formatla(s_detay[ft_gonullu_col]) if ft_gonullu_col else "—"

            val_pt_tum = oran_formatla(s_detay[pt_tum_col]) if pt_tum_col else "—"
            val_pt_gonullu = oran_formatla(s_detay[pt_gonullu_col]) if pt_gonullu_col else "—"

            # 3. Görsel Tasarım: Grup içi Full Time ve Part-Time Dağılımlarının Gösterilmesi (Sadece Tüm ve Gönüllü)
            st.markdown("#### 🎯 Mağaza İçi İstihdam Dağılım Oranları")
            kpi_col1, kpi_col2 = st.columns(2)

            with kpi_col1:
                st.markdown(f"""
                <div class="magaza-kart" style="border-left: 4px solid #85B7EB; padding: 18px; margin: 0;">
                    <p style="color:#aaaacc; font-size:12px; margin-bottom:5px; font-weight:bold; letter-spacing:0.5px;">GRUP İÇİ FULL TIME</p>
                    <div style="display: flex; justify-content: space-around; margin-top: 15px;">
                        <div style="text-align: center; flex: 1; border-right: 1px solid #2e3f7a;">
                            <p style="color:#85B7EB; font-size:11px; margin:0;">Tüm Turnover</p>
                            <h4 style="color:white; margin:6px 0 0 0; font-size:22px; font-weight:bold;">{val_ft_tum}</h4>
                        </div>
                        <div style="text-align: center; flex: 1;">
                            <p style="color:#ffd166; font-size:11px; margin:0;">Gönüllü Turnover</p>
                            <h4 style="color:white; margin:6px 0 0 0; font-size:22px; font-weight:bold;">{val_ft_gonullu}</h4>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            with kpi_col2:
                st.markdown(f"""
                <div class="magaza-kart" style="border-left: 4px solid #4a6fa5; padding: 18px; margin: 0;">
                    <p style="color:#aaaacc; font-size:12px; margin-bottom:5px; font-weight:bold; letter-spacing:0.5px;">GRUP İÇİ PART-TIME</p>
                    <div style="display: flex; justify-content: space-around; margin-top: 15px;">
                        <div style="text-align: center; flex: 1; border-right: 1px solid #2e3f7a;">
                            <p style="color:#85B7EB; font-size:11px; margin:0;">Tüm Turnover</p>
                            <h4 style="color:white; margin:6px 0 0 0; font-size:22px; font-weight:bold;">{val_pt_tum}</h4>
                        </div>
                        <div style="text-align: center; flex: 1;">
                            <p style="color:#ffd166; font-size:11px; margin:0;">Gönüllü Turnover</p>
                            <h4 style="color:white; margin:6px 0 0 0; font-size:22px; font-weight:bold;">{val_pt_gonullu}</h4>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            st.divider()

            # 4. Filtre: Unvan Filtresi (Oran Kartlarının Hemen Altında)
            st.markdown("👤 **Unvan Filtresi**")
            sec_unvan = st.selectbox("Unvan", list(unvan_secenekleri.keys()), label_visibility="collapsed")

            st.divider()

            # Seçilen unvana göre anahtarları alıyoruz
            keys, exclude = unvan_secenekleri[sec_unvan]

            # 5. Sütunları Unvana Göre Dinamik Bulma
            # Dönem Başı
            if sec_unvan == "Tümü (Toplam)":
                db_col = (
                        bul_kolon(df_detay.columns, ["nisan", "dönem başı"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"]) or
                        bul_kolon(df_detay.columns, ["dönem başı"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"]) or
                        bul_kolon(df_detay.columns, ["baş", "çalışan"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"])
                )
            else:
                db_col = (
                        bul_kolon(df_detay.columns, ["nisan", "dönem başı"] + keys, exclude) or
                        bul_kolon(df_detay.columns, ["dönem başı"] + keys, exclude) or
                        bul_kolon(df_detay.columns, ["baş", "çalışan"] + keys, exclude)
                )

            # İşe Giriş
            if sec_unvan == "Tümü (Toplam)":
                giris_col_detay = bul_kolon(df_detay.columns, ["giriş"],
                                            ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"])
            else:
                giris_col_detay = bul_kolon(df_detay.columns, ["giriş"] + keys, exclude)

            # İşten Çıkış
            if sec_unvan == "Tümü (Toplam)":
                cikis_col_detay = (
                        bul_kolon(df_detay.columns, ["işten çıkış"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy", "gönüllü", "gönülsüz",
                                   "zorunlu", "oran"]) or
                        bul_kolon(df_detay.columns, ["çıkış"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy", "gönüllü", "gönülsüz",
                                   "zorunlu", "oran"])
                )
            else:
                cikis_col_detay = (
                        bul_kolon(df_detay.columns, ["işten çıkış"] + keys,
                                  exclude + ["gönüllü", "gönülsüz", "zorunlu", "oran"]) or
                        bul_kolon(df_detay.columns, ["çıkış"] + keys,
                                  exclude + ["gönüllü", "gönülsüz", "zorunlu", "oran"])
                )

            # Gönüllü Çıkış
            if sec_unvan == "Tümü (Toplam)":
                gonullu_col = (
                        bul_kolon(df_detay.columns, ["gönüllü", "çıkış"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy", "oran"]) or
                        bul_kolon(df_detay.columns, ["gönüllü"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy", "oran"])
                )
            else:
                gonullu_col = (
                        bul_kolon(df_detay.columns, ["gönüllü", "çıkış"] + keys, exclude + ["oran"]) or
                        bul_kolon(df_detay.columns, ["gönüllü"] + keys, exclude + ["oran"])
                )

            # Gönülsüz Çıkış
            if sec_unvan == "Tümü (Toplam)":
                gonulsuz_col = (
                        bul_kolon(df_detay.columns, ["gönülsüz", "çıkış"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy", "oran"]) or
                        bul_kolon(df_detay.columns, ["gönülsüz"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy", "oran"])
                )
            else:
                gonulsuz_col = (
                        bul_kolon(df_detay.columns, ["gönülsüz", "çıkış"] + keys, exclude + ["oran"]) or
                        bul_kolon(df_detay.columns, ["gönülsüz"] + keys, exclude + ["oran"])
                )

            # Zorunlu Çıkış
            if sec_unvan == "Tümü (Toplam)":
                zorunlu_col = (
                        bul_kolon(df_detay.columns, ["zorunlu", "çıkış"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy", "oran"]) or
                        bul_kolon(df_detay.columns, ["zorunlu"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy", "oran"])
                )
            else:
                zorunlu_col = (
                        bul_kolon(df_detay.columns, ["zorunlu", "çıkış"] + keys, exclude + ["oran"]) or
                        bul_kolon(df_detay.columns, ["zorunlu"] + keys, exclude + ["oran"])
                )

            # Dönem Sonu
            if sec_unvan == "Tümü (Toplam)":
                ds_col = (
                        bul_kolon(df_detay.columns, ["nisan", "dönem sonu"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"]) or
                        bul_kolon(df_detay.columns, ["dönem sonu"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"]) or
                        bul_kolon(df_detay.columns, ["son", "çalışan"],
                                  ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"])
                )
            else:
                ds_col = (
                        bul_kolon(df_detay.columns, ["nisan", "dönem sonu"] + keys, exclude) or
                        bul_kolon(df_detay.columns, ["dönem sonu"] + keys, exclude) or
                        bul_kolon(df_detay.columns, ["son", "çalışan"] + keys, exclude)
                )

            # Değerlerin Hazırlanması (Unvana özel veriler yoksa "-" döndürülür)
            val_db = s_detay[db_col] if db_col else "—"
            val_giris = s_detay[giris_col_detay] if giris_col_detay else "—"
            val_cikis = s_detay[cikis_col_detay] if cikis_col_detay else "—"
            val_gonullu = s_detay[gonullu_col] if gonullu_col else "—"
            val_gonulsuz = s_detay[gonulsuz_col] if gonulsuz_col else "—"
            val_zorunlu = s_detay[zorunlu_col] if zorunlu_col else "—"
            val_ds = s_detay[ds_col] if ds_col else "—"

            # Tablo verilerinin oluşturulması
            tablo1_verileri = [
                ("Nisan Dönem Başı Çalışan Sayısı", val_db),
                ("İşe Giriş Sayısı", val_giris),
                ("İşten Çıkış Sayısı", val_cikis),
                ("Gönüllü Çıkış Sayısı", val_gonullu),
                ("Gönülsüz Çıkış Sayısı", val_gonulsuz),
                ("Zorunlu Çıkış Sayısı", val_zorunlu),
                ("Nisan Dönem Sonu Çalışan Sayısı", val_ds)
            ]

            df_tablo1 = pd.DataFrame(tablo1_verileri, columns=["Metrik", "Sayı / Değer"])
            df_tablo1["Sayı / Değer"] = df_tablo1["Sayı / Değer"].apply(
                lambda x: int(x) if isinstance(x, (int, float)) and pd.notna(x) else x
            )

            # Pandas Styler ile sayılar sütununu kalın font (bold) ve kurumsal lacivert renge (#0d1b3e) boyuyoruz
            style_func = lambda x: "color: #0d1b3e; font-weight: bold; font-size: 15px;"
            if hasattr(df_tablo1.style, "map"):
                styled_df_tablo1 = df_tablo1.style.map(style_func, subset=["Sayı / Değer"])
            else:
                styled_df_tablo1 = df_tablo1.style.applymap(style_func, subset=["Sayı / Değer"])

            # Turnover Sütunlarının Seçimi
            to_tum_col = (
                    bul_kolon(df_detay.columns, ["tüm", "turnover"]) or
                    bul_kolon(df_detay.columns, ["tüm", "to"]) or
                    bul_kolon(df_detay.columns, ["toplam", "turnover"]) or
                    bul_kolon(df_detay.columns, ["toplam", "to"]) or
                    bul_kolon(df_detay.columns, ["turnover"], haric_tutulacaklar=["gönüllü", "gönülsüz", "zorunlu"]) or
                    bul_kolon(df_detay.columns, ["to"], haric_tutulacaklar=["gönüllü", "gönülsüz", "zorunlu"])
            )
            to_gonullu_col = bul_kolon(df_detay.columns, ["gönüllü", "turnover"]) or bul_kolon(df_detay.columns,
                                                                                               ["gönüllü", "to"])
            to_gonulsuz_col = bul_kolon(df_detay.columns, ["gönülsüz", "turnover"]) or bul_kolon(df_detay.columns,
                                                                                                 ["gönülsüz", "to"])
            to_zorunlu_col = bul_kolon(df_detay.columns, ["zorunlu", "turnover"]) or bul_kolon(df_detay.columns,
                                                                                               ["zorunlu", "to"])

            to_tum_val = oran_formatla(s_detay[to_tum_col]) if to_tum_col else "—"
            to_gonullu_val = oran_formatla(s_detay[to_gonullu_col]) if to_gonullu_col else "—"
            to_gonulsuz_val = oran_formatla(s_detay[to_gonulsuz_col]) if to_gonulsuz_col else "—"
            to_zorunlu_val = oran_formatla(s_detay[to_zorunlu_col]) if to_zorunlu_col else "—"

            # Grid/Kolon Yapısıyla Tabloları Yanyana Gösteriyoruz
            t_col1, t_col2 = st.columns([3, 2])

            with t_col1:
                st.markdown(f"#### 📊 Dönem İçi Hareketler ({sec_unvan})")
                st.dataframe(styled_df_tablo1, use_container_width=True, hide_index=True, height=290)

            with t_col2:
                st.markdown(f"#### 📉 Turnover Oranları Özeti ({sec_unvan})")

                # 2x2 Grid Düzeni (2 Kare Üstte, 2 Kare Altta)
                r1_c1, r1_c2 = st.columns(2)
                with r1_c1:
                    st.markdown(f"""
                    <div class="turnover-kutu">
                        <p style="color:#aaaacc; font-size:11px; margin:0 0 4px 0; font-weight:bold;">TÜM TURNOVER</p>
                        <h3 style="color:white; margin:0; font-size:24px;">{to_tum_val}</h3>
                    </div>
                    """, unsafe_allow_html=True)
                with r1_c2:
                    st.markdown(f"""
                    <div class="turnover-kutu gonullu">
                        <p style="color:#aaaacc; font-size:11px; margin:0 0 4px 0; font-weight:bold;">GÖNÜLLÜ TURNOVER</p>
                        <h3 style="color:white; margin:0; font-size:24px;">{to_gonullu_val}</h3>
                    </div>
                    """, unsafe_allow_html=True)

                r2_c1, r2_c2 = st.columns(2)
                with r2_c1:
                    st.markdown(f"""
                    <div class="turnover-kutu gonulsuz">
                        <p style="color:#aaaacc; font-size:11px; margin:0 0 4px 0; font-weight:bold;">GÖNÜLSÜZ TURNOVER</p>
                        <h3 style="color:white; margin:0; font-size:24px;">{to_gonulsuz_val}</h3>
                    </div>
                    """, unsafe_allow_html=True)
                with r2_c2:
                    st.markdown(f"""
                    <div class="turnover-kutu zorunlu">
                        <p style="color:#aaaacc; font-size:11px; margin:0 0 4px 0; font-weight:bold;">ZORUNLU TURNOVER</p>
                        <h3 style="color:white; margin:0; font-size:24px;">{to_zorunlu_val}</h3>
                    </div>
                    """, unsafe_allow_html=True)

                st.info(
                    f"💡 **Not:** Yukarıdaki tüm veriler **{sec_magaza_detay}** mağazasının "
                    f"**{sec_unvan}** pozisyonuna ait değerlerini yansıtmaktadır."
                )