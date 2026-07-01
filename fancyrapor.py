import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import os
import traceback

st.set_page_config(
    page_title="Koton Mağazacılık - Turnover Analizi",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── SESSION STATE ────────
for key, default in [
    ('ay_key', "Mayıs"),  # Varsayılan başlangıç ayı sadece "Mayıs" yapıldı
    ('pm_key', "Tümü"),
    ('bm_key', "Tümü"),
    ('hrbp_key', "Tümü"),
    ('segment_key', "Tümü"),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ── AY HARİTASI ────────
ay_map = {
    "Ocak": ("Oca26_Cikis", "Oca26_Ort", "Oca26_TO", "Oca25_TO"),
    "Şubat": ("Sub26_Cikis", "Sub26_Ort", "Sub26_TO", "Sub25_TO"),
    "Mart": ("Mar26_Cikis", "Mar26_Ort", "Mar26_TO", "Mar25_TO"),
    "Nisan": ("Nis26_Cikis", "Nis26_Ort", "Nis26_TO", "Nis25_TO"),
    "Mayıs": ("Mayis26_Cikis", "Mayis26_Ort", "Mayis26_TO", "Mayis25_TO"),
}

# ── UNVAN SEÇENEKLERİ ────────
unvan_secenekleri = {
    "Tümü (Toplam)": ([], []),
    "Mağaza Müdürü": (["müdür"], ["yardım", "yrd", "vm", "mmy"]),
    "Müdür Yardımcısı": (["yardım", "yrd", "mmy"]),
    "Satış Danışmanı": (["satış", "danışman"], ["müdür", "yardım", "vm", "sorumlu"]),
}

# ── EXCEL DOSYASI BULUCU ────────
hedef_excel = None
try:
    for file in os.listdir("."):
        file_clean = str(file).lower().replace("ı", "i").strip()
        if "mayis" in file_clean and file.endswith(".xlsx") and not file.startswith("~$"):
            hedef_excel = file
            break
    if hedef_excel is None:
        for file in os.listdir("."):
            if file.endswith(".xlsx") and not file.startswith("~$"):
                hedef_excel = file
                break
    if hedef_excel is None:
        hedef_excel = "Mayis_TO.xlsx"
except Exception:
    hedef_excel = "Mayis_TO.xlsx"


# ── YARDIMCI FONKSİYONLAR ────────
def bul_kolon(columns, anahtar_kelimeler, haric_tutulacaklar=None):
    if haric_tutulacaklar is None:
        haric_tutulacaklar = []
    for col in columns:
        col_str = str(col).strip().lower()
        if all(kw.lower() in col_str for kw in anahtar_kelimeler):
            if not any(h.lower() in col_str for h in haric_tutulacaklar):
                return col
    for col in columns:
        col_str = str(col).strip().lower()
        if any(kw.lower() in col_str for kw in anahtar_kelimeler):
            if not any(h.lower() in col_str for h in haric_tutulacaklar):
                return col
    return None


def parse_percentage(val):
    if pd.isna(val) or val == "—":
        return None
    if isinstance(val, (int, float)):
        if 0 < val <= 1.0:
            return val * 100
        return val
    if isinstance(val, str):
        val_clean = val.replace("%", "").strip()
        try:
            v = float(val_clean)
            if 0 < v <= 1.0 and "%" not in val:
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


def genel_to(df_f, cikis_col, ort_col):
    c = pd.to_numeric(df_f[cikis_col], errors="coerce").sum()
    o = pd.to_numeric(df_f[ort_col], errors="coerce").sum()
    return round((c / o) * 100, 1) if o > 0 else 0.0


def risk_etiketi(to):
    if pd.isna(to):
        return "—", "⚪"
    if to > 20:
        return "Yüksek", "🔴"
    if to > 10:
        return "Normal", "🟡"
    return "Düşük", "🟢"


def sayi_col_bul(columns, anahtar_kelimeler, haric_tutulacaklar=None):
    """
    Sayı (count) içeren sütunları bulur; oran/yüzde sütunlarını hariç tutar.
    Otomatik olarak 'oran', 'turnover', 'to', '%', 'rate' gibi terimleri hariç tutar.
    """
    oran_haric = ["oran", "turnover", "to", "%", "rate", "yüzde"]
    haric = list(haric_tutulacaklar or []) + oran_haric
    return bul_kolon(columns, anahtar_kelimeler, haric)


# ── VERİ YÜKLEME ────────
@st.cache_data
def veri_yukle():
    df = pd.read_excel(hedef_excel, sheet_name="TO-Yeni Format MTD-YTD", header=2)

    # Sütunlardaki boşlukları temizleyelim
    df.columns = [str(c).strip() for c in df.columns]
    prev_hrbp_idx = None
    for idx, col in enumerate(df.columns):
        col_str = str(col).strip().lower()
        if ("önceki" in col_str or "onceki" in col_str) and "hrbp" in col_str:
            prev_hrbp_idx = idx
            break

    standard_cols = [
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

    if prev_hrbp_idx is not None:
        standard_cols.insert(prev_hrbp_idx, "Onceki_Donem_HRBP")

    if len(df.columns) == len(standard_cols):
        df.columns = standard_cols
    else:
        diff = len(df.columns) - len(standard_cols)
        if diff > 0:
            standard_cols = standard_cols + [f"Ekstra_Col_{i}" for i in range(diff)]
        else:
            standard_cols = standard_cols[:len(df.columns)]
        df.columns = standard_cols

    df = df.dropna(subset=["Masraf_Kodu"])
    df = df[df["Masraf_Kodu"].astype(str).str.strip() != "Toplam"]
    df = df.reset_index(drop=True)
    to_cols = [
        "Oca26_TO", "Sub26_TO", "Mar26_TO", "Nis26_TO", "Mayis26_TO", "YTD26_TO",
        "Oca25_TO", "Sub25_TO", "Mar25_TO", "Nis25_TO", "Mayis25_TO", "YTD25_TO"
    ]
    for col in to_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce") * 100
    return df


df = veri_yukle()

# ── EXCEL SAYFALARI ────────
try:
    xls = pd.ExcelFile(hedef_excel)
    sayfa_isimleri = xls.sheet_names

    hedef_sayfa = None
    for sheet in sayfa_isimleri:
        sc = str(sheet).strip().lower().replace("ı", "i")
        if "mayis" in sc and "detay" in sc:
            hedef_sayfa = sheet
            break
    if hedef_sayfa is None:
        for sheet in sayfa_isimleri:
            sc = str(sheet).strip().lower().replace("ı", "i")
            if "detay" in sc or "mayis" in sc:
                hedef_sayfa = sheet
                break
    if hedef_sayfa is None:
        hedef_sayfa = sayfa_isimleri[0]

    df_detay = pd.read_excel(hedef_excel, sheet_name=hedef_sayfa, header=2)
except Exception as e:
    st.error(f"🚨 Excel dosyası okunamadı: {str(e)}")
    df_detay = pd.DataFrame()

# Yasal bildirimler sayfası
try:
    hedef_yasal_sayfa = None
    for sheet in sayfa_isimleri:
        sc = str(sheet).strip().lower().replace("ı", "i")
        if "yasal" in sc and "bildirim" in sc:
            hedef_yasal_sayfa = sheet
            break
    if hedef_yasal_sayfa is None:
        for sheet in sayfa_isimleri:
            sc = str(sheet).strip().lower().replace("ı", "i")
            if "yasal" in sc or "bildirim" in sc:
                hedef_yasal_sayfa = sheet
                break

    if hedef_yasal_sayfa:
        df_yasal = None
        for h in [0, 1, 2, 3]:
            df_temp = pd.read_excel(hedef_excel, sheet_name=hedef_yasal_sayfa, header=h)
            df_temp = df_temp.dropna(how="all")
            has_exit = any("çıkış" in str(col).lower() or "cikis" in str(col).lower() for col in df_temp.columns)
            has_notice = any("bildirim" in str(col).lower() for col in df_temp.columns)
            if has_exit or has_notice:
                df_yasal = df_temp
                break
        if df_yasal is None:
            df_yasal = pd.read_excel(hedef_excel, sheet_name=hedef_yasal_sayfa)
    else:
        df_yasal = pd.DataFrame()
except Exception:
    df_yasal = pd.DataFrame()

# ── İŞTEN AYRILANLAR SAYFASI ────────
try:
    hedef_ayrilanlar_sayfa = None
    for sheet in sayfa_isimleri:
        sc = str(sheet).strip().lower().replace("ı", "i").replace("ş", "s").replace("ğ", "g")
        if "ayril" in sc or "ayrıl" in str(sheet).strip().lower():
            hedef_ayrilanlar_sayfa = sheet
            break

    if hedef_ayrilanlar_sayfa:
        df_ayrilanlar = None
        for h in [0, 1, 2, 3]:
            df_temp = pd.read_excel(hedef_excel, sheet_name=hedef_ayrilanlar_sayfa, header=h)
            df_temp = df_temp.dropna(how="all")
            if len(df_temp.columns) >= 3 and len(df_temp) > 0:
                df_ayrilanlar = df_temp
                break
        if df_ayrilanlar is None:
            df_ayrilanlar = pd.read_excel(hedef_excel, sheet_name=hedef_ayrilanlar_sayfa)
    else:
        df_ayrilanlar = pd.DataFrame()
except Exception:
    df_ayrilanlar = pd.DataFrame()

# ── GÜNCEL ÇALIŞANLAR SAYFASI ────────
try:
    hedef_guncel_sayfa = None
    for sheet in sayfa_isimleri:
        sc = str(sheet).strip().lower().replace("ı", "i").replace("ş", "s").replace("ğ", "g").replace("ç", "c")
        if "guncel" in sc or "güncel" in str(sheet).strip().lower() or "calisan" in sc:
            hedef_guncel_sayfa = sheet
            break

    if hedef_guncel_sayfa:
        df_guncel = None
        for h in [0, 1, 2, 3]:
            df_temp = pd.read_excel(hedef_excel, sheet_name=hedef_guncel_sayfa, header=h)
            df_temp = df_temp.dropna(how="all")
            if len(df_temp.columns) >= 3 and len(df_temp) > 0:
                df_guncel = df_temp
                break
        if df_guncel is None:
            df_guncel = pd.read_excel(hedef_excel, sheet_name=hedef_guncel_sayfa)
    else:
        df_guncel = pd.DataFrame()
except Exception:
    df_guncel = pd.DataFrame()

# ── DETAY TABLOSU SÜTUN EŞLEŞTİRMELERİ ────────
if not df_detay.empty:
    magaza_col_detay = (
            bul_kolon(df_detay.columns, ["mağaza"]) or
            bul_kolon(df_detay.columns, ["magaza"]) or
            bul_kolon(df_detay.columns, ["masraf"])
    )
    if magaza_col_detay is None and len(df_detay.columns) > 0:
        magaza_col_detay = df_detay.columns[min(2, len(df_detay.columns) - 1)]
    if magaza_col_detay:
        df_detay = df_detay.dropna(subset=[magaza_col_detay])
        df_detay = df_detay[df_detay[magaza_col_detay].astype(str).str.strip() != "Toplam"]
        df_detay = df_detay[df_detay[magaza_col_detay].astype(str).str.strip() != ""]
else:
    magaza_col_detay = None

# Yasal bildirim sütun eşleştirmeleri
if not df_yasal.empty:
    masraf_col_yasal = (
            bul_kolon(df_yasal.columns, ["masraf", "kod"]) or
            bul_kolon(df_yasal.columns, ["masraf"]) or
            bul_kolon(df_yasal.columns, ["kod"])
    )
    magaza_name_col_yasal = (
            bul_kolon(df_yasal.columns, ["masraf", "yeri"]) or
            bul_kolon(df_yasal.columns, ["mağaza", "ad"]) or
            bul_kolon(df_yasal.columns, ["magaza", "ad"]) or
            bul_kolon(df_yasal.columns, ["mağaza"], haric_tutulacaklar=["kod"]) or
            bul_kolon(df_yasal.columns, ["magaza"], haric_tutulacaklar=["kod"])
    )
    magaza_col_yasal = masraf_col_yasal or magaza_name_col_yasal or df_yasal.columns[0]
    df_yasal = df_yasal.dropna(subset=[magaza_col_yasal])
    df_yasal = df_yasal[df_yasal[magaza_col_yasal].astype(str).str.strip() != "Toplam"]
    df_yasal = df_yasal[df_yasal[magaza_col_yasal].astype(str).str.strip() != ""]
    df_yasal["Masraf_Kodu_Temiz"] = df_yasal[magaza_col_yasal].astype(str).str.strip()
    masraf_to_magaza = dict(zip(df["Masraf_Kodu"].astype(str).str.strip(), df["Magaza"].astype(str).str.strip()))
    if masraf_col_yasal and magaza_name_col_yasal and masraf_col_yasal != magaza_name_col_yasal:
        df_yasal["Display_Magaza"] = (
                df_yasal[masraf_col_yasal].astype(str).str.strip() + " — " +
                df_yasal[magaza_name_col_yasal].astype(str).str.strip()
        )
    else:
        df_yasal["Mapped_Magaza"] = df_yasal["Masraf_Kodu_Temiz"].map(masraf_to_magaza).fillna("Bilinmeyen Mağaza")
        df_yasal["Display_Magaza"] = df_yasal["Masraf_Kodu_Temiz"] + " — " + df_yasal["Mapped_Magaza"]
else:
    magaza_col_yasal = None

# ── İşten Ayrılanlar sütun eşleştirmeleri ────────
if not df_ayrilanlar.empty:
    ayr_masraf_col = (
            bul_kolon(df_ayrilanlar.columns, ["masraf", "kod"]) or
            bul_kolon(df_ayrilanlar.columns, ["masraf"]) or
            bul_kolon(df_ayrilanlar.columns, ["kod"])
    )
    ayr_magaza_col = (
            bul_kolon(df_ayrilanlar.columns, ["mağaza"]) or
            bul_kolon(df_ayrilanlar.columns, ["magaza"])
    )
    ayr_isim_col = (
            bul_kolon(df_ayrilanlar.columns, ["ad", "soyad"]) or
            bul_kolon(df_ayrilanlar.columns, ["isim"]) or
            bul_kolon(df_ayrilanlar.columns, ["çalışan"])
    )
    ayr_tarih_col = (
            bul_kolon(df_ayrilanlar.columns, ["çıkış", "tarih"]) or
            bul_kolon(df_ayrilanlar.columns, ["ayrılma", "tarih"]) or
            bul_kolon(df_ayrilanlar.columns, ["tarih"])
    )
    ayr_to_neden_col = (
            bul_kolon(df_ayrilanlar.columns, ["to", "neden"]) or
            bul_kolon(df_ayrilanlar.columns, ["turnover", "neden"]) or
            bul_kolon(df_ayrilanlar.columns, ["ayrılma", "neden"]) or
            bul_kolon(df_ayrilanlar.columns, ["neden"])
    )
    ayr_cikis_neden_col = (
            bul_kolon(df_ayrilanlar.columns, ["çıkış", "neden"]) or
            bul_kolon(df_ayrilanlar.columns, ["işten", "neden"]) or
            bul_kolon(df_ayrilanlar.columns, ["sebep"])
    )

    # Masraf kodunu filtreler ile eşleştirmek için JOIN sütunu bul
    ayr_join_col = ayr_masraf_col or ayr_magaza_col
    if ayr_join_col:
        df_ayrilanlar = df_ayrilanlar.dropna(subset=[ayr_join_col])
        df_ayrilanlar = df_ayrilanlar[df_ayrilanlar[ayr_join_col].astype(str).str.strip() != ""]

        # Mükerrer sütunları temizle
        cols_to_drop_ayr = [c for c in ["BM", "PM", "HRBP", "Segment", "Il", "Magaza", "Onceki_Donem_HRBP"] if
                            c in df_ayrilanlar.columns and c != ayr_join_col]
        if cols_to_drop_ayr:
            df_ayrilanlar = df_ayrilanlar.drop(columns=cols_to_drop_ayr)

        # Masraf kodu üzerinden BM/PM/HRBP/Segment eşleme
        merge_cols = ["Masraf_Kodu", "BM", "PM", "HRBP", "Segment", "Magaza", "Il"]
        if "Onceki_Donem_HRBP" in df.columns:
            merge_cols.append("Onceki_Donem_HRBP")

        masraf_info = df[merge_cols].copy()
        masraf_info["Masraf_Kodu"] = masraf_info["Masraf_Kodu"].astype(str).str.strip()
        df_ayrilanlar["_join_key"] = df_ayrilanlar[ayr_join_col].astype(str).str.strip()
        df_ayrilanlar = df_ayrilanlar.merge(masraf_info, left_on="_join_key", right_on="Masraf_Kodu", how="left")

# ── Güncel Çalışanlar sütun eşleştirmeleri ────────
if not df_guncel.empty:
    gunc_masraf_col = (
            bul_kolon(df_guncel.columns, ["masraf", "kod"]) or
            bul_kolon(df_guncel.columns, ["masraf"]) or
            bul_kolon(df_guncel.columns, ["kod"])
    )
    gunc_magaza_col = (
            bul_kolon(df_guncel.columns, ["mağaza"]) or
            bul_kolon(df_guncel.columns, ["magaza"])
    )
    # Tam ve Yarı zamanlı süzmelerinin sıfır gelmemesi için saptama kapsamı yasal Türkçe sütun adına ("çalışanaltgrubu" veya "tam zamanlı" / "yarı zamanlı") göre genişletildi
    gunc_altgrup_col = (
            bul_kolon(df_guncel.columns, ["çalışanaltgrubu"]) or
            bul_kolon(df_guncel.columns, ["calisanaltgrubu"]) or
            bul_kolon(df_guncel.columns, ["alt grup"]) or
            bul_kolon(df_guncel.columns, ["çalışan alt"]) or
            bul_kolon(df_guncel.columns, ["altgrup"]) or
            bul_kolon(df_guncel.columns, ["çalışma şekli"]) or
            bul_kolon(df_guncel.columns, ["calisma sekli"]) or
            bul_kolon(df_guncel.columns, ["çalışma"]) or
            bul_kolon(df_guncel.columns, ["calisma"])
    )
    gunc_cinsiyet_col = (
            bul_kolon(df_guncel.columns, ["cinsiyet"]) or
            bul_kolon(df_guncel.columns, ["cins"])
    )
    gunc_unvan_col = (
            bul_kolon(df_guncel.columns, ["pozisyon"]) or
            bul_kolon(df_guncel.columns, ["unvan"]) or
            bul_kolon(df_guncel.columns, ["görev"])
    )
    gunc_kidem_col = (
            bul_kolon(df_guncel.columns, ["kıdem sayı"]) or
            bul_kolon(df_guncel.columns, ["kıdem yıl"]) or
            bul_kolon(df_guncel.columns, ["kıdem"])
    )
    gunc_yas_col = (
            bul_kolon(df_guncel.columns, ["çalışan yaş"]) or
            bul_kolon(df_guncel.columns, ["yaş"]) or
            bul_kolon(df_guncel.columns, ["yas"])
    )

    gunc_join_col = gunc_masraf_col or gunc_magaza_col
    if gunc_join_col:
        df_guncel = df_guncel.dropna(subset=[gunc_join_col])
        df_guncel = df_guncel[df_guncel[gunc_join_col].astype(str).str.strip() != ""]

        # Mükerrer sütunları temizle (BM, PM, HRBP, Segment, Il, Magaza)
        # Böylece merge sonrası _x, _y takıları oluşmaz ve filtreleme kolonları temiz kalır!
        cols_to_drop_gunc = [c for c in ["BM", "PM", "HRBP", "Segment", "Il", "Magaza"] if
                             c in df_guncel.columns and c != gunc_join_col]
        if cols_to_drop_gunc:
            df_guncel = df_guncel.drop(columns=cols_to_drop_gunc)

        masraf_info = df[["Masraf_Kodu", "BM", "PM", "HRBP", "Segment", "Magaza", "Il"]].copy()
        masraf_info["Masraf_Kodu"] = masraf_info["Masraf_Kodu"].astype(str).str.strip()
        df_guncel["_join_key"] = df_guncel[gunc_join_col].astype(str).str.strip()
        df_guncel = df_guncel.merge(masraf_info, left_on="_join_key", right_on="Masraf_Kodu", how="left")

# ── STİL ────────
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
    .turnover-kutu {
        background-color: #1a2f5e;
        border-radius: 8px;
        padding: 22px;
        margin-bottom: 15px;
        text-align: center;
        border-bottom: 4px solid #85B7EB;
        box-shadow: 0 4px 6px rgba(0,0,0,0.15);
    }
    .turnover-kutu.gonullu  { border-bottom: 4px solid #4a6fa5; }
    .turnover-kutu.gonulsuz { border-bottom: 4px solid #e05c5c; }
    .turnover-kutu.zorunlu  { border-bottom: 4px solid #ffd166; }
</style>
""", unsafe_allow_html=True)

# ── SOL PANEL ────────
with st.sidebar:
    try:
        st.image("koton_siyah.png", width=180)
    except Exception:
        st.markdown("**KOTON**")
    st.markdown("**Karne Verisi - Turnover**")
    st.divider()

    if st.button("🔄 Tüm Filtreleri Temizle", use_container_width=True):
        for k in ["ay_key", "pm_key", "bm_key", "hrbp_key", "segment_key"]:
            default = "Mayıs" if k == "ay_key" else "Tümü"
            st.session_state[k] = default
        for k in ["ay_select", "pm_select", "bm_select", "hrbp_select", "segment_select"]:
            default = "Mayıs" if k == "ay_select" else "Tümü"
            st.session_state[k] = default
        st.rerun()

    # Varsayılan başlangıç ayı sadece "Mayıs" yapıldı
    ay_options = ["Mayıs", "Nisan", "Mart", "Şubat", "Ocak"]
    ay_idx = ay_options.index(st.session_state.ay_key) if st.session_state.ay_key in ay_options else 0
    st.markdown("📆 **2026 Yılı Ayları**")
    sec_ay = st.selectbox("Ay", ay_options, index=ay_idx, key="ay_select", label_visibility="collapsed")
    st.session_state.ay_key = sec_ay

    pm_listesi = ["Tümü"] + sorted(df["PM"].dropna().unique().tolist())
    pm_idx = pm_listesi.index(st.session_state.pm_key) if st.session_state.pm_key in pm_listesi else 0
    st.markdown("🏪 **Perakende Müdürlüğü (PM)**")
    sec_pm = st.selectbox("PM", pm_listesi, index=pm_idx, key="pm_select", label_visibility="collapsed")
    st.session_state.pm_key = sec_pm

    df_temp_pm = df[df["PM"] == sec_pm] if sec_pm != "Tümü" else df.copy()
    bm_listesi = ["Tümü"] + sorted(df_temp_pm["BM"].dropna().unique().tolist())
    bm_idx = bm_listesi.index(st.session_state.bm_key) if st.session_state.bm_key in bm_listesi else 0
    st.markdown("👤 **Bölge Müdürlüğü (BM)**")
    sec_bm = st.selectbox("BM", bm_listesi, index=bm_idx, key="bm_select", label_visibility="collapsed")
    st.session_state.bm_key = sec_bm

    df_temp_bm = df_temp_pm[df_temp_pm["BM"] == sec_bm] if sec_bm != "Tümü" else df_temp_pm.copy()
    hrbp_listesi = ["Tümü"] + sorted(df_temp_bm["HRBP"].dropna().unique().tolist())
    hrbp_idx = hrbp_listesi.index(st.session_state.hrbp_key) if st.session_state.hrbp_key in hrbp_listesi else 0
    st.markdown("🤝 **İK İş Ortağı (HRBP)**")
    sec_hrbp = st.selectbox("HRBP", hrbp_listesi, index=hrbp_idx, key="hrbp_select", label_visibility="collapsed")
    st.session_state.hrbp_key = sec_hrbp

    segment_sirasi = ["FS", "A++", "A+", "A", "B", "C", "D"]
    segment_listesi = ["Tümü"] + [s for s in segment_sirasi if s in df["Segment"].dropna().unique().tolist()]
    segment_idx = segment_listesi.index(
        st.session_state.segment_key) if st.session_state.segment_key in segment_listesi else 0
    st.markdown("🏷️ **Segment**")
    sec_segment = st.selectbox("Segment", segment_listesi, index=segment_idx, key="segment_select",
                               label_visibility="collapsed")
    st.session_state.segment_key = sec_segment

# ── FİLTRELEME ────────
df_f = df.copy()
if sec_bm != "Tümü":      df_f = df_f[df_f["BM"] == sec_bm]
if sec_pm != "Tümü":      df_f = df_f[df_f["PM"] == sec_pm]
if sec_hrbp != "Tümü":    df_f = df_f[df_f["HRBP"] == sec_hrbp]
if sec_segment != "Tümü": df_f = df_f[df_f["Segment"] == sec_segment]

# Sidebar filtresinden gelen mağaza kodları — tüm sekmelerde ortak kullanılır
filtered_masraf_list = df_f["Masraf_Kodu"].astype(str).str.strip().tolist()
filtered_magaza_list = df_f["Magaza"].astype(str).str.strip().tolist()

if sec_ay == "Mayıs":
    cikis26, ort26, to26_col, to25_col = "Mayis26_Cikis", "Mayis26_Ort", "Mayis26_TO", "Mayis25_TO"
    cikis25, ort25 = "Mayis25_Cikis", "Mayis25_Ort"
    ana_baslik26, ana_baslik25 = "2026 Mayıs TO MTD", "2025 Mayıs TO MTD"
    goster_mayis = True
else:
    cikis26, ort26, to26_col, to25_col = ay_map[sec_ay]
    cikis25 = cikis26.replace("26", "25")
    ort25 = ort26.replace("26", "25")
    ana_baslik26 = f"2026 {sec_ay} TO"
    ana_baslik25 = f"2025 {sec_ay} TO"
    goster_mayis = False

ytd26_cikis, ytd26_ort = "YTD26_Cikis", "YTD26_Ort"
ytd25_cikis, ytd25_ort = "YTD25_Cikis", "YTD25_Ort"
ytd26_col, ytd25_col = "YTD26_TO", "YTD25_TO"

kpi26 = genel_to(df_f, cikis26, ort26)
kpi25 = genel_to(df_f, cikis25, ort25)
kpi_ytd26 = genel_to(df_f, ytd26_cikis, ytd26_ort)
kpi_ytd25 = genel_to(df_f, ytd25_cikis, ytd25_ort)
kpi_may26 = genel_to(df_f, "Mayis26_Cikis", "Mayis26_Ort")
kpi_may25 = genel_to(df_f, "Mayis25_Cikis", "Mayis25_Ort")

# ── SEKMELER ────────
# 6. Sekme olarak Güncel Çalışanlar eklendi
sekme1, sekme2, sekme3, sekme4, sekme5, sekme6 = st.tabs([
    "📊 Genel Performans & Trendler",
    "📅 Mayıs Detay Analizi",
    "🏪 Mağaza Detay Analiz Kartı",
    "⚖️ Yasal Bildirimler",
    "🚪 İşten Ayrılanlar",
    "👥 Güncel Çalışanlar"
])

# ══════════════════════════════════════════════
# SEKME 1 — Genel Performans
# ══════════════════════════════════════════════
with sekme1:
    st.markdown("### Filtrelenmiş Mağaza Performans KPI Özetleri")

    st.markdown("#### 🚀 2026 Yılı Dönemsel Performans")
    if goster_mayis:
        k1, k2, k3 = st.columns(3)
        with k1:
            kpi_nis26 = genel_to(df_f, "Nis26_Cikis", "Nis26_Ort")
            kpi_nis25 = genel_to(df_f, "Nis25_Cikis", "Nis25_Ort")
            d = round(kpi_nis26 - kpi_nis25, 1)
            st.metric("📅 2026 Nisan TO", f"%{kpi_nis26:.1f}",
                      delta=f"%{d:.1f} (vs 2025 Nisan)", delta_color="normal" if d <= 0 else "inverse")
        with k2:
            d = round(kpi_may26 - kpi_may25, 1)
            st.metric("📆 2026 Mayıs TO MTD", f"%{kpi_may26:.1f}",
                      delta=f"%{d:.1f} (vs 2025 Mayıs)", delta_color="normal" if d <= 0 else "inverse")
        with k3:
            d = round(kpi_ytd26 - kpi_ytd25, 1)
            st.metric("📊 2026 YTD Toplam TO", f"%{kpi_ytd26:.1f}",
                      delta=f"%{d:.1f} (vs 2025 YTD)", delta_color="normal" if d <= 0 else "inverse")
    else:
        k1, k2 = st.columns(2)
        with k1:
            d = round(kpi26 - kpi25, 1)
            st.metric(f"📅 {ana_baslik26}", f"%{kpi26:.1f}",
                      delta=f"%{d:.1f} (vs 2025)", delta_color="normal" if d <= 0 else "inverse")
        with k2:
            d = round(kpi_ytd26 - kpi_ytd25, 1)
            st.metric("📊 2026 YTD Toplam TO", f"%{kpi_ytd26:.1f}",
                      delta=f"%{d:.1f} (vs 2025)", delta_color="normal" if d <= 0 else "inverse")

    st.markdown("#### 🌐 2025 Yılı Dönemsel Performans")
    if goster_mayis:
        k4, k5, k6 = st.columns(3)
        with k4:
            st.metric("📅 2025 Nisan TO", f"%{kpi_nis25:.1f}")
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
    st.markdown("### 🚨 En Yüksek YTD Turnover'a Sahip Mağazalar & Risk Haritası")

    df_top5 = df_f.nlargest(5, "YTD26_TO")
    if not df_top5.empty:
        df_top5_sorted = df_top5.sort_values("YTD26_TO", ascending=True)
        colors_top5 = ["#e05c5c" if v > 30 else "#ffd166" if v > 20 else "#85B7EB"
                       for v in df_top5_sorted["YTD26_TO"]]

        fig_top5 = go.Figure()
        fig_top5.add_trace(go.Bar(
            y=df_top5_sorted["Magaza"], x=df_top5_sorted["YTD26_TO"],
            orientation='h', marker_color=colors_top5,
            text=[f"%{v:.1f}" for v in df_top5_sorted["YTD26_TO"]],
            textposition='auto'
        ))
        fig_top5.update_layout(
            paper_bgcolor="#1a2f5e", plot_bgcolor="#1a2f5e", font=dict(color="white"),
            xaxis=dict(title="Turnover Oranı (%)", gridcolor="#2e3f7a"),
            yaxis=dict(gridcolor="#2e3f7a"),
            margin=dict(l=20, r=20, t=10, b=20), height=280
        )

        chart_col, table_col = st.columns([5, 4])
        with chart_col:
            st.markdown("#### 📊 Mağaza Turnover Risk Dağılımı")
            st.plotly_chart(fig_top5, use_container_width=True)
        with table_col:
            st.markdown("#### 📋 Kritik Risk Grubu Detay Listesi")
            top5_table = df_top5.copy()
            top5_table["Risk Durumu"] = top5_table["YTD26_TO"].apply(
                lambda x: "🔴 Yüksek (>%30)" if x > 30 else ("🟡 Normal (%20-30)" if x > 20 else "🟢 Düşük (<=%20)")
            )
            top5_table["Turnover (YTD)"] = top5_table["YTD26_TO"].apply(lambda x: f"%{x:.1f}")
            top5_goster = top5_table[["Masraf_Kodu", "Magaza", "Segment", "Turnover (YTD)", "Risk Durumu"]].copy()
            top5_goster.columns = ["Kritik Kod", "Mağaza Adı", "Segment", "YTD TO %", "Risk Durumu"]
            st.dataframe(top5_goster, use_container_width=True, hide_index=True)
    else:
        st.info("Filtrelere uygun mağaza verisi bulunamadı.")

    st.divider()
    st.markdown("### 🏷️ Segment Kırılımına Göre Turnover Oranları")

    segment_sirasi_map = ["FS", "A++", "A+", "A", "B", "C", "D"]
    unique_segments = [s for s in segment_sirasi_map if s in df_f["Segment"].dropna().unique()]
    segment_to_list = []
    for seg in unique_segments:
        df_seg_temp = df_f[df_f["Segment"] == seg]
        seg_cikis = pd.to_numeric(df_seg_temp[cikis26], errors="coerce").sum()
        seg_ort = pd.to_numeric(df_seg_temp[ort26], errors="coerce").sum()
        seg_to = round((seg_cikis / seg_ort) * 100, 1) if seg_ort > 0 else 0.0
        segment_to_list.append((seg, seg_to))

    df_segment_to = pd.DataFrame(segment_to_list, columns=["Segment", "Turnover Oranı"])
    if not df_segment_to.empty:
        fig_seg = go.Figure()
        fig_seg.add_trace(go.Bar(
            x=df_segment_to["Segment"], y=df_segment_to["Turnover Oranı"],
            marker_color="#4a6fa5",
            text=[f"%{v:.1f}" for v in df_segment_to["Turnover Oranı"]],
            textposition='auto'
        ))
        fig_seg.update_layout(
            paper_bgcolor="#1a2f5e", plot_bgcolor="#1a2f5e", font=dict(color="white"),
            xaxis=dict(title="Segment"),
            yaxis=dict(title="Turnover Oranı (%)", gridcolor="#2e3f7a"),
            margin=dict(l=20, r=20, t=10, b=20), height=280
        )
        seg_chart_col, seg_table_col = st.columns([5, 4])
        with seg_chart_col:
            ay_label = sec_ay
            st.markdown(f"#### 📊 Segment Turnover Dağılımı ({ay_label})")
            st.plotly_chart(fig_seg, use_container_width=True)
        with seg_table_col:
            st.markdown("#### 📋 Segment Oran Tablosu")
            df_seg_goster = df_segment_to.copy()
            df_seg_goster["Turnover Oranı"] = df_seg_goster["Turnover Oranı"].apply(lambda x: f"%{x:.1f}")
            df_seg_goster.columns = ["Segment Sınıfı", "Segment Turnover %"]
            st.dataframe(df_seg_goster, use_container_width=True, hide_index=True)

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

    if sec_ay == "Mayıs":
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
# (Sidebar filtresine göre mağaza listesi daralır)
# ── Mayıs Detay Analizi ────────
with sekme2:
    st.markdown("### 📅 Mayıs Detay Veri Analizi")
    st.markdown("Mağaza ve unvan seçimi yaparak dönem içi hareketlerini ve turnover oranlarını inceleyebilirsiniz.")

    if df_detay.empty or magaza_col_detay is None:
        st.warning("⚠️ Detay sayfası yüklenemedi. Lütfen Excel dosyasını kontrol edin.")
    else:
        # Sidebar filtresine göre detay sayfasını kısıtla
        detay_col_str = df_detay[magaza_col_detay].astype(str).str.strip()
        # Masraf kodu veya mağaza adı bazında eşleştir
        mask_detay = detay_col_str.isin(filtered_masraf_list) | detay_col_str.isin(filtered_magaza_list)
        df_detay_f = df_detay[mask_detay] if mask_detay.any() else df_detay

        st.markdown("🔍 **Mağaza Seçimi**")
        magaza_listesi_detay = sorted(df_detay_f[magaza_col_detay].astype(str).unique().tolist())
        if not magaza_listesi_detay:
            st.info("Seçili filtreye uygun mağaza bulunamadı.")
        else:
            sec_magaza_detay = st.selectbox("Mağaza (Detay)", magaza_listesi_detay, label_visibility="collapsed")

            if sec_magaza_detay:
                try:
                    s_detay = df_detay[df_detay[magaza_col_detay].astype(str) == sec_magaza_detay].iloc[0]

                    # Full-Time / Part-Time sütunlarını bul
                    ft_tum_col = ft_gonullu_col = pt_tum_col = pt_gonullu_col = None
                    for col in df_detay.columns:
                        cs = str(col).strip().lower()
                        if "sayısı" in cs or "sayisi" in cs:
                            continue
                        if "full time" in cs or "fulltime" in cs or "full" in cs:
                            if "tüm" in cs or "tum" in cs:
                                ft_tum_col = col
                            elif "gönüllü" in cs or "gonullu" in cs:
                                ft_gonullu_col = col
                        elif "part-time" in cs or "part time" in cs or "parttime" in cs or "part" in cs:
                            if "tüm" in cs or "tum" in cs:
                                pt_tum_col = col
                            elif "gönüllü" in cs or "gonullu" in cs:
                                pt_gonullu_col = col

                    if not ft_tum_col or not pt_tum_col:
                        tum_all = [c for c in df_detay.columns
                                   if ("tüm" in str(c).lower() or "tum" in str(c).lower())
                                   and ("turnover" in str(c).lower() or "to" in str(c).lower())
                                   and not any(u in str(c).lower() for u in
                                               ["müdür", "vm", "sorumlu", "satış", "danışman", "mmy",
                                                "yardım", "sayısı", "sayisi", "çıkış", "cikis"])]
                        gonullu_all = [c for c in df_detay.columns
                                       if "gönüllü" in str(c).lower()
                                       and ("turnover" in str(c).lower() or "to" in str(c).lower())
                                       and not any(u in str(c).lower() for u in
                                                   ["müdür", "vm", "sorumlu", "satış", "danışman", "mmy",
                                                    "yardım", "sayısı", "sayisi", "çıkış", "cikis"])]
                        if len(tum_all) >= 2:
                            ft_tum_col, pt_tum_col = tum_all[1], tum_all[0]
                        if len(gonullu_all) >= 2:
                            ft_gonullu_col, pt_gonullu_col = gonullu_all[1], gonullu_all[0]

                    val_ft_tum = oran_formatla(s_detay[ft_tum_col]) if ft_tum_col else "—"
                    val_ft_gonullu = oran_formatla(s_detay[ft_gonullu_col]) if ft_gonullu_col else "—"
                    val_pt_tum = oran_formatla(s_detay[pt_tum_col]) if pt_tum_col else "—"
                    val_pt_gonullu = oran_formatla(s_detay[pt_gonullu_col]) if pt_gonullu_col else "—"

                    st.markdown("#### 🎯 Mağaza İçi İstihdam Dağılım Oranları")
                    kpi_col1, kpi_col2 = st.columns(2)

                    with kpi_col1:
                        st.markdown(f"""
                        <div class="magaza-kart" style="border-left:4px solid #85B7EB; padding:18px; margin:0;">
                            <p style="color:#aaaacc; font-size:12px; margin-bottom:5px; font-weight:bold;">GRUP İÇİ FULL TIME</p>
                            <div style="display:flex; justify-content:space-around; margin-top:15px;">
                                <div style="text-align:center; flex:1; border-right:1px solid #2e3f7a;">
                                    <p style="color:#85B7EB; font-size:11px; margin:0;">Tüm Turnover</p>
                                    <h4 style="color:white; margin:6px 0 0 0; font-size:22px; font-weight:bold;">{val_ft_tum}</h4>
                                </div>
                                <div style="text-align:center; flex:1;">
                                    <p style="color:#ffd166; font-size:11px; margin:0;">Gönüllü Turnover</p>
                                    <h4 style="color:white; margin:6px 0 0 0; font-size:22px; font-weight:bold;">{val_ft_gonullu}</h4>
                                </div>
                            </div>
                        </div>""", unsafe_allow_html=True)

                    with kpi_col2:
                        st.markdown(f"""
                        <div class="magaza-kart" style="border-left:4px solid #4a6fa5; padding:18px; margin:0;">
                            <p style="color:#aaaacc; font-size:12px; margin-bottom:5px; font-weight:bold;">GRUP İÇİ PART-TIME</p>
                            <div style="display:flex; justify-content:space-around; margin-top:15px;">
                                <div style="text-align:center; flex:1; border-right:1px solid #2e3f7a;">
                                    <p style="color:#85B7EB; font-size:11px; margin:0;">Tüm Turnover</p>
                                    <h4 style="color:white; margin:6px 0 0 0; font-size:22px; font-weight:bold;">{val_pt_tum}</h4>
                                </div>
                                <div style="text-align:center; flex:1;">
                                    <p style="color:#ffd166; font-size:11px; margin:0;">Gönüllü Turnover</p>
                                    <h4 style="color:white; margin:6px 0 0 0; font-size:22px; font-weight:bold;">{val_pt_gonullu}</h4>
                                </div>
                            </div>
                        </div>""", unsafe_allow_html=True)

                    st.divider()

                    # ── UNVAN FİLTRESİ ────────
                    st.markdown("👤 **Unvan Filtresi**")
                    sec_unvan = st.selectbox("Unvan", list(unvan_secenekleri.keys()), label_visibility="collapsed")
                    keys, exclude = unvan_secenekleri[sec_unvan]

                    st.divider()

                    # Oran sütunlarını kesinlikle dışarıda bırakmak için sabit terimler
                    oran_haric = ["oran", "turnover", "to", "%", "rate", "yüzde"]
                    haric_genel_sayi = ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"] + oran_haric

                    if sec_unvan == "Tümü (Toplam)":
                        db_col = (
                                sayi_col_bul(df_detay.columns, ["dönem başı"],
                                             ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"]) or
                                sayi_col_bul(df_detay.columns, ["baş", "çalışan"],
                                             ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"])
                        )
                        gonullu_col = (
                                sayi_col_bul(df_detay.columns, ["gönüllü", "çıkış"],
                                             ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"]) or
                                sayi_col_bul(df_detay.columns, ["gönüllü"],
                                             ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"])
                        )
                        gonulsuz_col = (
                                sayi_col_bul(df_detay.columns, ["gönülsüz", "çıkış"],
                                             ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"]) or
                                sayi_col_bul(df_detay.columns, ["gönülsüz"],
                                             ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"])
                        )
                        zorunlu_col = (
                                sayi_col_bul(df_detay.columns, ["zorunlu", "çıkış"],
                                             ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"]) or
                                sayi_col_bul(df_detay.columns, ["zorunlu"],
                                             ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"])
                        )
                        ds_col = (
                                sayi_col_bul(df_detay.columns, ["dönem sonu"],
                                             ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"]) or
                                sayi_col_bul(df_detay.columns, ["son", "çalışan"],
                                             ["müdür", "vm", "sorumlu", "satış", "yardım", "yrd", "mmy"])
                        )
                    else:
                        db_col = (
                                sayi_col_bul(df_detay.columns, ["dönem başı"] + keys, exclude) or
                                sayi_col_bul(df_detay.columns, ["baş", "çalışan"] + keys, exclude)
                        )
                        gonullu_col = (
                                sayi_col_bul(df_detay.columns, ["gönüllü", "çıkış"] + keys, exclude) or
                                sayi_col_bul(df_detay.columns, ["gönüllü"] + keys, exclude)
                        )
                        gonulsuz_col = (
                                sayi_col_bul(df_detay.columns, ["gönülsüz", "çıkış"] + keys, exclude) or
                                sayi_col_bul(df_detay.columns, ["gönülsüz"] + keys, exclude)
                        )
                        zorunlu_col = (
                                sayi_col_bul(df_detay.columns, ["zorunlu", "çıkış"] + keys, exclude) or
                                sayi_col_bul(df_detay.columns, ["zorunlu"] + keys, exclude)
                        )
                        ds_col = (
                                sayi_col_bul(df_detay.columns, ["dönem sonu"] + keys, exclude) or
                                sayi_col_bul(df_detay.columns, ["son", "çalışan"] + keys, exclude)
                        )


                    def fmt_sayi(val):
                        if val is None or (isinstance(val, float) and pd.isna(val)):
                            return "—"
                        try:
                            return int(float(val))
                        except Exception:
                            return str(val)


                    # Çıkış sayılarını temiz tamsayılara çevirip toplamak için güvenli metot
                    def to_int_safe(val):
                        if val is None or (isinstance(val, float) and pd.isna(val)) or val == "—":
                            return 0
                        try:
                            return int(float(val))
                        except Exception:
                            return 0


                    val_db = fmt_sayi(s_detay[db_col]) if db_col else "—"
                    val_gonullu = fmt_sayi(s_detay[gonullu_col]) if gonullu_col else "—"
                    val_gonulsuz = fmt_sayi(s_detay[gonulsuz_col]) if gonulsuz_col else "—"
                    val_zorunlu = fmt_sayi(s_detay[zorunlu_col]) if zorunlu_col else "—"
                    val_ds = fmt_sayi(s_detay[ds_col]) if ds_col else "—"

                    # Gönüllü, gönülsüz ve zorunlu çıkışları toplayarak toplam çıkışı hesaplıyoruz
                    int_gonullu = to_int_safe(s_detay[gonullu_col]) if gonullu_col else 0
                    int_gonulsuz = to_int_safe(s_detay[gonulsuz_col]) if gonulsuz_col else 0
                    int_zorunlu = to_int_safe(s_detay[zorunlu_col]) if zorunlu_col else 0
                    val_cikis = int_gonullu + int_gonulsuz + int_zorunlu

                    # "İşe Giriş" satırı kaldırıldı ve "İşten Çıkış Sayısı" otomatik toplanarak yazıldı
                    tablo1_verileri = [
                        ("Dönem Başı Çalışan Sayısı", val_db),
                        ("İşten Çıkış Sayısı", val_cikis),
                        ("Gönüllü Çıkış Sayısı", val_gonullu),
                        ("Gönülsüz Çıkış Sayısı", val_gonulsuz),
                        ("Zorunlu Çıkış Sayısı", val_zorunlu),
                        ("Dönem Sonu Çalışan Sayısı", val_ds),
                    ]
                    df_tablo1 = pd.DataFrame(tablo1_verileri, columns=["Metrik", "Sayı / Değer"])

                    # Turnover oranları
                    to_tum_col = (bul_kolon(df_detay.columns, ["tüm", "turnover"]) or
                                  bul_kolon(df_detay.columns, ["turnover"],
                                            haric_tutulacaklar=["gönüllü", "gönülsüz", "zorunlu"]))
                    to_gonullu_col = (bul_kolon(df_detay.columns, ["gönüllü", "turnover"]) or
                                      bul_kolon(df_detay.columns, ["gönüllü", "to"]))
                    to_gonulsuz_col = (bul_kolon(df_detay.columns, ["gönülsüz", "turnover"]) or
                                       bul_kolon(df_detay.columns, ["gönülsüz", "to"]))
                    to_zorunlu_col = (bul_kolon(df_detay.columns, ["zorunlu", "turnover"]) or
                                      bul_kolon(df_detay.columns, ["zorunlu", "to"]))

                    to_tum_val = oran_formatla(s_detay[to_tum_col]) if to_tum_col else "—"
                    to_gonullu_val = oran_formatla(s_detay[to_gonullu_col]) if to_gonullu_col else "—"
                    to_gonulsuz_val = oran_formatla(s_detay[to_gonulsuz_col]) if to_gonulsuz_col else "—"
                    to_zorunlu_val = oran_formatla(s_detay[to_zorunlu_col]) if to_zorunlu_col else "—"

                    t_col1, t_col2 = st.columns([3, 2])
                    with t_col1:
                        st.markdown(f"#### 📊 Dönem İçi Hareketler ({sec_unvan})")
                        st.dataframe(df_tablo1, use_container_width=True, hide_index=True, height=290)
                    with t_col2:
                        st.markdown(f"#### 📉 Turnover Oranları ({sec_unvan})")
                        r1c1, r1c2 = st.columns(2)
                        with r1c1:
                            st.markdown(f"""<div class="turnover-kutu">
                                <p style="color:#aaaacc;font-size:11px;margin:0 0 4px 0;font-weight:bold;">TÜM TURNOVER</p>
                                <h3 style="color:white;margin:0;font-size:24px;">{to_tum_val}</h3></div>""",
                                        unsafe_allow_html=True)
                        with r1c2:
                            st.markdown(f"""<div class="turnover-kutu gonullu">
                                <p style="color:#aaaacc;font-size:11px;margin:0 0 4px 0;font-weight:bold;">GÖNÜLLÜ TURNOVER</p>
                                <h3 style="color:white;margin:0;font-size:24px;">{to_gonullu_val}</h3></div>""",
                                        unsafe_allow_html=True)
                        r2c1, r2c2 = st.columns(2)
                        with r2c1:
                            st.markdown(f"""<div class="turnover-kutu gonulsuz">
                                <p style="color:#aaaacc;font-size:11px;margin:0 0 4px 0;font-weight:bold;">GÖNÜLSÜZ TURNOVER</p>
                                <h3 style="color:white;margin:0;font-size:24px;">{to_gonulsuz_val}</h3></div>""",
                                        unsafe_allow_html=True)
                        with r2c2:
                            st.markdown(f"""<div class="turnover-kutu zorunlu">
                                <p style="color:#aaaacc;font-size:11px;margin:0 0 4px 0;font-weight:bold;">ZORUNLU TURNOVER</p>
                                <h3 style="color:white;margin:0;font-size:24px;">{to_zorunlu_val}</h3></div>""",
                                        unsafe_allow_html=True)
                        st.info(f"💡 **{sec_magaza_detay}** mağazası — **{sec_unvan}** verisi")

                        # Hata ayıklama: sütun eşleşmesini göster
                        with st.expander("🛠️ Sütun Eşleşme Detayı"):
                            st.write({
                                "Dönem Başı": db_col,
                                "Gönüllü": gonullu_col,
                                "Gönülsüz": gonulsuz_col,
                                "Zorunlu": zorunlu_col,
                                "Dönem Sonu": ds_col,
                            })

                except Exception as e:
                    st.error(f"❌ Veri okuma hatası: {str(e)}")
                    st.info("Lütfen Excel başlıklarını kontrol edin.")
                    with st.expander("🛠️ Teknik Hata Detayı"):
                        st.text(traceback.format_exc())

# ══════════════════════════════════════════════
# SEKME 3 — Mağaza Detay Analiz Kartı
# ══════════════════════════════════════════════
with sekme3:
    st.markdown("### Detaylı Mağaza Karnesi")

    col_il, col_mag = st.columns([1, 3])
    with col_il:
        st.markdown("📍 **İl Filtresi**")
        il_listesi = ["Tümü"] + sorted(df_f["Il"].dropna().unique().tolist())
        sec_il = st.selectbox("İl", il_listesi, key="il_detay_card", label_visibility="collapsed")

    df_magaza = df_f.copy()
    if sec_il != "Tümü":
        df_magaza = df_magaza[df_magaza["Il"] == sec_il]

    with col_mag:
        st.markdown("🔍 **İncelenecek Mağazayı Seçin**")
        magaza_listesi = list(df_magaza["Masraf_Kodu"] + " — " + df_magaza["Magaza"])
        if not magaza_listesi:
            st.warning("Seçili filtreye uygun mağaza bulunamadı.")
            st.stop()
        sec_magaza = st.selectbox("Mağaza", magaza_listesi, key="magaza_detay_card", label_visibility="collapsed")

    if sec_magaza:
        kod = sec_magaza.split(" — ")[0].strip()
        s = df[df["Masraf_Kodu"] == kod].iloc[0]

        # Önceki Dönem HRBP değeri var ise çekelim
        onceki_hrbp_val = s["Onceki_Donem_HRBP"] if "Onceki_Donem_HRBP" in s.index else "—"
        if pd.isna(onceki_hrbp_val) or onceki_hrbp_val == "" or onceki_hrbp_val == "nan":
            onceki_hrbp_val = "—"

        st.markdown(f"""
        <div class="magaza-kart">
            <p style="color:#aaaacc; font-size:12px; margin-bottom:8px;">MAĞAZA BİLGİLERİ</p>
            <h2 style="color:white; margin:0;">{s['Masraf_Kodu']} — {s['Magaza']}</h2><br>
            <span style="color:#e0e0e0; margin-right:24px;">📍 İl: <b>{s['Il']}</b></span>
            <span style="color:#e0e0e0; margin-right:24px;">➕ Segment: <b>{s['Segment']}</b></span>
            <br><br>
            <span style="color:#e0e0e0; margin-right:24px;">👤 BM: <b>{s['BM']}</b></span>
            <span style="color:#e0e0e0; margin-right:24px;">🏪 PM: <b>{s['PM']}</b></span>
            <br><br>
            <span style="color:#e0e0e0; margin-right:24px;">🤝 HRBP: <b>{s['HRBP']}</b></span>
            <span style="color:#e0e0e0;">⏪ Önceki Dönem HRBP: <b>{onceki_hrbp_val}</b></span>
        </div>""", unsafe_allow_html=True)

        col_sol, col_sag = st.columns(2)
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
                        st.markdown(f"<p style='color:#aaaacc;font-size:12px;'>{emoji} {etiket}</p>",
                                    unsafe_allow_html=True)

                st.markdown("#### 🌐 2025 Turnover Seviyeleri")
                k4, k5, k6 = st.columns(3)
                for kart, ay, to_val in zip([k4, k5, k6],
                                            ["Nisan 2025", "Mayıs 2025", "YTD 2025"],
                                            [s["Nis25_TO"], s["Mayis25_TO"], s["YTD25_TO"]]):
                    with kart:
                        etiket, emoji = risk_etiketi(to_val)
                        st.metric(ay, f"%{to_val:.1f}" if not pd.isna(to_val) else "—")
                        st.markdown(f"<p style='color:#aaaacc;font-size:12px;'>{emoji} {etiket}</p>",
                                    unsafe_allow_html=True)
            else:
                st.markdown(f"#### 🚀 {sec_ay} Karşılaştırması")
                k1, k2 = st.columns(2)
                for kart, ay, to_val in zip([k1, k2], [f"{sec_ay} 2026", f"{sec_ay} 2025"], [s[to26_col], s[to25_col]]):
                    with kart:
                        etiket, emoji = risk_etiketi(to_val)
                        st.metric(ay, f"%{to_val:.1f}" if not pd.isna(to_val) else "—")
                        st.markdown(f"<p style='color:#aaaacc;font-size:12px;'>{emoji} {etiket}</p>",
                                    unsafe_allow_html=True)
                st.markdown("#### 📊 YTD 2026")
                etiket, emoji = risk_etiketi(s["YTD26_TO"])
                st.metric("YTD 2026", f"%{s['YTD26_TO']:.1f}" if not pd.isna(s["YTD26_TO"]) else "—")
                st.markdown(f"<p style='color:#aaaacc;font-size:12px;'>{emoji} {etiket}</p>", unsafe_allow_html=True)

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

        st.divider()
        st.markdown(f"#### 🏷️ Segment Karşılaştırması — {s['Segment']}")
        df_seg = df[df["Segment"] == s["Segment"]]
        seg_nis26 = genel_to(df_seg, "Nis26_Cikis", "Nis26_Ort")
        seg_may26 = genel_to(df_seg, "Mayis26_Cikis", "Mayis26_Ort")
        seg_ytd26 = genel_to(df_seg, "YTD26_Cikis", "YTD26_Ort")
        seg_nis25 = genel_to(df_seg, "Nis25_Cikis", "Nis25_Ort")
        seg_may25 = genel_to(df_seg, "Mayis25_Cikis", "Mayis25_Ort")
        seg_ytd25 = genel_to(df_seg, "YTD25_Cikis", "YTD25_Ort")

        st.markdown(f"""<div class="segment-kart">
            <p style="color:#85B7EB; font-weight:bold; margin-bottom:12px;">
            Segment Ortalaması ({s['Segment']}) — {len(df_seg)} Mağaza</p></div>""",
                    unsafe_allow_html=True)

        sc1, sc2, sc3 = st.columns(3)
        for kart, baslik, mv, sv in zip(
                [sc1, sc2, sc3],
                ["📅 Nisan 2026", "📆 Mayıs 2026", "📊 YTD 2026"],
                [s["Nis26_TO"], s["Mayis26_TO"], s["YTD26_TO"]],
                [seg_nis26, seg_may26, seg_ytd26]
        ):
            with kart:
                mag_str = f"%{mv:.1f}" if not pd.isna(mv) else "—"
                delta = round(mv - sv, 1) if not pd.isna(mv) else 0
                delta_str = f"-{abs(delta):.1f} (Seg. Ort: %{sv:.1f})" if delta < 0 else f"+{delta:.1f} (Seg. Ort: %{sv:.1f})"
                st.metric(baslik, mag_str, delta=delta_str, delta_color="normal")

        sc4, sc5, sc6 = st.columns(3)
        for kart, baslik, mv, sv in zip(
                [sc4, sc5, sc6],
                ["📅 Nisan 2025", "📆 Mayıs 2025", "📊 YTD 2025"],
                [s["Nis25_TO"], s["Mayis25_TO"], s["YTD25_TO"]],
                [seg_nis25, seg_may25, seg_ytd25]
        ):
            with kart:
                mag_str = f"%{mv:.1f}" if not pd.isna(mv) else "—"
                delta = round(mv - sv, 1) if not pd.isna(mv) else 0
                st.metric(baslik, mag_str, delta=f"%{delta:.1f} (Seg. Ort: %{sv:.1f})", delta_color="inverse")

# ── YASAL BİLDİRİMLER ────────
with sekme4:
    st.markdown("### ⚖️ Yasal Bildirim Takip Analizi")

    if df_yasal.empty:
        st.warning("⚠️ Yasal Bildirim sayfası yüklenemedi.")
    else:
        # Sidebar filtresine göre yasal tabloyu kısıtla
        df_yasal_f = df_yasal[df_yasal["Masraf_Kodu_Temiz"].isin(filtered_masraf_list)]
        if df_yasal_f.empty:
            df_yasal_f = df_yasal

        st.markdown("🔍 **Mağaza Seçimi**")
        magaza_listesi_yasal = sorted(df_yasal_f["Display_Magaza"].astype(str).unique().tolist())
        if not magaza_listesi_yasal:
            st.info("Seçili filtreye uygun mağaza bulunamadı.")
        else:
            sec_magaza_yasal = st.selectbox("Mağaza (Yasal)", magaza_listesi_yasal, label_visibility="collapsed")

            if sec_magaza_yasal:
                try:
                    df_filtered_yasal = df_yasal[df_yasal["Display_Magaza"] == sec_magaza_yasal].copy().reset_index(
                        drop=True)

                    cikis_tarih_col = (bul_kolon(df_yasal.columns, ["çıkış", "tarih"]) or
                                       bul_kolon(df_yasal.columns, ["çıkış"]))
                    bildirim_tarih_col = (bul_kolon(df_yasal.columns, ["bildirim", "tarih"]) or
                                          bul_kolon(df_yasal.columns, ["bildirim"]))
                    fark_gun_col = (bul_kolon(df_yasal.columns, ["fark", "gün"]) or
                                    bul_kolon(df_yasal.columns, ["fark"]))
                    ad_soyad_col = (bul_kolon(df_yasal.columns, ["ad", "soyad"]) or
                                    bul_kolon(df_yasal.columns, ["çalışan"]) or
                                    bul_kolon(df_yasal.columns, ["isim"]))
                    sicil_col = (bul_kolon(df_yasal.columns, ["sicil"]) or
                                 bul_kolon(df_yasal.columns, ["no"]))

                    yasal_cols_raw = []
                    rename_dict = {}
                    for col, label in [
                        (sicil_col, "Sicil No"),
                        (ad_soyad_col, "Ad Soyad"),
                        (cikis_tarih_col, "Çıkış Tarihi"),
                        (bildirim_tarih_col, "Bildirim Tarihi"),
                        (fark_gun_col, "Fark Gün"),
                    ]:
                        if col:
                            yasal_cols_raw.append(col)
                            rename_dict[col] = label

                    yasal_cols = list(dict.fromkeys(yasal_cols_raw))
                    if not yasal_cols:
                        yasal_cols = df_yasal.columns.tolist()

                    for col in [cikis_tarih_col, bildirim_tarih_col]:
                        if col and col in df_filtered_yasal.columns:
                            try:
                                df_filtered_yasal[col] = pd.to_datetime(df_filtered_yasal[col]).dt.strftime('%d.%m.%Y')
                            except Exception:
                                pass

                    if fark_gun_col and fark_gun_col in df_filtered_yasal.columns:
                        df_filtered_yasal[fark_gun_col] = (
                            pd.to_numeric(df_filtered_yasal[fark_gun_col], errors="coerce").fillna(0).astype(int)
                        )

                    if df_filtered_yasal.empty:
                        st.success("🎉 Bu mağaza için geç bildirilen yasal çıkış bulunmamaktadır.")
                    else:
                        st.warning(
                            f"⚠️ Toplam **{len(df_filtered_yasal)}** adet geç bildirilen yasal çıkış tespit edildi.")
                        df_yasal_goster = df_filtered_yasal[yasal_cols].rename(columns=rename_dict)
                        st.dataframe(df_yasal_goster, use_container_width=True, hide_index=True)

                except Exception as e:
                    st.error(f"❌ Yasal bildirimler hatası: {str(e)}")
                    with st.expander("🛠️ Teknik Hata Detayı"):
                        st.text(traceback.format_exc())

# ══════════════════════════════════════════════
# SEKME 5 — İşten Ayrılanlar
# ══════════════════════════════════════════════
with sekme5:
    st.markdown("### 🚪 İşten Ayrılanlar Listesi")

    if df_ayrilanlar.empty:
        st.warning("⚠️ 'İşten Ayrılanlar' sayfası Excel dosyasında bulunamadı. Lütfen sayfa adını kontrol edin.")
        with st.expander("🛠️ Mevcut sayfa isimleri"):
            st.write(sayfa_isimleri)
    else:
        # Sidebar filtresini uygula (BM/PM/HRBP/Segment merge edilmişse)
        df_ayr_f = df_ayrilanlar.copy()
        if "BM" in df_ayr_f.columns and sec_bm != "Tümü":
            df_ayr_f = df_ayr_f[df_ayr_f["BM"] == sec_bm]
        if "PM" in df_ayr_f.columns and sec_pm != "Tümü":
            df_ayr_f = df_ayr_f[df_ayr_f["PM"] == sec_pm]
        if "HRBP" in df_ayr_f.columns and sec_hrbp != "Tümü":
            df_ayr_f = df_ayr_f[df_ayr_f["HRBP"] == sec_hrbp]
        if "Segment" in df_ayr_f.columns and sec_segment != "Tümü":
            df_ayr_f = df_ayr_f[df_ayr_f["Segment"] == sec_segment]

        # GÖNÜLLÜ, GÖNÜLSÜZ VE ZORUNLU ÇIKIŞ ÖZET KUTUCUKLARI (Ay filtresine tam bağlı olması sağlandı!)
        # Çıkış Tarihi sütununu kullanarak sadece aktif seçili ay verisini süzüyoruz
        ay_num_map = {"Ocak": 1, "Şubat": 2, "Mart": 3, "Nisan": 4, "Mayıs": 5}
        if ayr_tarih_col and ayr_tarih_col in df_ayr_f.columns:
            df_ayr_f["_temp_dt"] = pd.to_datetime(df_ayr_f[ayr_tarih_col], dayfirst=True, errors="coerce")
            selected_month_num = ay_num_map.get(sec_ay)
            if selected_month_num:
                df_ayr_f = df_ayr_f[
                    (df_ayr_f["_temp_dt"].dt.month == selected_month_num) &
                    (df_ayr_f["_temp_dt"].dt.year == 2026)
                    ]

        c_gonullu = 0
        c_gonulsuz = 0
        c_zorunlu = 0
        if ayr_to_neden_col and ayr_to_neden_col in df_ayr_f.columns:
            vals_neden = df_ayr_f[ayr_to_neden_col].astype(str).str.strip().str.lower()
            c_gonullu = vals_neden.str.contains("gönüllü|gonullu").sum()
            c_gonulsuz = vals_neden.str.contains("gönülsüz|gonulsuz").sum()
            c_zorunlu = vals_neden.str.contains("zorunlu").sum()

        col_g1, col_g2, col_g3 = st.columns(3)
        with col_g1:
            st.metric("🟢 Toplam Gönüllü Çıkış", f"{c_gonullu} Kişi")
        with col_g2:
            st.metric("🔴 Toplam Gönülsüz Çıkış", f"{c_gonulsuz} Kişi")
        with col_g3:
            st.metric("🟡 Toplam Zorunlu Çıkış", f"{c_zorunlu} Kişi")

        st.divider()

        ara_col, mag_col = st.columns([2, 2])
        with ara_col:
            st.markdown("🔍 **İsim veya Mağaza Ara**")
            arama_ayr = st.text_input("Ara", placeholder="İsim veya mağaza adı...",
                                      key="arama_ayrilanlar", label_visibility="collapsed")
        with mag_col:
            st.markdown("🏪 **Mağaza Filtresi**")
            mag_list_ayr = ["Tümü"] + sorted(df_ayr_f["Magaza"].dropna().astype(str).unique().tolist())
            sec_mag_ayr = st.selectbox("Mağaza", mag_list_ayr, key="mag_ayrilanlar", label_visibility="collapsed")
            if sec_mag_ayr != "Tümü":
                df_ayr_f = df_ayr_f[df_ayr_f["Magaza"].astype(str) == sec_mag_ayr]

        if arama_ayr:
            mask = pd.Series([False] * len(df_ayr_f), index=df_ayr_f.index)
            for col in [ayr_isim_col, "Magaza"]:
                if col and col in df_ayr_f.columns:
                    mask |= df_ayr_f[col].astype(str).str.lower().str.contains(arama_ayr.lower(), na=False)
            df_ayr_f = df_ayr_f[mask]

        st.markdown(f"**Toplam {len(df_ayr_f)} kayıt görüntüleniyor.**")

        goster_cols = []
        goster_rename = {}
        if ayr_isim_col and ayr_isim_col in df_ayr_f.columns:
            goster_cols.append(ayr_isim_col)
            goster_rename[ayr_isim_col] = "Ad Soyad"
        if "Magaza" in df_ayr_f.columns:
            goster_cols.append("Magaza")
            goster_rename["Magaza"] = "Mağaza"
        if ayr_masraf_col and ayr_masraf_col in df_ayr_f.columns:
            goster_cols.append(ayr_masraf_col)
            goster_rename[ayr_masraf_col] = "Masraf Kodu"
        if ayr_tarih_col and ayr_tarih_col in df_ayr_f.columns:
            goster_cols.append(ayr_tarih_col)
            goster_rename[ayr_tarih_col] = "Çıkış Tarihi"
        if ayr_to_neden_col and ayr_to_neden_col in df_ayr_f.columns:
            goster_cols.append(ayr_to_neden_col)
            goster_rename[ayr_to_neden_col] = "TO Nedeni"
        if ayr_cikis_neden_col and ayr_cikis_neden_col in df_ayr_f.columns:
            goster_cols.append(ayr_cikis_neden_col)
            goster_rename[ayr_cikis_neden_col] = "Çıkış Nedeni"

        if not goster_cols:
            merge_ekstra = ["_join_key", "Masraf_Kodu", "BM", "PM", "HRBP", "Segment", "Il", "Magaza",
                            "Onceki_Donem_HRBP"]
            goster_cols = [c for c in df_ayr_f.columns if c not in merge_ekstra]

        df_ayr_goster = df_ayr_f[goster_cols].rename(columns=goster_rename).copy()

        if "Çıkış Tarihi" in df_ayr_goster.columns:
            try:
                df_ayr_goster["Çıkış Tarihi"] = pd.to_datetime(df_ayr_goster["Çıkış Tarihi"]).dt.strftime('%d.%m.%Y')
            except Exception:
                pass

        # Çıkış Tarihi veya TO Nedeni sütunlarının görsel çakışmasını engellemek için mükerrer sütun isimlerini temizle
        df_ayr_goster = df_ayr_goster.loc[:, ~df_ayr_goster.columns.duplicated()].copy()

        st.dataframe(df_ayr_goster, use_container_width=True, hide_index=True)


        @st.cache_data
        def df_to_excel_bytes(df):
            import io
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="İşten Ayrılanlar")
            return buf.getvalue()


        st.download_button(
            label="📥 Listeyi Excel Olarak İndir",
            data=df_to_excel_bytes(df_ayr_goster),
            file_name="isten_ayrilanlar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ══════════════════════════════════════════════
# SEKME 6 — Güncel Çalışanlar
# (Yeni sekme - İstenen tüm mağaza kadro istatistikleri)
# ══════════════════════════════════════════════
with sekme6:
    st.markdown("### 👥 Güncel Çalışanlar Analizi")

    if df_guncel.empty:
        st.warning("⚠️ 'Güncel Çalışanlar' sayfası Excel dosyasında bulunamadı. Lütfen sayfa adını kontrol edin.")
    else:
        # Sol panel filtrelerini uygula
        df_gunc_f = df_guncel.copy()
        if "BM" in df_gunc_f.columns and sec_bm != "Tümü":
            df_gunc_f = df_gunc_f[df_gunc_f["BM"] == sec_bm]
        if "PM" in df_gunc_f.columns and sec_pm != "Tümü":
            df_gunc_f = df_gunc_f[df_gunc_f["PM"] == sec_pm]
        if "HRBP" in df_gunc_f.columns and sec_hrbp != "Tümü":
            df_gunc_f = df_gunc_f[df_gunc_f["HRBP"] == sec_hrbp]
        if "Segment" in df_gunc_f.columns and sec_segment != "Tümü":
            df_gunc_f = df_gunc_f[df_gunc_f["Segment"] == sec_segment]

        # Mağaza Seçimi
        st.markdown("🏪 **İncelenecek Mağazayı Seçin**")

        # Sütun adları uyumluluğu için _join_key kullanımı
        gunc_magaza_listesi = sorted(
            (df_gunc_f["_join_key"].astype(str) + " — " + df_gunc_f["Magaza"].astype(str)).unique().tolist()
        )
        if not gunc_magaza_listesi:
            st.info("Filtrelere uygun aktif çalışan kaydı olan mağaza bulunamadı.")
        else:
            sec_magaza_guncel = st.selectbox("Mağaza (Güncel)", gunc_magaza_listesi, label_visibility="collapsed")

            if sec_magaza_guncel:
                kod_guncel = sec_magaza_guncel.split(" — ")[0].strip()
                df_store_guncel = df_gunc_f[df_gunc_f["_join_key"] == kod_guncel]

                # Mağaza metadata'sını ana tablodan çekiyoruz
                s_meta = df[df["Masraf_Kodu"] == kod_guncel].iloc[0]

                # Mağaza Bilgileri Kartı
                st.markdown(f"""
                <div class="magaza-kart">
                    <p style="color:#aaaacc; font-size:12px; margin-bottom:8px;">MAĞAZA GENEL BİLGİLERİ</p>
                    <h2 style="color:white; margin:0;">{s_meta['Masraf_Kodu']} — {s_meta['Magaza']}</h2><br>
                    <span style="color:#e0e0e0; margin-right:24px;">📍 İl: <b>{s_meta['Il']}</b></span>
                    <span style="color:#e0e0e0; margin-right:24px;">➕ Segment: <b>{s_meta['Segment']}</b></span>
                    <span style="color:#e0e0e0; margin-right:24px;">🔢 Yeni Masraf Kodu: <b>{s_meta['Yeni_Masraf_Kodu']}</b></span>
                    <br><br>
                    <span style="color:#e0e0e0; margin-right:24px;">👤 BM: <b>{s_meta['BM']}</b></span>
                    <span style="color:#e0e0e0; margin-right:24px;">🏪 PM: <b>{s_meta['PM']}</b></span>
                    <span style="color:#e0e0e0;">🤝 HRBP: <b>{s['HRBP']}</b></span>
                </div>""", unsafe_allow_html=True)

                # Sayımları gerçekleştirelim
                toplam_calisan = len(df_store_guncel)

                # Tam/Yarı zamanlı sayımı (Belirttiğiniz ÇALIŞANALTGRUBU kelimesine göre güncellendi!)
                ft_count = 0
                pt_count = 0
                # Geliştirilmiş 'ÇALIŞANALTGRUBU' tespiti için ek kontrol (ÇALIŞANALTGRUBU'ndan tam zamanlı / yarı zamanlı okunur)
                gunc_altgrup_col_exact = (
                        bul_kolon(df_guncel.columns, ["çalışanaltgrubu"]) or
                        bul_kolon(df_guncel.columns, ["calisanaltgrubu"]) or
                        gunc_altgrup_col
                )

                if gunc_altgrup_col_exact and gunc_altgrup_col_exact in df_store_guncel.columns:
                    vals_grup = df_store_guncel[gunc_altgrup_col_exact].astype(str).str.strip().str.lower()
                    # Doğrudan "tam zamanlı" ve "yarı zamanlı" ifadelerine duyarlı sayım
                    ft_count = vals_grup.str.contains("tam zamanlı|tam zamanli|tam").sum()
                    pt_count = vals_grup.str.contains("yarı zamanlı|yari zamanli|yarı|yari").sum()

                # Cinsiyet sayımı
                kadin_count = 0
                erkek_count = 0
                if gunc_cinsiyet_col and gunc_cinsiyet_col in df_store_guncel.columns:
                    vals_cins = df_store_guncel[gunc_cinsiyet_col].astype(str).str.strip().str.lower()
                    kadin_count = vals_cins.str.startswith("k").sum()
                    erkek_count = vals_cins.str.startswith("e").sum()

                # Kıdem ortalaması
                kidem_ort = 0.0
                if gunc_kidem_col and gunc_kidem_col in df_store_guncel.columns:
                    # Virgüllü veya hatalı girilmiş değerleri temiz sayısal değerlere dönüştürüp ortalamasını alıyoruz
                    kidem_ort = pd.to_numeric(df_store_guncel[gunc_kidem_col], errors="coerce").mean()
                    if pd.isna(kidem_ort):
                        kidem_ort = 0.0

                # Yaş ortalaması
                yas_ort = 0.0
                if gunc_yas_col and gunc_yas_col in df_store_guncel.columns:
                    yas_ort = pd.to_numeric(df_store_guncel[gunc_yas_col], errors="coerce").mean()
                    if pd.isna(yas_ort):
                        yas_ort = 0.0

                st.markdown("#### 📊 Mağaza İstatistikleri")
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.metric("👥 Toplam Çalışan", f"{toplam_calisan} Kişi")
                with c2:
                    st.metric("⏳ Kıdem Ortalaması", f"{kidem_ort:.1f} Yıl")
                with c3:
                    st.metric("🎂 Yaş Ortalaması", f"{yas_ort:.1f} Yaş")
                with c4:
                    st.metric("👩‍💼 Kadın / Erkek", f"{kadin_count} K / {erkek_count} E")

                c5, c6 = st.columns(2)
                with c5:
                    st.metric("💼 Tam Zamanlı (Full Time)", f"{ft_count} Kişi")
                with c6:
                    st.metric("⏱️ Yarı Zamanlı (Part-Time)", f"{pt_count} Kişi")

                st.divider()

                st.markdown("#### 📋 Pozisyon Dağılımı ve Çalışan Listesi")
                col_pos, col_list = st.columns([2, 3])

                with col_pos:
                    st.markdown("**Pozisyona Göre Dağılım**")
                    if gunc_unvan_col and gunc_unvan_col in df_store_guncel.columns:
                        df_pos = df_store_guncel[gunc_unvan_col].value_counts().reset_index()
                        df_pos.columns = ["Pozisyon / Unvan", "Çalışan Sayısı"]
                        st.dataframe(df_pos, use_container_width=True, hide_index=True)
                    else:
                        st.info("Pozisyon kolon bilgisi saptanamadı.")

                with col_list:
                    st.markdown("**Aktif Çalışan Kadro Listesi**")
                    goster_guncel_cols = []
                    rename_guncel = {}

                    sicil_col_guncel = bul_kolon(df_guncel.columns, ["sicil"]) or bul_kolon(df_guncel.columns, ["no"])
                    isim_col_guncel = (bul_kolon(df_guncel.columns, ["ad", "soyad"]) or
                                       bul_kolon(df_guncel.columns, ["isim"]) or
                                       bul_kolon(df_guncel.columns, ["çalışan"]))

                    for col, label in [
                        (sicil_col_guncel, "Sicil No"),
                        (isim_col_guncel, "Ad Soyad"),
                        (gunc_unvan_col, "Pozisyon"),
                        (gunc_altgrup_col_exact, "Alt Grup"),
                        (gunc_kidem_col, "Kıdem (Yıl)"),
                        (gunc_yas_col, "Yaş"),
                    ]:
                        if col and col in df_store_guncel.columns:
                            goster_guncel_cols.append(col)
                            rename_guncel[col] = label

                    if goster_guncel_cols:
                        df_guncel_goster = df_store_guncel[goster_guncel_cols].rename(columns=rename_guncel).copy()
                        if "Kıdem (Yıl)" in df_guncel_goster.columns:
                            df_guncel_goster["Kıdem (Yıl)"] = df_guncel_goster["Kıdem (Yıl)"].apply(
                                lambda x: f"{x:.1f}" if pd.notna(x) else "—"
                            )
                        if "Yaş" in df_guncel_goster.columns:
                            df_guncel_goster["Yaş"] = df_guncel_goster["Yaş"].apply(
                                lambda x: f"{int(x)}" if pd.notna(x) else "—"
                            )
                        # Olası mükerrer sütun adlarını kaldırarak DataFrame çizim hatasını engelliyoruz
                        df_guncel_goster = df_guncel_goster.loc[:, ~df_guncel_goster.columns.duplicated()].copy()
                        st.dataframe(df_guncel_goster, use_container_width=True, hide_index=True)
                    else:
                        st.info("Kadro detay bilgisi bulunamadı.")