"""
Muhasebeci Veri Dönüştürücü - MVP
Dağınık Excel/CSV dosyalarını standart bir şablona dönüştürür.

Çalıştırma: streamlit run app.py
Kurulum: pip install streamlit pandas openpyxl
"""

import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="Veri Dönüştürücü", page_icon="📊", layout="centered")

# -----------------------------
# 1) KOLON EŞLEŞTİRME KURALLARI
# -----------------------------
# Buraya gerçek müşterilerden gördüğün kolon isim varyasyonlarını ekleyeceksin.
# Anahtar: standart kolon adı, Değer: karşılaşabileceğin varyasyonlar (küçük harf, boşluksuz karşılaştırılır)
COLUMN_MAP = {
    "Tarih": ["tarih", "işlem tarihi", "fatura tarihi", "date", "tarihi"],
    "Açıklama": ["açıklama", "aciklama", "işlem açıklaması", "description", "unvan", "cari unvan"],
    "Tutar": ["tutar", "toplam", "toplam tutar", "amount", "meblağ", "meblag"],
    "KDV": ["kdv", "kdv tutarı", "vat", "kdv tutari"],
    "Belge No": ["belge no", "fatura no", "belge numarası", "doc no", "belge numarasi"],
    "Kategori": ["kategori", "tür", "tur", "işlem türü", "category"],
}

def normalize(text: str) -> str:
    """Kolon adını karşılaştırma için sadeleştirir."""
    text = str(text).lower().strip()
    text = re.sub(r"[^a-zçğıöşü0-9 ]", "", text)
    return text

def match_columns(df_columns):
    """
    Gelen dosyadaki kolonları standart isimlere eşler.
    Dönüş: {orijinal_kolon_adı: standart_kolon_adı}
    """
    mapping = {}
    for col in df_columns:
        norm_col = normalize(col)
        matched = None
        for standard_name, variants in COLUMN_MAP.items():
            if norm_col in [normalize(v) for v in variants]:
                matched = standard_name
                break
        mapping[col] = matched  # None ise eşleşme bulunamadı demek
    return mapping

def clean_amount_column(series: pd.Series) -> pd.Series:
    """Tutar kolonundaki string/format sorunlarını (1.234,56 gibi TR formatı) düzeltir."""
    def parse_amount(val):
        if pd.isna(val):
            return None
        s = str(val).strip()
        s = s.replace(".", "").replace(",", ".")  # TR format -> float format
        s = re.sub(r"[^\d\.\-]", "", s)
        try:
            return float(s)
        except ValueError:
            return None
    return series.apply(parse_amount)

# -----------------------------
# 2) ARAYÜZ
# -----------------------------
st.title("📊 Muhasebe Veri Dönüştürücü")
st.write(
    "Dağınık Excel/CSV dosyanızı yükleyin, standart şablona uygun temiz "
    "dosyayı indirin."
)

uploaded_file = st.file_uploader("Excel veya CSV dosyası yükleyin", type=["xlsx", "xls", "csv"])

if uploaded_file:
    # Dosyayı oku
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("1. Yüklenen Veri (önizleme)")
    st.dataframe(df.head())

    # Kolon eşleştirme
    mapping = match_columns(df.columns)

    st.subheader("2. Kolon Eşleştirme")
    st.write("Sistem otomatik eşleştirdi. Yanlışsa aşağıdan düzeltebilirsiniz.")

    standard_options = ["(Eşleştirme yok)"] + list(COLUMN_MAP.keys())
    final_mapping = {}

    for col in df.columns:
        default = mapping[col] if mapping[col] else "(Eşleştirme yok)"
        choice = st.selectbox(
            f"'{col}' kolonu neye karşılık geliyor?",
            options=standard_options,
            index=standard_options.index(default) if default in standard_options else 0,
            key=col,
        )
        if choice != "(Eşleştirme yok)":
            final_mapping[col] = choice

    if st.button("3. Dönüştür ve İndir"):
        # Yeniden adlandır ve standart kolonlara göre yeni df oluştur
        renamed = df.rename(columns=final_mapping)
        output_columns = [c for c in COLUMN_MAP.keys() if c in renamed.columns]
        result_df = renamed[output_columns].copy()

        # Tutar/KDV kolonlarını temizle
        for money_col in ["Tutar", "KDV"]:
            if money_col in result_df.columns:
                result_df[money_col] = clean_amount_column(result_df[money_col])

        # Boş satırları at
        result_df = result_df.dropna(how="all")

        st.subheader("4. Sonuç")
        st.dataframe(result_df.head(20))

        # Excel olarak indirilebilir hale getir
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            result_df.to_excel(writer, index=False, sheet_name="Standart Rapor")
        buffer.seek(0)

        st.download_button(
            label="📥 Standart Excel'i İndir",
            data=buffer,
            file_name="standart_rapor.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Başlamak için bir dosya yükleyin.")
