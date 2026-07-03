"""
FinTemiz - Muhasebe Veri Dönüştürücü
Dağınık Excel/CSV dosyalarını standart bir şablona dönüştürür.

Çalıştırma: streamlit run app.py
Kurulum: pip install streamlit pandas openpyxl xlrd
"""

import streamlit as st
import pandas as pd
from io import BytesIO
import re
import time

st.set_page_config(
    page_title="FinTemiz | Veri Dönüştürücü",
    page_icon="📊",
    layout="centered",
)

# -----------------------------
# 1) KOLON EŞLEŞTİRME KURALLARI
# -----------------------------
COLUMN_MAP = {
    "Tarih": ["tarih", "işlem tarihi", "fatura tarihi", "date", "tarihi"],
    "Açıklama": ["açıklama", "aciklama", "işlem açıklaması", "description", "unvan", "cari unvan"],
    "Tutar": ["tutar", "toplam", "toplam tutar", "amount", "meblağ", "meblag"],
    "KDV": ["kdv", "kdv tutarı", "vat", "kdv tutari"],
    "Belge No": ["belge no", "fatura no", "belge numarası", "doc no", "belge numarasi"],
    "Kategori": ["kategori", "tür", "tur", "işlem türü", "category"],
}


def normalize(text: str) -> str:
    text = str(text).lower().strip()
    text = re.sub(r"[^a-zçğıöşü0-9 ]", "", text)
    return text


def match_columns(df_columns):
    mapping = {}
    for col in df_columns:
        norm_col = normalize(col)
        matched = None
        for standard_name, variants in COLUMN_MAP.items():
            if norm_col in [normalize(v) for v in variants]:
                matched = standard_name
                break
        mapping[col] = matched
    return mapping


def clean_amount_column(series: pd.Series) -> pd.Series:
    def parse_amount(val):
        if pd.isna(val):
            return None
        s = str(val).strip()
        s = s.replace(".", "").replace(",", ".")
        s = re.sub(r"[^\d\.\-]", "", s)
        try:
            return float(s)
        except ValueError:
            return None
    return series.apply(parse_amount)


def make_sample_file() -> BytesIO:
    """Kullanıcının sistemi denemesi için örnek dağınık bir dosya üretir."""
    sample = pd.DataFrame({
        "İşlem Tarihi": ["01.06.2026", "03.06.2026", "05.06.2026", "10.06.2026"],
        "Cari Unvan": ["ABC Ticaret Ltd.", "Mavi Elektronik", "Deniz Yapı A.Ş.", "Kaya Gıda San."],
        "Toplam Tutar": ["1.250,00", "3.400,50", "870,00", "12.300,75"],
        "KDV Tutarı": ["225,00", "612,00", "156,60", "2.214,00"],
        "Fatura No": ["F-2026-001", "F-2026-002", "F-2026-003", "F-2026-004"],
    })
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        sample.to_excel(writer, index=False, sheet_name="Örnek Veri")
    buf.seek(0)
    return buf


# -----------------------------
# 2) ÜST BAŞLIK / MARKA
# -----------------------------
st.markdown(
    """
    <div style="text-align:center; padding: 10px 0 20px 0;">
        <h1 style="margin-bottom:0;">📊 FinTemiz</h1>
        <p style="color:#6B7280; font-size:16px; margin-top:4px;">
            Dağınık Excel/CSV dosyalarınızı 3 adımda temiz, standart rapora dönüştürün
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.divider()

with st.expander("📎 Önce örnek bir dosyayla deneyin"):
    st.write("Kendi dosyanız elinizde değilse, aşağıdaki örnek dosyayı indirip sistemi test edebilirsiniz.")
    st.download_button(
        label="Örnek dosyayı indir",
        data=make_sample_file(),
        file_name="ornek_veri.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# -----------------------------
# 3) ADIM 1 - YÜKLEME
# -----------------------------
st.subheader("Adım 1 · Dosyanızı Yükleyin")
uploaded_file = st.file_uploader("Excel (.xlsx, .xls) veya CSV dosyası seçin", type=["xlsx", "xls", "csv"])

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    with st.expander("Yüklenen veriyi önizle", expanded=False):
        st.dataframe(df.head())

    st.divider()

    # -----------------------------
    # 4) ADIM 2 - EŞLEŞTİRME
    # -----------------------------
    st.subheader("Adım 2 · Kolonları Eşleştirin")
    st.caption("Sistem otomatik eşleştirdi. Yanlışsa aşağıdan düzeltebilirsiniz.")

    mapping = match_columns(df.columns)
    standard_options = ["(Eşleştirme yok)"] + list(COLUMN_MAP.keys())
    final_mapping = {}

    auto_matched_count = 0
    cols = st.columns(2)
    for i, col in enumerate(df.columns):
        default = mapping[col] if mapping[col] else "(Eşleştirme yok)"
        if mapping[col]:
            auto_matched_count += 1
        with cols[i % 2]:
            choice = st.selectbox(
                f"'{col}'",
                options=standard_options,
                index=standard_options.index(default) if default in standard_options else 0,
                key=col,
            )
        if choice != "(Eşleştirme yok)":
            final_mapping[col] = choice

    st.caption(f"✅ {auto_matched_count}/{len(df.columns)} kolon otomatik eşleştirildi")

    st.divider()

    # -----------------------------
    # 5) ADIM 3 - DÖNÜŞTÜRME
    # -----------------------------
    st.subheader("Adım 3 · Dönüştürün ve İndirin")

    if st.button("🔄 Dönüştür", type="primary", use_container_width=True):
        start_time = time.time()

        with st.spinner("Veriniz işleniyor..."):
            renamed = df.rename(columns=final_mapping)
            output_columns = [c for c in COLUMN_MAP.keys() if c in renamed.columns]
            result_df = renamed[output_columns].copy()

            for money_col in ["Tutar", "KDV"]:
                if money_col in result_df.columns:
                    result_df[money_col] = clean_amount_column(result_df[money_col])

            original_row_count = len(result_df)
            result_df = result_df.dropna(how="all")
            cleaned_row_count = len(result_df)
            removed_rows = original_row_count - cleaned_row_count

            elapsed = round(time.time() - start_time, 2)

        st.success("Dönüştürme tamamlandı!")

        stat_cols = st.columns(3)
        stat_cols[0].metric("İşlenen Satır", cleaned_row_count)
        stat_cols[1].metric("Temizlenen Boş Satır", removed_rows)
        stat_cols[2].metric("Süre", f"{elapsed} sn")

        with st.expander("Sonucu önizle", expanded=True):
            st.dataframe(result_df.head(20))

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            result_df.to_excel(writer, index=False, sheet_name="Standart Rapor")
        buffer.seek(0)

        st.download_button(
            label="📥 Standart Excel'i İndir",
            data=buffer,
            file_name="standart_rapor.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
else:
    st.info("Başlamak için bir dosya yükleyin, ya da yukarıdaki örnek dosyayı deneyin.")

st.divider()
st.caption("FinTemiz · Verileriniz sadece bu oturumda işlenir, sunucuda saklanmaz.")
