import pandas as pd


def fiyat_duzelt(deger):
    """
    '1.450,00 TL' gibi Türkçe metin formatındaki fiyatları temizler
    ve makinenin işlem yapabileceği ondalıklı sayıya (float) çevirir.
    """
    if pd.isna(deger):
        return 0.0

    metin = str(deger).strip()
    metin = metin.replace("TL", "").replace("tl", "").replace("₺", "").strip()

    if metin == "" or metin == "-":
        return 0.0

    # Hem nokta hem virgül varsa (Örn: "1.450,00") -> Noktayı sil, virgülü nokta yap
    if "." in metin and "," in metin:
        metin = metin.replace(".", "").replace(",", ".")
    # Sadece virgül varsa (Örn: "450,50") -> Virgülü nokta yap
    elif "," in metin:
        metin = metin.replace(",", ".")

    try:
        return float(metin)
    except ValueError:
        return 0.0


def stok_duzelt(deger):
    """
    Stok sütunundaki 'Tükendi', 'Yok', '-' veya boş (NaN) değerleri
    0 sayısına çevirir; kalanları tam sayı (int) yapar.
    """
    if pd.isna(deger):
        return 0

    metin = str(deger).strip().lower()
    sifir_kelimeleri = ["", "-", "tükendi", "tukendi", "yok", "stok yok", "none", "nan"]

    if metin in sifir_kelimeleri:
        return 0

    try:
        # Bazen stoklar "15.0" gibi gelebilir, önce float yapıp sonra int'e çeviriyoruz
        return int(float(metin.replace(",", ".")))
    except ValueError:
        return 0


def TEMİZLEME(
    dun_yolu,
    bugun_yolu,
    barkod_sutun="Barkod",
    fiyat_sutun="Alis_Fiyati",
    stok_sutun="Stok_Adedi",
):
    """
    Dünkü ve bugünkü Excel/CSV dosyalarını okur, barkodları tekilleştirir,
    fiyat ve stok sütunlarını sayısal formata çevirir ve iki temiz DataFrame döndürür.
    """
    temiz_tablolar = []

    for dosya_yolu in [dun_yolu, bugun_yolu]:
        # 1. Dosya uzantısına göre tüm sütunları METİN (str) olarak oku
        if dosya_yolu.lower().endswith(".csv"):
            df = pd.read_csv(dosya_yolu, dtype=str)
        else:
            df = pd.read_excel(dosya_yolu, dtype=str)

        # 2. Sütun isimlerindeki olası görünmez kenar boşluklarını temizle
        df.columns = df.columns.str.strip()

        # 3. Anahtar sütun (Barkod) boşsa o çöp satırları tamamen sil
        df = df.dropna(subset=[barkod_sutun])
        df[barkod_sutun] = df[barkod_sutun].str.strip()
        df = df[df[barkod_sutun] != ""]

        # 4. Aynı barkod Excel'de birden fazla kez geçiyorsa son satırı tut, diğerlerini sil
        df = df.drop_duplicates(subset=[barkod_sutun], keep="last")

        # 5. Fiyat ve stok sütunlarını yardımcı fonksiyonlarla sayısal tipe dönüştür
        df[fiyat_sutun] = df[fiyat_sutun].apply(fiyat_duzelt)
        df[stok_sutun] = df[stok_sutun].apply(stok_duzelt)

        # 6. Silinen satırlardan sonra indeks numaralarını 0, 1, 2... diye sıfırla
        df = df.reset_index(drop=True)

        temiz_tablolar.append(df)

    # Sırasıyla (df_dun_temiz, df_bugun_temiz) olarak iki DataFrame döndür
    return temiz_tablolar[0], temiz_tablolar[1]


# Kodu terminalden tek başına çalıştırıp test etmek için:
if __name__ == "__main__":
    df_dun, df_bugun = TEMİZLEME("dun.xlsx", "bugun.xlsx")

    print("--- DÜN TEMİZ TABLO ---")
    print("Satır ve Sütun Sayısı:", df_dun.shape)
    print(df_dun.dtypes)

    print("\n--- BUGÜN TEMİZ TABLO ---")
    print("Satır ve Sütun Sayısı:", df_bugun.shape)
    print(df_bugun.dtypes)
    print("0 ile başlayan barkod sayısı:", df_bugun["Barkod"].str.startswith("0869").sum())
    print("Stoğu 0 olan ürün sayısı:", (df_bugun["Stok_Adedi"] == 0).sum())