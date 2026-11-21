import pandas as pd
from cleaner import TEMİZLEME


def ek_sutunlari_ekle(hedef_df, kaynak_df, ek_sutunlar, son_ek):
    """
    Barkod, fiyat ve stok dışındaki (Urun_Adi, Marka, Kategori gibi) bilgi sütunlarını
    dosyada mevcutsa sonuç tablosuna güvenli şekilde kopyalar.
    """
    for sutun in ek_sutunlar:
        tam_ad = sutun + son_ek
        if tam_ad in kaynak_df.columns:
            hedef_df[sutun] = kaynak_df[tam_ad]


def KIYASLA(
    df_dun,
    df_bugun,
    barkod_sutun="Barkod",
    fiyat_sutun="Alis_Fiyati",
    stok_sutun="Stok_Adedi",
):
    """
    Temizlenmiş iki DataFrame'i barkod üzerinden outer join ile kıyaslar.
    Geriye 4 parça döndürür:
    1) acil_df  : Stoğu sıfırlanan veya listeden tamamen silinen ürünler
    2) fiyat_df : Fiyatı zamlanan veya indirime giren ürünler
    3) yeni_df  : Bugün kataloğa yeni eklenen ürünler
    4) ozet     : Toplam istatistikleri tutan sözlük (dict)
    """
    # 1. İki tabloyu barkod sütunu üzerinden FULL OUTER JOIN mantığıyla birleştir
    birlesik = pd.merge(
        df_dun,
        df_bugun,
        on=barkod_sutun,
        how="outer",
        suffixes=("_dun", "_bugun"),
        indicator=True,
    )

    # Fiyat ve stok sütunlarının birleşik tablodaki yeni isimleri
    fiyat_dun = fiyat_sutun + "_dun"
    fiyat_bugun = fiyat_sutun + "_bugun"
    stok_dun = stok_sutun + "_dun"
    stok_bugun = stok_sutun + "_bugun"

    # Barkod, fiyat ve stok dışında kalan tüm bilgi sütunlarını otomatik tespit et
    ana_sutunlar = [barkod_sutun, fiyat_sutun, stok_sutun]
    bilgi_sutunlari = []
    for col in df_dun.columns:
        if col not in ana_sutunlar:
            bilgi_sutunlari.append(col)

    # =========================================================================
    # SEPET 1: ACİL SATIŞA KAPATILACAKLAR (Listeden silinenler + Stoğu bitenler)
    # =========================================================================
    silinen_kosul = birlesik["_merge"] == "left_only"
    biten_kosul = (
        (birlesik["_merge"] == "both")
        & (birlesik[stok_dun] > 0)
        & (birlesik[stok_bugun] == 0)
    )

    acil_ham = birlesik[silinen_kosul | biten_kosul].copy()

    # Kapanma sebebini açıkça yazalım ki e-ticaretçi neden kapatacağını görsün
    acil_ham["Kapanma_Sebebi"] = "Stok Sıfırlandı"
    acil_ham.loc[acil_ham["_merge"] == "left_only", "Kapanma_Sebebi"] = "Listeden Silindi"

    acil_df = pd.DataFrame()
    acil_df[barkod_sutun] = acil_ham[barkod_sutun]
    ek_sutunlari_ekle(acil_df, acil_ham, bilgi_sutunlari, "_dun")
    acil_df["Dun_Stok"] = acil_ham[stok_dun].fillna(0).astype(int)
    acil_df["Bugun_Stok"] = acil_ham[stok_bugun].fillna(0).astype(int)
    acil_df["Son_Alis_Fiyati"] = acil_ham[fiyat_dun]
    acil_df["Kapanma_Sebebi"] = acil_ham["Kapanma_Sebebi"]
    acil_df = acil_df.reset_index(drop=True)

    # =========================================================================
    # SEPET 2: FİYATI DEĞİŞENLER (Zamlananlar ve İndirime Girenler)
    # =========================================================================
    fiyat_kosul = (birlesik["_merge"] == "both") & (
        birlesik[fiyat_dun] != birlesik[fiyat_bugun]
    )
    fiyat_ham = birlesik[fiyat_kosul].copy()

    fiyat_df = pd.DataFrame()
    fiyat_df[barkod_sutun] = fiyat_ham[barkod_sutun]
    ek_sutunlari_ekle(fiyat_df, fiyat_ham, bilgi_sutunlari, "_bugun")
    fiyat_df["Eski_Fiyat"] = fiyat_ham[fiyat_dun]
    fiyat_df["Yeni_Fiyat"] = fiyat_ham[fiyat_bugun]
    fiyat_df["Fark_TL"] = (fiyat_df["Yeni_Fiyat"] - fiyat_df["Eski_Fiyat"]).round(2)
    fiyat_df["Degisim_Yuzde"] = (
        (fiyat_df["Fark_TL"] / fiyat_df["Eski_Fiyat"]) * 100
    ).round(2)
    fiyat_df["Guncel_Stok"] = fiyat_ham[stok_bugun].astype(int)

    # En yüksek zam oranından en düşük indirime doğru sırala
    fiyat_df = fiyat_df.sort_values(by="Degisim_Yuzde", ascending=False).reset_index(drop=True)

    # =========================================================================
    # SEPET 3: YENİ EKLENEN ÜRÜNLER (Dün olmayıp bugün kataloğa girenler)
    # =========================================================================
    yeni_ham = birlesik[birlesik["_merge"] == "right_only"].copy()

    yeni_df = pd.DataFrame()
    yeni_df[barkod_sutun] = yeni_ham[barkod_sutun]
    ek_sutunlari_ekle(yeni_df, yeni_ham, bilgi_sutunlari, "_bugun")
    yeni_df["Alis_Fiyati"] = yeni_ham[fiyat_bugun]
    yeni_df["Stok_Adedi"] = yeni_ham[stok_bugun].astype(int)
    yeni_df = yeni_df.reset_index(drop=True)

    # =========================================================================
    # ÖZET İSTATİSTİKLER SÖZLÜĞÜ
    # =========================================================================
    ozet = {
        "dun_toplam": len(df_dun),
        "bugun_toplam": len(df_bugun),
        "acil_kapatilacak": len(acil_df),
        "stok_sifirlanan": int(biten_kosul.sum()),
        "listeden_silinen": int(silinen_kosul.sum()),
        "fiyat_degisen": len(fiyat_df),
        "zamlanan": int((fiyat_df["Fark_TL"] > 0).sum()),
        "indirim_giren": int((fiyat_df["Fark_TL"] < 0).sum()),
        "yeni_eklenen": len(yeni_df),
    }

    return acil_df, fiyat_df, yeni_df, ozet


# Kodu terminalden tek başına çalıştırıp test etmek için:
if __name__ == "__main__":
    df_dun, df_bugun = TEMİZLEME("dun.xlsx", "bugun.xlsx")
    acil_df, fiyat_df, yeni_df, ozet = KIYASLA(df_dun, df_bugun)

    print("=== KIYASLAMA SONUÇ ÖZETİ ===")
    for anahtar, deger in ozet.items():
        print(f"{anahtar}: {deger}")

    print("\n--- ÖRNEK FİYAT DEĞİŞENLER (İLK 3) ---")
    print(fiyat_df[["Barkod", "Urun_Adi", "Eski_Fiyat", "Yeni_Fiyat", "Degisim_Yuzde"]].head(3))
    