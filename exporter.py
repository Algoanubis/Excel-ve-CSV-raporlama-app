from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from cleaner import TEMİZLEME
from comparator import KIYASLA


def sutun_genisliklerini_ayarla(ws):
    """
    Sayfadaki her sütunu, içindeki en uzun metne göre otomatik genişletir.
    Böylece Excel açıldığında sayılar '###' şeklinde sıkışmaz.
    """
    for col in ws.columns:
        max_uzunluk = 0
        sutun_harfi = get_column_letter(col[0].column)

        for hucre in col:
            if hucre.value is not None:
                uzunluk = len(str(hucre.value))
                if uzunluk > max_uzunluk:
                    max_uzunluk = uzunluk

        # En fazla 50 karakter genişliğe kadar pay bırakarak ayarla
        ws.column_dimensions[sutun_harfi].width = min(max_uzunluk + 4, 50)


def baslik_stili_uygula(ws, baslik_rengi):
    """
    Sayfanın 1. satırındaki başlıklara koyu arka plan, beyaz kalın yazı
    ve dondurulmuş üst satır (freeze_panes) özelliği ekler.
    """
    dolgu = PatternFill(start_color=baslik_rengi, end_color=baslik_rengi, fill_type="solid")
    yazi = Font(bold=True, color="FFFFFF", size=11)
    hizalama = Alignment(horizontal="center", vertical="center")

    for hucre in ws[1]:
        hucre.fill = dolgu
        hucre.font = yazi
        hucre.alignment = hizalama

    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"  # Aşağı kaydırırken başlık satırı sabit kalır


def tabloyu_sayfaya_yaz(ws, df, baslik_rengi, fiyat_renklendir=False):
    """
    Bir DataFrame'i verilen Excel sekmesine yazar, barkodları metin olarak korur
    ve istenirse fiyat değişim satırlarını kırmızı/yeşil vurgular.
    """
    # 1. Sütun başlıklarını 1. satıra ekle
    basliklar = list(df.columns)
    ws.append(basliklar)
    baslik_stili_uygula(ws, baslik_rengi)

    # Renk tanımları (Fiyat artışı için açık kırmızı, düşüşü için açık yeşil)
    zam_dolgu = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
    indirim_dolgu = PatternFill(start_color="D4EFDF", end_color="D4EFDF", fill_type="solid")

    # 2. Satırları tek tek Excel'e aktar
    for satir_index in range(len(df)):
        satir_verisi = df.iloc[satir_index]
        excel_satiri = []

        for sutun in basliklar:
            deger = satir_verisi[sutun]
            # Barkod veya stok kodu gibi metin sütunlarını koru
            if sutun in ["Barkod", "Stok_Kodu"]:
                excel_satiri.append(str(deger))
            else:
                excel_satiri.append(deger)

        ws.append(excel_satiri)
        mevcut_satir_no = ws.max_row

        # Barkod hücresini Excel'de metin formatına ('@') sabitle
        ws.cell(row=mevcut_satir_no, column=1).number_format = "@"

        # Eğer Fiyat_Degisenler sekmesindeysek Fark_TL'ye göre satırı boya
        if fiyat_renklendir and "Fark_TL" in basliklar:
            fark = satir_verisi["Fark_TL"]
            satir_dolgu = zam_dolgu if fark > 0 else indirim_dolgu
            for col_no in range(1, len(basliklar) + 1):
                ws.cell(row=mevcut_satir_no, column=col_no).fill = satir_dolgu

    sutun_genisliklerini_ayarla(ws)


def RAPOR_OLUSTUR(acil_df, fiyat_df, yeni_df, ozet, cikti_yolu="karsilastirma_raporu.xlsx"):
    """
    4 sekmeli (Ozet_Rapor, Acil_Kapatilacaklar, Fiyat_Degisenler, Yeni_Urunler)
    biçimlendirilmiş Excel raporunu oluşturur ve kaydedilen dosya yolunu döndürür.
    """
    wb = Workbook()

    # =========================================================================
    # SEKME 1: ÖZET RAPOR (Yönetici Özeti)
    # =========================================================================
    ws_ozet = wb.active
    ws_ozet.title = "Ozet_Rapor"

    ws_ozet.append(["E-TİCARET TEDARİKÇİ STOK & FİYAT KARŞILAŞTIRMA RAPORU", ""])
    ws_ozet.append(["Rapor Oluşturma Tarihi", datetime.now().strftime("%d.%m.%Y %H:%M")])
    ws_ozet.append(["", ""])
    ws_ozet.append(["Metrik / Durum", "Ürün Adedi"])

    ozet_satirlari = [
        ("Dünkü Katalog Toplam Ürün", ozet["dun_toplam"]),
        ("Bugünkü Katalog Toplam Ürün", ozet["bugun_toplam"]),
        ("ACİL KAPATILACAK TOPLAM ÜRÜN", ozet["acil_kapatilacak"]),
        ("   • Stoğu Sıfıra Düşenler", ozet["stok_sifirlanan"]),
        ("   • Listeden Tamamen Silinenler", ozet["listeden_silinen"]),
        ("FİYATI DEĞİŞEN TOPLAM ÜRÜN", ozet["fiyat_degisen"]),
        ("   • Zamlanan Ürünler", ozet["zamlanan"]),
        ("   • İndirime Giren Ürünler", ozet["indirim_giren"]),
        ("KATALOĞA YENİ EKLENEN ÜRÜNLER", ozet["yeni_eklenen"]),
    ]

    for etiket, sayi in ozet_satirlari:
        ws_ozet.append([etiket, sayi])

    # Özet sayfası görsel düzeni
    ws_ozet["A1"].font = Font(bold=True, size=13, color="1B2631")
    tablo_baslik_dolgu = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    for hucre in ws_ozet[4]:
        hucre.fill = tablo_baslik_dolgu
        hucre.font = Font(bold=True, color="FFFFFF", size=11)

    ince_kenar = Side(style="thin", color="D5D8DC")
    cerceve = Border(left=ince_kenar, right=ince_kenar, top=ince_kenar, bottom=ince_kenar)

    for satir_no in range(4, ws_ozet.max_row + 1):
        hucre_a = ws_ozet.cell(row=satir_no, column=1)
        hucre_b = ws_ozet.cell(row=satir_no, column=2)
        hucre_a.border = cerceve
        hucre_b.border = cerceve
        hucre_b.alignment = Alignment(horizontal="center")

        if satir_no > 4 and not str(hucre_a.value).startswith("   •"):
            hucre_a.font = Font(bold=True)
            hucre_b.font = Font(bold=True)

    sutun_genisliklerini_ayarla(ws_ozet)

    # =========================================================================
    # SEKME 2: ACİL KAPATILACAKLAR (Koyu Kırmızı Başlık)
    # =========================================================================
    ws_acil = wb.create_sheet(title="Acil_Kapatilacaklar")
    tabloyu_sayfaya_yaz(ws_acil, acil_df, baslik_rengi="922B21", fiyat_renklendir=False)

    # =========================================================================
    # SEKME 3: FİYAT DEĞİŞENLER (Lacivert Başlık + Kırmızı/Yeşil Satır Vurgusu)
    # =========================================================================
    ws_fiyat = wb.create_sheet(title="Fiyat_Degisenler")
    tabloyu_sayfaya_yaz(ws_fiyat, fiyat_df, baslik_rengi="1F618D", fiyat_renklendir=True)

    # =========================================================================
    # SEKME 4: YENİ ÜRÜNLER (Koyu Yeşil Başlık)
    # =========================================================================
    ws_yeni = wb.create_sheet(title="Yeni_Urunler")
    tabloyu_sayfaya_yaz(ws_yeni, yeni_df, baslik_rengi="1E8449", fiyat_renklendir=False)

    wb.save(cikti_yolu)
    return cikti_yolu


# Tüm boru hattını (cleaner -> comparator -> exporter) uçtan uca test etmek için:
if __name__ == "__main__":
    df_dun, df_bugun = TEMİZLEME("dun.xlsx", "bugun.xlsx")
    acil_df, fiyat_df, yeni_df, ozet = KIYASLA(df_dun, df_bugun)
    kaydedilen_dosya = RAPOR_OLUSTUR(acil_df, fiyat_df, yeni_df, ozet)

    print("Rapor başarıyla oluşturuldu:", kaydedilen_dosya)