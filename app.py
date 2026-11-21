import os
import platform
import subprocess
import threading
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd

from cleaner import TEMİZLEME
from comparator import KIYASLA
from exporter import RAPOR_OLUSTUR


# =============================================================================
# 1. YARDIMCI FONKSİYONLAR (DOSYA BAŞLIK OKUMA VE SÜTUN TAHMİNİ)
# =============================================================================
def metni_sadelestir(metin):
    """
    Sütun isimlerini karşılaştırmak için Türkçe karakterleri dönüştürür
    ve küçük harfe çevirir.
    """
    donusum_tablosu = {
        "ı": "i",
        "İ": "i",
        "ğ": "g",
        "Ğ": "g",
        "ü": "u",
        "Ü": "u",
        "ş": "s",
        "Ş": "s",
        "ö": "o",
        "Ö": "o",
        "ç": "c",
        "Ç": "c",
    }
    sonuc = str(metin).strip()
    for eski, yeni in donusum_tablosu.items():
        sonuc = sonuc.replace(eski, yeni)
    return sonuc.lower()


def dosya_sutunlarini_oku(dosya_yolu):
    """
    Seçilen Excel veya CSV dosyasının sadece başlık satırını okur
    ve temizlenmiş sütun isimleri listesini döndürür.
    """
    if dosya_yolu.lower().endswith(".csv"):
        df_baslik = pd.read_csv(dosya_yolu, nrows=0, dtype=str)
    else:
        df_baslik = pd.read_excel(dosya_yolu, nrows=0, dtype=str)

    sutunlar = []
    for col in df_baslik.columns:
        sutunlar.append(str(col).strip())
    return sutunlar


def sutun_tahmin_et(sutunlar, anahtar_kelimeler):
    """
    Verilen sütun listesi içinde anahtar kelimelerle eşleşen ilk sütunu bulur.
    Eşleşme yoksa listenin ilk sütununu döndürür.
    """
    if not sutunlar:
        return ""

    for kelime in anahtar_kelimeler:
        for sutun in sutunlar:
            sade_sutun = metni_sadelestir(sutun)
            if kelime in sade_sutun:
                return sutun

    return sutunlar[0]


# =============================================================================
# 2. ANA UYGULAMA SINIFI (CUSTOMTKINTER ARAYÜZÜ)
# =============================================================================
class RaporlamaUygulamasi(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Pencere ve Tema Ayarları
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("Tedarikçi Stok & Fiyat Karşılaştırma")
        self.geometry("820x700")
        self.minsize(760, 660)

        # Uygulama Durum Değişkenleri
        self.dun_dosya_yolu = ""
        self.bugun_dosya_yolu = ""
        self.dun_sutunlar = []
        self.bugun_sutunlar = []
        self.olusturulan_rapor_yolu = ""

        self.arayuzu_olustur()

    def arayuzu_olustur(self):
        # Ana Taşıyıcı Çerçeve
        self.ana_cerceve = ctk.CTkFrame(self, fg_color="transparent")
        self.ana_cerceve.pack(fill="both", expand=True, padx=24, pady=20)

        # Başlık Alanı
        baslik = ctk.CTkLabel(
            self.ana_cerceve,
            text="Stok & Fiyat Karşılaştırma Paneli",
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        baslik.pack(anchor="w")

        alt_baslik = ctk.CTkLabel(
            self.ana_cerceve,
            text="Dünkü ve bugünkü tedarikçi dosyalarını seçip sütunları eşleştirin.",
            font=ctk.CTkFont(size=13),
            text_color="#9CA3AF",
        )
        alt_baslik.pack(anchor="w", pady=(2, 16))

        # ---------------------------------------------------------------------
        # BÖLÜM 1: DOSYA SEÇİM ALANI
        # ---------------------------------------------------------------------
        dosya_kart = ctk.CTkFrame(self.ana_cerceve, corner_radius=12)
        dosya_kart.pack(fill="x", pady=(0, 14))

        dosya_baslik = ctk.CTkLabel(
            dosya_kart,
            text="1. Dosya Seçimi",
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        dosya_baslik.pack(anchor="w", padx=16, pady=(12, 8))

        # Dünkü Dosya Satırı
        dun_satir = ctk.CTkFrame(dosya_kart, fg_color="transparent")
        dun_satir.pack(fill="x", padx=16, pady=(0, 8))

        self.btn_dun_sec = ctk.CTkButton(
            dun_satir,
            text="Dünkü Excel/CSV Seç",
            width=200,
            height=36,
            command=self.dun_dosyasi_sec,
        )
        self.btn_dun_sec.pack(side="left")

        self.lbl_dun_dosya = ctk.CTkLabel(
            dun_satir,
            text="Henüz dosya seçilmedi",
            text_color="#9CA3AF",
            anchor="w",
        )
        self.lbl_dun_dosya.pack(side="left", padx=(14, 0), fill="x", expand=True)

        # Bugünkü Dosya Satırı
        bugun_satir = ctk.CTkFrame(dosya_kart, fg_color="transparent")
        bugun_satir.pack(fill="x", padx=16, pady=(0, 14))

        self.btn_bugun_sec = ctk.CTkButton(
            bugun_satir,
            text="Bugünkü Excel/CSV Seç",
            width=200,
            height=36,
            command=self.bugun_dosyasi_sec,
        )
        self.btn_bugun_sec.pack(side="left")

        self.lbl_bugun_dosya = ctk.CTkLabel(
            bugun_satir,
            text="Henüz dosya seçilmedi",
            text_color="#9CA3AF",
            anchor="w",
        )
        self.lbl_bugun_dosya.pack(side="left", padx=(14, 0), fill="x", expand=True)

        # ---------------------------------------------------------------------
        # BÖLÜM 2: DİNAMİK SÜTUN EŞLEŞTİRME (COMBOBOX)
        # ---------------------------------------------------------------------
        sutun_kart = ctk.CTkFrame(self.ana_cerceve, corner_radius=12)
        sutun_kart.pack(fill="x", pady=(0, 14))

        sutun_baslik = ctk.CTkLabel(
            sutun_kart,
            text="2. Sütun Eşleştirme",
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        sutun_baslik.pack(anchor="w", padx=16, pady=(12, 10))

        eslestirme_grid = ctk.CTkFrame(sutun_kart, fg_color="transparent")
        eslestirme_grid.pack(fill="x", padx=16, pady=(0, 14))
        eslestirme_grid.columnconfigure((0, 1, 2), weight=1)

        # Barkod Sütunu
        ctk.CTkLabel(eslestirme_grid, text="Barkod Sütunu", anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 4)
        )
        self.cmb_barkod = ctk.CTkComboBox(
            eslestirme_grid, values=["Dosya seçiniz"], state="readonly"
        )
        self.cmb_barkod.set("Dosya seçiniz")
        self.cmb_barkod.grid(row=1, column=0, sticky="ew", padx=(0, 10))

        # Fiyat Sütunu
        ctk.CTkLabel(eslestirme_grid, text="Fiyat Sütunu", anchor="w").grid(
            row=0, column=1, sticky="w", padx=5, pady=(0, 4)
        )
        self.cmb_fiyat = ctk.CTkComboBox(
            eslestirme_grid, values=["Dosya seçiniz"], state="readonly"
        )
        self.cmb_fiyat.set("Dosya seçiniz")
        self.cmb_fiyat.grid(row=1, column=1, sticky="ew", padx=5)

        # Stok Sütunu
        ctk.CTkLabel(eslestirme_grid, text="Stok Sütunu", anchor="w").grid(
            row=0, column=2, sticky="w", padx=(10, 0), pady=(0, 4)
        )
        self.cmb_stok = ctk.CTkComboBox(
            eslestirme_grid, values=["Dosya seçiniz"], state="readonly"
        )
        self.cmb_stok.set("Dosya seçiniz")
        self.cmb_stok.grid(row=1, column=2, sticky="ew", padx=(10, 0))

        # ---------------------------------------------------------------------
        # BÖLÜM 3: İŞLEM BUTONU VE DURUM BİLGİSİ
        # ---------------------------------------------------------------------
        islem_alani = ctk.CTkFrame(self.ana_cerceve, fg_color="transparent")
        islem_alani.pack(fill="x", pady=(4, 14))

        self.btn_analiz = ctk.CTkButton(
            islem_alani,
            text="Analiz Et ve Rapor Oluştur",
            height=44,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            command=self.analizi_baslat,
        )
        self.btn_analiz.pack(fill="x")

        self.lbl_durum = ctk.CTkLabel(
            islem_alani,
            text="Hazır. Dosyaları seçtikten sonra analizi başlatabilirsiniz.",
            font=ctk.CTkFont(size=12),
            text_color="#9CA3AF",
        )
        self.lbl_durum.pack(anchor="w", pady=(8, 0))

        # ---------------------------------------------------------------------
        # BÖLÜM 4: 3 ADET ÖZET BİLGİ KARTI
        # ---------------------------------------------------------------------
        ozet_cerceve = ctk.CTkFrame(self.ana_cerceve, fg_color="transparent")
        ozet_cerceve.pack(fill="x", pady=(0, 16))
        ozet_cerceve.columnconfigure((0, 1, 2), weight=1)

        self.lbl_acil_sayi = self.ozet_karti_olustur(
            ozet_cerceve,
            sutun_no=0,
            baslik="Acil Kapatılacak Ürün Sayısı",
            renk="#EF4444",
            sag_bosluk=8,
        )
        self.lbl_fiyat_sayi = self.ozet_karti_olustur(
            ozet_cerceve,
            sutun_no=1,
            baslik="Fiyatı Değişen Ürün Sayısı",
            renk="#F59E0B",
            sol_bosluk=4,
            sag_bosluk=4,
        )
        self.lbl_yeni_sayi = self.ozet_karti_olustur(
            ozet_cerceve,
            sutun_no=2,
            baslik="Yeni Eklenen Ürün Sayısı",
            renk="#10B981",
            sol_bosluk=8,
        )

        # ---------------------------------------------------------------------
        # BÖLÜM 5: HIZLI AÇILIŞ BUTONU
        # ---------------------------------------------------------------------
        self.btn_rapor_ac = ctk.CTkButton(
            self.ana_cerceve,
            text="Oluşturulan Excel Raporunu Aç",
            height=42,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#059669",
            hover_color="#047857",
            state="disabled",
            command=self.raporu_ac,
        )
        self.btn_rapor_ac.pack(fill="x", pady=(4, 0))

    def ozet_karti_olustur(
        self, ebeveyn, sutun_no, baslik, renk, sol_bosluk=0, sag_bosluk=0
    ):
        kart = ctk.CTkFrame(ebeveyn, corner_radius=12)
        kart.grid(row=0, column=sutun_no, sticky="nsew", padx=(sol_bosluk, sag_bosluk))

        serit = ctk.CTkFrame(kart, height=4, fg_color=renk, corner_radius=2)
        serit.pack(fill="x", padx=14, pady=(12, 8))

        sayi_etiketi = ctk.CTkLabel(
            kart,
            text="-",
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color=renk,
        )
        sayi_etiketi.pack(pady=(4, 4))

        baslik_etiketi = ctk.CTkLabel(
            kart,
            text=baslik,
            font=ctk.CTkFont(size=12),
            text_color="#D1D5DB",
            wraplength=180,
            justify="center",
        )
        baslik_etiketi.pack(padx=10, pady=(0, 14))

        return sayi_etiketi

    # =========================================================================
    # 3. DOSYA SEÇİMİ VE DİNAMİK SÜTUN EŞLEŞTİRME MANTIĞI
    # =========================================================================
    def dun_dosyasi_sec(self):
        yol = filedialog.askopenfilename(
            title="Dünkü Excel veya CSV Dosyasını Seçin",
            filetypes=[
                ("Excel ve CSV Dosyaları", "*.xlsx *.xls *.csv"),
                ("Tüm Dosyalar", "*.*"),
            ],
        )
        if not yol:
            return

        try:
            sutunlar = dosya_sutunlarini_oku(yol)
        except Exception as hata:
            messagebox.showerror("Dosya Okuma Hatası", f"Dünkü dosya okunamadı:\n{hata}")
            return

        self.dun_dosya_yolu = yol
        self.dun_sutunlar = sutunlar
        self.lbl_dun_dosya.configure(text=os.path.basename(yol), text_color="#F3F4F6")
        self.sutun_listelerini_guncelle()

    def bugun_dosyasi_sec(self):
        yol = filedialog.askopenfilename(
            title="Bugünkü Excel veya CSV Dosyasını Seçin",
            filetypes=[
                ("Excel ve CSV Dosyaları", "*.xlsx *.xls *.csv"),
                ("Tüm Dosyalar", "*.*"),
            ],
        )
        if not yol:
            return

        try:
            sutunlar = dosya_sutunlarini_oku(yol)
        except Exception as hata:
            messagebox.showerror("Dosya Okuma Hatası", f"Bugünkü dosya okunamadı:\n{hata}")
            return

        self.bugun_dosya_yolu = yol
        self.bugun_sutunlar = sutunlar
        self.lbl_bugun_dosya.configure(text=os.path.basename(yol), text_color="#F3F4F6")
        self.sutun_listelerini_guncelle()

    def sutun_listelerini_guncelle(self):
        # İki dosya da seçiliyse ortak sütunları önceliklendir
        if self.dun_sutunlar and self.bugun_sutunlar:
            ortak_sutunlar = []
            for sutun in self.dun_sutunlar:
                if sutun in self.bugun_sutunlar:
                    ortak_sutunlar.append(sutun)
            aktif_sutunlar = ortak_sutunlar if ortak_sutunlar else self.dun_sutunlar
        elif self.dun_sutunlar:
            aktif_sutunlar = self.dun_sutunlar
        else:
            aktif_sutunlar = self.bugun_sutunlar

        if not aktif_sutunlar:
            return

        # Combobox seçeneklerini güncelle
        self.cmb_barkod.configure(values=aktif_sutunlar)
        self.cmb_fiyat.configure(values=aktif_sutunlar)
        self.cmb_stok.configure(values=aktif_sutunlar)

        # Barkod, Fiyat ve Stok sütunlarını otomatik tahmin et
        tahmin_barkod = sutun_tahmin_et(
            aktif_sutunlar, ["barkod", "barcode", "ean", "gtin", "urun_kodu"]
        )
        tahmin_fiyat = sutun_tahmin_et(
            aktif_sutunlar, ["alis_fiyat", "fiyat", "price", "maliyet", "tutar"]
        )
        tahmin_stok = sutun_tahmin_et(
            aktif_sutunlar, ["stok_adedi", "stok", "adet", "miktar", "stock", "qty"]
        )

        self.cmb_barkod.set(tahmin_barkod)
        self.cmb_fiyat.set(tahmin_fiyat)
        self.cmb_stok.set(tahmin_stok)

    # =========================================================================
    # 4. ANALİZ VE RAPOR OLUŞTURMA AKIŞI
    # =========================================================================
    def analizi_baslat(self):
        if not self.dun_dosya_yolu or not self.bugun_dosya_yolu:
            messagebox.showwarning(
                "Eksik Dosya", "Lütfen hem dünkü hem de bugünkü dosyayı seçin."
            )
            return

        barkod_sutun = self.cmb_barkod.get()
        fiyat_sutun = self.cmb_fiyat.get()
        stok_sutun = self.cmb_stok.get()

        if barkod_sutun == "Dosya seçiniz" or not barkod_sutun:
            messagebox.showwarning("Eksik Sütun", "Lütfen geçerli sütunları seçin.")
            return

        # İşlem sırasında butonları kilitle
        self.btn_analiz.configure(state="disabled", text="Analiz Ediliyor...")
        self.btn_rapor_ac.configure(state="disabled")
        self.lbl_durum.configure(
            text="Veriler temizleniyor, kıyaslanıyor ve Excel raporu oluşturuluyor...",
            text_color="#60A5FA",
        )

        is_parcacigi = threading.Thread(
            target=self.arka_plan_analiz,
            args=(barkod_sutun, fiyat_sutun, stok_sutun),
            daemon=True,
        )
        is_parcacigi.start()

    def arka_plan_analiz(self, barkod_sutun, fiyat_sutun, stok_sutun):
        try:
            # 1. Adım: Verileri temizle
            df_dun, df_bugun = TEMİZLEME(
                self.dun_dosya_yolu,
                self.bugun_dosya_yolu,
                barkod_sutun=barkod_sutun,
                fiyat_sutun=fiyat_sutun,
                stok_sutun=stok_sutun,
            )

            # 2. Adım: İki tabloyu kıyasla
            acil_df, fiyat_df, yeni_df, ozet = KIYASLA(
                df_dun,
                df_bugun,
                barkod_sutun=barkod_sutun,
                fiyat_sutun=fiyat_sutun,
                stok_sutun=stok_sutun,
            )

            # 3. Adım: Excel raporunu oluştur
            hedef_klasor = os.path.dirname(os.path.abspath(__file__))
            cikti_yolu = os.path.join(hedef_klasor, "karsilastirma_raporu.xlsx")
            kaydedilen_yol = RAPOR_OLUSTUR(
                acil_df, fiyat_df, yeni_df, ozet, cikti_yolu=cikti_yolu
            )

            self.after(0, self.analiz_tamamlandi, ozet, kaydedilen_yol)

        except Exception as hata:
            self.after(0, self.analiz_hatasi, str(hata))

    def analiz_tamamlandi(self, ozet, kaydedilen_yol):
        self.olusturulan_rapor_yolu = kaydedilen_yol

        # Özet kartlarını güncelle
        self.lbl_acil_sayi.configure(text=str(ozet["acil_kapatilacak"]))
        self.lbl_fiyat_sayi.configure(text=str(ozet["fiyat_degisen"]))
        self.lbl_yeni_sayi.configure(text=str(ozet["yeni_eklenen"]))

        # Butonları ve durum mesajını güncelle
        self.btn_analiz.configure(state="normal", text="Analiz Et ve Rapor Oluştur")
        self.btn_rapor_ac.configure(state="normal")
        self.lbl_durum.configure(
            text=f"İşlem tamamlandı! Rapor kaydedildi: {os.path.basename(kaydedilen_yol)}",
            text_color="#34D399",
        )

    def analiz_hatasi(self, hata_mesaji):
        self.btn_analiz.configure(state="normal", text="Analiz Et ve Rapor Oluştur")
        self.lbl_durum.configure(
            text="Analiz sırasında bir hata oluştu.", text_color="#F87171"
        )
        messagebox.showerror(
            "Analiz Hatası",
            f"İşlem sırasında bir hata meydana geldi:\n\n{hata_mesaji}",
        )

    # =========================================================================
    # 5. OLUŞTURULAN EXCEL RAPORUNU SİSTEM PROGRAMIYLA AÇMA
    # =========================================================================
    def raporu_ac(self):
        if not self.olusturulan_rapor_yolu or not os.path.exists(
            self.olusturulan_rapor_yolu
        ):
            messagebox.showerror(
                "Dosya Bulunamadı", "Oluşturulan Excel raporu bulunamadı."
            )
            return

        try:
            sistem = platform.system()
            if sistem == "Windows":
                os.startfile(self.olusturulan_rapor_yolu)
            elif sistem == "Darwin":
                subprocess.Popen(["open", self.olusturulan_rapor_yolu])
            else:
                subprocess.Popen(["xdg-open", self.olusturulan_rapor_yolu])
        except Exception as hata:
            messagebox.showerror(
                "Açma Hatası", f"Rapor dosyası açılamadı:\n{hata}"
            )


if __name__ == "__main__":
    uygulama = RaporlamaUygulamasi()
    uygulama.mainloop()
