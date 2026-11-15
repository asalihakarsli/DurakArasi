import heapq
import unicodedata
from pathlib import Path
import pandas as pd

# Temel Ayarlar
BASE_DIR = Path(__file__).resolve().parent
EXCEL_DOSYASI = BASE_DIR / "ankara_metro_kitaplar.xlsx"
CSV_DOSYASI = BASE_DIR / "ankara_metro_kitaplar.csv"

OKUMA_HIZI_SAAT = 40.0  # Yetişkin okuma hızı: saatte 40 sayfa
OKUMA_HIZI_DK = OKUMA_HIZI_SAAT / 60.0  # Dakikada ~0.667 sayfa

# Ankara Raylı Sistem Hatları (EGO Raylı Sistemler Haritası baz alınarak)
HAT_DURAKLARI = {
    "Ankaray (A1)": {
        "durak_arasi_dk": 1.8,  # 10 aralık ≈ 18 dk
        "duraklar": [
            "AŞTİ",
            "Emek",
            "Bahçelievler",
            "Beşevler",
            "Anadolu/Anıtkabir (Tandoğan)",
            "Maltepe",
            "Demirtepe",
            "15 Temmuz Kızılay Milli İrade",
            "Kolej",
            "Kurtuluş",
            "Dikimevi",
        ],
    },
    "Keçiören Metrosu (M4)": {
        "durak_arasi_dk": 2.5,  # 11 aralık ≈ 27.5 dk
        "duraklar": [
            "Şehitler",
            "Dutluk",
            "Kuyubaşı",
            "Mecidiye",
            "Belediye",
            "Meteoroloji",
            "Dışkapı",
            "ASKİ",
            "Atatürk Kültür Merkezi",
            "Gar",
            "Adliye",
            "15 Temmuz Kızılay Milli İrade",
        ],
    },
    "Normal Metro (M1-M2-M3)": {
        "durak_arasi_dk": 2.0,  # 33 aralık ≈ 66 dk (Koru - Kızılay - Batıkent - OSB-Törekent)
        "duraklar": [
            "Koru",
            "Çayyolu",
            "Ümitköy",
            "Beytepe",
            "Tarım Bakanlığı/Danıştay",
            "Bilkent",
            "ODTÜ",
            "MTA",
            "Söğütözü",
            "Milli Kütüphane",
            "Necatibey",
            "15 Temmuz Kızılay Milli İrade",
            "Sıhhiye",
            "Ulus",
            "Atatürk Kültür Merkezi",
            "Akköprü",
            "İvedik",
            "Yenimahalle",
            "Demetevler",
            "Hastane",
            "Macunköy",
            "OSTİM",
            "Batıkent",
            "Batı Merkez",
            "Mesa",
            "Botanik",
            "İstanbul Yolu",
            "Eryaman 1-2",
            "Eryaman 5",
            "Devlet Mah.",
            "Harikalar Diyarı",
            "Fatih",
            "GOP",
            "OSB-Törekent",
        ],
    },
}

# Hatlar Arası Aktarma Bağlantıları (Durak 1, Durak 2, Aktarma Süresi Dk)
AKTARMALAR = [
    ("15 Temmuz Kızılay Milli İrade", "15 Temmuz Kızılay Milli İrade", 4.0),
    ("Atatürk Kültür Merkezi", "Atatürk Kültür Merkezi", 3.5),
    ("Maltepe", "Gar", 5.0),  # Ankaray Maltepe - M4/YHT Gar yaya alt geçit bağlantısı
]

# Veri setindeki 8 aktif kütüphanenin en yakın metro/Ankaray durakları
KUTUPHANE_DURAK_ESLESMESI = {
    "Bahçelievler Genç Akademi": {
        "durak": "Bahçelievler",
        "hat": "Ankaray (A1)",
    },
    "Büyükşehir Genç Akademi Batıkent": {
        "durak": "Batıkent",
        "hat": "Normal Metro (M1-M2-M3)",
    },
    "Büyükşehir Genç Akademi Kafe Sıhhiye": {
        "durak": "Sıhhiye",
        "hat": "Normal Metro (M1-M2-M3)",
    },
    "Kolej TED Çalışma İstasyonu": {
        "durak": "Kolej",
        "hat": "Ankaray (A1)",
    },
    "Meteoroloji Çalışma İstasyonu": {
        "durak": "Meteoroloji",
        "hat": "Keçiören Metrosu (M4)",
    },
    "ODTÜ Çalışma İstasyonu": {
        "durak": "ODTÜ",
        "hat": "Normal Metro (M1-M2-M3)",
    },
    "Prof. Dr. Ruşen Keleş Kütüphanesi": {
        "durak": "15 Temmuz Kızılay Milli İrade",
        "hat": "Normal Metro (M1-M2-M3) / Ankaray (A1) / M4",
    },
    "YHT Gar Çalışma İstasyonu": {
        "durak": "Gar",
        "hat": "Keçiören Metrosu (M4) / Ankaray (Maltepe Bağlantılı)",
    },
}


def metin_normallestir(metin: str) -> str:
    """Türkçe karakter ve büyük/küçük harf duyarsız karşılaştırma için metni sadeleştirir."""
    tr_map = str.maketrans(
        {
            "ç": "c",
            "ğ": "g",
            "ı": "i",
            "İ": "i",
            "ö": "o",
            "ş": "s",
            "ü": "u",
            "Ç": "c",
            "Ğ": "g",
            "I": "i",
            "Ö": "o",
            "Ş": "s",
            "Ü": "u",
        }
    )
    sade = metin.translate(tr_map).lower().strip()
    sade = unicodedata.normalize("NFKD", sade)
    return "".join(ch for ch in sade if not unicodedata.combining(ch))


class AnkaraMetroAgi:
    """Ankara raylı sistem hatları üzerinde en kısa rota ve süre hesaplayıcı."""

    def __init__(self):
        self.graf: dict[tuple[str, str], list[tuple[tuple[str, str], float, int]]] = {}
        self.tum_duraklar: list[str] = []
        self.durak_hatlari: dict[str, list[str]] = {}
        self._agi_kur()

    def _agi_kur(self) -> None:
        gorulen_duraklar = set()

        for hat_adi, bilgi in HAT_DURAKLARI.items():
            duraklar = bilgi["duraklar"]
            adim_dk = bilgi["durak_arasi_dk"]

            for idx, durak in enumerate(duraklar):
                if durak not in gorulen_duraklar:
                    gorulen_duraklar.add(durak)
                    self.tum_duraklar.append(durak)

                self.durak_hatlari.setdefault(durak, []).append(hat_adi)
                dugum = (durak, hat_adi)
                self.graf.setdefault(dugum, [])

                if idx > 0:
                    onceki = (duraklar[idx - 1], hat_adi)
                    self.graf[dugum].append((onceki, adim_dk, 1))
                    self.graf[onceki].append((dugum, adim_dk, 1))

        # Aynı istasyondaki farklı hatlar arası aktarmalar
        for durak, hatlar in self.durak_hatlari.items():
            if len(hatlar) > 1:
                for i in range(len(hatlar)):
                    for j in range(i + 1, len(hatlar)):
                        u = (durak, hatlar[i])
                        v = (durak, hatlar[j])
                        aktarma_dk = 3.5 if durak == "Atatürk Kültür Merkezi" else 4.0
                        self.graf[u].append((v, aktarma_dk, 0))
                        self.graf[v].append((u, aktarma_dk, 0))

        # Özel istasyonlar arası yaya bağlantıları (Maltepe <-> Gar)
        for d1, d2, sure in AKTARMALAR:
            if d1 != d2:
                for h1 in self.durak_hatlari.get(d1, []):
                    for h2 in self.durak_hatlari.get(d2, []):
                        u = (d1, h1)
                        v = (d2, h2)
                        self.graf[u].append((v, sure, 0))
                        self.graf[v].append((u, sure, 0))

    def durak_cozumle(self, girdi: str) -> tuple[str | None, list[str]]:
        """
        Kullanıcı girdisini (durak numarası, tam ad veya kısmi ad) durak adına çevirir.
        Dönüş: (bulunan_durak, olasi_eslesmeler)
        """
        temiz = girdi.strip()
        if not temiz:
            return None, []

        # 1. Numara ile seçim kontrolü
        if temiz.isdigit():
            no = int(temiz)
            if 1 <= no <= len(self.tum_duraklar):
                secilen = self.tum_duraklar[no - 1]
                return secilen, [secilen]
            return None, []

        aranan = metin_normallestir(temiz)

        # 2. Tam eşleşme kontrolü
        for d in self.tum_duraklar:
            if metin_normallestir(d) == aranan:
                return d, [d]

        # 3. Kısmi eşleşme kontrolü (ör. "kızılay", "tandoğan", "akm")
        eslesenler = [d for d in self.tum_duraklar if aranan in metin_normallestir(d)]
        if len(eslesenler) == 1:
            return eslesenler[0], eslesenler
        return None, eslesenler

    def rota_hesapla(self, baslangic: str, varis: str) -> dict:
        """İki durak arasındaki en kısa süreyi (dk), durak sayısını ve güzergahı hesaplar."""
        if baslangic == varis:
            return {
                "baslangic": baslangic,
                "varis": varis,
                "sure_dk": 0.0,
                "durak_sayisi": 0,
                "aktarma_sayisi": 0,
                "hatlar": self.durak_hatlari.get(baslangic, []),
                "guzergah_duraklari": [baslangic],
            }

        kuyruk = []
        en_iyi = {}
        onceki_dugum = {}

        for hat in self.durak_hatlari[baslangic]:
            dugum = (baslangic, hat)
            en_iyi[dugum] = (0.0, 0, 0)  # (sure_dk, durak_sayisi, aktarma_sayisi)
            heapq.heappush(kuyruk, (0.0, 0, 0, dugum))

        hedef_dugum = None

        while kuyruk:
            sure, durak_say, aktarma_say, mevcut = heapq.heappop(kuyruk)
            if mevcut[0] == varis:
                hedef_dugum = mevcut
                break

            if (sure, durak_say, aktarma_say) > en_iyi.get(mevcut, (float("inf"), 0, 0)):
                continue

            for komsu, agirlik_dk, durak_adimi in self.graf.get(mevcut, []):
                yeni_sure = sure + agirlik_dk
                yeni_durak = durak_say + durak_adimi
                yeni_aktarma = aktarma_say + (1 if mevcut[1] != komsu[1] else 0)
                aday = (yeni_sure, yeni_durak, yeni_aktarma)

                if aday < en_iyi.get(komsu, (float("inf"), 0, 0)):
                    en_iyi[komsu] = aday
                    onceki_dugum[komsu] = mevcut
                    heapq.heappush(kuyruk, (yeni_sure, yeni_durak, yeni_aktarma, komsu))

        if hedef_dugum is None:
            raise ValueError(f"'{baslangic}' ile '{varis}' arasında rota bulunamadı.")

        # Yolu geriye doğru izle
        yol = []
        imlec = hedef_dugum
        while imlec is not None:
            yol.append(imlec)
            imlec = onceki_dugum.get(imlec)
        yol.reverse()

        guzergah_duraklari = []
        kullanilan_hatlar = []
        for durak, hat in yol:
            if not guzergah_duraklari or guzergah_duraklari[-1] != durak:
                guzergah_duraklari.append(durak)
            if not kullanilan_hatlar or kullanilan_hatlar[-1] != hat:
                kullanilan_hatlar.append(hat)

        toplam_sure, toplam_durak, toplam_aktarma = en_iyi[hedef_dugum]
        return {
            "baslangic": baslangic,
            "varis": varis,
            "sure_dk": round(toplam_sure, 1),
            "durak_sayisi": toplam_durak,
            "aktarma_sayisi": toplam_aktarma,
            "hatlar": kullanilan_hatlar,
            "guzergah_duraklari": guzergah_duraklari,
        }


def veri_setini_yukle() -> pd.DataFrame:
    """Veri setini okur: önce Excel, Excel yoksa CSV (ikisinde de başlık satırı 7. satırdadır)."""
    if EXCEL_DOSYASI.exists():
        df = pd.read_excel(EXCEL_DOSYASI, sheet_name="Temiz Kitap Listesi", header=6)
    elif CSV_DOSYASI.exists():
        df = pd.read_csv(CSV_DOSYASI, encoding="utf-8-sig", header=6)
    else:
        raise FileNotFoundError(
            f"Veri seti bulunamadı! '{EXCEL_DOSYASI.name}' veya '{CSV_DOSYASI.name}' dosyasını kod ile aynı klasöre koyun."
        )

    if "Katalog No" in df.columns:
        df["Link"] = df["Katalog No"].apply(
            lambda x: f"https://katalog.ankara.bel.tr/cgi-bin/koha/opac-detail.pl?biblionumber={int(x)}"
            if pd.notna(x)
            else ""
        )

    df = df.dropna(subset=["Kütüphane", "Kitap Adı", "Sayfa Sayısı"]).copy()
    df["Sayfa Sayısı"] = df["Sayfa Sayısı"].astype(int)
    df["Tahmini Okunma Süresi (Dk)"] = ((df["Sayfa Sayısı"] / OKUMA_HIZI_SAAT) * 60).round(1)
    df["Tahmini Okunma Süresi (Saat)"] = (df["Sayfa Sayısı"] / OKUMA_HIZI_SAAT).round(2)
    return df


def kitaplari_oner(
    df: pd.DataFrame,
    metro_agi: AnkaraMetroAgi,
    binis_duragi: str,
    inis_duragi: str,
    gidis_donus: bool = False,
    oneri_sayisi: int = 6,
) -> dict:
    """
    Seçilen biniş ve iniş durakları arasındaki yolculuk süresine göre kitap önerir:
    - Süre içinde bitirilebilen kitaplar varsa onları önceliklendirir (%100).
    - Bitirilebilecek kitap yoksa, her kitabın % kaçının okunabileceğini hesaplar.
    - Her kitap için kullanıcının bulunduğuna (biniş durağına) en yakın alış durağını bulur.
    """
    rota = metro_agi.rota_hesapla(binis_duragi, inis_duragi)
    yolculuk_suresi_dk = round(rota["sure_dk"] * (2 if gidis_donus else 1), 1)
    okunabilecek_sayfa = round(yolculuk_suresi_dk * OKUMA_HIZI_DK, 1)

    # Bulunduğumuz durağa göre her kütüphanenin mesafesini (durak ve dakika) hesapla
    kutuphane_mesafeleri = {}
    for ktp_adi, ktp_bilgi in KUTUPHANE_DURAK_ESLESMESI.items():
        ktp_durak = ktp_bilgi["durak"]
        mesafe_rota = metro_agi.rota_hesapla(binis_duragi, ktp_durak)
        kutuphane_mesafeleri[ktp_adi] = {
            "alinacak_durak": ktp_durak,
            "hat": ktp_bilgi["hat"],
            "uzaklik_durak": mesafe_rota["durak_sayisi"],
            "uzaklik_dk": mesafe_rota["sure_dk"],
            "guzergah_uzerinde": ktp_durak in rota["guzergah_duraklari"],
        }

    # Bulunduğumuz durağa genel olarak en yakın kütüphane durağı
    en_yakin_ktp_adi = min(
        kutuphane_mesafeleri,
        key=lambda k: (
            kutuphane_mesafeleri[k]["uzaklik_dk"],
            kutuphane_mesafeleri[k]["uzaklik_durak"],
        ),
    )
    en_yakin_ktp_bilgi = {
        "kutuphane": en_yakin_ktp_adi,
        **kutuphane_mesafeleri[en_yakin_ktp_adi],
    }

    # Veri setindeki her satıra alış durağı mesafesi ve okuma yüzdesi ekle
    calisma_df = df.copy()
    calisma_df["Alınacak Durak"] = calisma_df["Kütüphane"].map(
        lambda k: kutuphane_mesafeleri.get(k, {}).get("alinacak_durak", "Bilinmiyor")
    )
    calisma_df["Alış Durağı Uzaklığı (Durak)"] = calisma_df["Kütüphane"].map(
        lambda k: kutuphane_mesafeleri.get(k, {}).get("uzaklik_durak", 999)
    )
    calisma_df["Alış Durağı Uzaklığı (Dk)"] = calisma_df["Kütüphane"].map(
        lambda k: kutuphane_mesafeleri.get(k, {}).get("uzaklik_dk", 999.0)
    )
    calisma_df["Güzergah Üzerinde"] = calisma_df["Kütüphane"].map(
        lambda k: kutuphane_mesafeleri.get(k, {}).get("guzergah_uzerinde", False)
    )

    # Okuma yüzdesi (%) ve okunabilecek sayfa hesabı
    calisma_df["Okunabilecek Sayfa"] = calisma_df["Sayfa Sayısı"].apply(
        lambda s: min(float(s), okunabilecek_sayfa)
    )
    calisma_df["Okuma Yüzdesi (%)"] = calisma_df.apply(
        lambda r: min(100.0, round((yolculuk_suresi_dk / max(r["Tahmini Okunma Süresi (Dk)"], 0.1)) * 100, 1)),
        axis=1,
    )
    calisma_df["Tamamı Bitirilebilir"] = calisma_df["Tahmini Okunma Süresi (Dk)"] <= yolculuk_suresi_dk

    # Bulunduğumuz durağa en yakın kütüphanedeki en kısa eser
    en_yakin_sube_kitaplari = calisma_df[calisma_df["Kütüphane"] == en_yakin_ktp_adi].sort_values("Sayfa Sayısı")
    en_yakin_sube_en_kisa = (
        en_yakin_sube_kitaplari.iloc[0].to_dict() if not en_yakin_sube_kitaplari.empty else None
    )

    # Aynı kitap birden fazla şubede varsa, bulunduğumuz durağa EN YAKIN şubedeki kopyayı seç
    esiz_anahtar = calisma_df["Katalog No"].astype(str) if "Katalog No" in calisma_df.columns else calisma_df["Kitap Adı"]
    calisma_df["_eser_id"] = esiz_anahtar
    calisma_df = (
        calisma_df.sort_values(
            ["Alış Durağı Uzaklığı (Dk)", "Alış Durağı Uzaklığı (Durak)", "Kütüphane"]
        )
        .drop_duplicates(subset=["_eser_id"], keep="first")
        .drop(columns=["_eser_id"])
    )

    bitirilebilenler = calisma_df[calisma_df["Tamamı Bitirilebilir"]].copy()
    tamami_biten_var = not bitirilebilenler.empty

    if tamami_biten_var:
        # Süre içinde tamamen bitirilebilen kitapları:
        # 1) Alış durağı yakınlığı, 2) Yolculuk süresine en iyi uyum sırasıyla diz
        bitirilebilenler["_sure_farki"] = (
            yolculuk_suresi_dk - bitirilebilenler["Tahmini Okunma Süresi (Dk)"]
        ).abs()
        oneriler = (
            bitirilebilenler.sort_values(
                ["Alış Durağı Uzaklığı (Dk)", "_sure_farki", "Sayfa Sayısı"]
            )
            .drop(columns=["_sure_farki"])
            .head(oneri_sayisi)
        )
    else:
        # Seçilen duraklar arasında bitirilebilecek kitap yoksa:
        # En yüksek yüzdesi okunabilen kitapları (en kısa eserler + en yakın durak) sırala
        oneriler = (
            calisma_df.sort_values(
                ["Okuma Yüzdesi (%)", "Alış Durağı Uzaklığı (Dk)", "Sayfa Sayısı"],
                ascending=[False, True, True],
            )
            .head(oneri_sayisi)
        )

    return {
        "rota": rota,
        "gidis_donus": gidis_donus,
        "yolculuk_suresi_dk": yolculuk_suresi_dk,
        "okunabilecek_sayfa": okunabilecek_sayfa,
        "en_yakin_kutuphane": en_yakin_ktp_bilgi,
        "en_yakin_sube_en_kisa": en_yakin_sube_en_kisa,
        "tamami_biten_var": tamami_biten_var,
        "toplam_bitirilebilir_sayisi": len(bitirilebilenler),
        "oneriler": oneriler.reset_index(drop=True),
    }


def duraklari_listele(metro_agi: AnkaraMetroAgi) -> None:
    """Tüm metro/Ankaray duraklarını numaralı ve hat bilgili olarak ekrana yazdırır."""
    print("\n" + "=" * 78)
    print("ANKARA RAYLI SİSTEM DURAK LİSTESİ (Numara veya Durak Adı Yazabilirsiniz)")
    print("=" * 78)
    for idx, durak in enumerate(metro_agi.tum_duraklar, start=1):
        hatlar = " | ".join(metro_agi.durak_hatlari[durak])
        ktp_notu = ""
        for ktp_adi, bilgi in KUTUPHANE_DURAK_ESLESMESI.items():
            if bilgi["durak"] == durak:
                ktp_notu = f"  <-- [Kütüphane: {ktp_adi}]"
                break
        print(f"  {idx:2d}. {durak:<34} ({hatlar}){ktp_notu}")
    print("=" * 78)


def kullanicidan_durak_sec(metro_agi: AnkaraMetroAgi, istem_metni: str) -> str:
    """Kullanıcı geçerli bir durak seçinceye kadar girdi ister."""
    while True:
        girdi = input(istem_metni).strip()
        if girdi.lower() in {"liste", "duraklar", "l"}:
            duraklari_listele(metro_agi)
            continue

        bulunan, adaylar = metro_agi.durak_cozumle(girdi)
        if bulunan:
            return bulunan
        if adaylar:
            print(f"  Birden fazla durak eşleşti: {', '.join(adaylar)}. Lütfen daha net yazın.")
        else:
            print("  Durak bulunamadı. Durak numarasını veya adını tekrar girin (listeyi görmek için 'liste' yazın).")


def sonuclari_yazdir(sonuc: dict) -> None:
    """Hesaplanan rota, en yakın alış durağı ve kitap önerilerini düzenli biçimde gösterir."""
    rota = sonuc["rota"]
    yon_etiketi = "Gidiş-Dönüş" if sonuc["gidis_donus"] else "Tek Yön"
    en_yakin = sonuc["en_yakin_kutuphane"]
    yakin_kitap = sonuc.get("en_yakin_sube_en_kisa")

    print("\n" + "=" * 88)
    print(f"ROTA ÖZETİ: {rota['baslangic']}  --->  {rota['varis']} ({yon_etiketi})")
    print("=" * 88)
    print(f"  Kullanılan Hat(lar)       : {' -> '.join(rota['hatlar'])}")
    print(f"  Durak / Aktarma Sayısı    : {rota['durak_sayisi']} durak | {rota['aktarma_sayisi']} aktarma")
    print(f"  Toplam Yolculuk Süresi    : {sonuc['yolculuk_suresi_dk']} dakika")
    print(f"  Okuma Kapasiteniz (40s/sa): Yaklaşık {sonuc['okunabilecek_sayfa']} sayfa")
    print(
        f"  Size En Yakın Kütüphane   : {en_yakin['kutuphane']} "
        f"({en_yakin['alinacak_durak']} durağı — bulunduğunuz durağa {en_yakin['uzaklik_durak']} durak / {en_yakin['uzaklik_dk']} dk)"
    )
    if yakin_kitap:
        print(
            f"  En Yakın Duraktaki En Kısa: {yakin_kitap['Kitap Adı']} ({yakin_kitap['Sayfa Sayısı']} sayfa) "
            f"— Yolculukta okuma oranı: %{yakin_kitap['Okuma Yüzdesi (%)']}"
        )
    print("-" * 88)

    if sonuc["tamami_biten_var"]:
        print(
            f"  [DURUM] Bu yolculuk süresinde tamamını (%100) bitirebileceğiniz "
            f"{sonuc['toplam_bitirilebilir_sayisi']} farklı kitap bulundu:\n"
        )
    else:
        print(
            f"  [BİLGİ] Seçilen duraklar arasındaki {sonuc['yolculuk_suresi_dk']} dakikalık sürede "
            f"tamamı bitirilebilecek bir kitap bulunmamaktadır (en fazla ~{sonuc['okunabilecek_sayfa']} sayfa okunabilir)."
        )
        print("  Bu yolculukta en yüksek yüzdesini okuyabileceğiniz kitaplar:\n")

    for i, row in sonuc["oneriler"].iterrows():
        yuzde_metni = (
            "%100.0 (Tamamı Bitirilebilir)"
            if row["Tamamı Bitirilebilir"]
            else f"%{row['Okuma Yüzdesi (%)']} ({row['Okunabilecek Sayfa']:.0f}/{row['Sayfa Sayısı']} sayfa)"
        )
        uzaklik_metni = (
            "Bulunduğunuz durakta (0 durak)"
            if row["Alış Durağı Uzaklığı (Durak)"] == 0
            else f"Bulunduğunuz durağa {row['Alış Durağı Uzaklığı (Durak)']} durak / {row['Alış Durağı Uzaklığı (Dk)']} dk"
        )
        guzergah_notu = " [Güzergahınız üzerinde]" if row["Güzergah Üzerinde"] else ""

        print(f"{i + 1}. {row['Kitap Adı']} — {row['Yazar']}")
        print(f"   • Sayfa & Süre    : {row['Sayfa Sayısı']} sayfa | Tahmini Okuma: {row['Tahmini Okunma Süresi (Dk)']} dk")
        print(f"   • Yolculukta Okuma: {yuzde_metni}")
        print(f"   • Kitabı Alacağınız Durak: {row['Alınacak Durak']} ({row['Kütüphane']}) — {uzaklik_metni}{guzergah_notu}")
        if row.get("Link"):
            print(f"   • Katalog Linki   : {row['Link']}")
        print()


if __name__ == "__main__":
    df_kitaplar = veri_setini_yukle()
    metro_agi = AnkaraMetroAgi()

    print(f"Veri seti yüklendi: {len(df_kitaplar)} envanter kaydı ({df_kitaplar['Katalog No'].nunique()} benzersiz eser).")
    duraklari_listele(metro_agi)

    while True:
        print("\nÖneri almak için biniş ve iniş duraklarınızı seçin (Çıkmak için 'q' yazın):")
        binis_girdi = input("1) Hangi duraktan bineceksiniz? (Bulunduğunuz durak): ").strip()
        if binis_girdi.lower() in {"q", "cikis", "çıkış", "exit"}:
            print("İyi yolculuklar ve keyifli okumalar!")
            break

        bulunan_binis, aday_binis = metro_agi.durak_cozumle(binis_girdi)
        if not bulunan_binis:
            if aday_binis:
                print(f"  Birden fazla durak eşleşti: {', '.join(aday_binis)}. Lütfen daha net yazın.")
            else:
                print("  Durak bulunamadı. Durak numarasını veya adını tekrar girin.")
            continue

        inis_duragi = kullanicidan_durak_sec(metro_agi, "2) Hangi durağa gideceksiniz? (Varış durağı): ")

        if bulunan_binis == inis_duragi:
            print("  Biniş ve iniş durağı aynı seçildi! Lütfen farklı iki durak seçin.")
            continue

        yon_secimi = input("3) Yolculuk tipi [1: Tek Yön (Varsayılan) | 2: Gidiş-Dönüş]: ").strip()
        gidis_donus = yon_secimi == "2"

        sonuc = kitaplari_oner(
            df=df_kitaplar,
            metro_agi=metro_agi,
            binis_duragi=bulunan_binis,
            inis_duragi=inis_duragi,
            gidis_donus=gidis_donus,
            oneri_sayisi=6,
        )
        sonuclari_yazdir(sonuc)
