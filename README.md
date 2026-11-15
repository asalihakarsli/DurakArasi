# 🚇 DurakArası — Ankara Metro Kitap Önerici

Ankara raylı sistem (Ankaray, Keçiören Metrosu, M1-M2-M3) yolculuğunuzun süresine göre, **yolculuk bitmeden okuyabileceğiniz** kısa kitapları öneren bir araç. Kitaplar, Ankara Büyükşehir Belediyesi kütüphanelerinin metro hattı üzerindeki şubelerinden seçilir.

Proje iki parçadan oluşur:

- **Komut satırı uygulaması (Python):** Durak seçersiniz, rota ve kitap önerilerini terminalde görürsünüz.
- **Web arayüzü (HTML):** Etkileşimli şematik metro haritası üzerinden aynı mantıkla öneri verir.

## Nasıl çalışır?

1. Biniş ve iniş durakları arasındaki en kısa süre, **Dijkstra algoritması** ile hesaplanır (aktarma süreleri dahil).
2. Okuma hızı **saatte 40 sayfa** kabul edilerek yolculukta kaç sayfa okunabileceği bulunur.
3. Yolculuk süresinde tamamen bitirilebilen kitaplar varsa bunlar önerilir. Yoksa her kitabın yolculukta **yüzde kaçının** okunabileceği hesaplanır ve en yüksek orana sahip kitaplar listelenir.
4. Öneriler, bulunduğunuz durağa **en yakın kütüphane şubesine** göre sıralanır. Aynı eser birden fazla şubede varsa size en yakın şubedeki kopya gösterilir.
5. Her kitap için Ankara Büyükşehir Belediyesi katalog bağlantısı verilir.
6. Tek yön veya gidiş-dönüş seçilebilir.

## Dosyalar

| Dosya | Açıklama |
|---|---|
| `ankara_metro_kitap_onerici-v2.py` | Komut satırı uygulaması (rota hesaplama + kitap önerisi) |
| `ankara_metro_kitaplar.xlsx` | Veri seti: kitap listesi, kütüphane/hat özeti, metro hattı rehberi |
| `ankara_metro_kitaplar.csv` | Kitap listesinin CSV kopyası (Excel dosyası yoksa betik bunu kullanır) |
| `metro_kitap_onerici.html` | Web arayüzü (DurakArası), veri sayfanın içine gömülüdür |
| `requirements.txt` | Python bağımlılıkları |

## Kurulum

Python **3.10 veya üzeri** gerekir.

```bash
git clone <repo-adresi>
cd <repo-klasoru>

# (Önerilir) sanal ortam
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

pip install -r requirements.txt
```

## Kullanım

### Komut satırı

```bash
python ankara_metro_kitap_onerici-v2.py
```

Program sırasıyla şunları sorar:

1. Biniş durağı (numara, tam ad veya kısmi ad: `kızılay`, `odtü`, `akm` gibi)
2. Varış durağı
3. Yolculuk tipi (`1`: tek yön, `2`: gidiş-dönüş)

Durak listesini görmek için `liste` yazın, çıkmak için `q` yazın.

Örnek çıktı:

```
ROTA ÖZETİ: Şehitler  --->  ODTÜ (Tek Yön)
  Kullanılan Hat(lar)       : Keçiören Metrosu (M4) -> Normal Metro (M1-M2-M3)
  Toplam Yolculuk Süresi    : ... dakika
  Okuma Kapasiteniz (40s/sa): Yaklaşık ... sayfa
```

> Betik veri setini kendi klasöründen okur: önce `ankara_metro_kitaplar.xlsx`, bulamazsa `ankara_metro_kitaplar.csv` kullanılır. `.py` dosyası ve veri dosyası aynı klasörde olmalıdır.

### Web arayüzü

`metro_kitap_onerici.html` dosyasını tarayıcıda açmanız yeterlidir, kurulum gerekmez. Haritadan biniş ve iniş duraklarına tıklayın veya listeden seçin. Arayüz Tailwind CSS ve Google Fonts'u CDN üzerinden yüklediği için **internet bağlantısı** gerekir.

## Veri seti

- **8 aktif kütüphane / çalışma istasyonu**, 80 envanter kaydı (72 benzersiz eser)
- Yalnızca **80 sayfadan kısa** eserler (yaklaşık 2 saatten kısa okuma süresi) alınmıştır
- Kaynak: Ankara Büyükşehir Belediyesi Kütüphaneleri kataloğu (KOHA)
- Sütunlar: Kütüphane, Yakın Hat, Kitap Adı, Yazar, Kitap Dili, Sayfa Sayısı, Sayfa Kategorisi, Tahmini Okunma Süresi, Bulunduğu Şube Sayısı, Katalog No

| Kütüphane | Durak |
|---|---|
| Bahçelievler Genç Akademi | Bahçelievler |
| Büyükşehir Genç Akademi Batıkent | Batıkent |
| Büyükşehir Genç Akademi Kafe Sıhhiye | Sıhhiye |
| Kolej TED Çalışma İstasyonu | Kolej |
| Meteoroloji Çalışma İstasyonu | Meteoroloji |
| ODTÜ Çalışma İstasyonu | ODTÜ |
| Prof. Dr. Ruşen Keleş Kütüphanesi | 15 Temmuz Kızılay Milli İrade |
| YHT Gar Çalışma İstasyonu | Gar |

## Varsayımlar ve sınırlamalar

- Durak arası süreler tahminidir (Ankaray 1,8 dk, M4 2,5 dk, M1-M2-M3 2,0 dk). Gerçek seferlerle birebir örtüşmeyebilir.
- Aktarma süreleri: Kızılay 4 dk, AKM 3,5 dk, Maltepe ↔ Gar yaya bağlantısı 5 dk.
- Bekleme süresi hesaba katılmaz.
- Okuma hızı sabittir (40 sayfa/saat). Kod içindeki `OKUMA_HIZI_SAAT` değerini değiştirerek ayarlayabilirsiniz.
- Kitapların rafta olup olmadığı kontrol edilmez, katalogdan teyit edin.
- Raylı sistem hatları ve durakları zamanla değişebilir. Güncel bilgi için EGO resmi kaynaklarına bakın.

## Teknolojiler

Python (pandas, openpyxl, heapq), HTML, Tailwind CSS, vanilla JavaScript, SVG

## Lisans

Lisansı henüz belirtilmedi. Yayınlamadan önce uygun bir lisans (örn. MIT) ekleyin.
