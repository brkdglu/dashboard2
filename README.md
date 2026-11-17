# Hybrid (HB) Segment Raporu

Aylık müşteri portföyü CSV'lerinden (`YYYYMM.csv`) tek dosyalık, paylaşılabilir HTML rapor üreten Jupyter notebook.

## Kullanım
1. `hybrid_rapor.ipynb` dosyasını Jupyter'de açın.
2. **Ayarlar** hücresinde `VERI_KAYNAGI = "csv"` ve `VERI_KLASORU = "klasör_yolu"` yazın (klasörde `202603.csv`, `202604.csv` … bulunmalı).
3. **Run All** → `hybrid_rapor.html` oluşur (internetsiz açılır, e-postayla paylaşılabilir).
4. "Veri doğrulama" çıktısını kontrol edin (eksik kolonlar, cinsiyet/cihaz kolonu eşleşmesi, hareket adetleri).

Gerekli paketler: `numpy`, `pandas` (ve internetsiz HTML için `plotly`). Denemek için `VERI_KAYNAGI = "sentetik"` bırakın.

Kolon adları farklıysa Ayarlar hücresindeki `CINSIYET_KOLONU`, `CIHAZ_KOLONU`, `YAS_KOLONU`, `AKTIF_DURUM_IDS`, `SOZLESME_ISIM` vb. değerleri düzenleyin.

---

## Rahat Segment Panosu (yeni rapor)

`rapor_uret.py`, veri klasöründen dört sayfalık tek dosyalık HTML rapor üretir (Müşteri Profili · Aylık Karşılaştırma · Churn & Kazanım · Model & KPI).

```bash
python rapor_uret.py --veri data --cikti rahat_segment_raporu.html
```

Notebook içinden: `%run rapor_uret.py --veri data`. Tek dosyadır: rapor şablonu `rapor_uret.py` içine gömülüdür. Gerekenler: `numpy`, `pandas` (KPI Excel'i için `openpyxl`).

### Klasör yapısı

```
data/
  raw_data/    202604.csv …        aylık müşteri dosyaları
  kohort/      *.csv / *.xlsx      geniş kohort tablosu
  lifecycle/   *.csv / *.xlsx      lifecycle tablosu
  kpi/         *.xlsx              KPI tablosu (her dosya bir proje)
  post_model/  (isteğe bağlı)      müşteri bazlı skor / liste / kontrol tabloları
```

- **raw_data:** dosya adı `YYYYMM.csv`. Kullanılan kolonlar: `CAMPAIGN_SUBSCRIBER_ID`, `SUBSCRIBERS_LAST_TENURE_DAYS` (aya çevrilir), `OPTION_AND_PACKAGE_NAME` (tarife), `TARIFF_OPTION_PRICE`, `ARPU_TL`, `SUBS_AGE`, `CITY`, `COMPLAINT_FL`, `RAHAT_SEGMENT`, `DIGITAL_SEGMENT_NAME`, `TOT_VOI_MOU`, `TOT_DATA_VOL_GB`, `GENDER`, `DEVICE`, `*_SEARCH`, `TOTAL_FEE_TL`, `BUNDLE_FEE_TL`, `NON_BUNDLE_FEE_TL`. Dosyadaki her satır portföyde sayılır. Portföye giren / çıkan aboneler, kimliğin bir önceki ay dosyasında olup olmamasına göre bulunur; kimlik aylar arasında sabit değilse betik uyarır ve bu bölüm gizlenir.
- **kohort:** her satır bir liste ayı. Kolonlar: `YEARMONTH`, grup adetleri (`AramaL`, `AranmışL`, `UlasilmisL`, `KontrolL`, `RahatFırsatL`) ve `YearMonth_k_<grup>POUT` (k = liste ayından sonraki ay). POUT değerleri kümülatif kabul edilir (`KOHORT_KUMULATIF`); `YearMonth_1` liste ayından bir ay sonrasıdır (`KOHORT_AY_FARKI`).
- **lifecycle:** ay, lifecycle tipi ve adet kolonları. Kolon adları rapordan geldiği gibi olabilir (ör. `LIFECYCLE_TYPE_DESC`, `COUNTLIFECYCLESUBSCRIBERDISTIN`); ay kolonu adından değil içeriğinden (YYYYMM değerleri) bulunur. Yalnız adında "Hibrit" geçen tipler kullanılır: Yeni Tesis ve Port In yeni müşteri kazanımı, `…-to-Hibrit` diğer hat tiplerinden gelen, `Hibrit-to-…` diğer hat tiplerine giden sayılır.
- **kpi:** Excel (tüm sayfalar taranır) veya CSV olduğu gibi okunur: ay başlıkları satırı (`Mar.25`, `Ocak'26`, `202503` veya tarih), solunda grup ve KPI adı kolonları. Değerler hesaplanmaz, hücrelerin sayı biçimiyle gösterilir.

### Notlar

- **Gelir** `TOTAL_FEE_TL` üzerinden hesaplanır (toplam ve abone başı ortalama); ARPU ayrı bir ölçü olarak kalır. Data `TOT_DATA_VOL_GB`, ses `TOT_VOI_MOU`, tenure `SUBSCRIBERS_LAST_TENURE_DAYS` (aya çevrilerek) kolonundan gelir.
- Ana sayfadaki tarife listesinde yalnız adı `BULLET_TARIFE_DESEN` ("Rahat Efsane Fırsat") içeren paketler gösterilir; kırılım filtrelerinde tüm paketler vardır. Fatura dönemi `BC` kolonundan okunur.
- "Önceki ay" bir önceki takvim ayı, "geçen yıl" 12 ay önceki dosyadır; dosyası olmayan dönemde değişim boş görünür. Portföye giren / çıkan hesabı yalnız ardışık ay dosyaları arasında yapılır.
- Her kartın başlığındaki **Büyüt** düğmesi kartı tam ekranda gösterir (Esc ile kapanır).

- Rapora yalnızca adet ve toplamlar gömülür; müşteri satırı, kimlik ve telefon girmez (`PHONE` hiç okunmaz).
- Üç kırılımlı (filtreli) tablolar yalnızca son `TRI_AY` ay için gömülür (varsayılan 3); dosya boyutunu bu belirler.
- Çalıştırınca basılan özetleri (kohort dosyaları, bulunan kolonlar, ay bazında portföy ve hareket adetleri, tanınan KPI satırları) kontrol edin.
