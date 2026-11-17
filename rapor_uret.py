#!/usr/bin/env python3
"""Rahat segment raporu üretici.

Aylık müşteri dosyalarından (YYYYMM.csv) tek dosyalık, internetsiz açılan HTML rapor üretir.

    python rapor_uret.py --veri data --cikti rahat_segment_raporu.html

Beklenen klasör yapısı:
    data/raw_data/    202604.csv …            aylık müşteri dosyaları
    data/kohort/      kohort tablosu           YEARMONTH + AramaL, AranmışL, UlasilmisL, KontrolL, RahatFırsatL + YearMonth_k_<grup>POUT kolonları
    data/lifecycle/   lifecycle tablosu        yearmonth, lifecycle_type, counts (yalnız Hybrid ile ilgili tipler kullanılır)
    data/kpi/         KPI tablosu (Excel)      olduğu gibi rapora eklenir; her dosya bir proje
    data/post_model/  (isteğe bağlı)           müşteri bazlı skor / liste / kontrol tabloları

Notebook içinden:  %run rapor_uret.py --veri data

Rapora yalnızca toplulaştırılmış (adet / toplam) değerler gömülür; müşteri satırı, kimlik veya telefon gömülmez.
Tek dosyadır: rapor şablonu (HTML / CSS / JS) bu dosyanın sonundaki SABLON metnindedir.
Gerekli paketler: numpy, pandas, openpyxl (KPI Excel'i için).
"""
import argparse, base64, codecs, datetime as dt, glob, gzip, itertools, json, os, re, sys
import numpy as np
import pandas as pd

# ═══════════════════════════ AYARLAR ═══════════════════════════
DOSYA_DESENI = r"^(\d{6})\.csv$"            # ay dosyası adı (YYYYMM.csv)
ANAHTAR_KOLONU = None                        # müşteriyi aylar arası izleyen kimlik kolonu; None → adaylardan otomatik seçilir
ANAHTAR_ADAYLARI = ["CAMPAIGN_SUBSCRIBER_ID", "SUBSCRIBER_ID"]
AKTIF_DURUM_IDS = None                       # SUBSCRIBER_STATUS_ID içinde "aktif" sayılan değer(ler); None → son aydaki en sık değer
YAS_KOLONU = "SUBS_AGE"
CINSIYET_ADAYLARI = ["GENDER", "CINSIYET", "SEX", "GENDER_CD", "GENDER_ID", "GENDER_NAME"]
CINSIYET_ESLEME = {"K": "Kadın", "KADIN": "Kadın", "F": "Kadın", "FEMALE": "Kadın", "WOMAN": "Kadın",
                   "E": "Erkek", "ERKEK": "Erkek", "M": "Erkek", "MALE": "Erkek", "MAN": "Erkek"}
CIHAZ_ADAYLARI = ["DEVICE_BRAND", "DEVICE", "DEVICE_NAME", "HANDSET", "HANDSET_BRAND", "PHONE_BRAND", "DEVICE_MAKE", "BRAND", "CIHAZ_MARKA", "MARKA"]
TARIFE_ADI_ADAYLARI = ["OPTION_AND_PACKAGE_NAME", "TARIFF_NAME", "TARIFF_DESC"]
TARIFE_ISIM = {}                             # {TARIFF_ID: "ad"} — tarife adı kolonu yoksa buradan

# ── Model / arama listesi kolonları (aylık dosyalarda). İlk bulunan aday kullanılır; adınız farklıysa listeye ekleyin. ──
SKOR_ADAYLARI = ["DECISION_SCORE", "CHURN_SCORE", "MODEL_SCORE", "SCORE"]           # model skoru (yüksek = yüksek churn riski)
KOHORT_KOLON = {                                                                   # bayrak kolonları (1 / E / EVET / Y / TRUE = evet)
    '_L': ["LIST_FL", "LISTE_FL", "CALL_LIST_FL", "IN_LIST_FL", "TARGET_FL"],      # arama listesinde
    '_A': ["CALLED_FL", "ARANDI_FL", "CALL_FL", "ARAMA_FL"],                       # arandı
    '_U': ["REACHED_FL", "ULASILDI_FL", "CONTACT_FL", "ULASIM_FL"],                # ulaşıldı
    '_K': ["CONTROL_FL", "KONTROL_FL", "CONTROL_GROUP_FL", "KONTROL_GRUP_FL"],     # kontrol grubu
    '_E': ["EFSANE_FL", "RAHAT_EFSANE_FL"],                                        # Rahat Efsane paketi var (kolon yoksa EFSANE_DESEN aranır)
}
EFSANE_DESEN = r"EFSANE"                     # bayrak kolonu yoksa BONUS_NAME ve tarife adı içinde aranır (büyük-küçük harf duyarsız)
DOGRU = {"1", "1.0", "E", "EVET", "Y", "YES", "TRUE", "T", "VAR"}
AY_KOLON_ADAYLARI = ["MONTHLY_CALENDAR_ID", "YEARMONTH", "YEAR_MONTH", "AY", "DONEM", "PERIOD", "MONTH_ID"]   # kohort tablosundaki ay kolonu
SKOR_DILIM = 20                              # skor dilimi sayısı (20 → %5'lik dilimler)
TRI_AY = 3                                   # üç kırılımlı (filtreli) tablolar yalnızca son N ay için gömülür; dosya boyutunu belirler
CHURN_DURUM = (3, 4)                         # bir sonraki ay bu durum kodlarında olan (veya dosyadan çıkan) abone churn sayılır

BAYRAKLAR = ['HBACQ', 'HBPORTIN', 'PRE2HB', 'HB2PRE', 'POST2HB', 'HB2POST', 'PRE2POST', 'POST2PRE']
BAYRAK_RE = re.compile(r'^(%s)_\d{6}_FL$' % '|'.join(BAYRAKLAR))
SAYISAL = (['SUBSCRIBERS_LAST_TENURE_MONTHS', 'SUBS_AGE', 'TARIFF_OPTION_PRICE', 'ARPU_TL', 'COMPLAINT_FL', 'TOT_VOI_MOU', 'TOT_DATA_VOL_GB', 'POUT_COUNT_LM',
            'TOTAL_FEE_TL', 'BUNDLE_FEE_TL', 'NON_BUNDLE_FEE_TL']
           + [f'{b}_FL' for b in BAYRAKLAR])
METIN = ['CITY', 'BONUS_NAME', 'RAHAT_SEGMENT', 'DIGITAL_SEGMENT_NAME']
OLCU = [('arpu', 'ARPU_TL'), ('gb', 'TOT_DATA_VOL_GB'), ('mou', 'TOT_VOI_MOU'), ('ten', 'SUBSCRIBERS_LAST_TENURE_MONTHS'), ('age', YAS_KOLONU),
        ('fee', 'TOTAL_FEE_TL'), ('price', 'TARIFF_OPTION_PRICE'), ('comp', 'COMPLAINT_FL'), ('bfee', 'BUNDLE_FEE_TL'), ('nbfee', 'NON_BUNDLE_FEE_TL')]
#        gelir = TOTAL_FEE_TL · data = TOT_DATA_VOL_GB · ses = TOT_VOI_MOU · tenure = SUBSCRIBERS_LAST_TENURE_DAYS (aya çevrilir)
TENURE_GUN_ADAYLARI = ['SUBSCRIBERS_LAST_TENURE_DAYS', 'SUBSCRIBER_TENURE_DAYS']   # ay kolonu yoksa gün / 30,44 ile aya çevrilir
K2 = 6                                       # çoklu kırılımlarda tutulan ölçüler: OLCU'nun ilk K2 tanesi
BULLET_TARIFE_DESEN = 'Rahat Efsane Fırsat'  # ana sayfadaki tarife listesinde yalnız adı bunu içeren paketler gösterilir (filtrelerde tüm paketler)
BC_KOLONU = 'BC'                             # fatura dönemi (billing cycle)
ILGI = {'BANK': 'Bankacılık', 'SOCIALMEDIA': 'Sosyal medya', 'NEWS': 'Haber', 'SPORT': 'Spor', 'GAME': 'Oyun', 'MUSIC': 'Müzik', 'STREAMING': 'Video yayın',
        'ECOMMERCE': 'E-ticaret', 'HEALTH': 'Sağlık', 'RETAIL': 'Perakende', 'CLOTHES': 'Giyim', 'ACTUERIA': 'Sigorta', 'MAKEUP': 'Kozmetik', 'LUXURY': 'Lüks',
        'FOODORDER': 'Yemek siparişi', 'ART': 'Sanat', 'CHILDCARE': 'Çocuk bakımı', 'GAMBLING': 'Şans oyunları', 'CARRENT': 'Araç kiralama',
        'CARBUY': 'Araç alım-satım', 'TECH': 'Teknoloji', 'SECONDHAND': 'İkinci el', 'ESTATE': 'Emlak', 'TRAFFIC': 'Trafik / harita', 'TRAVEL': 'Seyahat'}
EV = ['in_acq', 'in_portin', 'in_pre', 'in_post', 'in_oth', 'out_pre', 'out_post', 'out_rival', 'out_close', 'out_ghost']
EV_LBL = ['Yeni hat', 'Numara taşıma ile gelen', "Faturasızdan gelen", "Faturalıdan gelen", 'Portföye giren',
          "Faturasıza geçen", "Faturalıya geçen", 'Rakibe giden (port-out)', 'Hat kapatan / pasif', 'Portföyden çıkan']


# ═══════════════════════════ CSV OKUMA ═══════════════════════════
def kolon_norm(c):
    n = re.sub(r'\s+', '_', str(c).replace('﻿', '').strip()).upper()
    return BAYRAK_RE.sub(lambda m: m.group(1) + '_FL', n) if BAYRAK_RE.match(n) else n


def dosyalari_bul(klasor):
    if not os.path.isdir(klasor):
        sys.exit(f"Veri klasörü bulunamadı: {os.path.abspath(klasor)}")
    pat = re.compile(DOSYA_DESENI)
    out = sorted((int(pat.match(f).group(1)), os.path.join(klasor, f)) for f in os.listdir(klasor) if pat.match(f))
    if not out:
        sys.exit(f"'{klasor}' içinde YYYYMM.csv biçiminde dosya yok.")
    return out


def _kodlama(yol):
    raw = open(yol, 'rb').read(2_000_000)
    for enc in ('utf-8-sig', 'utf-8', 'cp1254'):
        try:
            codecs.getincrementaldecoder(enc)().decode(raw, final=False); return enc
        except UnicodeDecodeError:
            continue
    return 'latin-1'


def _ayrac(yol, enc):
    with open(yol, 'r', encoding=enc, errors='replace') as f:
        satirlar = [f.readline() for _ in range(5)]
    sayim = {s: sum(l.count(s) for l in satirlar) for s in [',', ';', '\t', '|']}
    return max(sayim, key=sayim.get)


def _ondalik(yol, enc, sep):
    ornek = pd.read_csv(yol, sep=sep, encoding=enc, nrows=3000, dtype=str, usecols=lambda c: kolon_norm(c) in ('ARPU_TL', 'TOT_DATA_VOL_GB', 'TARIFF_OPTION_PRICE'))
    dec, th = '.', None
    for c in ornek.columns:
        s = ornek[c].dropna().astype(str)
        if len(s) and sep != ',' and s.str.match(r'^\s*-?\d{1,3}(\.\d{3})*,\d+\s*$').mean() > .2:
            dec = ','; th = '.' if s.str.contains(r'^\s*-?\d{1,3}\.\d{3},').mean() > .01 else None
    return dec, th


def anahtar_sec(dosyalar):
    """Ay içinde benzersiz (≥%98) ve önceki ayla en çok örtüşen aday kolonu seçer."""
    if ANAHTAR_KOLONU:
        return kolon_norm(ANAHTAR_KOLONU)
    oku = lambda y: (lambda enc: (lambda sep: pd.read_csv(y, sep=sep, encoding=enc, low_memory=False,
                     usecols=lambda c: kolon_norm(c) in ANAHTAR_ADAYLARI).rename(columns=kolon_norm))(_ayrac(y, enc)))(_kodlama(y))
    son = oku(dosyalar[-1][1]); onc = oku(dosyalar[-2][1]) if len(dosyalar) > 1 else None
    aday = []
    for a in ANAHTAR_ADAYLARI:
        if a in son.columns:
            u = son[a].nunique() / max(len(son), 1)
            ort = float(son[a].drop_duplicates().isin(set(onc[a].dropna())).mean()) if onc is not None and a in onc.columns else 0
            aday.append((u >= .98, ort, u, a))
    if not aday:
        sys.exit(f"Müşteri anahtarı bulunamadı; adaylar: {ANAHTAR_ADAYLARI}. ANAHTAR_KOLONU ayarını yazın.")
    return sorted(aday, reverse=True)[0][3]


KOHORT_VERI = {}                             # {ay: (anahtar kolonu, tablo)} — kohort klasöründen okunur
_kstr = lambda s: s.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)


def _kohort_std(df):
    """Skor ve bayrak kolonlarını standart adlara (_SKOR, _L, _A, _U, _K, _E) çevirir; bulunanları {standart ad: kaynak kolon} olarak döndürür."""
    bulunan = {}
    sk = next((c for c in SKOR_ADAYLARI if c in df.columns), None)
    if sk:
        df['_SKOR'] = pd.to_numeric(df[sk], errors='coerce'); bulunan['_SKOR'] = sk
    for hedef, adaylar in KOHORT_KOLON.items():
        k = next((c for c in adaylar if c in df.columns), None)
        if k:
            s = df[k]; bulunan[hedef] = k
            df[hedef] = s.astype(str).str.strip().str.upper().isin(DOGRU).to_numpy() if s.dtype == object else (pd.to_numeric(s, errors='coerce').fillna(0) > 0).to_numpy()
    return bulunan


def kohort_oku(klasor):
    """Kohort tablolarını okur: klasördeki her CSV'de müşteri anahtarı + (varsa) skor / liste / arandı / ulaşıldı / kontrol / Efsane kolonları.
    Ay, tablodaki ay kolonundan (AY_KOLON_ADAYLARI), yoksa dosya adındaki YYYYMM'den alınır. Liste bayrağı yoksa dosyada bulunan (kontrol grubu olmayan) her abone listede sayılır."""
    if not klasor or not os.path.isdir(klasor):
        return
    for yol in sorted(glob.glob(os.path.join(klasor, '*.csv'))):
        ad = os.path.basename(yol); enc = _kodlama(yol); sep = _ayrac(yol, enc)
        d = pd.read_csv(yol, sep=sep, encoding=enc, low_memory=False); d.columns = [kolon_norm(c) for c in d.columns]; d = d.loc[:, ~d.columns.duplicated()]
        kcol = next((c for c in ([kolon_norm(ANAHTAR_KOLONU)] if ANAHTAR_KOLONU else []) + ANAHTAR_ADAYLARI if c in d.columns), None)
        if not kcol:
            print(f"⚠ kohort/{ad}: müşteri anahtarı kolonu yok ({ANAHTAR_ADAYLARI}); dosya atlandı."); continue
        bul = _kohort_std(d)
        if '_L' not in d.columns:
            d['_L'] = ~d['_K'] if '_K' in d.columns else True; bul['_L'] = 'dosyada bulunma'
        ac = next((c for c in AY_KOLON_ADAYLARI if c in d.columns), None); m = re.search(r'(20\d{2})[-_]?(0[1-9]|1[0-2])', ad)
        if ac: d['_AY'] = pd.to_numeric(d[ac].astype(str).str.replace(r'\D', '', regex=True).str[:6], errors='coerce')
        elif m: d['_AY'] = int(m.group(1) + m.group(2))
        else:
            print(f"⚠ kohort/{ad}: ay bilgisi yok (ay kolonu veya dosya adında YYYYMM bekleniyor); dosya atlandı."); continue
        std = [c for c in ('_SKOR', '_L', '_A', '_U', '_K', '_E') if c in d.columns]; d[kcol] = _kstr(d[kcol]); aylar = []
        for ay, g in d.dropna(subset=['_AY']).groupby('_AY'):
            g = g.drop_duplicates(subset=[kcol]).set_index(kcol)[std]; ay = int(ay); aylar.append(ay)
            if ay in KOHORT_VERI and KOHORT_VERI[ay][0] == kcol:
                g = pd.concat([KOHORT_VERI[ay][1], g]); g = g[~g.index.duplicated(keep='last')]
            KOHORT_VERI[ay] = (kcol, g)
        print(f"Kohort dosyası {ad}: anahtar={kcol} · aylar={min(aylar) if aylar else '-'}–{max(aylar) if aylar else '-'} · kolonlar={bul}")


def csv_oku(yol, mid, anahtar):
    enc = _kodlama(yol); sep = _ayrac(yol, enc); dec, th = _ondalik(yol, enc, sep)
    baslik = list(pd.read_csv(yol, sep=sep, encoding=enc, nrows=0).columns); norm = {c: kolon_norm(c) for c in baslik}
    ister = (set(SAYISAL) | set(METIN) | set(CINSIYET_ADAYLARI) | set(CIHAZ_ADAYLARI) | set(TARIFE_ADI_ADAYLARI) | {anahtar, 'TARIFF_ID', 'SUBSCRIBER_STATUS_ID', YAS_KOLONU} | set(ANAHTAR_ADAYLARI) | set(TENURE_GUN_ADAYLARI) | {BC_KOLONU}
             | set(SKOR_ADAYLARI) | {c for v in KOHORT_KOLON.values() for c in v})
    kullan = [c for c in baslik if norm[c] in ister or norm[c].endswith('_SEARCH')]          # PHONE hiç okunmaz
    df = pd.read_csv(yol, sep=sep, encoding=enc, decimal=dec, thousands=th, usecols=kullan, low_memory=False)
    df.columns = [norm[c] for c in df.columns]; df = df.loc[:, ~df.columns.duplicated()]
    if anahtar not in df.columns:
        sys.exit(f"{os.path.basename(yol)}: anahtar kolon '{anahtar}' yok.")
    ham = len(df); df = df.drop_duplicates(subset=[anahtar], keep='first').reset_index(drop=True)
    gk_ = next((c for c in TENURE_GUN_ADAYLARI if c in df.columns), None)       # tenure: gün kolonu varsa her zaman o kullanılır (aya çevrilerek)
    if gk_: df['SUBSCRIBERS_LAST_TENURE_MONTHS'] = pd.to_numeric(df[gk_], errors='coerce') / 30.4375
    df['_BC'] = _kstr(df[BC_KOLONU]).where(df[BC_KOLONU].notna(), None) if BC_KOLONU in df.columns else None
    if 'TARIFF_ID' not in df.columns:                                          # tarife kimliği yoksa paket adı tarife olarak kullanılır
        tk_ = next((c for c in TARIFE_ADI_ADAYLARI if c in df.columns), None)
        if tk_: df['TARIFF_ID'] = df[tk_].map(lambda v: v.strip() if isinstance(v, str) and v.strip() else None)
    eksik = [c for c in SAYISAL + METIN + ['TARIFF_ID', 'SUBSCRIBER_STATUS_ID'] if c not in df.columns]
    for c in eksik:
        df[c] = (0 if c.endswith('_FL') or c == 'POUT_COUNT_LM' else np.nan) if c in SAYISAL else (1 if c == 'SUBSCRIBER_STATUS_ID' else None)
    for c in SAYISAL:
        df[c] = pd.to_numeric(df[c], errors='coerce')
    for c in METIN:
        df[c] = df[c].map(lambda v: v.strip() if isinstance(v, str) and v.strip() else None)
    gk = next((c for c in CINSIYET_ADAYLARI if c in df.columns), None)
    if gk:
        v = df[gk].astype(str).str.strip().str.upper(); df['_GEN'] = v.map(CINSIYET_ESLEME)
    ck = next((c for c in CIHAZ_ADAYLARI if c in df.columns), None)
    if ck:
        v = df[ck].astype(str).str.strip(); df['_DEV'] = v.where(~v.str.upper().isin(['NAN', '', 'NONE']), None)
    for c in ('_GEN', '_DEV'):                                                 # bu ayın dosyasında cinsiyet / cihaz kolonu yoksa boş bırak
        if c not in df.columns: df[c] = None
    tk = next((c for c in TARIFE_ADI_ADAYLARI if c in df.columns), None)
    if tk:
        df['_TAD'] = df[tk].map(lambda v: v.strip() if isinstance(v, str) and v.strip() else None)
    kohort = _kohort_std(df)
    kv = KOHORT_VERI.get(int(mid))                                             # <veri>/kohort klasöründen gelen liste / arama / kontrol bilgileri
    if kv is not None and kv[0] in df.columns:
        j = kv[1].reindex(_kstr(df[kv[0]]).to_numpy())
        for c in kv[1].columns:
            v = j[c].to_numpy()
            df[c] = pd.to_numeric(v, errors='coerce') if c == '_SKOR' else pd.Series(v).fillna(False).astype(bool).to_numpy()
            kohort[c] = 'kohort klasörü'
    if '_E' not in df.columns and EFSANE_DESEN:
        e = df['BONUS_NAME'].astype(str).str.contains(EFSANE_DESEN, case=False, regex=True, na=False)
        if '_TAD' in df.columns: e = e | df['_TAD'].astype(str).str.contains(EFSANE_DESEN, case=False, regex=True, na=False)
        if e.any(): df['_E'] = e.to_numpy(); kohort['_E'] = f"desen '{EFSANE_DESEN}'"
    bilgi = dict(ay=mid, ham_satir=ham, kohort=kohort, tekrar=ham - len(df), eksik=eksik, cinsiyet=gk, cihaz=ck, tarife_adi=tk,
                 arama=sum(c.endswith('_SEARCH') for c in df.columns))
    return df, bilgi


# ═══════════════════════════ HAREKET KURALLARI ═══════════════════════════
def aktif_mi(df):
    if AKTIF_DURUM_IDS is None:
        return np.ones(len(df), bool)
    s = df['SUBSCRIBER_STATUS_ID']
    return s.astype(str).str.replace(r'\.0$', '', regex=True).isin([str(x) for x in AKTIF_DURUM_IDS]).to_numpy()


def durum_kodu(df):
    """0 portföyde · 1 faturasıza geçti (HB2PRE) · 2 faturalıya geçti (HB2POST) · 3 rakibe gitti (pasif + port-out) · 4 hat kapatma / pasif"""
    aktif = aktif_mi(df); d = np.zeros(len(df), np.int8)
    d[~aktif] = 4; d[(~aktif) & (df['POUT_COUNT_LM'].fillna(0).to_numpy() > 0)] = 3
    d[df['HB2POST_FL'].fillna(0).to_numpy() == 1] = 2; d[df['HB2PRE_FL'].fillna(0).to_numpy() == 1] = 1
    return d


def giris_kodu(df):
    """Portföye giriş türü: 0 yeni hat · 1 numara taşıma · 2 PRE2HB · 3 POST2HB · 4 sınıflanamayan"""
    g = np.full(len(df), 4, np.int8)
    g[df['POST2HB_FL'].fillna(0).to_numpy() == 1] = 3; g[df['PRE2HB_FL'].fillna(0).to_numpy() == 1] = 2
    g[df['HBACQ_FL'].fillna(0).to_numpy() == 1] = 0; g[df['HBPORTIN_FL'].fillna(0).to_numpy() == 1] = 1
    return g


# ═══════════════════════════ KIRILIMLAR ═══════════════════════════
def _tr(v):
    return f"{v:,.0f}".replace(',', '.') if float(v).is_integer() else f"{v:,.1f}".replace(',', '§').replace('.', ',').replace('§', '.')


def kat_dim(key, name, s, col, n, labfn=str, bos='Belirsiz'):
    """En sık n değer + (gerekirse) 'Diğer'. Boş değerler ayrı kategori olur."""
    v = s[col].dropna(); top = list(v.value_counts().index[:n])
    labels = [labfn(x) for x in top]; mp = {x: i for i, x in enumerate(top)}
    diger = v.nunique() > len(top); bosvar = s[col].isna().any()
    if diger: labels.append('Diğer')
    if bosvar: labels.append(bos)
    oi = len(top); bi = oi + (1 if diger else 0)
    def fn(d, col=col, mp=mp, oi=oi, bi=bi):
        c = d[col].map(mp)
        return c.fillna(oi).where(d[col].notna(), bi).clip(upper=len(labels) - 1).to_numpy().astype(np.int64)
    return dict(key=key, name=name, labels=labels, fn=fn)


def bant_dim(key, name, s, col, birim, hedef=8, kenar=None):
    """Sayısal değişkeni yuvarlak sınırlı bantlara böler (histogram)."""
    x = s[col].to_numpy(float); x = x[np.isfinite(x)]
    if kenar is None:
        lo, hi = np.quantile(x, [.02, .97]) if len(x) else (0, 1)
        ham = max((hi - lo) / hedef, 1e-9); p = 10 ** np.floor(np.log10(ham)); adim = min([1, 2, 2.5, 5, 10], key=lambda m: abs(m * p - ham)) * p
        kenar = np.arange(np.floor(lo / adim) * adim + adim, hi, adim)
        if len(kenar) == 0: kenar = np.array([np.floor(lo) + 1])
    kenar = np.asarray(kenar, float)
    labels = [f'< {_tr(kenar[0])}'] + [f'{_tr(kenar[i])}–{_tr(kenar[i + 1])}' for i in range(len(kenar) - 1)] + [f'{_tr(kenar[-1])}+']
    labels = [l + birim for l in labels]
    return dict(key=key, name=name, labels=labels, fn=lambda d, col=col, e=kenar: np.digitize(d[col].fillna(e[0] - 1).to_numpy(float), e).astype(np.int64))


def dims_tanimla(df):
    s = df[durum_kodu(df) == 0]
    if not len(s): s = df
    D = [kat_dim('rah', 'Rahat segment', s, 'RAHAT_SEGMENT', 8),
         kat_dim('dig', 'Dijital segment', s, 'DIGITAL_SEGMENT_NAME', 10)]
    tmap = dict(TARIFE_ISIM)
    if not tmap and '_TAD' in s and s['_TAD'].notna().any():
        tmap = s.dropna(subset=['_TAD']).groupby('TARIFF_ID')['_TAD'].agg(lambda x: x.value_counts().index[0]).to_dict()
    tl = lambda v: tmap.get(v, f'Tarife {v:g}' if isinstance(v, (int, float, np.integer, np.floating)) else f'Tarife {v}')
    if s['TARIFF_ID'].notna().any(): D.append(kat_dim('tar', 'Tarife', s, 'TARIFF_ID', 10, tl))
    if '_GEN' in s and s['_GEN'].notna().any(): D.append(kat_dim('gen', 'Cinsiyet', s, '_GEN', 2))
    if '_DEV' in s and s['_DEV'].notna().any(): D.append(kat_dim('dev', 'Cihaz markası', s, '_DEV', 8))
    if s['CITY'].notna().any(): D.append(kat_dim('city', 'Şehir', s, 'CITY', 10))
    if s[YAS_KOLONU].notna().any(): D.append(bant_dim('age', 'Yaş', s, YAS_KOLONU, '', kenar=[25, 35, 45, 55, 65]))
    if s['SUBSCRIBERS_LAST_TENURE_MONTHS'].notna().any(): D.append(bant_dim('ten', 'Tenure', s, 'SUBSCRIBERS_LAST_TENURE_MONTHS', ' ay', kenar=[7, 13, 25, 37, 61]))
    if '_BC' in s and s['_BC'].notna().any(): D.append(kat_dim('bc', 'Fatura dönemi (BC)', s, '_BC', 14, lambda v: f'BC {v}'))
    if s['ARPU_TL'].notna().any(): D.append(bant_dim('ab', 'ARPU bandı', s, 'ARPU_TL', ' TL'))
    if s['TOTAL_FEE_TL'].notna().any(): D.append(bant_dim('fb', 'Gelir bandı', s, 'TOTAL_FEE_TL', ' TL'))
    if s['TOT_DATA_VOL_GB'].notna().any(): D.append(bant_dim('gbb', 'Data kullanım bandı', s, 'TOT_DATA_VOL_GB', ' GB'))
    if s['TOT_VOI_MOU'].notna().any(): D.append(bant_dim('mb', 'Ses kullanım bandı', s, 'TOT_VOI_MOU', ' dk'))
    if s['BONUS_NAME'].notna().any(): D.append(kat_dim('bon', 'Ek paket', s, 'BONUS_NAME', 6, bos='Ek paket yok'))
    return [d for d in D if len(d['labels']) > 0]


# ═══════════════════════════ TOPLULAŞTIRMA ═══════════════════════════
def _liste(a):
    return [int(x) if float(x).is_integer() else x for x in np.round(np.asarray(a, float), 1).ravel().tolist()]


def _anahtar(C, cols, sizes):
    key = np.zeros(len(C), np.int64)
    for c, sz in zip(cols, sizes):
        key = key * sz + C[:, c]
    return key


def topla(C, X, cols, sizes, k, gecerli=False):
    """Kırılım hücresi başına [adet, ölçü toplamları (, geçerli değer adetleri)] — satır öncelikli düz liste."""
    size = int(np.prod(sizes)) if sizes else 1; key = _anahtar(C, cols, sizes)
    out = [np.bincount(key, minlength=size)]
    out += [np.bincount(key, weights=np.nan_to_num(X[:, j]), minlength=size) for j in range(k)]
    if gecerli:
        out += [np.bincount(key, weights=np.isfinite(X[:, j]).astype(float), minlength=size) for j in range(k)]
    return _liste(np.stack(out, 1))


GRUPLAR = [('p', 'Portföy (tümü)', None, []), ('l', 'Listedekiler', 'p', ['_L']), ('a', 'Arananlar', 'l', ['_A']), ('u', 'Ulaşılanlar', 'a', ['_A', '_U']),
           ('e', 'Rahat Efsane paketi olanlar (listede)', 'l', ['_E']), ('k', 'Kontrol grubu', 'p', ['_K']), ('x', 'Kontrol grubu dışı', 'p', ['_K']), ('d', 'Liste ve kontrol dışı', 'p', ['_L'])]


def grup_maskeleri(df, gruplar):
    n = len(df); f = lambda c: df[c].to_numpy(bool) if c in df.columns else np.zeros(n, bool)
    L, A, U, K, E = f('_L'), f('_A'), f('_U'), f('_K'), f('_E')
    if '_L' not in df.columns: L = A | U                                      # liste kolonu yoksa aranan / ulaşılanlar listedir
    m = dict(p=np.ones(n, bool), l=L, a=L & A, u=L & A & U, e=L & E, k=K, x=~K, d=~L & ~K)
    return np.stack([m[g[0]] for g in gruplar], 1)


def skor_dilimi(df, s0):
    """Portföydeki skorlu aboneleri skora göre (yüksekten düşüğe) eşit büyüklükte dilimlere ayırır; 0 = en yüksek skorlu dilim, -1 = skor yok."""
    b = np.full(len(df), -1, np.int64)
    if '_SKOR' in df.columns:
        ok = s0 & df['_SKOR'].notna().to_numpy()
        if ok.any():
            sira = (-df['_SKOR'][ok]).rank(method='first').to_numpy() - 1
            b[ok] = np.minimum((sira * SKOR_DILIM // ok.sum()).astype(np.int64), SKOR_DILIM - 1)
    return b


def kohort_topla(C, G, sb, churn, sz, dims):
    """Her grup için [adet, churn adedi]: toplam, tek kırılım ve skor dilimi bazında."""
    ch = np.zeros(len(C)) if churn is None else churn.astype(float)
    def ag(key, size, gecerli=None):
        out = []
        for g in range(G.shape[1]):
            m = G[:, g] if gecerli is None else G[:, g] & gecerli
            out += [np.bincount(key[m], minlength=size), np.bincount(key[m], weights=ch[m], minlength=size)]
        return _liste(np.stack(out, 1))
    return dict(known=churn is not None, tot=ag(np.zeros(len(C), np.int64), 1), one={d['key']: ag(C[:, i], sz[i]) for i, d in enumerate(dims)},
                skor=ag(np.clip(sb, 0, None), SKOR_DILIM, sb >= 0) if (sb >= 0).any() else [])


def uret(dosyalar):
    anahtar = anahtar_sec(dosyalar); print(f"Müşteri anahtarı: {anahtar}")
    global AKTIF_DURUM_IDS
    son, _ = csv_oku(dosyalar[-1][1], dosyalar[-1][0], anahtar)
    if AKTIF_DURUM_IDS is None and son['SUBSCRIBER_STATUS_ID'].nunique(dropna=True) > 1:
        vc = son['SUBSCRIBER_STATUS_ID'].value_counts()
        AKTIF_DURUM_IDS = [str(vc.index[0]).replace('.0', '')]
        print(f"Aktif durum kimliği otomatik seçildi: {AKTIF_DURUM_IDS} (dağılım: {dict(vc.head(6))})")
    dims = dims_tanimla(son); ND = len(dims); sz = [len(d['labels']) for d in dims]; KO = len(OLCU); CI = [k for k, _ in OLCU].index('comp')
    srch = []                                                                  # arama kolonları: tüm ay dosyalarının başlıklarındaki *_SEARCH kolonlarının birleşimi
    for _m, _y in dosyalar:
        _enc = _kodlama(_y)
        for c in pd.read_csv(_y, sep=_ayrac(_y, _enc), encoding=_enc, nrows=0).columns:
            c = kolon_norm(c)
            if c.endswith('_SEARCH') and c not in srch: srch.append(c)
    srch.sort(); NI = len(srch); aramasiz = []; paket = []
    kk = dict(_['kohort'])
    for _k, _d in KOHORT_VERI.values():
        for c in _d.columns: kk.setdefault(c, 'kohort klasörü')
    gruplar = [g for g in GRUPLAR if all(c in kk for c in g[3])] if any(c in kk for c in ('_L', '_A', '_K')) or '_SKOR' in kk else []
    if kk: print("Model / liste kolonları:", kk)
    R = dict(tot=[], one=[], pair=[], tri=[], srch_tot=[], srch_one=[], ev_tot=[], ev_one=[], ev_srch=[], coh=[], fc=[], ev_ok=[]); onceki = None; onceki_ay = 0; ozet = []; ortusme = []
    izlenen = [(g[0], i) for i, g in enumerate(gruplar) if g[0] in ('l', 'k')]; kohortlar = []
    for t, (mid, yol) in enumerate(dosyalar):
        df, bilgi = (son, None) if t == len(dosyalar) - 1 else csv_oku(yol, mid, anahtar)
        C = np.stack([d['fn'](df) for d in dims], 1); X = np.stack([df[c].to_numpy(float) for _, c in OLCU], 1)
        X[:, CI] = (np.nan_to_num(X[:, CI]) > 0).astype(float)
        avar = NI > 0 and any(c in df.columns for c in srch)                   # bu ayın dosyasında arama kolonu var mı
        if NI and not avar: aramasiz.append(mid)
        S = df.reindex(columns=srch).apply(pd.to_numeric, errors='coerce').fillna(0).to_numpy(np.float32) if NI else np.zeros((len(df), 0), np.float32)
        idv = df[anahtar].to_numpy(); st = durum_kodu(df); gk = giris_kodu(df); s0 = st == 0
        Cs, Xs, Ss = C[s0], X[s0], S[s0]
        if not s0.any():
            sys.exit(f"{mid}: portföyde hiç satır yok. SUBSCRIBER_STATUS_ID dağılımı: {dict(df['SUBSCRIBER_STATUS_ID'].value_counts().head(8))}")
        R['tot'].append(topla(Cs, Xs, [], [], KO, True))
        R['one'].append({d['key']: topla(Cs, Xs, [i], [sz[i]], KO, True) for i, d in enumerate(dims)})
        R['pair'].append({f"{dims[a]['key']}|{dims[b]['key']}": topla(Cs, Xs, [a, b], [sz[a], sz[b]], K2) for a, b in itertools.combinations(range(ND), 2)})
        _p = pd.DataFrame(dict(ad=df['TARIFF_ID'].to_numpy()[s0], arpu=Xs[:, 0])).dropna(subset=['ad']); _p = _p[_p['ad'].map(lambda v: _sade(BULLET_TARIFE_DESEN) in _sade(v))]
        paket.append(_p.groupby('ad')['arpu'].agg(['size', 'sum']))
        R['srch_tot'].append(_liste(np.concatenate([(Ss > 0).sum(0), Ss.sum(0)])) if avar else [])
        R['tri'].append({'|'.join(dims[i]['key'] for i in k): topla(Cs, Xs, list(k), [sz[i] for i in k], K2) for k in itertools.combinations(range(ND), 3)}
                        if t >= len(dosyalar) - TRI_AY else {})
        R['srch_one'].append({d['key']: (_liste(np.stack([np.concatenate([(Ss[Cs[:, i] == c] > 0).sum(0), Ss[Cs[:, i] == c].sum(0)]) for c in range(sz[i])])) if avar else [])
                              for i, d in enumerate(dims)})
        # ── kohort: önceki ayın gruplarının bu aydaki churn durumu ──
        ardisik = onceki is not None and (mid // 100 * 12 + mid % 100) - (onceki_ay // 100 * 12 + onceki_ay % 100) == 1   # önceki dosya bir önceki takvim ayı mı
        R['ev_ok'].append(bool(ardisik))
        if gruplar and ardisik:
            pos = pd.Index(idv).get_indexer(onceki['id']); churn = pos < 0
            churn[pos >= 0] = np.isin(st[pos[pos >= 0]], CHURN_DURUM)
            R['coh'].append(kohort_topla(onceki['C'], onceki['G'], onceki['sb'], churn, sz, dims))
        Gs = grup_maskeleri(df, gruplar)[s0] if gruplar else None; sbs = skor_dilimi(df, s0)[s0] if gruplar else None
        if gruplar and t == len(dosyalar) - 1:
            R['coh'].append(kohort_topla(Cs, Gs, sbs, None, sz, dims))                # son ay: churn sonucu henüz yok
        # ── ileriye dönük kohort takibi: listeye / kontrol grubuna girenlerin sonraki aylardaki aranma, ulaşılma, paket ve churn durumu ──
        if izlenen:
            bay = lambda c: df[c].to_numpy(bool) if c in df.columns else np.zeros(len(df), bool)
            fA, fU, fE = bay('_A'), bay('_U'), bay('_E'); ix = pd.Index(idv); Gt = grup_maskeleri(df, gruplar)
            def kaydet(kh, a_m, u_m, ch_m):                                     # [kohort adedi, o ay aranan, küm. aranan, o ay ulaşılan, küm. ulaşılan, küm. Efsane, o ay churn, küm. churn]
                V = np.stack([np.ones(len(kh['id'])), a_m, kh['a'], u_m, kh['u'], kh['e'], ch_m, kh['churn']], 1).astype(float)
                R['fc'].append(dict(g=kh['g'], c=kh['c'], t=t, tot=_liste(V.sum(0))))
            for kh in kohortlar:
                pos = ix.get_indexer(kh['id']); var = pos >= 0; p0 = np.where(var, pos, 0)
                ch_m = ~kh['churn'] & (~var | np.isin(st[p0], CHURN_DURUM)); kh['churn'] |= ch_m
                a_m = var & fA[p0]; u_m = var & fU[p0]; kh['a'] |= a_m; kh['u'] |= u_m; kh['e'] |= var & fE[p0]
                kaydet(kh, a_m, u_m, ch_m)
            for gk, gi in izlenen:
                m = s0 & Gt[:, gi]
                if m.any():
                    kh = dict(g=gk, c=t, id=idv[m], churn=np.zeros(int(m.sum()), bool), a=fA[m].copy(), u=fU[m].copy(), e=fE[m].copy())
                    kohortlar.append(kh); kaydet(kh, fA[m], fU[m], np.zeros(int(m.sum()), bool))
        # ── hareketler: gelenler (bu ay portföyde, önceki ay değil), gidenler (bu ayın çıkış durumları + dosyadan çıkanlar) ──
        et, eo, es = {}, {}, {}
        if ardisik:
            ortusme.append(float(np.isin(idv[s0], onceki['id']).mean()))
            yeni = s0 & ~np.isin(idv, onceki['id'])
            mask = {EV[k]: (yeni & (gk == k), C, X, S) for k in range(5)}
            mask.update({EV[5 + k]: (st == k + 1, C, X, S) for k in range(4)})
            mask['out_ghost'] = (~np.isin(onceki['id'], idv), onceki['C'], onceki['X'], onceki['S'])
            for e, (m, c_, x_, s_) in mask.items():
                et[e] = topla(c_[m], x_[m], [], [], K2)
                eo[e] = {d['key']: topla(c_[m], x_[m], [i], [sz[i]], K2) for i, d in enumerate(dims)}
                es[e] = _liste((s_[m] > 0).sum(0)) if NI else []
        R['ev_tot'].append(et); R['ev_one'].append(eo); R['ev_srch'].append(es)
        onceki = dict(id=idv[s0], C=Cs, X=Xs, S=Ss, G=Gs, sb=sbs); onceki_ay = mid
        bos = {k: round(100 * float(np.isnan(X[s0][:, j]).mean()), 2) for j, (k, _) in enumerate(OLCU) if j != CI and np.isnan(X[s0][:, j]).any()}
        ozet.append(dict(ay=mid, satir=len(df), portfoy=int(s0.sum()), **{e: (et[e][0] if e in et else 0) for e in EV}, bos_oran=bos or ''))
    print(pd.DataFrame(ozet).to_string(index=False))
    if aramasiz: print(f"Not: {aramasiz} dosyalarında *_SEARCH kolonları yok; bu aylarda arama ilgi alanı analizleri boş bırakıldı.")
    _ad = list(paket[-1].sort_values('size', ascending=False).index) + sorted({a for p_ in paket for a in p_.index} - set(paket[-1].index))
    R['pk'] = dict(desen=BULLET_TARIFE_DESEN, labels=[str(a) for a in _ad], n=[[int(p_['size'].get(a, 0)) for a in _ad] for p_ in paket],
                   arpu=[[round(float(p_['sum'].get(a, 0)), 1) for a in _ad] for p_ in paket])
    R['hareket_ok'] = not ortusme or float(np.median(ortusme)) >= .5
    if not R['hareket_ok']:
        print(f"⚠ '{anahtar}' aylar arasında örtüşmüyor (ortanca örtüşme %{100 * np.median(ortusme):.0f}): bu kimlik her ay yeniden üretiliyor olabilir. "
              "Portföye giren / çıkan analizi raporda gizlendi; aylar arasında sabit kalan bir abone kimliği kolonu ekleyip ANAHTAR_ADAYLARI'na yazın.")
    if any(o['bos_oran'] for o in ozet):
        print("⚠ Boş ölçü değerleri var (bos_oran, %). Tek kırılımlı ortalamalar boşlar hariç hesaplanır; çoklu kırılım tablolarında boşlar 0 sayılır.")
    return dict(dims=[dict(key=d['key'], name=d['name'], labels=d['labels']) for d in dims], srch_keys=srch, coh_groups=[dict(key=g[0], name=g[1], base=g[2]) for g in gruplar],
                coh_skor='_SKOR' in kk, skor_dilim=SKOR_DILIM,
                srch_labels=[ILGI.get(c[:-7], c[:-7].title()) for c in srch], **R)


# ═══════════════════════════ MODEL KPI TABLOLARI ═══════════════════════════
_TRMAP = str.maketrans('çğıöşüâî', 'cgiosuai')
_AYLAR = {'jan': 1, 'oca': 1, 'feb': 2, 'sub': 2, 'mar': 3, 'apr': 4, 'nis': 4, 'may': 5, 'jun': 6, 'haz': 6, 'jul': 7, 'tem': 7, 'aug': 8, 'agu': 8,
          'sep': 9, 'eyl': 9, 'oct': 10, 'eki': 10, 'nov': 11, 'kas': 11, 'dec': 12, 'ara': 12}


def _sade(s):
    return re.sub(r'[^a-z0-9%]', '', str(s).replace('İ', 'i').replace('I', 'i').lower().translate(_TRMAP))


def _kpi_anahtar(etiket):
    """KPI satır başlığını alan adına eşler (sıra önemli: özelden genele)."""
    s = _sade(etiket); oran = any(x in s for x in ('rate', 'rr', 'oran'))
    if 'lift' in s: return 'lift5'
    if re.search(r'top%?5', s): return 'top5_kontrol_rr' if oran else 'top5_kontrol_n'
    if 'eleme' in s: return 'eleme_rr' if oran else 'eleme_baz'
    if 'kontrol' in s: return 'kontrol_rr' if oran else 'kontrol_n'
    if 'paylasilan' in s: return 'paylasilan_rr' if oran else 'paylasilan_n'
    if 'tutundurma' in s: return 'tut3y' if '3y' in s else 'tut1y' if any(x in s for x in ('1y', 'incremental', 'gelir')) else 'tut'
    if 'gelir' in s: return 'ek_gelir'
    if 'etki' in s: return 'etki_ay'
    if 'baz' in s: return 'baz_rr' if oran else 'baz'
    return None


def _ay_coz(h):
    """Kolon başlığından YYYYMM: 202608, 2026-08, tarih, 'Ağu 26', 'Ağustos 2026'."""
    if isinstance(h, (dt.datetime, dt.date, pd.Timestamp)): return f"{h.year}{h.month:02d}"
    s = str(h).strip(); m = re.match(r'^(20\d{2})[-/. ]?(\d{2})(?:\.0)?$', s)
    if m and 1 <= int(m.group(2)) <= 12: return m.group(1) + m.group(2)
    m = re.match(r'^([a-z]{3})[a-z]*(\d{4}|\d{2})$', _sade(s).replace('%', ''))      # Mar.25, Ocak'26, Ağustos 2026
    if m and m.group(1) in _AYLAR:
        y = m.group(2); return ('20' + y if len(y) == 2 else y) + f"{_AYLAR[m.group(1)]:02d}"
    return None


def _sayi(v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return None
    if isinstance(v, (int, float, np.integer, np.floating)): return float(v)
    s = re.sub(r'[^0-9,.\-]', '', str(v))
    if not s or s in '-.,': return None
    if ',' in s: s = s.replace('.', '').replace(',', '.')                      # 1.234,5 → 1234.5
    elif re.match(r'^-?\d{1,3}(\.\d{3})+$', s): s = s.replace('.', '')         # 31.400 → 31400
    try: return float(s)
    except ValueError: return None


def _bicim(v, fmt):
    """Excel hücresini, hücrenin sayı biçimine göre Türkçe gösterim metnine çevirir (ek hesap yapılmaz)."""
    if v is None or (isinstance(v, float) and np.isnan(v)): return ''
    if not isinstance(v, (int, float, np.integer, np.floating)): return str(v).strip()
    fmt = str(fmt or ''); nd = len(fmt.split('.')[1].rstrip('%').rstrip()) if '.' in fmt and re.search(r'\.0+', fmt) else None
    if '%' in fmt:
        nd = 2 if nd is None else nd; return '%' + f"{v * 100:,.{nd}f}".replace(',', '§').replace('.', ',').replace('§', '.')
    if nd is None: nd = 0 if abs(v - round(v)) < 1e-9 or abs(v) >= 100 else 2
    return f"{v:,.{nd}f}".replace(',', '§').replace('.', ',').replace('§', '.')


def _kpi_izgara(H):
    """Hücre ızgarasından ([(değer, sayı biçimi)]) KPI tablosunu çıkarır: ay başlıkları satırını bulur, solundaki metin kolonlarını grup ve KPI adı kabul eder."""
    say = lambda i: sum(_ay_coz(v) is not None for v, _ in H[i] if v is not None and str(v).strip() != '')
    if not H: return [], [], 0
    en = max(range(min(len(H), 60)), key=say); aylar = [(j, _ay_coz(v)) for j, (v, _) in enumerate(H[en]) if v is not None and _ay_coz(v)]
    if len(aylar) < 2: return [], [], 0
    j0 = aylar[0][0]; satirlar = []; grup = ''
    for row in H[en + 1:]:
        metin = [str(v).strip() for v, _ in row[:j0] if isinstance(v, str) and v.strip()]
        hucre = [row[j] if j < len(row) else (None, '') for j, _ in aylar]
        sayi = []
        for v, _ in hucre:
            if isinstance(v, (int, float)) and not isinstance(v, bool): sayi.append(None if isinstance(v, float) and np.isnan(v) else float(v))
            elif isinstance(v, str) and v.strip(): x = _sayi(v); sayi.append(None if x is None else x / 100 if '%' in v else x)
            else: sayi.append(None)
        if not metin: continue
        if len(metin) > 1: grup = metin[0]
        if all(v is None for v in sayi):
            if len(metin) == 1: grup = metin[0]                                 # yalnız grup başlığı olan satır (Model / Aksiyon / Gelir)
            continue
        ad = metin[-1]
        satirlar.append(dict(grup=grup, ad=ad, key=_kpi_anahtar(ad), v=[None if v is None else round(v, 6) for v in sayi],
                             s=[(v.strip() if isinstance(v, str) else _bicim(v, f)) if sayi[i] is not None else '' for i, (v, f) in enumerate(hucre)]))
    return [a for _, a in aylar], satirlar, len(aylar)


def kpi_excel(yol):
    """KPI tablosunu olduğu gibi okur (Excel'de tüm sayfalar taranır, ay başlığı en çok olan sayfa alınır). Değerlere dokunulmaz;
    her hücre Excel'deki sayı biçimiyle (yüzde, binlik ayraç) gösterim metnine çevrilir."""
    if os.path.splitext(yol)[1].lower() == '.csv':
        enc = _kodlama(yol); d = pd.read_csv(yol, sep=_ayrac(yol, enc), encoding=enc, header=None, dtype=str)
        return _kpi_izgara([[(None if isinstance(v, float) else v, '') for v in row] for row in d.itertuples(index=False)])[:2]
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError("openpyxl paketi kurulu değil (pip install openpyxl) — ya da tabloyu CSV olarak kaydedip kpi klasörüne koyun")
    wb = openpyxl.load_workbook(yol, data_only=True); en_iyi = ([], [], 0); bos_formul = False
    for ws in wb.worksheets:
        H = [[(c.value, c.number_format) for c in row] for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 500))]
        s = _kpi_izgara(H)
        if s[2] > en_iyi[2] and s[1]: en_iyi = s
        elif s[2] and not s[1]: bos_formul = True
    if not en_iyi[1]:
        raise ValueError(f"sayfalar {wb.sheetnames} içinde ay başlıkları ve değer satırları bulunamadı"
                         + (" — başlıklar bulundu ama hücre değerleri boş: dosya formül içeriyorsa Excel'de açıp kaydedin" if bos_formul else
                            " — başlık satırında Mar.25, 202503 veya tarih biçiminde en az iki ay bekleniyor"))
    return en_iyi[:2]


def modelleri_oku(klasor):
    """kpi klasöründeki her Excel (veya CSV) dosyası bir proje; tablo olduğu gibi rapora eklenir."""
    out = []
    if not klasor or not os.path.isdir(klasor):
        print(f"KPI klasörü yok: {klasor}"); return out
    dosyalar = [y for y in sorted(glob.glob(os.path.join(klasor, '*'))) if os.path.splitext(y)[1].lower() in ('.xlsx', '.xlsm', '.csv') and not os.path.basename(y).startswith(('~', '.'))]
    if not dosyalar: print(f"KPI klasöründe .xlsx / .csv dosyası yok: {klasor} → {os.listdir(klasor)}")
    for yol in dosyalar:
        ad, uz = os.path.splitext(os.path.basename(yol))
        try:
            aylar, satirlar = kpi_excel(yol)
            if not satirlar: raise ValueError("ay başlıkları veya değer satırları bulunamadı")
        except Exception as e:
            print(f"⚠ kpi/{ad}{uz} okunamadı: {e}"); continue
        out.append(dict(id=ad, ad=ad.replace('_', ' ').title(), aylar=aylar, satirlar=satirlar))
        print(f"KPI tablosu {ad}{uz}: {aylar[0]}–{aylar[-1]} ({len(aylar)} ay) · {len(satirlar)} satır · tanınan: {[s['key'] for s in satirlar if s['key']]}")
    return out


# ═══════════════════════════ LIFECYCLE VE KOHORT TABLOLARI ═══════════════════════════
def _tablo_oku(yol):
    uz = os.path.splitext(yol)[1].lower()
    if uz in ('.xlsx', '.xlsm', '.xls'): return pd.read_excel(yol)
    enc = _kodlama(yol); return pd.read_csv(yol, sep=_ayrac(yol, enc), encoding=enc, low_memory=False)


def _dosyalar(klasor):
    return [y for y in sorted(glob.glob(os.path.join(klasor, '*'))) if os.path.splitext(y)[1].lower() in ('.csv', '.xlsx', '.xlsm', '.xls') and not os.path.basename(y).startswith(('~', '.'))] if klasor and os.path.isdir(klasor) else []


def _ay6(v):
    s = re.sub(r'\D', '', str(v).split('.')[0]) if not isinstance(v, (dt.date, pd.Timestamp)) else f"{v.year}{v.month:02d}"
    return s[:6] if len(s) >= 6 and s[:2] == '20' else None


HIBRIT_DESEN = r'hibrit|hybrid'              # lifecycle tiplerinden yalnız bu deseni içerenler (Hybrid ile ilgili olanlar) alınır


def lifecycle_oku(klasor):
    """lifecycle tablosu: ay, lifecycle tipi ve adet kolonları. Kolon adları rapordan geldiği gibi olabilir; ay kolonu adından değil
    içeriğinden (YYYYMM değerleri) bulunur. Yalnız Hybrid ile ilgili tipler ay × tip tablosuna çevrilir."""
    par = []
    for yol in _dosyalar(klasor):
        ad = os.path.basename(yol)
        if os.path.splitext(yol)[1].lower() == '.csv':
            enc = _kodlama(yol); d = pd.read_csv(yol, sep=_ayrac(yol, enc), encoding=enc, dtype=str)
        else: d = pd.read_excel(yol, dtype=str)
        d = d.dropna(how='all'); sade = {c: _sade(c) for c in d.columns}
        ayoran = {c: float(d[c].ffill().map(lambda v: _ay6(v) is not None and 1 <= int(_ay6(v)[4:6]) <= 12 and len(re.sub(r'\D', '', str(v).split('.')[0])) in (6, 8)).mean()) for c in d.columns}
        ac = next((c for c in d.columns if 'yearmonth' in sade[c] and ayoran[c] > .8), None) or (max(ayoran, key=ayoran.get) if max(ayoran.values()) > .8 else None)
        tc = (next((c for c in d.columns if c != ac and 'type' in sade[c] and 'group' not in sade[c] and 'yearmonth' not in sade[c]), None)
              or next((c for c in d.columns if c != ac and d[c].map(lambda v: bool(re.search(HIBRIT_DESEN, _sade(v)))).any()), None))
        aday = [c for c in d.columns if c not in (ac, tc) and d[c].map(lambda v: _sayi(v) is not None).mean() > .8]
        nc = next((c for c in aday if any(x in sade[c] for x in ('distin', 'subscriber', 'adet'))), None) or next((c for c in aday if 'count' in sade[c]), None) or (aday[-1] if aday else None)
        if not (tc and nc):
            print(f"⚠ lifecycle/{ad}: tip / adet kolonları bulunamadı. Kolonlar: {list(d.columns)}"); continue
        d = pd.DataFrame(dict(ay=d[ac].ffill().map(_ay6) if ac else 'Toplam', tip=d[tc].astype(str).str.strip(), n=d[nc].map(_sayi))).dropna()
        par.append(d); print(f"Lifecycle {ad}: ay={ac or 'yok (tek toplam)'} · tip={tc} · adet={nc} · {len(d)} satır · {d['ay'].min()}–{d['ay'].max()}")
    if not par:
        if klasor and os.path.isdir(klasor) and not _dosyalar(klasor): print(f"Lifecycle klasöründe .csv / .xlsx dosyası yok: {os.listdir(klasor)}")
        return None
    d = pd.concat(par); d = d[d['tip'].map(lambda s: bool(re.search(HIBRIT_DESEN, _sade(s))))]
    if d.empty:
        print("⚠ lifecycle: Hybrid ile ilgili tip bulunamadı (HIBRIT_DESEN)."); return None
    p = d.groupby(['tip', 'ay'])['n'].sum().unstack('ay').fillna(0); aylar = sorted(p.columns)
    def tur(ad):
        s = _sade(ad); cik = re.match(r'^(hibrit|hybrid)to', s) is not None; gecis = cik or re.search(r'to(hibrit|hybrid)$', s) is not None
        hedef = 'pre' if 'prepaid' in s or 'faturasiz' in s else 'post' if 'postpaid' in s or 'faturali' in s else ''
        return dict(ad=ad, yon='out' if cik else 'in', hat=hedef, sinif='gecis' if gecis else 'yeni')       # yeni = yeni müşteri kazanımı (Yeni Tesis, Port In)
    tipler = sorted(p.index, key=lambda a: (tur(a)['yon'] != 'in', tur(a)['sinif'] != 'yeni', a))
    print("Lifecycle (Hybrid) tipleri:", [(a, tur(a)['yon'], tur(a)['sinif']) for a in tipler])
    return dict(aylar=aylar, tipler=[tur(a) for a in tipler], v=[[int(p.loc[a, m]) for m in aylar] for a in tipler])


KOHORT_GRUP = {'arama': 'Arama listesi', 'aranmis': 'Arananlar', 'ulasilmis': 'Ulaşılanlar', 'kontrol': 'Kontrol grubu', 'rahatfirsat': 'Rahat Fırsat'}
KOHORT_KUMULATIF = True                      # YearMonth_k_…POUT kolonları kümülatif mi? (False → aylık adetler; rapor kümülatife çevirir)
KOHORT_AY_FARKI = 1                          # YearMonth_1 kolonu, liste ayından kaç ay sonrasıdır


def kohort_tablosu(klasor):
    """Geniş kohort tablosu: YEARMONTH + grup adetleri (AramaL, AranmışL, UlasilmisL, KontrolL, RahatFırsatL) + YearMonth_k_<grup>POUT kolonları."""
    par = []
    for yol in _dosyalar(klasor):
        d = _tablo_oku(yol); d.columns = [_sade(c) for c in d.columns]
        if 'yearmonth' not in d.columns or not any(re.match(r'^yearmonth\d+.+pout$', c) for c in d.columns):
            print(f"⚠ kohort/{os.path.basename(yol)}: YEARMONTH ve YearMonth_k_…POUT kolonları bekleniyor; dosya atlandı."); continue
        par.append(d)
    if not par: return None
    d = pd.concat(par); d['yearmonth'] = d['yearmonth'].map(_ay6); d = d.dropna(subset=['yearmonth']).drop_duplicates('yearmonth', keep='last').sort_values('yearmonth')
    kol = {}
    for c in d.columns:
        m = re.match(r'^yearmonth(\d+)(.+?)l?pout$', c)
        if m: kol.setdefault(m.group(2), {})[int(m.group(1))] = c
    K = max(max(v) for v in kol.values()); gruplar = [g for g in KOHORT_GRUP if g in kol] + [g for g in kol if g not in KOHORT_GRUP]
    sayi = lambda v: None if v is None or (isinstance(v, float) and np.isnan(v)) else float(v) if isinstance(v, (int, float, np.integer, np.floating)) else _sayi(v)
    taban = {g: next((c for c in (g + 'l', g) if c in d.columns), None) for g in gruplar}
    out = dict(aylar=list(d['yearmonth']), K=K, kum=KOHORT_KUMULATIF, fark=KOHORT_AY_FARKI, gruplar=[dict(key=g, name=KOHORT_GRUP.get(g, g)) for g in gruplar],
               taban={g: [sayi(v) for v in d[taban[g]]] if taban[g] else [None] * len(d) for g in gruplar},
               pout={g: [[sayi(row[kol[g][k]]) if k in kol[g] else None for k in range(1, K + 1)] for _, row in d.iterrows()] for g in gruplar})
    print(f"Kohort tablosu: {out['aylar'][0]}–{out['aylar'][-1]} · gruplar={[KOHORT_GRUP.get(g, g) for g in gruplar]} · izleme {K} ay")
    return out


# ═══════════════════════════ HTML ═══════════════════════════
def html_yaz(veri, cikti, sablon=None):
    """Veriyi şablona gömer. Şablon bu dosyanın sonundaki SABLON metnidir; --sablon ile dışarıdan bir dosya da verilebilir."""
    tpl = SABLON
    if sablon:
        with open(sablon, encoding='utf-8') as f:
            tpl = f.read()
    ham = json.dumps(veri, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
    blok = base64.b64encode(gzip.compress(ham, 9)).decode('ascii')
    with open(cikti, 'w', encoding='utf-8') as f:
        f.write(tpl.replace('__VERI__', blok))
    print(f"✔ {cikti} oluşturuldu ({os.path.getsize(cikti) / 1e6:.1f} MB; gömülü veri {len(blok) / 1e6:.2f} MB)")


def main():
    burasi = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser(description="Rahat segment raporu üretici")
    ap.add_argument('--veri', default='data', help="veri klasörü: raw_data/, kohort/, lifecycle/, kpi/ ve (isteğe bağlı) post_model/ alt klasörleri")
    ap.add_argument('--cikti', default='rahat_segment_raporu.html')
    ap.add_argument('--sablon', default=None, help="isteğe bağlı: gömülü şablon yerine kullanılacak HTML şablon dosyası")
    a = ap.parse_args(); alt = lambda ad: os.path.join(a.veri, ad)
    ham = alt('raw_data') if os.path.isdir(alt('raw_data')) else a.veri
    dosyalar = dosyalari_bul(ham); print(f"Aylık dosyalar ({ham}):", [os.path.basename(p) for _, p in dosyalar])
    kohort_oku(alt('post_model'))                                              # müşteri bazlı skor / liste / kontrol tabloları (varsa)
    R = uret(dosyalar)
    veri = dict(meta=dict(uretim=dt.datetime.now().strftime('%d.%m.%Y %H:%M'), k1=len(OLCU), k2=K2, olcu=[k for k, _ in OLCU]), months=[str(m) for m, _ in dosyalar],
                ev=EV, evl=EV_LBL, model=modelleri_oku(alt('kpi')), lc=lifecycle_oku(alt('lifecycle')), kh=kohort_tablosu(alt('kohort')), **R)
    html_yaz(veri, a.cikti, a.sablon)


# ═══════════════════════════ RAPOR ŞABLONU (HTML / CSS / JS) ═══════════════════════════
# Raporun görünümü ve sayfa kodu. __VERI__ işaretinin yerine özetlenmiş veri gömülür.
SABLON = r'''<!doctype html>
<html lang="tr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="light dark">
<title>Rahat Segment Panosu</title>
<style>
:root{
  color-scheme:light;
  --page:#f6f6f3; --surface:#ffffff; --surface2:#f3f3ef;
  --ink:#0b0b0b; --ink2:#52514e; --muted:#6b6a65;
  --grid:#e6e5df; --axis:#c3c2b7; --border:rgba(11,11,11,.10);
  --s1:#2a78d6; --s2:#eb6834; --s3:#1baf7a; --s4:#eda100; --s5:#e87ba4; --s6:#008300; --s7:#6250d6; --s8:#e34948;
  --neg:#e34948; --neutral:#a3a199;
  --good:#0ca30c; --warn:#fab219; --bad:#d03b3b;
  --good-text:#006300; --warn-text:#8a5a00; --bad-text:#b42c2c;
  --accent:#2a78d6; --accent-ink:#1c5cab; --accent-wash:rgba(42,120,214,.10);
  --shadow:0 1px 2px rgba(11,11,11,.04),0 4px 16px rgba(11,11,11,.04);
  --radius:14px;
}
:root[data-theme="dark"]{
  color-scheme:dark;
  --page:#0d0d0d; --surface:#1a1a19; --surface2:#232321;
  --ink:#f0efec; --ink2:#c3c2b7; --muted:#9a988f;
  --grid:#2c2c2a; --axis:#44443f; --border:rgba(255,255,255,.10);
  --s1:#3987e5; --s2:#d95926; --s3:#199e70; --s4:#c98500; --s5:#d55181; --s6:#008300; --s7:#9085e9; --s8:#e66767;
  --neg:#e66767; --neutral:#6f6e68;
  --good-text:#35c435; --warn-text:#fab219; --bad-text:#f08a8a;
  --accent:#3987e5; --accent-ink:#86b6ef; --accent-wash:rgba(57,135,229,.16);
  --shadow:none;
}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--page);color:var(--ink);font:15px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
h1,h2,h3,p{margin:0}
button,select{font:inherit;color:inherit}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px;border-radius:6px}
.skip{position:absolute;left:-999px;top:8px;background:var(--surface);padding:8px 12px;border-radius:8px;z-index:50}
.skip:focus{left:8px}
.wrap{max-width:1280px;margin:0 auto;padding:0 20px}

/* Üst bar */
.top{position:sticky;top:0;z-index:20;background:color-mix(in srgb,var(--page) 88%,transparent);backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);border-bottom:1px solid var(--border)}
.top-row{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:14px 0 10px;flex-wrap:wrap}
.brand{display:flex;align-items:center;gap:12px;min-width:0}
.logo{width:38px;height:38px;border-radius:11px;background:var(--accent);color:#fff;display:grid;place-items:center;font-weight:700;font-size:18px;flex:none}
.brand h1{font-size:18px;line-height:1.2;letter-spacing:-.01em}
.brand p{font-size:13px;color:var(--ink2)}
.controls{display:flex;align-items:center;gap:10px}
.field{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--ink2)}
select{min-height:40px;padding:0 34px 0 12px;border:1px solid var(--border);border-radius:10px;background:var(--surface) url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8'%3E%3Cpath d='M1 1.5l5 5 5-5' fill='none' stroke='%23898781' stroke-width='1.8' stroke-linecap='round'/%3E%3C/svg%3E") no-repeat right 12px center;appearance:none;-webkit-appearance:none;font-weight:600;color:var(--ink);cursor:pointer}
.icon-btn{width:40px;height:40px;border-radius:10px;border:1px solid var(--border);background:var(--surface);cursor:pointer;display:grid;place-items:center}
.icon-btn:hover,select:hover{border-color:var(--axis)}
.tabs{display:flex;gap:4px;overflow-x:auto;scrollbar-width:none;margin:0 -4px;padding:0 4px}
.tabs::-webkit-scrollbar{display:none}
.tab{flex:none;min-height:44px;padding:0 14px;border:0;background:none;cursor:pointer;color:var(--ink2);font-weight:600;font-size:14px;border-bottom:2px solid transparent;display:flex;align-items:center;gap:8px;white-space:nowrap}
.tab:hover{color:var(--ink)}
.tab[aria-selected="true"]{color:var(--accent-ink);border-bottom-color:var(--accent)}
.tab .n{font-size:11px;width:20px;height:20px;border-radius:6px;background:var(--surface2);display:grid;place-items:center;color:var(--ink2)}
.tab[aria-selected="true"] .n{background:var(--accent-wash);color:var(--accent-ink)}

/* Sayfa */
main{padding:22px 0 40px;overflow-x:clip}
.banner{display:flex;gap:10px;align-items:flex-start;background:var(--accent-wash);border-radius:12px;padding:10px 14px;font-size:13px;color:var(--ink2);margin-bottom:18px}
.banner b{color:var(--ink)}
.page-head{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;flex-wrap:wrap;margin-bottom:16px}
.page-head h2{font-size:24px;letter-spacing:-.02em;line-height:1.2}
.page-head p{color:var(--ink2);max-width:70ch;margin-top:4px}
.grid{display:grid;gap:16px;grid-template-columns:repeat(12,minmax(0,1fr));margin-bottom:16px}
.c12{grid-column:span 12}.c8{grid-column:span 8}.c7{grid-column:span 7}.c6{grid-column:span 6}.c5{grid-column:span 5}.c4{grid-column:span 4}
@media(max-width:980px){.c8,.c7,.c6,.c5,.c4{grid-column:span 12}}
.tiles{display:grid;gap:12px;grid-template-columns:repeat(auto-fit,minmax(168px,1fr));margin-bottom:16px}
.card,.tile{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow)}
.tile{padding:14px 16px;min-width:0}
.tile .lbl{font-size:13px;color:var(--ink2)}
.tile .val{font-size:26px;font-weight:650;letter-spacing:-.02em;line-height:1.25;margin-top:2px}
.tile .sub{font-size:12px;color:var(--muted);margin-top:2px}
.delta{font-size:12px;font-weight:650;white-space:nowrap}
.delta.good{color:var(--good-text)}.delta.bad{color:var(--bad-text)}.delta.flat{color:var(--muted)}
.ref{font-size:12px;color:var(--muted)}
.card{padding:16px 18px 18px;min-width:0}
.card>header{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:flex-start;gap:8px 12px;margin-bottom:12px}
.card>header>div:first-child{flex:1 1 160px;min-width:0}
.card>header .controls{flex-wrap:wrap}
.card h3{font-size:15px;line-height:1.3}
.card header p{font-size:13px;color:var(--ink2);margin-top:2px}
.tbl-btn{flex:none;min-height:32px;padding:0 10px;border-radius:8px;border:1px solid var(--border);background:none;font-size:12px;font-weight:600;color:var(--ink2);cursor:pointer}
.tbl-btn:hover{background:var(--surface2)}
.tbl-btn[aria-pressed="true"]{background:var(--accent-wash);color:var(--accent-ink);border-color:transparent}
.chart{min-width:0}
.chart svg{display:block;width:100%;overflow:visible}
.scroll-x{overflow-x:auto}
svg text{font:11px system-ui,-apple-system,"Segoe UI",sans-serif;fill:var(--muted);font-variant-numeric:tabular-nums}
svg text.strong{fill:var(--ink);font-weight:650}
svg text.lab{fill:var(--ink2)}
.gl{stroke:var(--grid);stroke-width:1}.ax{stroke:var(--axis);stroke-width:1}
.hit .x{opacity:0}.hit:hover .x{opacity:1}
.hit rect.h{fill:transparent}
.wash{fill:var(--ink);opacity:.05}
.markband{fill:var(--accent);opacity:.07}
.legend{display:flex;flex-wrap:wrap;gap:4px 14px;list-style:none;padding:0;margin:0 0 8px;font-size:12px;color:var(--ink2)}
.legend li{display:flex;align-items:center;gap:6px}
.sw{width:10px;height:10px;border-radius:3px;flex:none;display:inline-block}

/* Yatay çubuklar */
.hb{display:grid;grid-template-columns:minmax(90px,34%) 1fr auto;gap:6px 10px;align-items:center;font-size:13px}
.hb .l{color:var(--ink2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.hb .t{height:12px;position:relative}
.hb .b{position:absolute;left:0;top:0;bottom:0;border-radius:0 4px 4px 0;min-width:2px}
.hb .v{font-variant-numeric:tabular-nums;font-weight:600;text-align:right}
.hb .row{display:contents}
.hb .row:hover .l,.hb .row:focus .l{color:var(--ink)}
.db{display:grid;grid-template-columns:minmax(90px,30%) 1fr auto;gap:6px 10px;align-items:center;font-size:13px}
.db .l{color:var(--ink2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.db .t{height:12px;position:relative;background:linear-gradient(var(--axis),var(--axis)) 50% 0/1px 100% no-repeat}
.db .b{position:absolute;top:0;bottom:0;min-width:2px}
.db .b.p{left:50%;background:var(--good);border-radius:0 4px 4px 0}
.db .b.n{right:50%;background:var(--neg);border-radius:4px 0 0 4px}
.db .v{font-variant-numeric:tabular-nums;font-weight:600;text-align:right}
.db .row{display:contents}

/* Segment çubuğu */
.seg{display:flex;gap:2px;height:28px}
.seg span{border-radius:4px;color:#fff;font-size:12px;font-weight:650;display:flex;align-items:center;padding-left:7px;min-width:3px;overflow:hidden;white-space:nowrap}
.seg span.dk{color:#0b0b0b}
.seg+.legend{margin:10px 0 0}
.block+.block{margin-top:18px}
.block h4{font-size:13px;font-weight:600;color:var(--ink2);margin:0 0 8px}

/* Tablolar */
.tblwrap{overflow-x:auto;position:relative}
.tblwrap th:first-child{position:sticky;left:0;z-index:1;background:var(--surface)}
.tblwrap tr.grp th:first-child{background:var(--surface2)}
table{border-collapse:collapse;width:100%;font-size:13px}
th,td{padding:8px 10px;text-align:right;border-bottom:1px solid var(--grid);white-space:nowrap;font-variant-numeric:tabular-nums}
th:first-child,td:first-child{text-align:left}
thead th{font-size:12px;font-weight:600;color:var(--ink2);border-bottom-color:var(--axis)}
tbody tr:last-child td,tbody tr:last-child th{border-bottom:0}
tbody th{font-weight:600}
td.wrap-ok,th.wrap-ok{white-space:normal}
.heat td{text-align:center;border:2px solid var(--surface);border-radius:6px;padding:10px 6px}
.heat th{border:0}

/* Öne çıkanlar */
.insights{display:grid;gap:10px;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));margin-bottom:16px}
.ins{background:var(--surface);border:1px solid var(--border);border-left:3px solid var(--accent);border-radius:10px;padding:10px 14px;font-size:13px;color:var(--ink2)}
.ins b{color:var(--ink)}

/* KPI */
.kpis{display:grid;gap:12px;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));margin-bottom:16px}
.kpi{padding:14px 16px}
.kpi .hd{display:flex;justify-content:space-between;gap:8px;align-items:flex-start}
.kpi .lbl{font-size:13px;color:var(--ink2)}
.kpi .row2{display:flex;justify-content:space-between;align-items:flex-end;gap:10px;margin-top:4px}
.kpi .val{font-size:26px;font-weight:650;letter-spacing:-.02em;line-height:1.2}
.kpi .tg{font-size:12px;color:var(--muted)}
.chip{display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:650;padding:3px 8px;border-radius:999px;white-space:nowrap}
.chip i{font-style:normal;width:14px;height:14px;border-radius:50%;display:grid;place-items:center;font-size:10px;color:#fff}
.chip.ok{background:color-mix(in srgb,var(--good) 14%,transparent);color:var(--good-text)}.chip.ok i{background:var(--good)}
.chip.warn{background:color-mix(in srgb,var(--warn) 20%,transparent);color:var(--warn-text)}.chip.warn i{background:var(--warn);color:#0b0b0b}
.chip.bad{background:color-mix(in srgb,var(--bad) 14%,transparent);color:var(--bad-text)}.chip.bad i{background:var(--bad)}
.meter{height:8px;border-radius:999px;margin:10px 0 6px;position:relative;overflow:hidden}
.meter b{position:absolute;left:0;top:0;bottom:0;border-radius:999px}
.meter.ok{background:color-mix(in srgb,var(--good) 18%,var(--surface))}.meter.ok b{background:var(--good)}
.meter.warn{background:color-mix(in srgb,var(--warn) 22%,var(--surface))}.meter.warn b{background:var(--warn)}
.meter.bad{background:color-mix(in srgb,var(--bad) 18%,var(--surface))}.meter.bad b{background:var(--bad)}
.hero{display:flex;align-items:center;gap:18px;flex-wrap:wrap;padding:18px 20px;margin-bottom:16px}
.hero .big{font-size:52px;font-weight:650;letter-spacing:-.03em;line-height:1}
.hero .big small{font-size:22px;color:var(--muted);font-weight:600}
.hero p{color:var(--ink2);max-width:60ch}
.hero .chips{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px}

.model-meta{display:flex;flex-wrap:wrap;gap:6px 18px;font-size:13px;color:var(--ink2);margin-bottom:14px}
.model-meta b{color:var(--ink)}
.persona{display:grid;gap:12px;grid-template-columns:repeat(auto-fit,minmax(210px,1fr))}
.persona article{border:1px solid var(--border);border-radius:12px;padding:12px 14px}
.persona h4{margin:0 0 2px;font-size:14px;display:flex;align-items:center;gap:8px}
.persona .sh{font-size:22px;font-weight:650;letter-spacing:-.02em}
.persona dl{margin:8px 0 0;display:grid;grid-template-columns:auto 1fr;gap:2px 10px;font-size:12px}
.persona dt{color:var(--muted)}.persona dd{margin:0;text-align:right;font-weight:600;font-variant-numeric:tabular-nums}
.tags{display:flex;flex-wrap:wrap;gap:4px;margin-top:8px}
.tags span{font-size:11px;background:var(--surface2);border-radius:6px;padding:2px 7px;color:var(--ink2)}

.segs{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 18px}
@media(max-width:700px){.segs{flex-wrap:nowrap;overflow-x:auto;scrollbar-width:none;margin:0 -20px 16px;padding:2px 20px}.segs::-webkit-scrollbar{display:none}}
.segs::-webkit-scrollbar{display:none}
.seg-chip{flex:none;display:flex;align-items:center;gap:8px;min-height:44px;padding:6px 16px 6px 12px;border-radius:999px;border:1px solid var(--border);background:var(--surface);cursor:pointer;text-align:left;box-shadow:var(--shadow)}
.seg-chip .dot{width:10px;height:10px;border-radius:50%;flex:none}
.seg-chip .nm{font-weight:650;font-size:14px;white-space:nowrap}
.seg-chip .ct{font-size:12px;color:var(--muted);font-variant-numeric:tabular-nums;white-space:nowrap}
.seg-chip:hover{border-color:var(--axis)}
.seg-chip[aria-checked="true"]{background:var(--accent-wash);border-color:var(--accent);box-shadow:none}
.seg-chip[aria-checked="true"] .ct{color:var(--ink2)}
.chart>svg text.strong,.chart>svg text.lab{paint-order:stroke;stroke:var(--surface);stroke-width:3px;stroke-linejoin:round}
.page-head .controls{flex-wrap:wrap}
#tip{position:fixed;z-index:90;pointer-events:none;background:var(--surface);color:var(--ink);border:1px solid var(--border);border-radius:10px;box-shadow:0 6px 24px rgba(0,0,0,.16);padding:8px 10px;font-size:12px;max-width:280px;opacity:0;transition:opacity .08s}
#tip.on{opacity:1}
#tip .tt{font-weight:650;margin-bottom:4px}
#tip .tr{display:flex;align-items:center;gap:6px;justify-content:space-between}
#tip .tr span:first-child{display:flex;align-items:center;gap:6px;color:var(--ink2)}
#tip .tr b{font-variant-numeric:tabular-nums}
footer{border-top:1px solid var(--border);padding:16px 0 28px;font-size:12px;color:var(--muted)}
footer details{margin-top:6px}
footer summary{cursor:pointer;color:var(--ink2);font-weight:600;min-height:28px}
footer li{margin:2px 0}
.sr{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap}
@media(prefers-reduced-motion:reduce){*{transition:none!important;scroll-behavior:auto!important}}
@media print{.top{position:static}.controls,.tbl-btn,.banner{display:none}.card,.tile{break-inside:avoid;box-shadow:none}}

.period{font-size:13px;color:var(--ink2);white-space:nowrap}.period b{color:var(--ink);font-weight:650;margin-left:6px;font-size:14px}
h3.sec{font-size:16px;margin:22px 0 10px;letter-spacing:-.01em}
.tiles.g2{grid-template-columns:repeat(auto-fit,minmax(270px,1fr))}
.duo{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:6px}
.duo .k{font-size:11px;color:var(--muted);font-weight:600;letter-spacing:.03em;text-transform:uppercase}
.val.sm{font-size:19px}
.tiles.g8{grid-template-columns:repeat(10,minmax(0,1fr));gap:8px;margin-bottom:12px}
@media(max-width:1180px){.tiles.g8{grid-template-columns:repeat(5,minmax(0,1fr))}}
@media(max-width:560px){.tiles.g8{grid-template-columns:repeat(2,minmax(0,1fr))}}
.tile.ct{padding:8px 11px}.tile.ct .lbl{font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tile.ct .val{font-size:18px;margin-top:0;white-space:nowrap}.tile.ct .dl{font-size:11px;color:var(--muted);white-space:nowrap}.tile.ct .delta{font-size:11px}
.g3{display:grid;grid-template-columns:1fr 1fr 1.05fr;gap:10px;margin-bottom:16px;align-items:start}
@media(max-width:1100px){.g3{grid-template-columns:1fr}}
.card.gl{padding:9px 12px 8px;overflow-x:auto}.card.gl h3{font-size:13px;margin-bottom:2px}
table.gt{font-size:12px}table.gt th,table.gt td{padding:2.5px 5px;border-bottom:1px solid var(--grid)}
table.gt thead th{font-size:10.5px;padding-bottom:3px}
table.gt tbody th{width:100%;max-width:0;overflow:hidden;text-overflow:ellipsis;font-weight:600;color:var(--ink2)}
table.gt td.d{color:var(--muted);font-size:11px}table.gt .delta{font-size:11px}
.db .v{white-space:nowrap}.db .v small{font-weight:500;color:var(--muted);font-size:12px;margin-left:8px}
.card.grp{padding:12px 14px;margin-bottom:10px}.card.grp>header{margin-bottom:8px;align-items:center}.card.grp h3{font-size:14px}
.minis{display:grid;gap:6px;grid-template-columns:repeat(auto-fill,minmax(172px,1fr))}
.mini{background:var(--surface2);border-radius:9px;padding:6px 9px;min-width:0}
.mini .nm{font-size:12px;font-weight:600;color:var(--ink2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.mini .ln{display:flex;justify-content:space-between;align-items:baseline;gap:6px;font-size:13px;line-height:1.45;font-variant-numeric:tabular-nums}
.mini .ln span{font-size:11px;color:var(--muted);white-space:nowrap}.mini .delta{font-size:11px}
.tiles.g4{grid-template-columns:repeat(4,minmax(0,1fr))}
@media(max-width:900px){.tiles.g4{grid-template-columns:repeat(2,minmax(0,1fr))}}
.card>header .controls{flex-wrap:wrap;justify-content:flex-end;gap:6px}
.card>header .controls select{min-height:34px;font-size:13px;max-width:200px;padding:0 30px 0 10px}
input[type=number]{min-height:34px;width:82px;padding:0 8px;border:1px solid var(--border);border-radius:10px;background:var(--surface);color:var(--ink);font:inherit;font-weight:600}
th.cur,td.cur{background:var(--accent-wash)}
tr.grp th{background:var(--surface2);font-size:12px;color:var(--ink2);text-transform:uppercase;letter-spacing:.04em}
tr.em th,tr.em td,.tblwrap tr.em th:first-child{background:color-mix(in srgb,var(--surface2) 55%,var(--surface))}
.loading{padding:60px 0;text-align:center;color:var(--ink2)}
body.fs-on{overflow:hidden}
body.fs-on::before{content:"";position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:70}
.card.full{position:fixed;inset:12px;z-index:80;overflow:auto;margin:0;box-shadow:0 24px 70px rgba(0,0,0,.4)}
.card.full .tblwrap{max-height:none!important;overflow:visible}
.card.full table{font-size:14px}.card.full table.heat td{padding:12px 8px}
.card.full>header{position:sticky;top:-16px;background:var(--surface);z-index:3;padding:6px 0;margin-top:-6px}
@media print{.fs-btn{display:none}}
code{font-size:.92em;background:var(--surface2);padding:1px 5px;border-radius:5px}
</style>
</head>
<body>
<a class="skip" href="#main">İçeriğe geç</a>
<header class="top">
  <div class="wrap">
    <div class="top-row">
      <div class="brand">
        <div class="logo" aria-hidden="true">R</div>
        <div><h1>Rahat Segment Panosu</h1><p>Hybrid (HB) müşteri segmenti · aylık rapor</p></div>
      </div>
      <div class="controls">
        <label class="field" for="month">Dönem <select id="month"></select></label>
        <button class="icon-btn" id="theme" type="button" aria-label="Açık / koyu tema değiştir" title="Tema">
          <svg width="18" height="18" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3a9 9 0 1 0 0 18z" fill="currentColor"/><circle cx="12" cy="12" r="8.2" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>
        </button>
      </div>
    </div>
    <div class="tabs" role="tablist" aria-label="Rapor sekmeleri" id="tabs"></div>
  </div>
</header>
<main id="main" class="wrap">
  <div id="panel" role="tabpanel" tabindex="-1"><p class="loading">Rapor yükleniyor…</p></div>
</main>
<footer><div class="wrap"><span id="foot"></span></div></footer>
<div id="tip" aria-hidden="true"></div>
<script id="veri" type="application/octet-stream">__VERI__</script>
<script>
document.documentElement.dataset.theme=(()=>{try{return localStorage.getItem('rahat-theme')}catch(e){return null}})()||(matchMedia('(prefers-color-scheme: dark)').matches?'dark':'light');
(async()=>{
const yukle=async()=>{const b=atob(document.getElementById('veri').textContent.trim()),u=new Uint8Array(b.length);for(let i=0;i<b.length;i++)u[i]=b.charCodeAt(i);
  return JSON.parse(await new Response(new Blob([u]).stream().pipeThrough(new DecompressionStream('gzip'))).text())};
let D;try{D=await yukle()}catch(e){document.getElementById('panel').innerHTML='<p class="loading">Rapor verisi açılamadı. Güncel bir Chrome, Edge, Safari veya Firefox ile açın.</p>';return}
let T=D.months.length-1;const MONTHS=D.months,TN=MONTHS.length-1,K1=D.meta.k1,K2=D.meta.k2,arr=(n,fn)=>Array.from({length:n},(_,k)=>fn(k));
/* ================= Yardımcılar ================= */
const $=(s,r=document)=>r.querySelector(s);
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const num=(v,d=0)=>v.toLocaleString('tr-TR',{minimumFractionDigits:d,maximumFractionDigits:d});
const f0=v=>num(Math.round(v)), f1=v=>num(v,1), f2=v=>num(v,2);
const pct=v=>'%'+num(v,1), pct2=v=>'%'+num(v,2), tl=v=>num(v,1)+' TL', gbf=v=>num(v,1)+' GB', dk=v=>f0(v)+' dk', ay=v=>num(v,1)+' ay';
const compact=v=>Math.abs(v)>=1e6?num(v/1e6,1)+' Mn':Math.abs(v)>=1e4?num(v/1e3,0)+' B':f0(v);
const axisF=v=>Math.abs(v)>=1e4?compact(v):v.toLocaleString('tr-TR',{maximumFractionDigits:2});
const MN=['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran','Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık'];
const mLong=id=>MN[+id.slice(4)-1]+' '+id.slice(0,4), mShort=id=>MN[+id.slice(4)-1].slice(0,3)+' '+id.slice(2,4);
const LBL=MONTHS.map(mShort);
const C=k=>`var(--s${k})`;
function ticks(lo,hi,n=4){if(hi===lo)hi=lo+1;const raw=(hi-lo)/n,p=10**Math.floor(Math.log10(raw)),e=raw/p,st=(e<=1?1:e<=2?2:e<=2.5?2.5:e<=5?5:10)*p,out=[];
  for(let v=Math.floor(lo/st+1e-9)*st;v<hi+st*(1-1e-9);v+=st)out.push(+v.toFixed(10));return out}
const tipH=(title,rows)=>`<div class="tt">${esc(title)}</div>`+rows.map(([c,n,v])=>`<div class="tr"><span>${c?`<i class="sw" style="background:${c}"></i>`:''}${esc(n)}</span><b>${esc(v)}</b></div>`).join('');
const legend=series=>series.length<2?'':`<ul class="legend">${series.map(s=>`<li><i class="sw" style="background:${s.color}"></i>${esc(s.name)}</li>`).join('')}</ul>`;
const colPath=(x,y,w,h,r,down)=>{if(h<=0)return'';r=Math.min(r,w/2,h);return down?`M${x},${y}h${w}v${h-r}a${r},${r} 0 0 1 ${-r},${r}h${-(w-2*r)}a${r},${r} 0 0 1 ${-r},${-r}z`:`M${x},${y+h}v${-(h-r)}a${r},${r} 0 0 1 ${r},${-r}h${w-2*r}a${r},${r} 0 0 1 ${r},${r}v${h-r}z`};

/* ---- Çizgi grafik ---- */
function lineChart(el,o){
  const {labels,series,h:h0=230,mark=null,target=null,zero=false,endLabel=true}=o,W=el.clientWidth||600,h=FH(el,h0),n=labels.length,fmt=v=>isFinite(v)?(o.fmt||f0)(v):'–';
  let all=series.flatMap(s=>s.values).concat(target?[target.value]:[]).filter(v=>isFinite(v));if(!all.length)all=[0,1];
  let lo=zero?0:Math.min(...all),hi=Math.max(...all);const pad=(hi-lo)*.08;if(!zero)lo-=pad;hi+=pad;
  const tk=ticks(lo,hi,4);lo=tk[0];hi=tk[tk.length-1];
  const m={t:14,r:series.length===1&&endLabel?Math.max(18,fmt(series[0].values[n-1]).length*3.6+6):14,b:24,l:Math.max(...tk.map(v=>axisF(v).length))*6.4+12};
  const pw=W-m.l-m.r,ph=h-m.t-m.b,x=i=>m.l+(n>1?pw*i/(n-1):pw/2),y=v=>m.t+ph*(1-(v-lo)/(hi-lo));
  let g='';tk.forEach(v=>{g+=`<line class="${v===0?'ax':'gl'}" x1="${m.l}" x2="${W-m.r}" y1="${y(v)}" y2="${y(v)}"/><text x="${m.l-8}" y="${y(v)+4}" text-anchor="end">${axisF(v)}</text>`});
  const step=Math.ceil(n*58/pw);
  labels.forEach((l,i)=>{if(i===mark||((n-1-i)%step===0&&(mark==null||Math.abs(i-mark)>=step)))g+=`<text x="${x(i)}" y="${h-6}" text-anchor="middle" class="${i===mark?'strong':''}">${esc(l)}</text>`});
  if(mark!=null)g+=`<line x1="${x(mark)}" x2="${x(mark)}" y1="${m.t}" y2="${m.t+ph}" stroke="var(--accent)" stroke-width="1" opacity=".35"/>`;
  if(target)g+=`<line x1="${m.l}" x2="${W-m.r}" y1="${y(target.value)}" y2="${y(target.value)}" stroke="var(--ink2)" stroke-width="1"/><text class="lab" x="${m.l+4}" y="${y(target.value)-5}">${esc(target.label)}</text>`;
  series.forEach(s=>{const pts=s.values.map((v,i)=>isFinite(v)?`${x(i)},${y(v)}`:'').filter(Boolean).join(' ');if(!pts)return;
    if(series.length===1)g+=`<polygon points="${x(0)},${m.t+ph} ${pts} ${x(n-1)},${m.t+ph}" fill="${s.color}" opacity=".08"/>`;
    g+=`<polyline points="${pts}" fill="none" stroke="${s.color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
    const k=mark??n-1;if(isFinite(s.values[k]))g+=`<circle cx="${x(k)}" cy="${y(s.values[k])}" r="4.5" fill="${s.color}" stroke="var(--surface)" stroke-width="2"/>`;
    if(series.length===1&&endLabel&&isFinite(s.values[mark??n-1])){const e=mark??n-1;g+=`<text class="strong" x="${x(e)}" y="${y(s.values[e])-10}" text-anchor="${e===n-1?'end':e===0?'start':'middle'}">${fmt(s.values[e])}</text>`}});
  labels.forEach((l,i)=>{const bw=n>1?pw/(n-1):pw;
    g+=`<g class="hit" data-tip="${esc(tipH(o.tipTitle?o.tipTitle(i):l,series.map(s=>[s.color,s.name,fmt(s.values[i])]).concat(target?[[null,target.label,fmt(target.value)]]:[])))}"><rect class="h" x="${x(i)-bw/2}" y="${m.t}" width="${bw}" height="${ph+m.b}"/><line class="x" x1="${x(i)}" x2="${x(i)}" y1="${m.t}" y2="${m.t+ph}" stroke="var(--axis)"/>${series.filter(s=>isFinite(s.values[i])).map(s=>`<circle class="x" cx="${x(i)}" cy="${y(s.values[i])}" r="4.5" fill="${s.color}" stroke="var(--surface)" stroke-width="2"/>`).join('')}</g>`});
  el.innerHTML=legend(series)+`<svg role="img" aria-label="${esc(o.aria||'Çizgi grafik')}" height="${h}" viewBox="0 0 ${W} ${h}">${g}</svg>`;
  return {head:[o.xName||'Dönem',...series.map(s=>s.name)],rows:labels.map((l,i)=>[l,...series.map(s=>fmt(s.values[i]))])};
}
/* ---- Sütun grafik (gruplu / yığılmış / işaretli) ---- */
function colChart(el,o){
  const {labels,series,stacked=false,fmt=f0,h:h0=230,mark=null,pct100=false,line=null}=o,W=el.clientWidth||600,h=FH(el,h0),n=labels.length,K=series.length;
  const tot=labels.map((_,i)=>series.reduce((s,r)=>s+r.values[i],0));let lo=0,hi;
  if(stacked)hi=pct100?100:Math.max(...tot);else{const all=series.flatMap(s=>s.values);hi=Math.max(0,...all);lo=Math.min(0,...all)}
  if(line)hi=Math.max(hi,line.value);
  const tk=pct100?[0,25,50,75,100]:ticks(lo,hi,4);lo=tk[0];hi=tk[tk.length-1];
  const m={t:12,r:8,b:24,l:Math.max(...tk.map(v=>axisF(v).length))*6.4+12},pw=W-m.l-m.r,ph=h-m.t-m.b,band=pw/n,y=v=>m.t+ph*(1-(v-lo)/(hi-lo));
  let g='';tk.forEach(v=>{g+=`<line class="gl" x1="${m.l}" x2="${W-m.r}" y1="${y(v)}" y2="${y(v)}"/><text x="${m.l-8}" y="${y(v)+4}" text-anchor="end">${axisF(v)}${pct100?'':''}</text>`});
  const step=Math.ceil(n*58/pw);
  labels.forEach((l,i)=>{const cx=m.l+band*(i+.5);let bars='';
    if(stacked){const bw=Math.min(24,band*.62);let acc=0,last=-1;series.forEach((s,k)=>{if(s.values[i]>0)last=k});
      series.forEach((s,k)=>{const v=s.values[i];if(v<=0)return;const y1=y(acc+v),y0=y(acc);const hh=Math.max(1,y0-y1-(acc>0?2:0));acc+=v;
        bars+=k===last?`<path d="${colPath(cx-bw/2,y1,bw,hh,4)}" fill="${s.color}"/>`:`<rect x="${cx-bw/2}" y="${y1}" width="${bw}" height="${hh}" fill="${s.color}"/>`})}
    else{const bw=Math.max(3,Math.min(24,band*.72/K-2)),x0=cx-(K*bw+(K-1)*2)/2;
      series.forEach((s,k)=>{const v=s.values[i],hh=Math.abs(y(v)-y(0));bars+=`<path d="${colPath(x0+k*(bw+2),v>=0?y(v):y(0),bw,Math.max(1,hh),4,v<0)}" fill="${v<0&&s.neg?s.neg:s.color}"/>`})}
    const rows=series.map(s=>[s.color,s.name,fmt(s.values[i])]);if(stacked&&!pct100&&K>1)rows.push([null,'Toplam',fmt(tot[i])]);
    g+=`<g class="hit" data-tip="${esc(tipH(o.tipTitle?o.tipTitle(i):l,rows))}">${i===mark?`<rect class="markband" x="${cx-band/2}" y="${m.t}" width="${band}" height="${ph}"/>`:''}<rect class="h" x="${cx-band/2}" y="${m.t}" width="${band}" height="${ph+m.b}"/><rect class="x wash" x="${cx-band/2}" y="${m.t}" width="${band}" height="${ph}"/>${bars}</g>`;
    if(i===mark||((n-1-i)%step===0&&(mark==null||Math.abs(i-mark)>=step)))g+=`<text x="${cx}" y="${h-6}" text-anchor="middle" class="${i===mark?'strong':''}">${esc(l)}</text>`});
  g+=`<line class="ax" x1="${m.l}" x2="${W-m.r}" y1="${y(0)}" y2="${y(0)}"/>`;
  if(line)g+=`<line x1="${m.l}" x2="${W-m.r}" y1="${y(line.value)}" y2="${y(line.value)}" stroke="var(--ink2)" stroke-width="1"/><text class="lab" x="${W-m.r}" y="${y(line.value)-5}" text-anchor="end">${esc(line.label)}</text>`;
  el.innerHTML=legend(series)+`<svg role="img" aria-label="${esc(o.aria||'Sütun grafik')}" height="${h}" viewBox="0 0 ${W} ${h}">${g}</svg>`;
  return {head:[o.xName||'Dönem',...series.map(s=>s.name)],rows:labels.map((l,i)=>[l,...series.map(s=>fmt(s.values[i]))])};
}
/* ---- Yatay çubuk listesi ---- */
function hbars(el,o){const {items,fmt=pct,color=C(1)}=o,max=o.max||Math.max(...items.map(d=>d.value));
  el.innerHTML=`<div class="hb">${items.map(d=>`<div class="row" tabindex="0" data-tip="${esc(tipH(d.label,[[color,o.name||'Değer',fmt(d.value)],...(d.extra||[])]))}"><span class="l" title="${esc(d.label)}">${esc(d.label)}</span><span class="t"><span class="b" style="width:${d.value/max*100}%;background:${color}"></span></span><span class="v">${fmt(d.value)}</span></div>`).join('')}</div>`;
  return {head:[o.xName||'Kategori',o.name||'Değer'],rows:items.map(d=>[d.label,fmt(d.value)])}}
/* ---- Fark çubukları (iki dönem arası puan değişimi) ---- */
function deltaBars(el,o){const {items,fmt=v=>(v>0?'▲ ':v<0?'▼ ':'')+f2(Math.abs(v))+' puan'}=o,vf=o.vf||pct,max=Math.max(.1,...items.map(d=>Math.abs(d.value)));
  el.innerHTML=`<ul class="legend"><li><i class="sw" style="background:var(--good)"></i>Arttı ▲</li><li><i class="sw" style="background:var(--neg)"></i>Azaldı ▼</li></ul><div class="db">${items.map(d=>`<div class="row" tabindex="0" data-tip="${esc(tipH(d.label,d.extra||[]))}"><span class="l" title="${esc(d.label)}">${esc(d.label)}</span><span class="t"><span class="b ${d.value>=0?'p':'n'}" style="width:${Math.abs(d.value)/max*50}%"></span></span><span class="v">${fmt(d.value)}${d.note?`<small>${d.note}</small>`:''}</span></div>`).join('')}</div>`;
  return o.tbl||{head:[o.xName||'Kategori',o.aName,o.bName,'Değişim'],rows:items.map(d=>[d.label,vf(d.a),vf(d.b),fmt(d.value)])}}
/* ---- Segment çubuğu ---- */
function segBar(names,vals,colors){const dark={'var(--s4)':1};
  return `<div class="seg" role="img" aria-label="${esc(names.map((n,i)=>n+' '+pct(vals[i])).join(', '))}">${vals.map((v,i)=>`<span class="${dark[colors[i]]?'dk':''}" tabindex="0" style="flex:${v} 0 0;background:${colors[i]}" data-tip="${esc(tipH(names[i],[[colors[i],'Pay',pct(v)]]))}">${v>=11?pct(v):''}</span>`).join('')}</div><ul class="legend">${names.map((n,i)=>`<li><i class="sw" style="background:${colors[i]}"></i>${esc(n)}</li>`).join('')}</ul>`}
function spark(vals,target,color='var(--accent)'){const w=120,h=34,lo=Math.min(...vals,target??Infinity),hi=Math.max(...vals,target??-Infinity),x=i=>3+(w-6)*i/(vals.length-1),y=v=>4+(h-8)*(1-(v-lo)/((hi-lo)||1));
  return `<svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" aria-hidden="true" style="width:${w}px;flex:none">${target!=null?`<line x1="0" x2="${w}" y1="${y(target)}" y2="${y(target)}" stroke="var(--axis)"/>`:''}<polyline points="${vals.map((v,i)=>x(i)+','+y(v)).join(' ')}" fill="none" stroke="var(--neutral)" stroke-width="1.5" stroke-linejoin="round"/><circle cx="${x(vals.length-1)}" cy="${y(vals[vals.length-1])}" r="3.5" fill="${color}" stroke="var(--surface)" stroke-width="1.5"/></svg>`}
const table=(t,cls='')=>`<div class="tblwrap"><table class="${cls}"><thead><tr>${t.head.map(h=>`<th scope="col">${esc(h)}</th>`).join('')}</tr></thead><tbody>${t.rows.map(r=>`<tr>${r.map((c,i)=>i?`<td>${esc(c)}</td>`:`<th scope="row">${esc(c)}</th>`).join('')}</tr>`).join('')}</tbody></table></div>`;

/* ---- Kart ve gösterge yardımcıları ---- */
let queue=[];const tableMode=new Set();
function card(id,title,sub,fn,cls='c6'){queue.push([id,fn]);const on=tableMode.has(id);
  return `<section class="card ${cls}" aria-labelledby="${id}-t"><header><div><h3 id="${id}-t">${title}</h3>${sub?`<p>${sub}</p>`:''}</div><button type="button" class="tbl-btn" id="${id}-b" data-tbl="${id}" aria-pressed="${on}">${on?'Grafik':'Tablo'}</button></header><div class="chart" id="${id}"></div></section>`}
const plain=(title,sub,body,cls='c6')=>`<section class="card ${cls}"><header><div><h3>${title}</h3>${sub?`<p>${sub}</p>`:''}</div></header>${body}</section>`;
const tile=(lbl,val,dl,sub)=>(val=/NaN|Infinity|undefined/.test(String(val))?'–':val,_tile(lbl,val,dl,sub));
const _tile=(lbl,val,dl,sub)=>`<div class="tile"><div class="lbl">${lbl}</div><div class="val">${val}</div><div>${dl||''}</div>${sub?`<div class="sub">${sub}</div>`:''}</div>`;
const head=(t,p,extra='')=>`<div class="page-head"><div><h2>${t}</h2><p>${p}</p></div>${extra}</div>`;

/* ================= Veri erişimi ================= */
const DIX={};D.dims.forEach((d,i)=>DIX[d.key]=i);
const dim=k=>D.dims[DIX[k]],has=k=>k in DIX;
const MK=D.meta.olcu||['arpu','gb','mou','ten','age','price','comp'],W1=1+2*K1,W2=1+K2;
const FI=MK.indexOf('fee'),st1=v=>{const o={n:v[0],rev:FI>=0&&v[1+K1+FI]?v[1+FI]:NaN};MK.slice(0,K1).forEach((k,j)=>{const c=v[1+K1+j];o[k]=c?v[1+j]/c*(k==='comp'?100:1):NaN});return o};
const st2=v=>{const o={n:v?v[0]:0};MK.slice(0,K2).forEach((k,j)=>o[k]=v&&v[0]?v[1+j]/v[0]:NaN);return o};
const tot=t=>t>=0&&t<=TN?st1(D.tot[t]):null;
const one=(t,k)=>{if(t<0||t>TN)return null;const a=D.one[t][k];return arr(dim(k).labels.length,c=>st1(a.slice(c*W1,(c+1)*W1)))};
function cube(t,keys){            // 2 veya 3 kırılımlı hücre erişimi: g([i,j(,k)]) → {n, arpu, gb, mou, ten, age}
  const ord=keys.map((k,i)=>[DIX[k],i]).sort((a,b)=>a[0]-b[0]),sz=ord.map(o=>D.dims[o[0]].labels.length),nm=ord.map(o=>D.dims[o[0]].key).join('|');
  const src=keys.length===2?D.pair[t][nm]:D.tri[t][nm];
  return ix=>{let p=0;ord.forEach((o,q)=>{p=p*sz[q]+ix[o[1]]});return st2(src.slice(p*W2,p*W2+W2))}}
const yoyT=t=>MONTHS.indexOf(String(+MONTHS[t]-100));
const fMn=v=>num(v/1e6,2)+' Mn TL',sgn=v=>(v>0?'+':v<0?'−':'')+'%'+f1(Math.abs(v));
const MS={n:['Abone adedi',f0,f0],fee:['Ort. gelir',tl,f0],arpu:['Ort. ARPU',tl,f0],gb:['Ort. data kullanımı',gbf,f1],mou:['Ort. ses kullanımı',dk,f0],ten:['Ort. tenure',ay,f1],age:['Ort. yaş',f1,f1]};
const msOpts=Object.entries(MS).map(([k,v])=>[k,v[0]]),dimOpts=(ex=[])=>D.dims.filter(d=>!ex.includes(d.key)).map(d=>[d.key,d.name]);
const lc=x=>x==='ARPU'?x:x.toLocaleLowerCase('tr');
const shortLab=l=>l.replace(/ (TL|GB|dk|ay)$/,'');

/* ================= Durum ve ortak bileşenler ================= */
const TABS=[['profil','Müşteri Profili'],['karsilastirma','Aylık Karşılaştırma'],['hareket','Churn & Kazanım'],['model','Model & KPI']];
const store={get(k){try{return localStorage.getItem(k)}catch(e){return null}},set(k,v){try{localStorage.setItem(k,v)}catch(e){}}};
const S={tab:TABS.some(t=>'#'+t[0]===location.hash)?location.hash.slice(1):'profil',t:TN,c:{}};
const cs=(id,def)=>S.c[id]||(S.c[id]={...def});
function delta(cur,prev,o={}){const {up=true,pp=false,d=1,ref='önceki aya göre',neutral=false}=o;
  if(prev==null||!isFinite(prev)||!isFinite(cur)||(!pp&&!prev))return `<span class="delta flat">–</span> <span class="ref">${ref}</span>`;
  const v=pp?cur-prev:(cur-prev)/Math.abs(prev)*100,flat=Math.abs(v)<(pp?.5*10**-d:.05),good=(v>0)===up;
  return `<span class="delta ${flat||neutral?'flat':good?'good':'bad'}">${flat?'■':v>0?'▲':'▼'} ${pp?num(Math.abs(v),d)+' puan':'%'+num(Math.abs(v),1)}</span> <span class="ref">${ref}</span>`}
const dd=(cur,prev,yo,o={})=>`<div>${delta(cur,prev,o)}</div><div>${delta(cur,yo,{...o,ref:'geçen yıla göre'})}</div>`;
const dk2=(cur,prev,yo)=>`<div>${delta(cur,prev,{ref:'önceki ay'})}</div><div>${delta(cur,yo,{ref:'geçen yıl'})}</div>`;
const sel=(card,prop,opts,cur,label,vis)=>`<label class="field">${vis?label:`<span class="sr">${label}</span>`}<select id="s-${card}-${prop}" data-card="${card}" data-prop="${prop}" title="${label}">${opts.map(([v,l])=>`<option value="${esc(v)}" ${String(v)===String(cur)?'selected':''}>${esc(l)}</option>`).join('')}</select></label>`;
const numIn=(card,prop,cur,label,step=1)=>`<label class="field">${label} <input type="number" id="s-${card}-${prop}" data-card="${card}" data-prop="${prop}" data-num="1" value="${cur}" min="0" step="${step}"></label>`;
function xcard(id,title,sub,controls,fn,cls='c6'){queue.push([id,fn]);const on=tableMode.has(id);
  return `<section class="card ${cls}" aria-labelledby="${id}-t"><header><div><h3 id="${id}-t">${title}</h3>${sub?`<p>${sub}</p>`:''}</div><div class="controls">${controls||''}<button type="button" class="tbl-btn" id="${id}-b" data-tbl="${id}" aria-pressed="${on}">${on?'Grafik':'Tablo'}</button></div></header><div class="chart" id="${id}"></div></section>`}
function heat(el,o){const {rows,cols,get,fmt,corner='',div=false,minN=0,name='Değer'}=o,cells=rows.map((_,r)=>cols.map((_,c)=>get(r,c)));
  const ok=x=>x&&isFinite(x.v)&&(x.n==null||x.n>=minN),vs=cells.flat().filter(ok).map(x=>x.v),mx=Math.max(...vs),mn=Math.min(...vs),am=Math.max(Math.abs(mx),Math.abs(mn))||1;
  const td=(x,r,c)=>{if(!ok(x))return '<td style="color:var(--muted)">–</td>';
    const q=div?Math.round(4+70*Math.abs(x.v)/am):Math.round(6+70*(mx>mn?(x.v-mn)/(mx-mn):.5)),col=div?((x.v>=0)!==!!o.inv?'var(--good)':'var(--neg)'):'var(--s1)';
    return `<td tabindex="0" style="background:color-mix(in srgb,${col} ${q}%,var(--surface));color:${q>52?'#fff':'var(--ink)'}" data-tip="${esc(tipH(rows[r]+' · '+cols[c],[[null,name,fmt(x.v)]].concat(x.n!=null?[[null,'Abone',f0(x.n)]]:[],x.extra||[])))}">${fmt(x.v)}</td>`};
  el.innerHTML=`<div class="tblwrap"><table class="heat"><thead><tr><th scope="col">${esc(corner)}</th>${cols.map(c=>`<th scope="col" style="text-align:center">${esc(c)}</th>`).join('')}</tr></thead><tbody>${rows.map((rw,r)=>`<tr><th scope="row">${esc(rw)}</th>${cells[r].map((x,c)=>td(x,r,c)).join('')}</tr>`).join('')}</tbody></table></div>`;
  return {head:[corner,...cols],rows:rows.map((rw,r)=>[rw,...cells[r].map(x=>ok(x)?fmt(x.v):'–')])}}
const empty=t=>`<section class="card"><p style="color:var(--ink2)">${t}</p></section>`;

/* ================= Sayfa 1 — Müşteri Profili ================= */
function pProfil(){
  const pT=MONTHS.indexOf(addM(MONTHS[T],-1)),m=tot(T),p=tot(pT),yi=yoyT(T),y=tot(yi);
  const G=[['Aktif müşteri',f0,'n',{}],['Toplam gelir',fMn,'rev',{}],['Ort. gelir',tl,'fee',{}],['Ort. ARPU',tl,'arpu',{}],['Ort. data',gbf,'gb',{}],['Ort. ses',dk,'mou',{}],['Ort. tenure',ay,'ten',{}],['Ort. yaş',f1,'age',{neutral:true}]];
  const cd=(x,q,o={})=>{if(q==null||!isFinite(q)||!isFinite(x)||(!o.pp&&!q))return '<span class="delta flat">–</span>';const v=o.pp?x-q:(x-q)/Math.abs(q)*100,fl=Math.abs(v)<.05,good=(v>0)===(o.up!==false);return `<span class="delta ${fl||o.neutral?'flat':good?'good':'bad'}">${fl?'■':v>0?'▲':'▼'}${f1(Math.abs(v))}</span>`};
  const ct=(l,val,x,q,w,o={})=>`<div class="tile ct"><div class="lbl">${l}</div><div class="val">${/NaN|Infinity/.test(val)?'–':val}</div><div class="dl">${cd(x,q,o)} ay · ${cd(x,w,o)} yıl</div></div>`;
  const gth='<thead><tr><th scope="col"></th><th scope="col">Adet</th><th scope="col">% ay · yıl</th><th scope="col">ARPU</th><th scope="col">% ay · yıl</th></tr></thead>';
  const pkList=()=>{const P=D.pk;if(!P||!P.labels.length||!P.n[T].some(v=>v))return grp('tar','Tarife seçenekleri');
    const n=P.n[T],s=P.arpu[T],np=pT>=0?P.n[pT]:null,sp=pT>=0?P.arpu[pT]:null,ny=yi>=0?P.n[yi]:null,sy=yi>=0?P.arpu[yi]:null,av=(a,q,i)=>a&&q&&q[i]?a[i]/q[i]:null;
    return `<section class="card gl"><h3>Tarife seçenekleri · ${esc(P.desen)}</h3><table class="gt">${gth}<tbody>${P.labels.map((l,i)=>n[i]?`<tr tabindex="0" data-tip="${esc(tipH(l,[[null,'Adet',f0(n[i])],[null,'Pay',pct(n[i]/m.n*100)],[null,'Ort. ARPU',tl(s[i]/n[i])]].concat(np?[[null,'Önceki ay adet',f0(np[i])]]:[],ny?[[null,'Geçen yıl adet',f0(ny[i])]]:[])))}"><th scope="row" title="${esc(l)}">${esc(l.replace(P.desen,'').trim()||l)}</th><td><b>${f0(n[i])}</b></td><td class="d">${cd(n[i],np&&np[i])} · ${cd(n[i],ny&&ny[i])}</td><td><b>${f0(s[i]/n[i])}</b></td><td class="d">${cd(s[i]/n[i],av(sp,np,i))} · ${cd(s[i]/n[i],av(sy,ny,i))}</td></tr>`:'').join('')}</tbody></table></section>`};
  const bcCard=()=>{if(!has('bc'))return'';const a=one(T,'bc'),b=one(pT,'bc'),c=one(yi,'bc');
    return `<section class="card gl" style="margin:-6px 0 16px"><h3 style="margin-bottom:6px">Fatura dönemi (BC) · abone adedi</h3><div class="minis">${dim('bc').labels.map((l,i)=>a[i].n?`<div class="mini" tabindex="0" data-tip="${esc(tipH(l,[[null,'Adet',f0(a[i].n)],[null,'Pay',pct(a[i].n/m.n*100)],[null,'Ort. gelir',isFinite(a[i].fee)?tl(a[i].fee):'–'],[null,'Ort. ARPU',tl(a[i].arpu)]]))}"><div class="nm">${esc(l)}</div><div class="ln"><b>${f0(a[i].n)}</b><span>${pct(a[i].n/m.n*100)}</span></div><div class="ln"><span>ay · yıl</span><span>${cd(a[i].n,b&&b[i].n)} · ${cd(a[i].n,c&&c[i].n)}</span></div></div>`:'').join('')}</div></section>`};
  let gen='';if(has('gen')){const sh=t=>{const o=one(t,'gen');if(!o)return null;const k=dim('gen').labels.indexOf('Kadın'),e=dim('gen').labels.indexOf('Erkek'),N=o.reduce((s,v)=>s+v.n,0);return [k<0?NaN:o[k].n/N*100,e<0?NaN:o[e].n/N*100]},a=sh(T),b=sh(pT),c=sh(yi);
    gen=ct('Kadın / Erkek',`<span style="font-size:14.5px">%${f1(a[0])} / %${f1(a[1])}</span>`,a[0],b&&b[0],c&&c[0],{pp:true,neutral:true})}
  const grp=(k,title)=>{if(!has(k))return'';const a=one(T,k),b=one(pT,k),c=one(yi,k);
    return `<section class="card gl"><h3>${title}</h3><table class="gt">${gth}<tbody>${dim(k).labels.map((l,i)=>a[i].n?`<tr tabindex="0" data-tip="${esc(tipH(l,[[null,'Adet',f0(a[i].n)],[null,'Pay',pct(a[i].n/m.n*100)],[null,'Ort. ARPU',tl(a[i].arpu)]].concat(b?[[null,'Önceki ay adet',f0(b[i].n)],[null,'Önceki ay ARPU',tl(b[i].arpu)]]:[],c?[[null,'Geçen yıl adet',f0(c[i].n)],[null,'Geçen yıl ARPU',tl(c[i].arpu)]]:[])))}"><th scope="row">${esc(l)}</th><td><b>${f0(a[i].n)}</b></td><td class="d">${cd(a[i].n,b&&b[i].n)} · ${cd(a[i].n,c&&c[i].n)}</td><td><b>${f0(a[i].arpu)}</b></td><td class="d">${cd(a[i].arpu,b&&b[i].arpu)} · ${cd(a[i].arpu,c&&c[i].arpu)}</td></tr>`:'').join('')}</tbody></table></section>`};
  const numCard=(id,title,dk_,mk)=>{if(!has(dk_))return'';const st=cs(id,{b:''}),[ml,mf]=MS[mk];
    return xcard(id,title,st.b?`${ml} · ${dim(st.b).name} kırılımında`:`${dim(dk_).name} dağılımı · abone adedi`,sel(id,'b',[['','Kırılım yok']].concat(dimOpts([dk_])),st.b,'Kırılım'),el=>{
      if(!st.b){const o=one(T,dk_);return colChart(el,{labels:dim(dk_).labels.map(shortLab),series:[{name:'Abone adedi',color:C(1),values:o.map(v=>v.n)}],xName:dim(dk_).name,h:230,tipTitle:i=>dim(dk_).labels[i]+' · '+pct(o[i].n/m.n*100)})}
      const o=one(T,st.b);return hbars(el,{items:dim(st.b).labels.map((l,i)=>({label:l,value:o[i][mk],extra:[[null,'Abone',f0(o[i].n)]]})).filter(d=>isFinite(d.value)),fmt:mf,name:ml,xName:dim(st.b).name})})};
  const catCard=(id,title,k)=>{if(!has(k))return'';const st=cs(id,{m:'n',b:''}),[ml,mf,mh]=MS[st.m];
    return xcard(id,title,st.b?`${ml} · ${dim(k).name} × ${dim(st.b).name}`:`${ml} · ${dim(k).name} bazında`,sel(id,'m',msOpts,st.m,'Ölçü')+sel(id,'b',[['','Kırılım yok']].concat(dimOpts([k])),st.b,'Kırılım'),el=>{
      if(!st.b){const o=one(T,k);return hbars(el,{items:dim(k).labels.map((l,i)=>({label:l,value:o[i][st.m],n:o[i].n,extra:[[null,'Abone',f0(o[i].n)],[null,'Pay',pct(o[i].n/m.n*100)]]})).filter(d=>d.n&&isFinite(d.value)),fmt:mf,name:ml,xName:dim(k).name})}
      const g=cube(T,[k,st.b]);return heat(el,{rows:dim(k).labels,cols:dim(st.b).labels,corner:dim(k).name+' \\ '+dim(st.b).name,name:ml,fmt:mh,minN:st.m==='n'?0:5,get:(r,c)=>{const x=g([r,c]);return {v:x[st.m],n:x.n}}})})};
  const NI=D.srch_labels.length;let srch='';
  if(NI&&!D.srch_tot[T].length)srch=plain('Arama ilgi alanları','',`<p style="color:var(--ink2)">${mLong(MONTHS[T])} dosyasında arama (*_SEARCH) kolonları yok.</p>`,'c6');
  else if(NI){const st=cs('srch',{m:'pen',b:''}),pen=st.m==='pen',ml=pen?'Arama yapan abone oranı':'Abone başı ort. arama',mf=pen?pct:f1,a=D.srch_tot[T];
    const ord=arr(NI,i=>i).sort((q,w)=>a[w]-a[q]);
    srch=xcard('srch','Arama ilgi alanları',st.b?`${ml} · ${dim(st.b).name} kırılımında`:ml,sel('srch','m',[['pen','Yaygınlık (%)'],['avg','Ort. arama adedi']],st.m,'Ölçü')+sel('srch','b',[['','Kırılım yok']].concat(dimOpts()),st.b,'Kırılım'),el=>{
      if(!st.b)return hbars(el,{items:ord.map(i=>({label:D.srch_labels[i],value:pen?a[i]/m.n*100:a[NI+i]/m.n,extra:[[null,'Arama yapan abone',f0(a[i])]]})),fmt:mf,name:ml,xName:'İlgi alanı'});
      const o=one(T,st.b),src=D.srch_one[T][st.b];return heat(el,{rows:ord.map(i=>D.srch_labels[i]),cols:dim(st.b).labels,corner:'İlgi alanı \\ '+dim(st.b).name,name:ml,fmt:pen?f1:f1,minN:5,get:(r,c)=>({v:src[c*2*NI+(pen?0:NI)+ord[r]]/o[c].n*(pen?100:1),n:o[c].n})})},st.b?'c12':'c6')}
  const dk0=D.dims.map(d=>d.key),mx=cs('mx',{r:has('gen')?'gen':dk0[0],c:has('dev')?'dev':dk0[1],f:'',fv:0,m:'arpu'});
  const triOk=!!(D.tri[T]&&Object.keys(D.tri[T]).length);if(!triOk)mx.f='';if(mx.c===mx.r)mx.c=dk0.find(k=>k!==mx.r);if(mx.f===mx.r||mx.f===mx.c)mx.f='';if(mx.f&&+mx.fv>=dim(mx.f).labels.length)mx.fv=0;
  const multi=xcard('mx','Çok değişkenli kırılım',`${MS[mx.m][0]} · ${dim(mx.r).name} × ${dim(mx.c).name}${mx.f?` · ${dim(mx.f).name}: ${dim(mx.f).labels[+mx.fv]}`:''}${triOk?'':' · filtre yalnızca son dönemlerde kullanılabilir'}`,
    sel('mx','r',dimOpts([mx.c]),mx.r,'Satır')+sel('mx','c',dimOpts([mx.r]),mx.c,'Sütun')+(triOk?sel('mx','f',[['','Filtre yok']].concat(dimOpts([mx.r,mx.c])),mx.f,'Filtre'):'')+(mx.f?sel('mx','fv',dim(mx.f).labels.map((l,i)=>[i,l]),mx.fv,'Filtre değeri'):'')+sel('mx','m',msOpts,mx.m,'Ölçü'),
    el=>{const g=cube(T,[mx.r,mx.c].concat(mx.f?[mx.f]:[]));return heat(el,{rows:dim(mx.r).labels,cols:dim(mx.c).labels,corner:dim(mx.r).name+' \\ '+dim(mx.c).name,name:MS[mx.m][0],fmt:MS[mx.m][2],minN:mx.m==='n'?0:5,get:(r,c)=>{const x=g(mx.f?[r,c,+mx.fv]:[r,c]);return {v:x[mx.m],n:x.n}}})},'c12');
  return head('Müşteri Profili',`${mLong(MONTHS[T])} dönemi müşteri portföyü.`)+
  `<div class="tiles g8">${G.map(([l,f,k,o])=>ct(l,f(m[k]),m[k],p&&p[k],y&&y[k],o)).join('')}${gen}${ct('Şikayet oranı',pct(m.comp),m.comp,p&&p.comp,y&&y.comp,{up:false,pp:true})}</div>
  <p class="ref" style="margin:-8px 0 8px">Kutulardaki oklar: önceki aya (ay) ve geçen yıla (yıl) göre değişim, % · oranlarda puan</p><div class="g3">${grp('rah','Rahat segmentleri')}${grp('dig','Dijital segmentler')}${pkList()}</div>${bcCard()}
  <h3 class="sec">Dağılımlar ve kırılımlar</h3>
  <div class="grid">
   ${numCard('c_fee','Gelir (TOTAL_FEE_TL)','fb','fee')}${numCard('c_arpu','ARPU','ab','arpu')}${numCard('c_gb','Data (GB) kullanımı','gbb','gb')}${catCard('c_tar','Tarife türü','tar')}${numCard('c_mou','Ses kullanımı','mb','mou')}
   ${numCard('c_ten','Tenure','ten','ten')}${numCard('c_age','Yaş dağılımı','age','age')}${catCard('c_gen','Cinsiyet','gen')}${catCard('c_city','Şehir','city')}
   ${catCard('c_dev','Cihaz markası','dev')}${catCard('c_dig','Dijital segment','dig')}${catCard('c_rah','Rahat segment','rah')}${catCard('c_bc','Fatura dönemi (BC)','bc')}${catCard('c_bon','Ek paket','bon')}${srch}${multi}
  </div>`;
}

/* ================= Sayfa 2 — Aylık Karşılaştırma ================= */
function pKars(){
  if(TN<1)return head('Aylık Karşılaştırma','')+empty('Karşılaştırma için veri klasöründe en az iki ay bulunmalı.');
  const st=cs('kx',{b:T-1,d:has('rah')?'rah':D.dims[0].key,d2:has('dig')?'dig':D.dims[1].key,mm:'n',metric:'arpu',mode:'adet',thA:100,thP:3,minN:300});
  if(st.d2===st.d)st.d2=D.dims.find(d=>d.key!==st.d).key;
  if(+st.b===T||+st.b<0||+st.b>TN)st.b=T>0?T-1:1;
  const Bi=+st.b,a=tot(T),b=tot(Bi),A=mLong(MONTHS[T]),B=mLong(MONTHS[Bi]),oa=one(T,st.d),ob=one(Bi,st.d),L=dim(st.d).labels,rel=(x,y)=>y?(x-y)/Math.abs(y)*100:NaN;
  const GM=[['n','Aktif müşteri',f0,{}],['rev','Toplam gelir',fMn,{}],['fee','Ort. gelir (abone başı)',tl,{}],['arpu','Ortalama ARPU',tl,{}],['price','Ort. tarife ücreti',tl,{}],['gb','Ort. data kullanımı',gbf,{}],['mou','Ort. ses kullanımı',dk,{}],['ten','Ort. tenure',ay,{}],['age','Ortalama yaş',f1,{neutral:true}],['comp','Şikayet oranı',pct,{up:false,pp:true,d:2}]].concat(isFinite(a.bfee)||isFinite(a.nbfee)?[['bfee','Ort. paket içi ücret',tl,{}],['nbfee','Ort. paket dışı ücret',tl,{}]]:[]);
  const gm=GM.find(g=>g[0]===st.metric)||GM[2],sh=o=>{const N=o.reduce((s,v)=>s+v.n,0);return o.map(v=>v.n/N*100)},sa=sh(oa),sb=sh(ob);
  const HC=[['n','Adet'],['fee','Gelir'],['arpu','ARPU'],['gb','Data'],['mou','Ses'],['ten','Tenure'],['age','Yaş']],ga=cube(T,[st.d,st.d2]),gb_=cube(Bi,[st.d,st.d2]);
  // eşik üstü değişim listesi: önce mod (adet / oran) seçilir, sonra o modun eşiği uygulanır
  const sgnN=v=>(v>0?'+':v<0?'−':'')+f0(Math.abs(v)),adet=st.mode==='adet',items=[];
  D.dims.forEach(dm=>{const xa=one(T,dm.key),xb=one(Bi,dm.key);dm.labels.forEach((l,i)=>{if(Math.min(xa[i].n,xb[i].n)<+st.minN)return;
    (adet?[['n','Adet',f0]]:[['n','Adet',f0],['fee','Ort. gelir',tl],['arpu','Ort. ARPU',tl],['gb','Ort. data',gbf],['mou','Ort. ses',dk]]).forEach(([k,kl,f])=>{const x=xa[i][k],q=xb[i][k],r=rel(x,q),df=x-q;
      if(isFinite(r)&&(adet?Math.abs(df)>=+st.thA:Math.abs(r)>=+st.thP))items.push({dm:dm.name,l,kl,f,a:x,b:q,r,df,sc:adet?Math.abs(df):Math.abs(r)})})})});
  items.sort((q,w)=>w.sc-q.sc);
  const tg=(v,l)=>`<button type="button" class="tbl-btn" id="kx-mode-${v}" data-card="kx" data-prop="mode" data-val="${v}" aria-pressed="${st.mode===v}">${l}</button>`;
  const NIs=D.srch_labels.length,dcard=(id,k)=>{if(!has(k))return'';const xa=one(T,k),xb=one(Bi,k),pa=sh(xa),pb=sh(xb),Lk=dim(k).labels;
    return card(id,`${dim(k).name} — pay ve adet değişimi`,`${B} → ${A} · çubuk: pay değişimi (puan) · yanında adet değişimi`,el=>deltaBars(el,{items:Lk.map((l,i)=>({label:l,value:pa[i]-pb[i],note:sgnN(xa[i].n-xb[i].n)+' abone',extra:[[null,B,f0(xb[i].n)+' abone · '+pct(pb[i])],[null,A,f0(xa[i].n)+' abone · '+pct(pa[i])],[null,'Adet değişimi',sgnN(xa[i].n-xb[i].n)]]})),
      tbl:{head:[dim(k).name,B+' adet',B+' pay',A+' adet',A+' pay','Adet değişimi','Pay değişimi'],rows:Lk.map((l,i)=>[l,f0(xb[i].n),pct(pb[i]),f0(xa[i].n),pct(pa[i]),sgnN(xa[i].n-xb[i].n),(pa[i]-pb[i]>0?'+':pa[i]-pb[i]<0?'−':'')+f2(Math.abs(pa[i]-pb[i]))+' puan'])}}))};
  const scard=()=>{if(!NIs||!D.srch_tot[T].length||!D.srch_tot[Bi].length)return'';const ua=D.srch_tot[T],ub=D.srch_tot[Bi],it=D.srch_labels.map((l,i)=>{const x=ua[i]/a.n*100,q=ub[i]/b.n*100;return {label:l,value:x-q,ua:ua[i],ub:ub[i],x,q,note:sgnN(ua[i]-ub[i])+' abone',extra:[[null,B,f0(ub[i])+' abone · '+pct(q)],[null,A,f0(ua[i])+' abone · '+pct(x)],[null,'Adet değişimi',sgnN(ua[i]-ub[i])]]}}).sort((q,w)=>Math.abs(w.value)-Math.abs(q.value)).slice(0,12).sort((q,w)=>w.value-q.value);
    return card('kd_s','Arama ilgi alanları — en çok değişenler',`${B} → ${A} · arama yapan abone oranı (puan) · yanında adet değişimi`,el=>deltaBars(el,{items:it,tbl:{head:['İlgi alanı',B+' adet',B+' oran',A+' adet',A+' oran','Adet değişimi','Oran değişimi'],rows:it.map(d=>[d.label,f0(d.ub),pct(d.q),f0(d.ua),pct(d.x),sgnN(d.ua-d.ub),(d.value>0?'+':d.value<0?'−':'')+f2(Math.abs(d.value))+' puan'])}}))};
  return head('Aylık Karşılaştırma',`<b>${A}</b> dönemi, <b>${B}</b> dönemiyle karşılaştırılıyor. Değişimler ${B} → ${A} yönündedir.`,`<div class="controls">${sel('kx','b',MONTHS.map((id,i)=>[i,mLong(id)]).filter(x=>x[0]!==T),st.b,'Karşılaştırma ayı',1)}</div>`)+
  `<div class="grid">
   ${plain('Temel göstergeler',`${B} ve ${A} değerleri`,`<div class="tblwrap"><table><thead><tr><th scope="col">Gösterge</th><th scope="col">${B}</th><th scope="col">${A}</th><th scope="col">Değişim</th></tr></thead><tbody>${GM.map(([k,l,f,o])=>`<tr><th scope="row">${l}</th><td>${isFinite(b[k])?f(b[k]):'–'}</td><td><b>${isFinite(a[k])?f(a[k]):'–'}</b></td><td>${delta(a[k],b[k],{...o,ref:''})}</td></tr>`).join('')}</tbody></table></div>`,'c5')}
   ${xcard('k0','Aylık seyir',`${gm[1]} · tüm dönemler`,sel('kx','metric',GM.map(g=>[g[0],g[1]]),gm[0],'Gösterge'),el=>lineChart(el,{labels:LBL,mark:T,fmt:gm[2],series:[{name:gm[1],color:C(1),values:MONTHS.map((_,t)=>tot(t)[gm[0]])}]}),'c7')}
   ${['rah','dig','tar','age','ten','dev','city','bon'].map(k=>dcard('kd_'+k,k)).join('')}${scard()}
   ${xcard('k3',`${dim(st.d).name} × gösterge değişimi`,`${B} → ${A} · yüzde değişim (yeşil artış, kırmızı azalış)`,sel('kx','d',dimOpts(),st.d,'Kırılım'),el=>heat(el,{rows:L,cols:HC.map(h=>h[1]),corner:dim(st.d).name,div:true,name:'Değişim',fmt:sgn,get:(r,c)=>{const k=HC[c][0];return {v:rel(oa[r][k],ob[r][k]),n:Math.min(oa[r].n,ob[r].n),extra:[[null,B,MS[k][1](ob[r][k])],[null,A,MS[k][1](oa[r][k])],[null,'Abone ('+B+' → '+A+')',f0(ob[r].n)+' → '+f0(oa[r].n)]]}}}),'c12')}
   ${xcard('k4','İki kırılımlı değişim matrisi',`${MS[st.mm][0]} · ${dim(st.d).name} × ${dim(st.d2).name} · ${B} → ${A} yüzde değişim`,sel('kx','d2',dimOpts([st.d]),st.d2,'İkinci kırılım')+sel('kx','mm',msOpts,st.mm,'Ölçü'),el=>heat(el,{rows:L,cols:dim(st.d2).labels,corner:dim(st.d).name+' \\ '+dim(st.d2).name,div:true,name:'Değişim',fmt:sgn,minN:20,get:(r,c)=>{const x=ga([r,c]),y=gb_([r,c]);return {v:rel(x[st.mm],y[st.mm]),n:Math.min(x.n,y.n),extra:[[null,B,MS[st.mm][1](y[st.mm])],[null,A,MS[st.mm][1](x[st.mm])]]}}}),'c12')}
   <section class="card c12" aria-labelledby="k5-t"><header><div><h3 id="k5-t">Eşik üstü değişimler</h3><p>${B} → ${A} · ${adet?'abone adedi en az '+f0(+st.thA)+' değişen gruplar':'adedi veya ortalaması en az %'+f1(+st.thP)+' değişen gruplar'} · ${items.length} kayıt</p></div>
     <div class="controls">${tg('adet','Adete göre')}${tg('oran','Orana göre')}${adet?numIn('kx','thA',st.thA,'En az adet değişimi',10):numIn('kx','thP',st.thP,'En az değişim %',.5)}${numIn('kx','minN',st.minN,'En az abone',50)}</div></header>
     ${items.length?`<div class="tblwrap" style="max-height:520px;overflow-y:auto"><table><thead><tr><th scope="col">Kırılım</th><th scope="col" style="text-align:left">Grup</th><th scope="col" style="text-align:left">Gösterge</th><th scope="col">${B}</th><th scope="col">${A}</th><th scope="col">Fark</th><th scope="col">Değişim</th></tr></thead><tbody>${items.slice(0,80).map(x=>`<tr><th scope="row">${esc(x.dm)}</th><td style="text-align:left">${esc(x.l)}</td><td style="text-align:left">${x.kl}</td><td>${x.f(x.b)}</td><td><b>${x.f(x.a)}</b></td><td>${(x.df>0?'+':x.df<0?'−':'')+x.f(Math.abs(x.df))}</td><td>${delta(x.a,x.b,{ref:''})}</td></tr>`).join('')}</tbody></table></div>`:'<p style="color:var(--ink2)">Bu eşiklerle listelenecek değişim yok.</p>'}</section>
  </div>`;
}

/* ================= Sayfa 3 — Churn & Kazanım ================= */
const EVOK=q=>q>=1&&(!D.ev_ok||D.ev_ok[q]),INS=D.ev.slice(0,5),OUTS=D.ev.slice(5),CH=['out_rival','out_close','out_ghost'],evL=e=>D.evl[D.ev.indexOf(e)];
const evVec=(t,es)=>{const v=new Array(W2).fill(0);es.forEach(e=>{const x=D.ev_tot[t]&&D.ev_tot[t][e];if(x)x.forEach((q,i)=>v[i]+=q)});return v};
const evS=(t,es)=>st2(evVec(t,es));
const evOne=(t,es,k)=>{const n=dim(k).labels.length,v=new Array(n*W2).fill(0);es.forEach(e=>{const x=D.ev_one[t]&&D.ev_one[t][e]&&D.ev_one[t][e][k];if(x)x.forEach((q,i)=>v[i]+=q)});return arr(n,c=>st2(v.slice(c*W2,(c+1)*W2)))};
const churnRate=t=>t>=1?evS(t,CH).n/tot(t-1).n*100:NaN;
const EVG=[['in','Portföye girenler',INS],['out','Portföyden çıkanlar',OUTS],['churn','Churn (rakip + kapatan + dosyadan çıkan)',CH]].concat(D.ev.map((e,i)=>[e,D.evl[i],[e]]));
const addM=(id,k)=>{const m=+id.slice(0,4)*12+(+id.slice(4)-1)+k;return Math.floor(m/12)+String(m%12+1).padStart(2,'0')},difM=(a,b)=>(+a.slice(0,4)-+b.slice(0,4))*12+(+a.slice(4)-+b.slice(4));
const LC=D.lc,lcSum=(yon,i)=>LC.tipler.reduce((q,t,k)=>q+(t.yon===yon?LC.v[k][i]:0),0),lcVal=(yon,hat,i)=>LC.tipler.reduce((q,t,k)=>q+(t.yon===yon&&t.hat===hat?LC.v[k][i]:0),0);
function flow(el,i){
  const W=Math.max(560,el.clientWidth||600),H=150,nw=124,nh=54,cy=H/2,xs=[nw/2+8,W/2,W-nw/2-8],nm=['Faturasız','Hybrid','Faturalı'],sub=['Prepaid','Hibrit · Rahat','Postpaid'];
  const F=[['in','pre',0,1,-12,1,'Faturasızdan Hybrid\'e'],['out','pre',1,0,12,2,'Hybrid\'den faturasıza'],['in','post',2,1,-12,1,'Faturalıdan Hybrid\'e'],['out','post',1,2,12,2,'Hybrid\'den faturalıya']].map(f=>[...f,lcVal(f[0],f[1],i)]),mx=Math.max(1,...F.map(f=>f[7])),e=nw/2+6;
  const mk=(id,c)=>`<marker id="${id}" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse" markerUnits="userSpaceOnUse"><path d="M0,0L10,5L0,10z" fill="${c}"/></marker>`;
  el.innerHTML=`<ul class="legend"><li><i class="sw" style="background:var(--s1)"></i>Hybrid'e giriş</li><li><i class="sw" style="background:var(--s2)"></i>Hybrid'den çıkış</li></ul><div class="scroll-x"><svg role="img" aria-label="Faturasız, Hybrid ve Faturalı hatlar arasındaki geçiş adetleri" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="min-width:${W}px"><defs>${mk('mk-a','var(--s1)')}${mk('mk-b','var(--s2)')}</defs>
   ${F.map(([,,a,b,dy,c,lab,v])=>{const x1=xs[a]+(b>a?e:-e),x2=xs[b]+(b>a?-e-4:e+4),yy=cy+dy;return `<g class="hit" data-tip="${esc(tipH(lab,[[C(c),'Adet',f0(v)]]))}"><line x1="${x1}" x2="${x2}" y1="${yy}" y2="${yy}" stroke="${C(c)}" stroke-width="${2+7*v/mx}" marker-end="url(#mk-${c===1?'a':'b'})"/><rect class="h" x="${Math.min(x1,x2)}" y="${yy-16}" width="${Math.abs(x2-x1)}" height="32"/><text class="strong" x="${(x1+x2)/2}" y="${yy+(dy<0?-11:21)}" text-anchor="middle" style="font-size:13px">${f0(v)}</text></g>`}).join('')}
   ${xs.map((x,k)=>`<rect x="${x-nw/2}" y="${cy-nh/2}" width="${nw}" height="${nh}" rx="12" fill="${k===1?'var(--accent-wash)':'var(--surface2)'}" stroke="${k===1?'var(--accent)':'var(--border)'}"/><text class="strong" x="${x}" y="${cy-2}" text-anchor="middle" style="font-size:14px">${nm[k]}</text><text x="${x}" y="${cy+15}" text-anchor="middle">${sub[k]}</text>`).join('')}
  </svg></div>`;
  return {head:['Geçiş','Adet'],rows:F.map(f=>[f[6],f0(f[7])])};
}
/* ---- Lifecycle: Hybrid'e giriş ve Hybrid'den geçişler ---- */
function lcSection(){
  if(!LC)return `<h3 class="sec">Hybrid giriş ve çıkışları (lifecycle)</h3>`+empty('Lifecycle tablosu bulunamadı: <code>data/lifecycle</code> klasörüne yearmonth, lifecycle_type ve counts kolonlu tabloyu koyun.');
  const cur=MONTHS[T],i=LC.aylar.indexOf(cur),ip=LC.aylar.indexOf(addM(cur,-1)),lcLab=x=>/^\d{6}$/.test(x)?mShort(x):x,lcS=(yon,sn,q)=>LC.tipler.reduce((w,t_,k)=>w+(t_.yon===yon&&t_.sinif===sn?LC.v[k][q]:0),0),lab=LC.aylar.map(lcLab),ser=yon=>LC.tipler.map((t,k)=>[t,k]).filter(x=>x[0].yon===yon).map(([t,k],q)=>({name:t.ad,color:C(q+1),values:LC.v[k]}));
  const pv=(f,o)=>ip>=0?delta(f(i),f(ip),o):'',net=q=>lcSum('in',q)-lcSum('out',q),gone=EVOK(T)&&D.hareket_ok!==false?evS(T,OUTS).n:NaN,est=gone-lcSum('out',i);
  return `<h3 class="sec">Hybrid giriş ve çıkışları (lifecycle)</h3>`+(i<0?`<div class="ins" style="margin-bottom:12px">${mLong(cur)} için lifecycle verisi yok; grafikler tablodaki tüm ayları gösterir.</div>`:
   `<div class="tiles">${tile('Yeni müşteri kazanımı',f0(lcS('in','yeni',i)),pv(q=>lcS('in','yeni',q)),'Yeni Tesis Hibrit + Port In Hibrit')}${tile('Diğer hat tiplerinden gelen',f0(lcS('in','gecis',i)),pv(q=>lcS('in','gecis',q)),'Prepaid ve Postpaid\'den Hybrid\'e geçenler')}${tile('Diğer hat tiplerine giden',f0(lcSum('out',i)),pv(q=>lcSum('out',q),{up:false}),'Hybrid\'den Prepaid ve Postpaid\'e geçenler')}${tile('Net (gelen − giden)',(net(i)>0?'+':'')+f0(net(i)),pv(net))}${isFinite(est)?tile('Tahmini churn',f0(Math.max(0,est)),'',`Portföyden çıkan ${f0(gone)} − Hybrid'den geçiş ${f0(lcSum('out',i))}`):''}</div>`)+
  `<div class="grid">
   ${card('lc1','Hybrid\'e nasıl geldiler?','Aylık giriş adetleri · yeni kazanım (Yeni Tesis, Port In) ve diğer hat tiplerinden geçiş',el=>colChart(el,{labels:lab,mark:i<0?null:i,stacked:true,series:ser('in')}))}
   ${card('lc2','Hybrid\'den nereye geçtiler?','Aylık geçiş adetleri · lifecycle tipine göre',el=>colChart(el,{labels:lab,mark:i<0?null:i,stacked:true,series:ser('out')}))}
   ${i>=0?card('lc3','Faturasız ↔ Hybrid ↔ Faturalı geçişleri',`${mLong(cur)} · çizgi kalınlığı geçiş adediyle orantılıdır`,el=>flow(el,i),'c6'):''}
   ${card('lc4','Net Hybrid hareketi','Giriş − geçiş (lifecycle)',el=>colChart(el,{labels:lab,mark:i<0?null:i,series:[{name:'Net',color:'var(--good)',neg:'var(--neg)',values:LC.aylar.map((_,q)=>net(q))}],h:i>=0?190:230}),i>=0?'c6':'c12')}
   ${plain('Lifecycle tablosu','Hybrid ile ilgili tipler · aylara göre adet',`<div class="tblwrap"><table><thead><tr><th scope="col">Lifecycle tipi</th>${LC.aylar.map((m,q)=>`<th scope="col"${q===i?' class="cur"':''}>${lcLab(m)}</th>`).join('')}</tr></thead><tbody>${[['in','yeni','Yeni müşteri kazanımı'],['in','gecis','Diğer hat tiplerinden gelen'],['out','gecis','Diğer hat tiplerine giden']].filter(g=>LC.tipler.some(t=>t.yon===g[0]&&t.sinif===g[1])).map(([yon,sn,gl])=>`<tr class="grp"><th scope="row" colspan="${LC.aylar.length+1}">${gl}</th></tr>`+LC.tipler.map((t,k)=>t.yon===yon&&t.sinif===sn?`<tr><th scope="row">${esc(t.ad)}</th>${LC.v[k].map((v,q)=>`<td${q===i?' class="cur"':''}>${f0(v)}</td>`).join('')}</tr>`:'').join('')+`<tr class="em"><th scope="row">Toplam</th>${LC.aylar.map((_,q)=>`<td${q===i?' class="cur"':''}><b>${f0(lcS(yon,sn,q))}</b></td>`).join('')}</tr>`).join('')}</tbody></table></div>`,'c12')}
  </div>`;
}
/* ---- Kohort tablosu (liste ayına göre port-out takibi) ---- */
const KH=D.kh;
function khSection(){
  if(!KH)return `<h3 class="sec">Arama listesi kohort takibi</h3>`+empty('Kohort tablosu bulunamadı: <code>data/kohort</code> klasörüne YEARMONTH ve YearMonth_k_…POUT kolonlu tabloyu koyun.');
  const G=KH.gruplar,hasK=G.some(g=>g.key==='kontrol'),gn=k=>(G.find(g=>g.key===k)||{name:k}).name,N=KH.aylar.length;
  const st=cs('kh',{g:(G.find(g=>g.key==='aranmis')||G[0]).key,m:'kum',v:'oran',c:'',k:'',h:''});
  const MM=[['kum','Port-out (kümülatif)'],['ay','Port-out (o ay)']].concat(hasK?[['fark','Kontrol grubuna göre fark (kümülatif)'],['farkay','Kontrol grubuna göre fark (o ay)']]:[]);
  if(!MM.some(x=>x[0]===st.m)||(st.g==='kontrol'&&st.m.startsWith('fark')))st.m='kum';
  const fark=st.m.startsWith('fark'),aylik=st.m==='ay'||st.m==='farkay',adet=st.v==='adet'&&!fark,ml=MM.find(x=>x[0]===st.m)[1],sp=v=>(v>0?'+':v<0?'−':'')+f2(Math.abs(v));
  // kümülatif ve aylık port-out dizileri (tablo aylık geliyorsa kümülatife, kümülatif geliyorsa aylığa çevrilir)
  const cumA=(g,ci)=>{let a=0;return KH.pout[g][ci].map(v=>v==null?null:KH.kum?v:(a+=v))},monA=(g,ci)=>{const c=cumA(g,ci);return c.map((v,k)=>v==null?null:k&&c[k-1]!=null?v-c[k-1]:v)};
  const base=(g,ci)=>KH.taban[g]?KH.taban[g][ci]:null,val=(g,ci,k,ay)=>{const b=base(g,ci),v=(ay?monA:cumA)(g,ci)[k];return b&&v!=null?{n:v,b,r:v/b*100}:null};
  const obs=KH.aylar.map((c,ci)=>{let k=0;G.forEach(g=>KH.pout[g.key][ci].forEach((v,q)=>{if(v!=null)k=Math.max(k,q+1)}));return k}),first=KH.aylar[0];
  const endCal=KH.aylar.reduce((mx,c,ci)=>{const e=addM(c,obs[ci]-1+KH.fark);return e>mx?e:mx},first),cal0=addM(first,KH.fark),cols=arr(Math.max(1,difM(endCal,cal0)+1),k=>addM(cal0,k));
  // seçili liste ayı ve izleme ayı
  let ci0=KH.aylar.indexOf(st.c);if(ci0<0)ci0=KH.aylar.indexOf(MONTHS[T]);if(ci0<0||!obs[ci0]){ci0=N-1;while(ci0>0&&!obs[ci0])ci0--}
  const K0=obs[ci0];let k0=+st.k-1;if(!(k0>=0&&k0<K0))k0=K0-1;
  const oran=(g,ci)=>{const b=base(g.key,ci),m={aranmis:['arama','listenin'],ulasilmis:['aranmis','arananların'],rahatfirsat:['ulasilmis','ulaşılanların']}[g.key];if(!m||!base(m[0],ci)||b==null)return '–';return pct(b/base(m[0],ci)*100)+` <span class="ref">${m[1]}</span>`};
  const kx=K0?val('kontrol',ci0,k0):null,ozet=K0?`<div class="tblwrap"><table><thead><tr><th scope="col">Grup</th><th scope="col">Adet</th><th scope="col">Oran</th><th scope="col">Port-out adedi</th><th scope="col">Port-out oranı</th>${hasK?'<th scope="col">Kontrole göre fark</th>':''}</tr></thead><tbody>${G.map(g=>{const x=val(g.key,ci0,k0);
      return `<tr${g.key==='kontrol'?' class="em"':''}><th scope="row">${esc(g.name)}</th><td><b>${base(g.key,ci0)==null?'–':f0(base(g.key,ci0))}</b></td><td>${oran(g,ci0)}</td><td>${x?f0(x.n):'–'}</td><td><b>${x?pct2(x.r):'–'}</b></td>${hasK?`<td>${x&&kx&&g.key!=='kontrol'?delta(x.r,kx.r,{pp:true,d:2,up:false,ref:''}):'–'}</td>`:''}</tr>`}).join('')}</tbody></table></div><p class="ref" style="margin-top:8px">Port-out: liste ayından sonraki ${k0+KH.fark}. aya (${mShort(addM(KH.aylar[ci0],k0+KH.fark))}) kadar kümülatif. Fark: grubun port-out oranı − kontrol grubunun port-out oranı; yeşil = kontrol grubundan düşük.</p>`
    :'<p style="color:var(--ink2)">Bu liste ayı için henüz port-out gözlemi yok.</p>';
  // tüm kohortların karşılaştırması (sabit izleme ayında)
  const Kmax=Math.max(...obs,1);let h0=+st.h-1;if(!(h0>=0&&h0<Kmax))h0=Math.min(2,Kmax-1);
  const og=G.filter(g=>g.key!=='arama'),rt=(k,ci)=>{const b=base(k,ci),m={aranmis:'arama',ulasilmis:'aranmis',rahatfirsat:'ulasilmis'}[k];return m&&b!=null&&base(m,ci)?pct(b/base(m,ci)*100):'–'};
  const genel=`<div class="tblwrap"><table><thead><tr><th scope="col">Liste ayı</th><th scope="col">Arama listesi</th>${og.map(g=>`<th scope="col">${esc(g.name)}</th>`).join('')}${og.filter(g=>g.key!=='kontrol').map(g=>`<th scope="col">${esc(g.name)} oranı</th>`).join('')}${og.map(g=>`<th scope="col">${esc(g.name)}<br><span class="ref">port-out</span></th>`).join('')}${hasK?'<th scope="col">Arananlar − kontrol</th>':''}</tr></thead><tbody>${KH.aylar.map((c,ci)=>{const xa=val('aranmis',ci,h0),xk=val('kontrol',ci,h0);
      return `<tr><th scope="row">${mShort(c)}</th><td>${base('arama',ci)==null?'–':f0(base('arama',ci))}</td>${og.map(g=>`<td>${base(g.key,ci)==null?'–':f0(base(g.key,ci))}</td>`).join('')}${og.filter(g=>g.key!=='kontrol').map(g=>`<td>${rt(g.key,ci)}</td>`).join('')}${og.map(g=>{const x=val(g.key,ci,h0);return `<td>${x?pct2(x.r):'–'}</td>`}).join('')}${hasK?`<td>${xa&&xk?delta(xa.r,xk.r,{pp:true,d:2,up:false,ref:''}):'–'}</td>`:''}</tr>`}).join('')}</tbody></table></div>`;
  return `<h3 class="sec">Arama listesi kohort takibi</h3><div class="grid">
   ${xcard('kh1','İleriye dönük kohort takibi',`${fark?gn(st.g)+' − Kontrol grubu · '+(aylik?'o ayki':'kümülatif')+' port-out oranı farkı (puan; yeşil = kontrol grubundan düşük)':gn(st.g)+' · '+ml+(adet?' · abone adedi':' · kohort içindeki oran (%)')} · satır: liste ayı · sütun: takvim ayı`,
     sel('kh','g',G.map(g=>[g.key,g.name]),st.g,'Grup')+sel('kh','m',st.g==='kontrol'?MM.slice(0,2):MM,st.m,'Ölçü')+(fark?'':sel('kh','v',[['oran','Oran (%)'],['adet','Adet']],st.v,'Gösterim')),
     el=>heat(el,{rows:KH.aylar.map((c,ci)=>`${mShort(c)} · ${base(st.g,ci)==null?'–':f0(base(st.g,ci))} abone`),cols:cols.map(mShort),corner:'Liste ayı · kohort',name:ml,div:fark,inv:true,fmt:fark?sp:adet?f0:v=>'%'+f2(v),get:(r,cx)=>{const k=difM(cols[cx],KH.aylar[r])-KH.fark;if(k<0||k>=KH.K)return null;const x=val(st.g,r,k,aylik);if(!x)return null;
       if(fark){const y=val('kontrol',r,k,aylik);if(!y)return null;return {v:x.r-y.r,n:x.b,extra:[[null,gn(st.g)+' port-out',pct2(x.r)],[null,'Kontrol grubu port-out',pct2(y.r)],[null,'Kontrol grubu adedi',f0(y.b)],[null,'Liste ayından sonra',(k+KH.fark)+'. ay']]}}
       return {v:adet?x.n:x.r,n:x.b,extra:[[null,'Adet',f0(x.n)],[null,'Oran',pct2(x.r)],[null,'Liste ayından sonra',(k+KH.fark)+'. ay']]}}}),'c12')}
   <section class="card c7" aria-labelledby="kh2-t"><header><div><h3 id="kh2-t">Seçili kohortun özeti</h3><p>${mLong(KH.aylar[ci0])} listesi · grup adetleri, oranlar ve port-out durumu</p></div><div class="controls">${sel('kh','c',KH.aylar.map(c=>[c,mLong(c)]),KH.aylar[ci0],'Liste ayı')}${K0?sel('kh','k',arr(K0,k=>[k+1,`+${k+KH.fark} ay`]),k0+1,'İzleme ayı'):''}</div></header>${ozet}</section>
   ${K0?xcard('kh3','Grupların port-out eğrisi',`${mLong(KH.aylar[ci0])} listesi · liste ayından sonraki aylarda kümülatif port-out oranı`,'',el=>lineChart(el,{labels:arr(K0,k=>`+${k+KH.fark} ay`),xName:'Liste ayından sonra',fmt:pct2,zero:true,mark:k0,tipTitle:k=>`+${k+KH.fark} ay · ${mShort(addM(KH.aylar[ci0],k+KH.fark))}`,series:G.map((g,q)=>({name:g.name,color:g.key==='kontrol'?'var(--ink2)':C(q+1),values:arr(K0,k=>{const x=val(g.key,ci0,k);return x?x.r:NaN})}))}),'c5'):''}
   <section class="card c12" aria-labelledby="kh4-t"><header><div><h3 id="kh4-t">Kohortların karşılaştırması</h3><p>Her liste ayı için grup adetleri, aranma / ulaşılma oranları ve liste ayından sonraki ${h0+KH.fark}. ayda kümülatif port-out oranı</p></div><div class="controls">${sel('kh','h',arr(Kmax,k=>[k+1,`+${k+KH.fark} ay`]),h0+1,'İzleme ayı')}</div></header>${genel}</section>
  </div>`;
}
/* ---- Kohortlar: arama listesi, arananlar, ulaşılanlar, kontrol grubu ---- */
const CG=D.coh_groups||[],NG=CG.length,NB=D.skor_dilim||20,gix=k=>CG.findIndex(g=>g.key===k),gname=k=>(CG[gix(k)]||{name:''}).name;
const cv=(a,i,g)=>({n:a[(i*NG+g)*2]||0,c:a[(i*NG+g)*2+1]||0}),rate=x=>x.n?x.c/x.n*100:NaN,coh=t=>NG&&D.coh&&D.coh[t]?D.coh[t]:null;
const fr=v=>isFinite(v)?pct2(v):'–',noOut=()=>`<p class="ref" style="margin-top:10px">${mLong(MONTHS[T])} kohortunun churn sonucu henüz yok: sonuç bir sonraki ayın dosyasından hesaplanır. Churn oranlarını görmek için üstten önceki bir dönem seçin.</p>`;
function cohTable(co){const kn=co.known,ki=gix('k'),kr=ki>=0?rate(cv(co.tot,0,ki)):NaN;
  return `<div class="tblwrap"><table><thead><tr><th scope="col">Grup</th><th scope="col">Adet</th><th scope="col">Oran</th><th scope="col">Churn adedi</th><th scope="col">Churn oranı</th>${ki>=0?'<th scope="col">Kontrol grubuna göre fark</th>':''}</tr></thead><tbody>${CG.map((g,i)=>{const x=cv(co.tot,0,i),bi=g.base?gix(g.base):-1,bs=bi>=0?cv(co.tot,0,bi):null,r=rate(x);
    return `<tr${g.key==='p'?' class="em"':''}><th scope="row">${esc(g.name)}</th><td><b>${f0(x.n)}</b></td><td>${bs&&bs.n?pct(x.n/bs.n*100)+` <span class="ref">${esc(CG[bi].name.toLowerCase())} içinde</span>`:'–'}</td><td>${kn?f0(x.c):'–'}</td><td><b>${kn?fr(r):'–'}</b></td>${ki>=0?`<td>${kn&&'laue'.includes(g.key)&&isFinite(r)&&isFinite(kr)?delta(r,kr,{pp:true,d:2,up:false,ref:''}):'–'}</td>`:''}</tr>`}).join('')}</tbody></table></div>${kn?'':noOut()}`}
const FC=D.fc||[],FCM={},FCV=['n','a_m','a_cum','u_m','u_cum','e_cum','ch_m','ch_cum'];FC.forEach(r=>FCM[r.g+'|'+r.c+'|'+r.t]=r);
function fcCard(){
  if(!FC.length)return'';
  const gs=[...new Set(FC.map(r=>r.g))],st=cs('fc',{g:gs.includes('l')?'l':gs[0],m:'ch_cum',v:'oran'});
  const both=gs.includes('l')&&gs.includes('k'),MM=[['ch_cum','Churn (kümülatif)'],['ch_m','Churn (o ay)'],['a_cum','Aranma (kümülatif)'],['a_m','Aranma (o ay)'],['u_cum','Ulaşılma (kümülatif)'],['u_m','Ulaşılma (o ay)'],['e_cum','Rahat Efsane paketi (kümülatif)']].concat(both?[['diff','Kontrol grubuna göre churn farkı']]:[]);
  if(!MM.some(x=>x[0]===st.m))st.m='ch_cum';
  const diff=st.m==='diff',gg=diff?'l':st.g,adet=st.v==='adet'&&!diff,ix=FCV.indexOf(st.m),ml=MM.find(x=>x[0]===st.m)[1];
  const vec=(g,c,t)=>{const r=FCM[g+'|'+c+'|'+t];return r?r.tot:null};
  const cohs=[...new Set(FC.filter(r=>r.g===gg).map(r=>r.c))].sort((a,b)=>a-b),c0=cohs[0];
  const cols=arr(TN-c0+1,k=>LBL[c0+k]),rows=cohs.map(c=>{const v=vec(gg,c,c);return `${LBL[c]} · ${f0(v?v[0]:0)} abone`}),sp=v=>(v>0?'+':v<0?'−':'')+f2(Math.abs(v));
  return xcard('fc1','İleriye dönük kohort takibi',`${diff?'Listedekiler − kontrol grubu · kümülatif churn oranı farkı (puan; yeşil = liste daha düşük)':gname(gg)+' · '+ml+(adet?' · abone adedi':' · kohort içindeki oran (%)')} · satır: listeye girilen ay · sütun: takvim ayı`,
    (diff?'':sel('fc','g',gs.map(g=>[g,gname(g)]),st.g,'Grup'))+sel('fc','m',MM,st.m,'Ölçü')+(diff?'':sel('fc','v',[['oran','Oran (%)'],['adet','Adet']],st.v,'Gösterim')),
    el=>heat(el,{rows,cols,corner:'Liste ayı · kohort',name:ml,div:diff,inv:true,fmt:diff?sp:adet?f0:v=>'%'+f1(v),minN:adet?0:1,get:(r,ci)=>{const c=cohs[r],t=c0+ci;if(t<c||t>TN)return null;const v=vec(gg,c,t);if(!v||!v[0])return null;
      if(diff){const k=vec('k',c,t);if(!k||!k[0])return null;const a=v[7]/v[0]*100,b=k[7]/k[0]*100;return {v:a-b,n:v[0],extra:[[null,'Listedekiler churn',pct2(a)],[null,'Kontrol grubu churn',pct2(b)],[null,'Kontrol grubu adedi',f0(k[0])]]}}
      return {v:adet?v[ix]:v[ix]/v[0]*100,n:v[0],extra:[[null,'Adet',f0(v[ix])],[null,'Oran',pct2(v[ix]/v[0]*100)]]}}}),'c12')}
function cohSection(){
  if(!NG)return plain('Arama listesi kohort analizi','Listedekiler, arananlar, ulaşılanlar ve kontrol grubu',`<p style="color:var(--ink2)">Aylık dosyalarda liste, arama, ulaşma ve kontrol grubu kolonları bulunamadı. Kolon adlarını <code>rapor_uret.py</code> içindeki <code>KOHORT_KOLON</code> ve <code>SKOR_ADAYLARI</code> ayarlarına ekleyin.</p>`,'c12');
  const co=coh(T),st=cs('cx',{d:has('rah')?'rah':D.dims[0].key,m:'rate'}),src=co.one[st.d],rt=st.m==='rate'&&co.known;
  return `<h3 class="sec c12" style="margin:8px 0 -4px">Arama listesi kohort analizi</h3>`+fcCard()+
   plain('Seçili dönemin kohort özeti',`${mLong(MONTHS[T])} portföyü · adet, oran ve bir sonraki aydaki churn durumu`,cohTable(co),'c12')+
   xcard('cx1','Kohort × kırılım',`${rt?'Churn oranı (%)':'Abone adedi'} · ${dim(st.d).name} · ${mLong(MONTHS[T])}`,sel('cx','d',dimOpts(),st.d,'Kırılım')+sel('cx','m',[['rate','Churn oranı'],['n','Abone adedi']],st.m,'Ölçü'),
     el=>heat(el,{rows:dim(st.d).labels,cols:CG.map(g=>g.name),corner:dim(st.d).name+' \\ Grup',name:rt?'Churn oranı':'Abone adedi',fmt:rt?v=>'%'+f2(v):f0,minN:rt?20:0,get:(r,c)=>{const x=cv(src,r,c);return {v:rt?rate(x):x.n,n:x.n,extra:co.known?[[null,'Churn eden',f0(x.c)]]:[]}}}),'c12')}

function pfSection(){
  if(D.hareket_ok===false)return `<h3 class="sec">Portföy hareketi (aylık dosyalardan)</h3>`+empty('Aylık dosyalardaki abone kimliği aylar arasında örtüşmüyor; bu yüzden portföye giren ve çıkan aboneler hesaplanamadı. Aylar arasında sabit kalan bir kimlik kolonu gerekir.');
  if(!EVOK(T))return `<h3 class="sec">Portföy hareketi (aylık dosyalardan)</h3>`+empty(`${mLong(MONTHS[T])} için bir önceki ayın dosyası yok; portföye giren ve çıkanlar hesaplanamadı.`);
  const st=cs('hx',{m:'fee',e:'out',d:has('rah')?'rah':D.dims[0].key,mm:'pay'}),t=T,ts=arr(TN,i=>i+1).filter(EVOK),LB=ts.map(i=>LBL[i]),mi=ts.indexOf(t);
  const I=evS(t,INS),O=evS(t,OUTS),Ip=EVOK(t-1)?evS(t-1,INS):null,Op=EVOK(t-1)?evS(t-1,OUTS):null,pf=tot(t),net=I.n-O.n,cr=O.n/tot(t-1).n*100,crp=EVOK(t-1)?evS(t-1,OUTS).n/tot(t-2).n*100:null;
  const ML={fee:['Gelir',tl],arpu:['ARPU',tl],gb:['Data kullanımı',gbf],mou:['Ses kullanımı',dk],ten:['Tenure',ay],age:['Yaş',f1]},[ml,mf]=ML[st.m];
  const EG=EVG.filter(g=>ts.some(q=>evS(q,g[2]).n>0)&&g[0]!=='churn'),eg=EG.find(g=>g[0]===st.e)||EG[0],mo=evOne(t,eg[2],st.d),po=one(t,st.d),pv=one(t-1,st.d),L=dim(st.d).labels,mN=mo.reduce((q,v)=>q+v.n,0);
  const MM=[['pay','Pay (%)'],['oran','Hareket oranı (%)']].concat(Object.entries(ML).map(([k,v])=>[k,'Ort. '+lc(v[0])]));
  const rows=[['Portföy',pf],['Portföye girenler',I],['Portföyden çıkanlar',O]].concat(D.ev.filter(e=>!['in_oth','out_ghost'].includes(e)).map(e=>[evL(e),evS(t,[e])]).filter(r=>r[1].n));
  return `<h3 class="sec">Portföy hareketi (aylık dosyalardan)</h3>
  <div class="tiles">
    ${tile('Portföye giren',f0(I.n),delta(I.n,Ip&&Ip.n),'Önceki ay dosyasında olmayan aboneler')}${tile('Portföyden çıkan',f0(O.n),delta(O.n,Op&&Op.n,{up:false}),'Bu ay dosyasında olmayan aboneler')}
    ${tile('Net değişim',(net>0?'+':'')+f0(net),delta(net,Ip&&Ip.n-Op.n))}${tile('Çıkış oranı',pct2(cr),delta(cr,crp,{up:false,pp:true,d:2}),'Çıkanlar / önceki ay portföyü')}
    ${tile('Girenlerin ort. geliri',tl(I.fee),delta(I.fee,Ip&&Ip.fee),'Portföy: '+tl(pf.fee))}${tile('Çıkanların ort. geliri',tl(O.fee),delta(O.fee,Op&&Op.fee,{neutral:true}),'Portföy: '+tl(pf.fee))}
  </div>
  <div class="grid">
   ${card('h1','Portföye giren ve çıkan aboneler','Aylık adetler',el=>colChart(el,{labels:LB,mark:mi,series:[{name:'Giren',color:C(1),values:ts.map(q=>evS(q,INS).n)},{name:'Çıkan',color:C(2),values:ts.map(q=>evS(q,OUTS).n)}]}))}
   ${card('h4','Çıkış oranı','Portföyden çıkanlar / önceki ay portföyü',el=>lineChart(el,{labels:LB,mark:mi,series:[{name:'Çıkış oranı',color:C(1),values:ts.map(q=>evS(q,OUTS).n/tot(q-1).n*100)}],fmt:pct2}))}
   ${xcard('h6','Giren, çıkan ve portföy karşılaştırması',`Ortalama ${lc(ml)} · aylara göre`,sel('hx','m',Object.entries(ML).map(([k,v])=>[k,v[0]]),st.m,'Gösterge'),el=>lineChart(el,{labels:LB,mark:mi,fmt:mf,series:[{name:'Girenler',color:C(1),values:ts.map(q=>evS(q,INS)[st.m])},{name:'Çıkanlar',color:C(2),values:ts.map(q=>evS(q,OUTS)[st.m])},{name:'Portföy',color:'var(--neutral)',values:ts.map(q=>tot(q)[st.m])}]}),'c5')}
   ${plain('Hareket edenlerin profili',`${mLong(MONTHS[t])} · adet ve ortalamalar`,`<div class="tblwrap"><table><thead><tr><th scope="col">Grup</th><th scope="col">Adet</th><th scope="col">Ort. gelir</th><th scope="col">Ort. ARPU</th><th scope="col">Ort. data</th><th scope="col">Ort. ses</th><th scope="col">Ort. tenure</th><th scope="col">Ort. yaş</th></tr></thead><tbody>${rows.map(([l,x],q)=>`<tr${q<1?' class="em"':''}><th scope="row">${esc(l)}</th><td>${f0(x.n)}</td>${[[tl,'fee'],[tl,'arpu'],[gbf,'gb'],[dk,'mou'],[ay,'ten'],[f1,'age']].map(([f,k])=>`<td>${isFinite(x[k])?f(x[k]):'–'}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`,'c7')}
   ${xcard('h7','Hareket edenlerin kırılımı',`${eg[1]} · ${dim(st.d).name} · ${mLong(MONTHS[t])}`,sel('hx','e',EG.map(g=>[g[0],g[1]]),eg[0],'Hareket')+sel('hx','d',dimOpts(),st.d,'Kırılım')+sel('hx','mm',MM,st.mm,'Ölçü'),el=>{
     if(st.mm==='oran')return hbars(el,{items:L.map((l,i)=>({label:l,value:pv[i].n?mo[i].n/pv[i].n*100:NaN,extra:[[null,'Hareket eden',f0(mo[i].n)],[null,'Önceki ay portföyü',f0(pv[i].n)]]})).filter(d=>isFinite(d.value)),fmt:pct2,name:'Hareket oranı',xName:dim(st.d).name});
     const pN=po.reduce((q,v)=>q+v.n,0),pay=st.mm==='pay',f=pay?pct:ML[st.mm][1];
     return colChart(el,{labels:L.map(shortLab),xName:dim(st.d).name,fmt:f,h:250,tipTitle:i=>L[i]+' · '+f0(mo[i].n)+' hareket',series:[{name:eg[1],color:C(1),values:L.map((_,i)=>pay?(mN?mo[i].n/mN*100:0):(mo[i].n?mo[i][st.mm]:0))},{name:'Portföy',color:'var(--neutral)',values:L.map((_,i)=>pay?po[i].n/pN*100:(po[i][st.mm]||0))}]})},'c12')}
  </div>`;
}
function pHareket(){
  return head('Churn & Kazanım',`${mLong(MONTHS[T])} · Hybrid'e giriş ve çıkışlar, portföye giren ve çıkan abonelerin profili, arama listesi kohortları.`)+lcSection()+pfSection()+khSection()+(NG?`<div class="grid" style="margin-top:16px">${cohSection()}</div>`:'');
}

/* ================= Sayfa 4 — Model & KPI ================= */
function pModel(){
  const P0=D.model||[],co=coh(T),kn=co&&co.known,A=mLong(MONTHS[T]);
  if(!P0.length&&!NG)return head('Model & KPI','Model sonuçları ve parasal kazanımlar.')+empty('KPI tablosu bulunamadı: Excel dosyasını <code>data/kpi</code> klasörüne koyun.');
  const st=cs('md',{p:0,r:'',d:has('ten')?'ten':D.dims[0].key}),P=P0[+st.p]||P0[0];
  let tiles='',kpi='',seyir='';
  if(P){const R=P.satirlar,qi=P.aylar.indexOf(MONTHS[T]),by=k=>R.find(r=>r.key===k),yz=r=>r.s.some(x=>x.includes('%'));
    const tl_=(k,l,o)=>{const r=by(k);if(!r||qi<0||r.v[qi]==null)return'';const p_=qi>0?r.v[qi-1]:null,pp=yz(r);return tile(l,r.s[qi],delta(pp?r.v[qi]*100:r.v[qi],p_==null?null:pp?p_*100:p_,{ref:'önceki döneme göre',pp,d:2,...o}))};
    tiles=qi<0?`<div class="ins" style="margin-bottom:16px">${A} için KPI tablosunda değer yok. Tablo ${mShort(P.aylar[0])} – ${mShort(P.aylar[P.aylar.length-1])} dönemlerini içeriyor.</div>`
      :`<div class="tiles">${tl_('lift5','%5 lift',{})}${tl_('paylasilan_n','Paylaşılan kitle',{neutral:true})}${tl_('paylasilan_rr','Paylaşılan kitle response rate',{neutral:true})}${tl_('kontrol_rr','Kontrol grubu response rate',{neutral:true})}${tl_('tut','Tutundurma',{})}${tl_('tut1y','Tutundurma geliri (1 yıl)',{})}</div>`;
    let g='';kpi=`<div class="tblwrap"><table><thead><tr><th scope="col">KPI</th>${P.aylar.map((m,i)=>`<th scope="col"${i===qi?' class="cur"':''}>${mShort(m)}</th>`).join('')}</tr></thead><tbody>${R.map(r=>(r.grup!==g?(g=r.grup,r.grup?`<tr class="grp"><th scope="row" colspan="${P.aylar.length+1}">${esc(r.grup)}</th></tr>`:''):'')+`<tr><th scope="row">${esc(r.ad)}</th>${r.s.map((x,i)=>`<td${i===qi?' class="cur"':''}>${i===qi?'<b>':''}${x===''?'–':esc(x)}${i===qi?'</b>':''}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
    let ri=R.findIndex((r,i)=>String(i)===String(st.r));if(ri<0)ri=Math.max(0,R.findIndex(r=>r.key==='lift5'));const r=R[ri],pz=yz(r);
    seyir=xcard('m0','KPI seyri',`${esc(r.ad)} · ${mShort(P.aylar[0])} – ${mShort(P.aylar[P.aylar.length-1])}`,sel('md','r',R.map((x,i)=>[i,(x.grup?x.grup+' · ':'')+x.ad]),ri,'KPI'),el=>lineChart(el,{labels:P.aylar.map(mShort),mark:qi<0?null:qi,fmt:pz?pct2:v=>Math.abs(v)<100&&v%1?f2(v):f0(v),series:[{name:r.ad,color:C(1),values:r.v.map(v=>v==null?NaN:pz?v*100:v)}]}),'c12')}
  // skor dilimleri ve kırılımlar: kontrol grubu ile kontrol grubu dışı
  const ia=gix('k'),ib=gix('x'),KG=ia>=0&&ib>=0,na=KG?gname('k'):'',nb=KG?gname('x'):'';
  let skor='';if(D.coh_skor&&co&&co.skor.length){const gp=gix('p'),bl=arr(NB,b=>`${Math.round(b*100/NB)}–${Math.round((b+1)*100/NB)}`),B=arr(NB,b=>cv(co.skor,b,gp)),tn=B.reduce((q,x)=>q+x.n,0),tc=B.reduce((q,x)=>q+x.c,0);
    if(kn){let cn=0,cc=0;const gain=B.map(x=>{cn+=x.n;cc+=x.c;return [cn/tn*100,tc?cc/tc*100:0]}),avg=tn?tc/tn*100:0;
      skor=card('m1','Decision score dilimine göre churn oranı',`Skora göre sıralı %${f0(100/NB)}'lik dilimler (0–${f0(100/NB)} = en yüksek skor) · ${A} skoru, bir sonraki aydaki churn`,el=>colChart(el,{labels:bl,xName:'Skor dilimi (%)',fmt:pct2,series:[{name:'Churn oranı',color:C(1),values:B.map(x=>rate(x)||0)}],line:{value:avg,label:'Ortalama '+pct2(avg)},tipTitle:i=>'Skor dilimi %'+bl[i]+' · '+f0(B[i].n)+' abone · '+f0(B[i].c)+' churn'}))
       +card('m2','Kümülatif kazanım (gain)','En yüksek skorlu aboneler hedeflendiğinde yakalanan churn oranı',el=>lineChart(el,{labels:gain.map(g=>'%'+f0(g[0])),xName:'Hedeflenen kitle',fmt:pct,zero:true,endLabel:false,series:[{name:'Model',color:C(1),values:gain.map(g=>g[1])},{name:'Rastgele seçim',color:'var(--neutral)',values:gain.map(g=>g[0])}]}))
       +(KG?card('m3','Skor dilimine göre kontrol grubu ve kontrol grubu dışı','Churn oranı · aynı skor diliminde aksiyon alınmayan (kontrol) ve diğer aboneler',el=>{colChart(el,{labels:bl,xName:'Skor dilimi (%)',fmt:pct2,tipTitle:i=>'Skor dilimi %'+bl[i]+' · kontrol '+f0(cv(co.skor,i,ia).n)+' · kontrol dışı '+f0(cv(co.skor,i,ib).n)+' abone',series:[{name:na,color:C(2),values:arr(NB,b=>rate(cv(co.skor,b,ia))||0)},{name:nb,color:C(1),values:arr(NB,b=>rate(cv(co.skor,b,ib))||0)}]});
          return {head:['Skor dilimi (%)',na+' adet',na+' churn',na+' churn oranı',nb+' adet',nb+' churn',nb+' churn oranı'],rows:bl.map((l,b)=>{const x=cv(co.skor,b,ia),y=cv(co.skor,b,ib);return [l,f0(x.n),f0(x.c),fr(rate(x)),f0(y.n),f0(y.c),fr(rate(y))]})}},'c12'):'')}
    else skor=`<section class="card c12">${noOut()}</section>`}
  let grp='';if(co&&KG){const src=co.one[st.d],L=dim(st.d).labels,X=L.map((_,i)=>[cv(src,i,ia),cv(src,i,ib)]),ta=cv(co.tot,0,ia),tb=cv(co.tot,0,ib);
    const ayr=(id,q,nm)=>card(id,`${nm}: churn olan ve olmayanların ayrışması`,`${dim(st.d).name} · her grubun kendi içindeki dağılımı (%) · ${A}`,el=>{const tc=X.reduce((s_,x)=>s_+x[q].c,0),ts=X.reduce((s_,x)=>s_+x[q].n-x[q].c,0);return colChart(el,{labels:L.map(shortLab),xName:dim(st.d).name,fmt:pct,h:230,tipTitle:i=>L[i]+' · '+f0(X[i][q].c)+' churn / '+f0(X[i][q].n)+' abone',series:[{name:'Churn olanlar',color:C(2),values:X.map(x=>tc?x[q].c/tc*100:0)},{name:'Churn olmayanlar',color:C(1),values:X.map(x=>ts?(x[q].n-x[q].c)/ts*100:0)}]})});
    grp=xcard('m4','Kontrol grubu ve kontrol grubu dışı: kırılım bazında churn',`${kn?'Churn oranı':'Abone adedi'} · ${dim(st.d).name} · ${A} · ${na}: ${f0(ta.n)} abone${kn?' (churn '+fr(rate(ta))+')':''} · ${nb}: ${f0(tb.n)} abone${kn?' (churn '+fr(rate(tb))+')':''}`,sel('md','d',dimOpts(),st.d,'Kırılım'),el=>{colChart(el,{labels:L.map(shortLab),xName:dim(st.d).name,fmt:kn?pct2:f0,h:250,tipTitle:i=>L[i]+' · kontrol '+f0(X[i][0].n)+' · kontrol dışı '+f0(X[i][1].n)+' abone',series:[{name:na,color:C(2),values:X.map(x=>kn?(rate(x[0])||0):x[0].n)},{name:nb,color:C(1),values:X.map(x=>kn?(rate(x[1])||0):x[1].n)}]});
        return {head:[dim(st.d).name,na+' adet',na+' churn',na+' churn oranı',nb+' adet',nb+' churn',nb+' churn oranı','Fark (puan)'],rows:L.map((l,i)=>{const [x,y]=X[i],d=rate(x)-rate(y);return [l,f0(x.n),kn?f0(x.c):'–',kn?fr(rate(x)):'–',f0(y.n),kn?f0(y.c):'–',kn?fr(rate(y)):'–',kn&&isFinite(d)?(d>0?'+':d<0?'−':'')+f2(Math.abs(d)):'–']})}},'c12')
     +(kn?ayr('m5',0,na)+ayr('m6',1,nb):'')}
  else if(co&&!KG)grp=`<section class="card c12"><p style="color:var(--ink2)">Kontrol grubu kolonu bulunamadı; kırılım bazındaki karşılaştırma için kohort tablosunda kontrol grubu bayrağı gerekir.</p></section>`;
  return head('Model & KPI',`${A} · model sonuçları ve iş birimine etkisi.`,P0.length>1?`<div class="controls">${sel('md','p',P0.map((p,i)=>[i,p.ad||p.id]),st.p,'Proje',1)}</div>`:'')+tiles+
  `<div class="grid">
   ${P?plain('KPI tablosu',`${esc(P.ad||P.id)} · seçili dönem vurgulu`,kpi,'c12')+seyir:''}
   ${skor}${grp}
  </div>`;
}

/* ================= Çizim döngüsü ================= */
const PAGES={profil:pProfil,karsilastirma:pKars,hareket:pHareket,model:pModel};
function paint(){queue.forEach(([id,fn])=>{const el=document.getElementById(id);if(!el)return;
  if(tableMode.has(id)){const tmp=document.createElement('div');el.innerHTML=table(fn(tmp))}else fn(el)})}
/* Kart büyütme: geniş tablo ve grafikleri kaydırmadan tam ekranda göster */
const FH=(el,h)=>el.closest&&el.closest('.card.full')?Math.max(h,Math.min(innerHeight-250,640)):h;
function buyut(){
  const cards=[...document.querySelectorAll('#panel section.card')];if(S.fs!=null&&!cards[S.fs])S.fs=null;
  cards.forEach((c,i)=>{const hd=c.querySelector(':scope>header');if(!hd)return;const on=S.fs===i,ctl=hd.querySelector(':scope>.controls')||hd;
    ctl.insertAdjacentHTML('beforeend',`<button type="button" class="tbl-btn fs-btn" id="fs-${i}" data-fs="${i}" aria-pressed="${on}" title="${on?'Normal görünüme dön (Esc)':'Tam ekranda göster'}">${on?'✕ Küçült':'⤢ Büyüt'}</button>`);
    c.classList.toggle('full',on)});
  document.body.classList.toggle('fs-on',S.fs!=null);
}
function render(){
  const fid=document.activeElement&&document.activeElement.id,panel=$('#panel');
  T=+S.t;queue=[];$('#foot').textContent=`Rahat Segment Panosu · Dönem: ${mLong(MONTHS[T])} · Veri: ${mLong(MONTHS[0])} – ${mLong(MONTHS[TN])} · Üretim: ${D.meta.uretim}`;panel.setAttribute('aria-labelledby','tab-'+S.tab);panel.innerHTML=PAGES[S.tab]();buyut();paint();
  document.querySelectorAll('.tab').forEach(t=>{const on=t.dataset.tab===S.tab;t.setAttribute('aria-selected',on);t.tabIndex=on?0:-1});
  if(fid){const f=document.getElementById(fid);f&&f.focus({preventScroll:true})}
}
$('#month').innerHTML=MONTHS.map((id,k)=>`<option value="${k}">${mLong(id)}</option>`).join('');$('#month').value=T;
$('#month').addEventListener('change',e=>{S.t=+e.target.value;render()});
$('#tabs').innerHTML=TABS.map(([id,l],k)=>`<button class="tab" role="tab" type="button" id="tab-${id}" data-tab="${id}" aria-controls="panel"><span class="n" aria-hidden="true">${k+1}</span>${l}</button>`).join('');
const go=t=>{S.fs=null;S.tab=t;history.replaceState(null,'','#'+t);render()};
$('#tabs').addEventListener('click',e=>{const t=e.target.closest('.tab');if(!t)return;go(t.dataset.tab);window.scrollTo({top:0})});
$('#tabs').addEventListener('keydown',e=>{const k=TABS.findIndex(t=>t[0]===S.tab),n=TABS.length,j=e.key==='ArrowRight'?(k+1)%n:e.key==='ArrowLeft'?(k+n-1)%n:e.key==='Home'?0:e.key==='End'?n-1:-1;if(j<0)return;e.preventDefault();go(TABS[j][0]);$('#tab-'+S.tab).focus()});
document.addEventListener('change',e=>{const d=e.target.dataset;if(!d||!d.card)return;const v=e.target.value;cs(d.card,{})[d.prop]=d.num?Math.max(0,+v||0):v;render()});
addEventListener('keydown',e=>{if(e.key==='Escape'&&S.fs!=null){S.fs=null;render()}});
document.addEventListener('click',e=>{const fsb=e.target.closest('[data-fs]');if(fsb){S.fs=S.fs===+fsb.dataset.fs?null:+fsb.dataset.fs;render();return}const v=e.target.closest('button[data-val]');if(v){cs(v.dataset.card,{})[v.dataset.prop]=v.dataset.val;render();return}const b=e.target.closest('[data-tbl]');if(!b)return;const id=b.dataset.tbl;tableMode.has(id)?tableMode.delete(id):tableMode.add(id);render()});
const setTheme=t=>{document.documentElement.dataset.theme=t;store.set('rahat-theme',t)};
$('#theme').addEventListener('click',()=>setTheme(document.documentElement.dataset.theme==='dark'?'light':'dark'));
const tip=$('#tip');
function showTip(t,x,y){tip.innerHTML=t.dataset.tip;tip.classList.add('on');const r=tip.getBoundingClientRect();let L=x+14,Tp=y+14;if(L+r.width>innerWidth-8)L=x-r.width-14;if(Tp+r.height>innerHeight-8)Tp=y-r.height-14;tip.style.left=Math.max(8,L)+'px';tip.style.top=Math.max(8,Tp)+'px'}
document.addEventListener('pointermove',e=>{const t=e.target.closest&&e.target.closest('[data-tip]');t?showTip(t,e.clientX,e.clientY):tip.classList.remove('on')});
document.addEventListener('focusin',e=>{const t=e.target.closest&&e.target.closest('[data-tip]');if(t){const r=t.getBoundingClientRect();showTip(t,r.left+r.width/2,r.bottom-8)}});
document.addEventListener('focusout',()=>tip.classList.remove('on'));
addEventListener('scroll',()=>tip.classList.remove('on'),{passive:true});
let rw=innerWidth,rt;addEventListener('resize',()=>{clearTimeout(rt);rt=setTimeout(()=>{if(innerWidth!==rw){rw=innerWidth;paint()}},150)});
addEventListener('hashchange',()=>{const t=location.hash.slice(1);if(PAGES[t]&&t!==S.tab){S.tab=t;render()}});
window.__rapor={D,S,render,TABS};
render();
})();
</script>
</body>
</html>
'''


if __name__ == '__main__':
    main()
