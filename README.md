# Kasa Takip Dashboard

Kurumsal kasa, banka bakiyesi, döviz kuru, kredi portföyü, ödeme takvimi ve nakit projeksiyonunu izleyen **gerçek zamanlı, çok-PC dağıtık dashboard**. Tek bir paylaşımlı Excel dosyasını (`xlsm`) kaynak olarak kullanır; muhasebe ekibi dosyayı kaydeder, ağdaki tüm bilgisayarlarda açık dashboard otomatik güncellenir.

Gerçek bir iş ihtiyacı için yapıldı: muhasebe Excel'de kayıt tutuyor, yönetim canlı bakiye/akış görmek istiyor ama herkesin Excel'i açıp kapatması verimsiz ve birden fazla kullanıcı aynı anda açamıyor. Bu uygulama Excel'i tek doğruluk kaynağı olarak korurken, üzerine canlı görselleştirme katmanı ekler.

**No-build, no-framework:** tek Python sunucusu + birkaç bağımsız HTML sayfası. Vanilla JS + Chart.js. Frontend için npm/webpack/vite yok.

## Ne sunuyor

Standart bir masaüstü uygulaması değil — **tarayıcıdan açılan bir web
uygulaması**. Kurulum gerektirmez, herkes aynı anda kendi bilgisayarından
veya telefonundan aynı adrese girer; hepsi aynı canlı veriyi görür. 4 sayfası
var:

**Ana Dashboard** (`/`) — Kasadaki nakit, banka hesap bakiyeleri
(kullanılabilir ve bloke ayrı ayrı) ve güncel döviz kurları tek ekranda.
KPI kartları ve trend grafikleriyle "bugün elimizde ne var" sorusunun
cevabı tek bakışta.

**Kredi Portföyü** (`/kredi`) — Şirketin kullandığı kredi, leasing ve
teminat mektuplarının tamamı tek listede: kalan bakiye, vade, hangi
yükümlülüğün ne durumda olduğu — muhasebeye sormaya gerek kalmadan.

**Ödeme Takip Panosu** (`/odeme-takip`) — Yaklaşan ve geciken ödemeler
takvim görünümünde; bu hafta/gelecek ay hangi ödemelerin çıkacağını
filtreleyerek görün. Beklenen gelirler de aynı ekranda.

**Nakit Raporu** (`/nakit-rapor`) — Önümüzdeki 6 ay ile 1 yıl için nakit
pozisyonu projeksiyonu; sistem geçmiş hareketlere bakarak olağandışı
(anomali) harcamaları ve düzenli tekrar eden kalemleri (kira, taksit gibi)
kendiliğinden işaretler.

### Faydaları

- **Tek bakışta güncel durum** — Excel dosyasını açıp satır satır
  taramaya gerek yok
- **Herkes aynı anda bakabilir** — muhasebe Excel'i düzenlerken yönetim
  aynı anda, başka bir bilgisayardan canlı veriyi izleyebilir; dosya
  "kilitli, açılamıyor" derdi yok
- **Anında güncellenir** — muhasebe kaydı kaydettiği an (birkaç saniye
  içinde) açık tüm ekranlar kendiliğinden yenilenir, "yenile" tuşuna
  basmaya gerek yok
- **Tek doğruluk kaynağı korunur** — veri hâlâ Excel'de tutulur, çift
  giriş veya senkron sorunu yok; dashboard sadece üzerine bakılan bir
  pencere
- **Öngörü sağlar** — Nakit Raporu geçmişe değil geleceğe bakar: önümüzdeki
  aylarda nakit sıkışıklığı olur mu, olağandışı bir harcama var mı,
  otomatik uyarır
- **Her yerden erişim** — sadece bir tarayıcı yeterli; PWA desteğiyle
  "ana ekrana ekle" ile telefondan da kullanılabilir

## Mimari

Uygulama iki dağıtım modunu da destekler. **Üretimde A modu kullanılıyor.**

### A) Tek sunucu (üretimdeki kurulum)

Tek bir Linux sunucu servisi, Excel'i CIFS/SMB mount üzerinden okur; LAN'daki
tüm tarayıcılar aynı sunucuya bağlanır.

```
   ┌──────────────────────────┐
   │  Dosya Sunucusu (SMB)    │
   │  Kasa_Takip.xlsm         │◀── Muhasebe kaydediyor
   └────────────┬─────────────┘
                │ CIFS mount (/mnt/...)
                ▼
   ┌──────────────────────────┐
   │  Linux Sunucu (7/24)     │
   │  systemd: kasa.service   │
   │  python3 server.py       │
   │  --polling  :8765        │
   └────────────┬─────────────┘
                │ LAN
                ▼
   ┌──────────────────────────────────────┐
   │  Tarayıcılar — LAN'daki her cihaz    │
   │  http://<sunucu-ip>:8765             │
   │  HTTP Basic Auth ile korumalı        │
   └──────────────────────────────────────┘
```

Tek servis; `systemctl restart kasa` ile yönetilir. Dosya SMB üzerinden
okunduğu için `--polling` şart (watchdog inotify olayları SMB'de tetiklenmez).
Kurulum için `kasa.service.example` dosyasına bakın.

### B) Çok-PC dağıtık (alternatif / merkezi sunucusuz)

Sunucu ayırmak istemeyen kurulumlar için: her PC kendi `server.py`'ını
çalıştırır, hepsi aynı paylaşımlı Excel'i izler, mDNS ile `kasa.local`
olarak kendini duyurur.

```
              ┌─────────────────────────┐
              │  Paylaşımlı Network     │
              │  Drive (SMB/UNC)        │
              │  Kasa_Takip.xlsm        │◀── Muhasebe kaydediyor
              │  stop.flag / restart.flag│
              └────────────┬────────────┘
                           │ Tüm PC'ler aynı dosyayı izler
        ┌──────────────────┼──────────────────┐
        ▼                  ▼                  ▼
   ┌──────────┐       ┌──────────┐       ┌──────────┐
   │ PC #1    │       │ PC #2    │       │ PC #N    │
   │ server.py│       │ server.py│       │ server.py│
   │ :8765    │       │ :8765    │       │ :8765    │
   └──────────┘       └──────────┘       └──────────┘
         mDNS: kasa.local — her PC kendi tarayıcılarına servis eder
```

Her PC bağımsız çalışır, ortak Excel'i izlediği için **veri tutarlılığı dosya
seviyesinde** sağlanır. Sinyal dosyaları (`stop.flag`, `restart.flag`) tek
komutla tüm PC'lerdeki sunucuları durdurmaya/yeniden başlatmaya yarar — bu
mod için `kurulum.bat` / `durdur_hepsi.bat` / `trigger_yenile.bat` betikleri
kullanılır.

## Teknik özellikler

- **Gerçek zamanlı güncelleme** — Excel kaydedildiği anda WebSocket ile tüm bağlı tarayıcılara push (3 sn debounce, `~$` lock dosyalarını yoksay)
- **HTTP Basic Auth** — `KASA_USER` / `KASA_PASS` env var'ları ile koruma (`secrets.compare_digest` — timing attack korumalı)
- **mDNS / Zeroconf** — sunucu kendini `kasa.local` olarak duyurur, IP bilmek gerekmez
- **Çok-PC dağıtık çalışma** — birden fazla PC aynı anda aynı Excel'i izleyebilir; her PC kendi tarayıcılarına servis eder
- **Sinyal dosyaları** — paylaşımlı network drive üzerinden `stop.flag` / `restart.flag` ile koordineli kontrol
- **Üç ayrı veri kaynağı** tek dashboard'da:
  - Kasa hareketleri (tutar, ödeme türü, ödeme aracı, durum, tarih)
  - Banka bakiyeleri (kullanılabilir + bloke)
  - Döviz kuru tarihçesi (EUR/USD, son 60 gün)
- **Fuzzy column matching** — Excel başlık ismi değişse bile (büyük/küçük harf, Türkçe karakter, boşluk) eşler
- **Otomatik agregasyon** — KPI'lar, gruplamalar, aylık tarihçe (son 18 ay)
- **Polling fallback** — WebSocket düşerse 20 sn'de bir REST polling devreye girer
- **Embedded Python desteği** — sistemde Python yoksa `_ARSIV/python` klasöründen kullanır
- **Sıfır build** — vanilla HTML + Chart.js (lokal servis edilir, CDN bağımlılığı yok)
- **Kredi Portföyü** — kredi/leasing/teminat mektubu takibi; Excel tarayıcıda (client-side) parse edilir, sunucu sadece dosyayı ve değişiklik zamanını sunar
- **Ödeme Takip Panosu** — bekleyen/gelecek ödemeler için filtreli liste + takvim görünümü, gelir beklentileri kaydı
- **Nakit Raporu** — bütçe planlama (6 ay / 1 yıl projeksiyon), anomali tespiti ve tekrarlayan kalem tanıma
- **HTTPS desteği** — `--ssl-cert` / `--ssl-key` ile opsiyonel TLS
- **PWA manifest** — tarayıcıya "ana ekrana ekle" desteği

## Hızlı başlangıç

```bash
pip install -r requirements.txt

# Auth'u ortam değişkenleri ile ayarla (varsayılanı kullanma!)
set KASA_USER=<kullanici-adi>
set KASA_PASS=<guclu-sifre>

# Çalıştır
python server.py --file "/path/to/Kasa_Takip.xlsm" --port 8765 --polling
```

> ⚠️ `server.py` içinde fallback default değerler vardır ama bunları ASLA üretimde kullanmayın. `KASA_USER` ve `KASA_PASS` env var'larını mutlaka tanımlayın.

Ya da `baslat.bat`'taki `EXCEL_PATH` değişkenini düzenleyip çift tıkla.

Dashboard: `http://kasa.local:8765` (mDNS yayını) ya da `http://<bilgisayar-ip>:8765`

### Komut satırı argümanları

| Argüman | Açıklama | Varsayılan |
|---------|----------|------------|
| `--file` | Excel dosya yolu (zorunlu) | — |
| `--kredi-file` | Kredi/leasing/teminat mektubu Excel yolu (opsiyonel) | — |
| `--port` | HTTP/WebSocket portu | `8765` |
| `--host` | Bind adresi | `0.0.0.0` |
| `--polling` | Watchdog polling modu (network drive'larda gerekli) | kapalı |
| `--ssl-cert` / `--ssl-key` | HTTPS için sertifika/anahtar yolu (opsiyonel) | — |

### Ortam değişkenleri

| Değişken | Açıklama |
|----------|----------|
| `KASA_USER` | Basic Auth kullanıcı adı (mutlaka ayarla) |
| `KASA_PASS` | Basic Auth şifresi (mutlaka ayarla) |

## Yardımcı betikler

Aşağıdakiler **B modu** (çok-PC dağıtık) içindir — üretimdeki A modu
(systemd) bunları kullanmaz, servis doğrudan `systemctl` ile yönetilir.

| Dosya | Ne yapar |
|-------|----------|
| `baslat.bat` | Manuel başlatma — etkileşimli, log konsola |
| `baslat_servis.bat` | Hizmet kapatma + startup linkini temizleme |
| `kurulum.bat` | İlk kurulum: Python kopyala, startup linki ekle, firewall portu aç (yönetici izni ister) |
| `durdur.bat` | Lokal makineyi durdur |
| `durdur_hepsi.bat` | Sinyal dosyası ile tüm PC'lerdeki sunucuları durdur |
| `trigger_yenile.bat` | Sinyal dosyası ile tüm PC'lerdeki sunucuları yeniden başlat |
| `gizli.vbs` | `baslat_servis.bat`'ı gizli pencerede çalıştırır (startup için) |

## Excel yapısı

`server.py` üç sheet bekler. Eksik sheet'ler hata vermez, ilgili bölüm boş kalır. Sütun isimleri fuzzy matching ile bulunur.

### `VERI_GIRISI` (header=4. satır)

| Beklenen sütun pattern | Eşlenen alan |
|------------------------|--------------|
| `TUTAR TL` | `tutar_tl` |
| `ÖDEME TÜRÜ` / `ODEME TURU` | `odeme_turu` |
| `ÖDEME ARACI` / `ODEME ARACI` | `odeme_araci` |
| `DURUM` | `durum` |
| `TARİH` / `TARIH` | `tarih` |

### `KASA DURUM` (header=1. satır)

`BANKA ADI`, `BAKİYE TL`, `KULLANILABİLİR BAKİYE`, `BLOKE`

### `KUR_GECMISI` (header=3. satır)

`TARİH`, `EUR`, `USD` — son 60 gün gösterilir.

## API

| Endpoint | Auth | Açıklama |
|----------|------|----------|
| `GET /` | ✓ | `dashboard.html` |
| `GET /api/data` | ✓ | Son parse edilmiş veri (JSON) |
| `GET /api/refresh` | ✓ | Manuel cache invalidation + reparse |
| `GET /api/health` | — | Sunucu durumu, bağlı WS sayısı, son güncelleme |
| `GET /api/test` | — | Diagnostic — KPI sayısı, grup sayısı |
| `GET /kredi` · `/odeme-takip` · `/nakit-rapor` | ✓ | Ek dashboard sayfaları |
| `GET /api/kredi-file` | ✓ | Ham kredi Excel'i — parse tarayıcıda yapılır |
| `GET/POST /api/kredi-notlar` | ✓ | Kredi kalemlerine not ekleme |
| `GET/POST /api/gelir-beklentileri` | ✓ | Gelir beklentisi kayıtları |
| `GET/POST /api/manuel-giderler` | ✓ | Elle girilen gider kayıtları |
| `WS /ws` | — | Canlı güncelleme stream'i |
| `GET /chart.min.js` · `/xlsx.full.min.js` | — | Lokal JS bundle'ları (Chart.js, SheetJS) |

## Stack

- **Backend:** FastAPI · uvicorn · WebSocket · watchdog (PollingObserver) · pandas · openpyxl · zeroconf (mDNS) · HTTP Basic Auth
- **Frontend:** Vanilla HTML/JS + Chart.js
- **Platform:** Windows (mapped drive üzerinden Excel okuma), Linux/macOS'ta da çalışır

## Güvenlik notu

- **Basic Auth** ile korumalı, ama HTTP üzerinden çalıştığı için trafiği güvenilir LAN'da tutmak gerekir
- **HTTPS** için ya `--ssl-cert`/`--ssl-key` ile native TLS ya da reverse proxy önerilir
- **CORS `*`** — şu an açık; whitelist'e çevirmek isteyenler `server.py` içinden değiştirebilir
- İnternete açmadan önce port forwarding yapmayın; bu uygulama dahili kullanım içindir

## Notlar

- Excel kaydedildiğinde Office önce `~$dosyaadi.xlsm` adlı lock dosyası oluşturur; `FH.on_modified` bunları filtreler
- `watchdog` SMB üzerinde olay yakalayamayabilir — bu yüzden `--polling` flag'i şart
- 3 saniyelik debounce, hızlı ardışık kayıtlarda gereksiz parse'i önler
- `NaN`/`Inf` float değerleri `SafeEncoder` ile `null`'a çevrilir (JSON spec uyumu)
- mDNS kurulamazsa sunucu IP ile erişilebilir kalır
