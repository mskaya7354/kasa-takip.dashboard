# Kasa Takip

Kurumsal kasa hareketlerini izlemek için **gerçek zamanlı web dashboard'u**. Excel dosyasını (`xlsm`) kaynak olarak kullanır; dosya her kaydedildiğinde dashboard otomatik güncellenir.

No-build, no-framework: tek Python sunucusu + tek HTML dosyası.

## Nasıl çalışır

1. `server.py` Excel dosyasını açar, verileri parse eder
2. Dosya değişikliklerini `watchdog` ile izler
3. Değişiklik olduğunda bağlı tüm tarayıcılara **WebSocket** ile yeni veri gönderir
4. `dashboard.html` gerçek zamanlı güncellenir — sayfa yenilemesi gerekmez

## Kullanım

```bash
pip install -r requirements.txt
python server.py --file "Z:\RAPOR\Kasa\Kasa_Takip_v3.xlsm" --port 8765 --polling
```

Ya da `baslat.bat` dosyasındaki `EXCEL_PATH` değişkenini düzenleyip çift tıkla.

Dashboard: `http://localhost:8765` (veya LAN IP'si ile ağdaki diğer cihazlardan)

## Stack

- **Backend:** FastAPI, uvicorn, WebSocket, watchdog, pandas, openpyxl
- **Frontend:** Vanilla HTML/JS + Chart.js (CDN) — build adımı yok

## Excel yapısı

`server.py` şu sheet isimlerini bekler:

| Sheet | İçerik |
|-------|--------|
| `VERI_GIRISI` | Kasa hareketleri (tarih, tutar, tür, araç, durum) |

Kolon isimleri `TUTAR TL`, `ÖDEME TÜRÜ`, `ÖDEME ARACI`, `DURUM`, `TARİH` pattern'larıyla eşleniyor.
