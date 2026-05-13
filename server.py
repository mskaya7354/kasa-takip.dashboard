"""
Kasa Takip — Dashboard Sunucusu (v4)
- /api/data  ... tüm özet veriyi döndürür (KPI + bakiyeler + bekleyen + takvim + işlemler + kurlar)
- /ws        ... Excel değişince gerçek zamanlı push
Excel şeması değişmiş olursa, kolon map'leri aşağıdan ayarlanabilir.
"""
import asyncio, argparse, json, logging, math, os, sys, time, traceback, threading, secrets
from collections import defaultdict
from contextlib import asynccontextmanager
from datetime import datetime, date, timedelta
from pathlib import Path

import pandas as pd
import uvicorn
from fastapi import FastAPI, WebSocket, WebSocketDisconnect, Depends, HTTPException, status
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, PlainTextResponse, FileResponse
from watchdog.events import FileSystemEventHandler
from watchdog.observers.polling import PollingObserver

_log_file = Path(__file__).parent / "server.log"
_handlers: list = [logging.FileHandler(_log_file, encoding="utf-8")]
try:
    sys.stdout.write("")  # pythonw altında exception fırlatır
    _handlers.append(logging.StreamHandler())
except Exception:
    pass
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", handlers=_handlers)
log = logging.getLogger("kasa")


# ---------- JSON helpers ----------
class SafeEncoder(json.JSONEncoder):
    def default(self, o):
        if isinstance(o, float) and (math.isnan(o) or math.isinf(o)): return None
        if hasattr(o, "isoformat"): return o.isoformat()
        return str(o)

def safe_val(v):
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)): return None
    return v

def s(v):
    """NaN/None güvenli string dönüşümü — 'nan' stringine dönmez."""
    try:
        if v is None: return ""
        if isinstance(v, float) and math.isnan(v): return ""
        if pd.isna(v): return ""
    except Exception:
        pass
    return str(v).strip()

def clean(records):
    return [{k: safe_val(v) for k, v in r.items()} for r in records]


# ---------- Column matching helpers ----------
def find_col(cols, keys):
    """Case-insensitive partial match — ilk eşleşen kolon adını döndürür."""
    for c in cols:
        u = str(c).upper()
        if all(k.upper() in u for k in keys):
            return c
    return None


# ---------- Main parser ----------
def parse_excel(path):
    log.info(f"Excel okunuyor: {path}")
    t0 = time.time()
    xl = pd.ExcelFile(path, engine="openpyxl", engine_kwargs={"read_only": True, "data_only": True})
    today = date.today()
    data = {
        "guncelleme": datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
        "dosya": os.path.basename(path),
        "kpi": {},
        "bakiyeler": [],
        "aylik": [],
        "odeme_turleri": [],
        "odeme_araclari": [],
        "kurlar": [],
        "islemler": [],
        "bekleyen": [],
        "takvim": [],
        "nakit_akisi": [],
        "karsilastirma": {"bu_ay": [], "gecen_ay": []},
        "kredi_kartlari": [],
        "hakedisler": [],
        "cekler_verilen": [],
        "cekler_alinan": [],
        "bakiye_log": [],
    }

    # ---------- VERI_GIRISI ----------
    df = None
    try:
        if "VERI_GIRISI" in xl.sheet_names:
            df = xl.parse("VERI_GIRISI", header=3)
            df.columns = df.columns.str.strip()
            cm = {}
            for c in df.columns:
                u = str(c).upper()
                val = None
                if "TUTAR" in u and "TL" in u: val = "tutar_tl"
                elif any(x in u for x in ["ODEME TURU", "ÖDEME TÜRÜ"]): val = "odeme_turu"
                elif any(x in u for x in ["ODEME ARACI", "ÖDEME ARACI"]): val = "odeme_araci"
                elif "DURUM" in u: val = "durum"
                elif any(x in u for x in ["VADE"]): val = "vade"
                elif any(x in u for x in ["TARİH", "TARIH"]) and "VADE" not in u: val = "tarih"
                elif "CARİ" in u or "CARI" in u or "MÜŞTERİ" in u or "MUSTERI" in u or "PROJE" in u: val = "cari"
                elif "AÇIKLAMA" in u or "ACIKLAMA" in u: val = "aciklama"
                elif "BANKA" in u or "ÖDEME ARACI" in u or "ODEME ARACI" in u: val = "banka"
                elif "YÖN" in u or "YON" in u or "TİP" in u or "TIP" in u or "GİRİŞ" in u or "ÇIKIŞ" in u: val = "yon"
                elif "DÖVİZ" in u or "DOVIZ" in u or "PARA BİRİMİ" in u or "PARA BIRIMI" in u or u == "CCY": val = "doviz"
                elif "TUTAR" in u and "TL" not in u: val = "tutar_orj"

                if val and val not in cm.values():
                    cm[c] = val

            df = df.rename(columns=cm)

            # Odeme turunu her zaman B sutunundan (2. sutun) al
            try:
                raw = xl.parse("VERI_GIRISI", header=3)
                if raw.shape[1] > 1:
                    df["odeme_turu"] = raw.iloc[:, 1].astype(str).str.strip().replace({"nan": "", "None": ""})
                    log.info(f"  odeme_turu: VERI_GIRISI B sutunundan ({raw.columns[1]})")
            except Exception as ex:
                log.error(f"  odeme_turu B sutun okuma hatasi: {ex}")

            # ODEME ARACI hem odeme_araci hem banka olarak kullanilir
            if "banka" not in df.columns and "odeme_araci" in df.columns:
                df["banka"] = df["odeme_araci"]

            if "vade" not in df.columns and "tarih" in df.columns:
                df["vade"] = df["tarih"]

            if "tutar_tl" in df.columns:
                df["tutar_tl"] = pd.to_numeric(df["tutar_tl"], errors="coerce")
                df = df[df["tutar_tl"].notna() & (df["tutar_tl"] != 0)]

                # ---------- KPI ----------
                data["kpi"].update({
                    "toplam_tutar": round(float(df["tutar_tl"].sum()), 2),
                    "islem_sayisi": int(len(df)),
                    "ortalama_tutar": round(float(df["tutar_tl"].mean()), 2),
                    "en_buyuk": round(float(df["tutar_tl"].max()), 2),
                    "bekleyen_adet": int(df["durum"].astype(str).str.upper().str.contains("BEKL").sum()) if "durum" in df.columns else 0,
                })

                # ---------- Tahsilat / Odeme ayrimi ----------
                if "yon" in df.columns:
                    yu = df["yon"].astype(str).str.upper()
                    df["_is_in"] = yu.str.contains("GİRİŞ|GIRIS|TAHSİLAT|TAHSILAT|GELEN|IN|\\+", regex=True, na=False)
                elif "tutar_tl" in df.columns:
                    df["_is_in"] = df["tutar_tl"] > 0

                tahsilat_sum = float(df.loc[df["_is_in"], "tutar_tl"].abs().sum()) if "_is_in" in df.columns else 0
                odeme_sum = float(df.loc[~df["_is_in"], "tutar_tl"].abs().sum()) if "_is_in" in df.columns else 0
                data["kpi"]["toplam_tahsilat"] = round(tahsilat_sum, 2)
                data["kpi"]["toplam_odeme"] = round(odeme_sum, 2)
                data["kpi"]["net_akis"] = round(tahsilat_sum - odeme_sum, 2)

                # ---------- Odeme turleri ----------
                if "odeme_turu" in df.columns:
                    g = df.groupby("odeme_turu")["tutar_tl"].agg(["sum", "count"]).reset_index()
                    g.columns = ["odeme_turu", "tutar", "adet"]
                    g["tutar"] = g["tutar"].abs().round(2)
                    g["adet"] = g["adet"].astype(int)
                    data["odeme_turleri"] = clean(g.sort_values("tutar", ascending=False).to_dict("records"))

                if "odeme_araci" in df.columns:
                    g = df.groupby("odeme_araci")["tutar_tl"].agg(["sum", "count"]).reset_index()
                    g.columns = ["odeme_araci", "tutar", "adet"]
                    g["tutar"] = g["tutar"].abs().round(2)
                    g["adet"] = g["adet"].astype(int)
                    data["odeme_araclari"] = clean(g.sort_values("tutar", ascending=False).head(12).to_dict("records"))

                # ---------- Aylik ve nakit akisi ----------
                if "tarih" in df.columns:
                    df["tarih"] = pd.to_datetime(df["tarih"], errors="coerce")
                    df2 = df.dropna(subset=["tarih"]).copy()
                    if len(df2) > 0:
                        df2["ay_key"] = df2["tarih"].dt.to_period("M")
                        df2["ay"] = df2["tarih"].dt.strftime("%b %y")

                        current_month_key = pd.Timestamp(today).to_period("M")
                        df2_chart = df2[df2["ay_key"] <= current_month_key]

                        m = df2_chart.groupby(["ay_key", "ay"])["tutar_tl"].sum().abs().reset_index().sort_values("ay_key").tail(18)
                        data["aylik"] = [{"ay": r["ay"], "tutar": round(float(r["tutar_tl"]), 2)} for _, r in m.iterrows()]

                        past_cf = []
                        if "_is_in" in df2_chart.columns:
                            g = df2_chart.groupby("ay_key").apply(lambda x: pd.Series({
                                "tahsilat": float(x.loc[x["_is_in"], "tutar_tl"].abs().sum()),
                                "odeme":    float(x.loc[~x["_is_in"], "tutar_tl"].abs().sum()),
                            })).reset_index().sort_values("ay_key")
                            for _, r in g.iterrows():
                                past_cf.append({
                                    "ay_key": r["ay_key"],
                                    "ay": r["ay_key"].strftime("%b %y"),
                                    "tahsilat": round(r["tahsilat"], 2),
                                    "odeme": round(r["odeme"], 2)
                                })
                        data["nakit_akisi"] = past_cf

                        bu_ay_key = pd.Timestamp(today).to_period("M")
                        gecen_ay_key = (pd.Timestamp(today) - pd.DateOffset(months=1)).to_period("M")
                        for key, dst in [(bu_ay_key, "bu_ay"), (gecen_ay_key, "gecen_ay")]:
                            sub = df2[df2["ay_key"] == key].copy()
                            if len(sub) == 0: continue
                            sub["gun"] = sub["tarih"].dt.day
                            sub["net"] = sub.apply(lambda r: abs(r["tutar_tl"]) * (1 if r.get("_is_in", True) else -1), axis=1)
                            daily = sub.groupby("gun")["net"].sum().sort_index()
                            cum = daily.cumsum()
                            data["karsilastirma"][dst] = [{"gun": int(g), "tutar": round(float(v), 2)} for g, v in cum.items()]

                # ---------- Islemler ----------
                if "tarih" in df.columns:
                    sub = df.dropna(subset=["tarih"]).sort_values("tarih", ascending=False).head(2000)
                    islemler = []
                    for _, r in sub.iterrows():
                        cari = s(r.get("cari")) or s(r.get("odeme_turu"))
                        islemler.append({
                            "tarih": r["tarih"].strftime("%d.%m") if pd.notna(r["tarih"]) else "",
                            "tarih_tam": r["tarih"].strftime("%d.%m.%Y") if pd.notna(r["tarih"]) else "",
                            "cari": cari,
                            "aciklama": s(r.get("aciklama")),
                            "banka": s(r.get("banka")),
                            "odeme_araci": s(r.get("odeme_araci")),
                            "odeme_turu": s(r.get("odeme_turu")),
                            "yon": "in" if r.get("_is_in", True) else "out",
                            "tutar_tl": round(abs(float(r["tutar_tl"])), 2) if pd.notna(r["tutar_tl"]) else 0,
                            "tutar_orj": float(r["tutar_orj"]) if ("tutar_orj" in df.columns and pd.notna(r.get("tutar_orj"))) else None,
                            "doviz": s(r.get("doviz")) or "TRY",
                            "durum": s(r.get("durum")),
                        })
                    data["islemler"] = clean(islemler)

                # ---------- Bekleyen tahsilat / odeme ----------
                if "durum" in df.columns:
                    bek = df[df["durum"].astype(str).str.upper().str.contains("BEKL", na=False)].copy()
                    if "vade" in bek.columns:
                        bek["vade"] = pd.to_datetime(bek["vade"], errors="coerce")
                    vade_ici = 0; vade_gec = 0; tutar_ici = 0.0; tutar_gec = 0.0
                    vade_ici_out = 0; vade_gec_out = 0; tutar_ici_out = 0.0; tutar_gec_out = 0.0
                    adet_in = 0; adet_out = 0
                    for _, r in bek.iterrows():
                        t = abs(float(r.get("tutar_tl", 0) or 0))
                        v = r.get("vade")
                        if r.get("_is_in", True):
                            adet_in += 1
                            if pd.notna(v) and v.date() < today:
                                vade_gec += 1; tutar_gec += t
                            else:
                                vade_ici += 1; tutar_ici += t
                        else:
                            adet_out += 1
                            if pd.notna(v) and v.date() < today:
                                vade_gec_out += 1; tutar_gec_out += t
                            else:
                                vade_ici_out += 1; tutar_ici_out += t
                    data["kpi"]["bekleyen_adet"] = adet_in
                    data["kpi"]["bekleyen_tutar"] = round(tutar_ici + tutar_gec, 2)
                    data["kpi"]["bekleyen_vade_ici"] = int(vade_ici)
                    data["kpi"]["bekleyen_gecikmis"] = int(vade_gec)
                    data["kpi"]["bekleyen_gecikmis_tutar"] = round(tutar_gec, 2)
                    data["kpi"]["bekleyen_odeme_adet"] = adet_out
                    data["kpi"]["bekleyen_odeme_tutar"] = round(tutar_ici_out + tutar_gec_out, 2)
                    data["kpi"]["bekleyen_odeme_vade_ici"] = int(vade_ici_out)
                    data["kpi"]["bekleyen_odeme_gecikmis"] = int(vade_gec_out)
                    data["kpi"]["bekleyen_odeme_gecikmis_tutar"] = round(tutar_gec_out, 2)

                    if "vade" in bek.columns:
                        bek = bek.sort_values("vade", na_position="last").head(2000)
                    bekleyen = []
                    for _, r in bek.iterrows():
                        v = r.get("vade")
                        cari = s(r.get("cari")) or s(r.get("odeme_turu"))
                        bekleyen.append({
                            "tarih": r.get("tarih").strftime("%d.%m.%Y") if pd.notna(r.get("tarih")) else "",
                            "vade": v.strftime("%d.%m.%Y") if pd.notna(v) else "",
                            "cari": cari,
                            "odeme_turu": s(r.get("odeme_turu")),
                            "aciklama": s(r.get("aciklama")),
                            "banka": s(r.get("banka")),
                            "tutar_tl": round(abs(float(r.get("tutar_tl", 0) or 0)), 2),
                            "yon": "in" if r.get("_is_in", True) else "out",
                            "doviz": s(r.get("doviz")) or "TRY",
                            "gecikmis": bool(pd.notna(v) and v.date() < today),
                            "gun_fark": int((v.date() - today).days) if pd.notna(v) else None,
                        })
                    data["bekleyen"] = clean(bekleyen)

                    # Projeksiyon (Gelecek Nakit Akisi)
                    proj_cf = {}
                    current_month_key = pd.Timestamp(today).to_period("M")
                    for _, r in bek.iterrows():
                        v = r.get("vade")
                        t = abs(float(r.get("tutar_tl", 0) or 0))
                        is_in = r.get("_is_in", True)
                        if pd.isna(v): continue

                        v_key = pd.Timestamp(v).to_period("M")
                        if v_key < current_month_key:
                            v_key = current_month_key

                        if v_key not in proj_cf:
                            proj_cf[v_key] = {"tahsilat": 0.0, "odeme": 0.0}

                        if is_in: proj_cf[v_key]["tahsilat"] += t
                        else: proj_cf[v_key]["odeme"] += t

                    cf_dict = {row["ay_key"]: row for row in data.get("nakit_akisi", [])}
                    for k, v in proj_cf.items():
                        if k not in cf_dict:
                            cf_dict[k] = {"ay_key": k, "ay": k.strftime("%b %y"), "tahsilat": 0.0, "odeme": 0.0}
                        cf_dict[k]["tahsilat"] += v["tahsilat"]
                        cf_dict[k]["odeme"] += v["odeme"]

                    merged = []
                    for k in sorted(cf_dict.keys()):
                        r = cf_dict[k]
                        merged.append({
                            "ay": r["ay"], "ay_str": str(k),
                            "tahsilat": round(r["tahsilat"], 2),
                            "odeme": round(r["odeme"], 2),
                            "net": round(r["tahsilat"] - r["odeme"], 2)
                        })
                    data["nakit_akisi"] = merged

                # ---------- Takvim (bu ay) ----------
                if "tarih" in df.columns or "vade" in df.columns:
                    ay_basi = today.replace(day=1)
                    ay_sonu = (ay_basi + pd.DateOffset(months=1)).date() - timedelta(days=1)
                    takvim = defaultdict(lambda: {"tahsilat": 0.0, "odeme": 0.0, "gecikmis": 0.0, "adet": 0, "detaylar": []})
                    if "tarih" in df.columns:
                        sub = df.dropna(subset=["tarih"])
                        sub = sub[(sub["tarih"].dt.date >= ay_basi) & (sub["tarih"].dt.date <= ay_sonu)]
                        for _, r in sub.iterrows():
                            d = r["tarih"].day
                            t = abs(float(r.get("tutar_tl", 0) or 0))
                            key = "tahsilat" if r.get("_is_in", True) else "odeme"
                            takvim[d][key] += t
                            takvim[d]["adet"] += 1
                            takvim[d]["detaylar"].append({
                                "tutar": round(t, 2), "yon": "in" if r.get("_is_in", True) else "out",
                                "cari": (s(r.get("cari")) or s(r.get("odeme_turu")))[:30]
                            })
                    if "vade" in df.columns:
                        vsub = df.dropna(subset=["vade"]) if "vade" in df.columns else pd.DataFrame()
                        if len(vsub):
                            vsub = vsub[(vsub["vade"].dt.date >= ay_basi) & (vsub["vade"].dt.date <= ay_sonu)]
                            for _, r in vsub.iterrows():
                                d = r["vade"].day
                                t = abs(float(r.get("tutar_tl", 0) or 0))
                                gec = r["vade"].date() < today
                                if gec:
                                    takvim[d]["gecikmis"] += t
                                else:
                                    key = "tahsilat" if r.get("_is_in", True) else "odeme"
                                    takvim[d][key] += t
                                takvim[d]["adet"] += 1
                                takvim[d]["detaylar"].append({
                                    "tutar": round(t, 2),
                                    "yon": "in" if r.get("_is_in", True) else "out",
                                    "cari": (s(r.get("cari")) or s(r.get("odeme_turu")))[:30],
                                    "bekleyen": True, "gecikmis": gec
                                })
                    data["takvim"] = [
                        {"gun": int(g), "tahsilat": round(v["tahsilat"], 2), "odeme": round(v["odeme"], 2),
                         "gecikmis": round(v["gecikmis"], 2), "adet": v["adet"],
                         "detaylar": v["detaylar"][:5]}
                        for g, v in sorted(takvim.items())
                    ]
                    data["takvim_ay"] = today.strftime("%B %Y")
                    data["takvim_ay_no"] = today.month
                    data["takvim_yil"] = today.year
                    data["takvim_gun_sayisi"] = ay_sonu.day
                    data["takvim_baslangic"] = ay_basi.weekday()
                    data["bugun"] = today.day

    except Exception as e:
        log.error(f"VERI_GIRISI: {e}\n{traceback.format_exc()}")

    # ---------- KASA DURUM ----------
    try:
        if "KASA DURUM" in xl.sheet_names:
            kd = xl.parse("KASA DURUM", header=0)
            kd.columns = [str(c).strip() for c in kd.columns]

            # -- Planlanan Hakedisler bolum siniri (ham okuma ile tespit) --
            # kd_raw satir 0 = baslik satiri -> kd satir 0 = kd_raw satir 1
            kd_raw = xl.parse("KASA DURUM", header=None)
            hakedis_kd_cutoff = None   # kd (header=0) index'inde kredi karti siniri
            hakedis_hdr_raw   = None   # kd_raw'da hakedis alt-baslik satiri pozisyonu
            for raw_i in range(len(kd_raw)):
                row_str = " ".join(
                    str(v).upper() for v in kd_raw.iloc[raw_i]
                    if pd.notna(v) and str(v).strip()
                )
                if "PLANLANAN" in row_str and "HAKED" in row_str:
                    hakedis_kd_cutoff = raw_i - 1
                    for j in range(raw_i + 1, min(raw_i + 6, len(kd_raw))):
                        hrow = " ".join(
                            str(v).upper() for v in kd_raw.iloc[j]
                            if pd.notna(v) and str(v).strip()
                        )
                        if "PROJE" in hrow or "BEKLENE" in hrow:
                            hakedis_hdr_raw = j
                            break
                    if hakedis_hdr_raw is None:
                        hakedis_hdr_raw = raw_i + 1
                    break

            # Sol blok: banka hesaplari (banka x para birimi)
            banka_col = find_col(kd.columns, ["BANKA ADI"]) or find_col(kd.columns, ["BANKA"])
            hes_col = find_col(kd.columns, ["HESAP", "TÜR"]) or find_col(kd.columns, ["HESAP", "TUR"])
            bak_col = find_col(kd.columns, ["BAKİYE", "TL"]) or find_col(kd.columns, ["BAKIYE", "TL"])
            kul_col = find_col(kd.columns, ["KULLANILAB"])
            blk_col = find_col(kd.columns, ["BLOKE"])
            tem_col = find_col(kd.columns, ["TAHSİL", "ÇEK"]) or find_col(kd.columns, ["TAHSIL", "CEK"])

            if banka_col and bak_col:
                pick = [banka_col, bak_col]
                names = ["banka", "bakiye_tl"]
                for col, nm in [(hes_col, "hesap_turu"), (kul_col, "kullanilabilir"),
                                (blk_col, "bloke"), (tem_col, "tahsil_cek")]:
                    if col: pick.append(col); names.append(nm)
                kd2 = kd[pick].copy()
                kd2.columns = names
                kd2 = kd2[kd2["banka"].notna() & ~kd2["banka"].astype(str).str.upper().str.contains("TOPLAM", na=False)]
                for c in ["bakiye_tl", "kullanilabilir", "bloke", "tahsil_cek"]:
                    if c in kd2.columns:
                        kd2[c] = pd.to_numeric(kd2[c], errors="coerce").fillna(0).round(2)
                    else:
                        kd2[c] = 0.0
                if "hesap_turu" not in kd2.columns:
                    kd2["hesap_turu"] = "TL"
                kd2["hesap_turu"] = kd2["hesap_turu"].fillna("TL").astype(str).str.strip().str.upper()
                kd2["hesap_turu"] = kd2["hesap_turu"].replace({"EURO": "EUR", "": "TL"})
                kd2 = kd2[kd2["hesap_turu"].isin(["TL", "USD", "EUR"])]
                kd2["banka"] = kd2["banka"].astype(str).str.strip()
                kd2 = kd2[kd2["bakiye_tl"] != 0]

                grouped = []
                for banka, grp in kd2.groupby("banka", sort=False):
                    para_birimleri = [
                        {"tur": str(r["hesap_turu"]),
                         "bakiye_tl": round(float(r["bakiye_tl"]), 2),
                         "kullanilabilir": round(float(r["kullanilabilir"]), 2),
                         "bloke": round(float(r["bloke"]), 2)}
                        for _, r in grp.iterrows()
                    ]
                    grp_tl = grp[grp["hesap_turu"] == "TL"]
                    grouped.append({
                        "banka": str(banka),
                        "bakiye_tl": round(float(grp_tl["bakiye_tl"].sum()), 2),
                        "kullanilabilir": round(float(grp_tl["kullanilabilir"].sum()), 2),
                        "bloke": round(float(grp_tl["bloke"].sum()), 2),
                        "tahsil_cek": round(float(grp["tahsil_cek"].sum()), 2),
                        "para_birimleri": para_birimleri,
                    })
                grouped.sort(key=lambda b: b["bakiye_tl"], reverse=True)
                data["bakiyeler"] = clean(grouped)
                kd2_tl = kd2[kd2["hesap_turu"] == "TL"]
                data["kpi"].update({
                    "toplam_bakiye": round(float(kd2_tl["bakiye_tl"].sum()), 2),
                    "toplam_kullanilabilir": round(float(kd2_tl["kullanilabilir"].sum()), 2),
                    "toplam_bloke": round(float(kd2_tl["bloke"].sum()), 2),
                    "toplam_tahsil_cek": round(float(kd2_tl["tahsil_cek"].sum()), 2),
                    "banka_sayisi": int(len(grouped)),
                })

            # Sag blok: kredi kartlari
            kk_banka = find_col(kd.columns, ["BANKA ADI.1"]) or find_col(kd.columns, ["BANKA", ".1"])
            kk_kul = find_col(kd.columns, ["KULLAN", "LİMİT"]) or find_col(kd.columns, ["KREDİ", "LİMİT"])
            kk_lim = None
            for c in kd.columns:
                u = str(c).upper().strip()
                if u == "LİMİT" or u == "LIMIT":
                    kk_lim = c; break
            kk_kes = find_col(kd.columns, ["HESAP", "KESİM"]) or find_col(kd.columns, ["HESAP", "KESIM"])
            kk_ode = find_col(kd.columns, ["ÖDEME", "TARİH"]) or find_col(kd.columns, ["ODEME", "TARIH"])
            kk_borc = find_col(kd.columns, ["BORÇ"]) or find_col(kd.columns, ["BORC"])
            kk_asg = find_col(kd.columns, ["ASGARİ"]) or find_col(kd.columns, ["ASGARI"])

            if kk_banka and (kk_lim or kk_borc):
                kk_cols = [kk_banka]
                kk_names = ["banka"]
                for col, nm in [(kk_kul, "kullanim"), (kk_lim, "limit"),
                                (kk_kes, "hesap_kesim"), (kk_ode, "odeme_tarihi"),
                                (kk_borc, "borc"), (kk_asg, "asgari")]:
                    if col: kk_cols.append(col); kk_names.append(nm)
                kk = kd[kk_cols].copy()
                kk.columns = kk_names
                # Hakedis bolumu satirlarini kredi kartindan disla
                if hakedis_kd_cutoff is not None:
                    kk = kk[kk.index < hakedis_kd_cutoff]
                kk = kk[kk["banka"].notna() & ~kk["banka"].astype(str).str.upper().str.contains("TOPLAM|YETKI|YETKİ", na=False)]
                for c in ["kullanim", "limit", "borc", "asgari"]:
                    if c in kk.columns:
                        kk[c] = pd.to_numeric(kk[c], errors="coerce").fillna(0).round(2)
                for c in ["hesap_kesim", "odeme_tarihi"]:
                    if c in kk.columns:
                        kk[c] = pd.to_datetime(kk[c], errors="coerce")
                kk = kk[(kk.get("limit", 0) > 0) | (kk.get("borc", 0) > 0)]

                kredi_kartlari = []
                for _, r in kk.iterrows():
                    lim = float(r.get("limit", 0) or 0)
                    kul = float(r.get("kullanim", 0) or 0)
                    borc = float(r.get("borc", 0) or 0)
                    ode = r.get("odeme_tarihi")
                    kes = r.get("hesap_kesim")
                    gun_fark = None
                    if pd.notna(ode):
                        gun_fark = (ode.date() - today).days
                    kredi_kartlari.append({
                        "banka": str(r["banka"]).strip(),
                        "limit": round(lim, 2),
                        "kullanim": round(kul, 2),
                        "kalan": round(max(lim - kul, 0), 2),
                        "kullanim_oran": round((kul / lim * 100) if lim > 0 else 0, 1),
                        "borc": round(borc, 2),
                        "asgari": round(float(r.get("asgari", 0) or 0), 2),
                        "hesap_kesim": kes.strftime("%d.%m.%Y") if pd.notna(kes) else "",
                        "odeme_tarihi": ode.strftime("%d.%m.%Y") if pd.notna(ode) else "",
                        "gun_fark": int(gun_fark) if gun_fark is not None else None,
                        "gecikmis": bool(pd.notna(ode) and ode.date() < today and borc > 0),
                    })
                kredi_kartlari.sort(key=lambda k: k["borc"], reverse=True)
                data["kredi_kartlari"] = clean(kredi_kartlari)
                data["kpi"].update({
                    "kk_toplam_limit": round(sum(k["limit"] for k in kredi_kartlari), 2),
                    "kk_toplam_kullanim": round(sum(k["kullanim"] for k in kredi_kartlari), 2),
                    "kk_toplam_borc": round(sum(k["borc"] for k in kredi_kartlari), 2),
                    "kk_sayisi": len(kredi_kartlari),
                })

            # -- Planlanan Hakedisler ayri bolum --
            if hakedis_hdr_raw is not None:
                hak_hdrs = [str(v).strip() for v in kd_raw.iloc[hakedis_hdr_raw].tolist()]
                log.info(f"  Hakedis baslik satiri ({hakedis_hdr_raw}): {[h for h in hak_hdrs if h and h != 'nan']}")
                hak_df = kd_raw.iloc[hakedis_hdr_raw + 1:].copy()
                hak_df.columns = hak_hdrs[:len(hak_df.columns)]
                hak_df = hak_df.reset_index(drop=True)

                hcols = hak_df.columns.tolist()
                proje_col   = find_col(hcols, ["PROJE"])
                acik_col    = find_col(hcols, ["AÇIKLAMA"]) or find_col(hcols, ["ACIKLAMA"])
                bek_col     = find_col(hcols, ["BEKLENE"])
                pltarih_col = find_col(hcols, ["PLANLANAN"])
                engec_col   = find_col(hcols, ["EN", "GEÇ"]) or find_col(hcols, ["EN", "GEC"])
                yap_col     = find_col(hcols, ["YAPILAN"])
                kal_col     = find_col(hcols, ["KALAN"])
                dur_col     = find_col(hcols, ["DURUM"])
                log.info(f"  Hakedis kolonlar -> proje={proje_col} beklenen={bek_col} "
                         f"tarih={pltarih_col} en_gec={engec_col} durum={dur_col}")

                for c in [bek_col, yap_col, kal_col]:
                    if c:
                        hak_df[c] = pd.to_numeric(hak_df[c], errors="coerce").fillna(0)
                for c in [pltarih_col, engec_col]:
                    if c:
                        hak_df[c] = pd.to_datetime(hak_df[c], errors="coerce")

                # Para birimi etiketi veya baska bolume ait satirlari disla
                _PARA_BIRIMLERI = {"TL", "USD", "EUR", "EURO", "DOLAR", "DOVIZ", "DÖVİZ"}
                _STOP_KEYS = {"TOPLAM", "GÖSTERGE", "GOSTERGE", "GENEL", "OZET", "ÖZET"}

                hakedisler = []
                for _, r in hak_df.iterrows():
                    proje    = s(r.get(proje_col, "")) if proje_col else ""
                    aciklama = s(r.get(acik_col, "")) if acik_col else ""
                    beklenen = float(r.get(bek_col, 0) or 0) if bek_col else 0.0

                    # Bos satirlari atla
                    if not proje and not beklenen:
                        continue

                    # TOPLAM / GÖSTERGE satirinda dur
                    tum_metin = (proje + " " + aciklama).upper().strip()
                    if any(k in tum_metin for k in _STOP_KEYS):
                        break

                    # Para birimi / banka veri satirlarini atla
                    if proje.upper().strip() in _PARA_BIRIMLERI:
                        continue
                    en_gec_dt = r.get(engec_col) if engec_col else None
                    plan_dt   = r.get(pltarih_col) if pltarih_col else None
                    durum     = s(r.get(dur_col, "")) if dur_col else ""
                    is_done   = any(x in durum.upper() for x in ["TAMAMLANDI", "TAMAMLANMIS", "TAMAMLANDI"])
                    gun_fark  = None
                    gecikmis  = False
                    if engec_col and pd.notna(en_gec_dt):
                        try:
                            gun_fark = int((en_gec_dt.date() - today).days)
                            gecikmis = gun_fark < 0 and not is_done
                        except Exception:
                            pass
                    hakedisler.append({
                        "proje":            proje,
                        "aciklama":         aciklama,
                        "beklenen_tutar":   round(beklenen, 2),
                        "planlanan_tarih":  plan_dt.strftime("%d.%m.%Y") if (pltarih_col and pd.notna(plan_dt)) else "",
                        "en_gec_tarih":     en_gec_dt.strftime("%d.%m.%Y") if (engec_col and pd.notna(en_gec_dt)) else "",
                        "yapilan_tahsilat": round(float(r.get(yap_col, 0) or 0), 2) if yap_col else 0.0,
                        "kalan_alacak":     round(float(r.get(kal_col, 0) or 0), 2) if kal_col else 0.0,
                        "durum":            durum,
                        "gun_fark":         gun_fark,
                        "gecikmis":         gecikmis,
                    })

                data["hakedisler"] = clean(hakedisler)
                if hakedisler:
                    data["kpi"].update({
                        "hak_toplam_beklenen": round(sum(h["beklenen_tutar"] for h in hakedisler), 2),
                        "hak_toplam_kalan":    round(sum(h["kalan_alacak"]    for h in hakedisler), 2),
                        "hak_gecikmis_adet":   sum(1 for h in hakedisler if h["gecikmis"]),
                        "hak_sayisi":          len(hakedisler),
                    })
                    log.info(f"  Hakedisler: {len(hakedisler)} satir, "
                             f"beklenen {data['kpi']['hak_toplam_beklenen']:,.0f} TL, "
                             f"kalan {data['kpi']['hak_toplam_kalan']:,.0f} TL")

    except Exception as e:
        log.error(f"KASA DURUM: {e}\n{traceback.format_exc()}")

    # ---------- VERILEN / ALINAN CEKLER ----------
    def _parse_cek_sheet(sheet_name, header_row, yon):
        rows_out = []
        try:
            if sheet_name not in xl.sheet_names:
                return rows_out
            cdf = xl.parse(sheet_name, header=header_row)
            cdf.columns = [str(c).strip() for c in cdf.columns]
            vade_col = find_col(cdf.columns, ["VADE"])
            tur_col = find_col(cdf.columns, ["TÜRÜ"]) or find_col(cdf.columns, ["TURU"])
            banka_col = find_col(cdf.columns, ["KEŞİDECİ", "BANKA"]) or find_col(cdf.columns, ["KESIDECI", "BANKA"]) or find_col(cdf.columns, ["BANKA"])
            sube_col = find_col(cdf.columns, ["ŞUBE"]) or find_col(cdf.columns, ["SUBE"])
            tutar_col = find_col(cdf.columns, ["TUTAR"])
            no_col = find_col(cdf.columns, ["ÇEK", "NO"]) or find_col(cdf.columns, ["CEK", "NO"])
            ciro_col = find_col(cdf.columns, ["CİRO"]) or find_col(cdf.columns, ["CIRO"])
            ack_col = find_col(cdf.columns, ["AÇIKLAMA"]) or find_col(cdf.columns, ["ACIKLAMA"])
            ver_col = find_col(cdf.columns, ["VERİLİŞ"]) or find_col(cdf.columns, ["VERILIS"])
            dur_col = find_col(cdf.columns, ["DURUM"])
            if not (vade_col and tutar_col):
                return rows_out
            sub = cdf[cdf[tutar_col].notna()].copy()
            sub[vade_col] = pd.to_datetime(sub[vade_col], errors="coerce")
            if ver_col: sub[ver_col] = pd.to_datetime(sub[ver_col], errors="coerce")
            sub[tutar_col] = pd.to_numeric(sub[tutar_col], errors="coerce")
            sub = sub[sub[tutar_col].notna() & (sub[tutar_col] != 0)]
            for _, r in sub.iterrows():
                v = r.get(vade_col)
                durum = str(r.get(dur_col, "") or "").strip().upper() if dur_col else ""
                kapali = bool(durum and ("ÖDEN" in durum or "ODEN" in durum or "TAHSİL" in durum or "TAHSIL" in durum or "ÇIKIŞ" in durum))
                gec = bool(pd.notna(v) and v.date() < today and not kapali)
                rows_out.append({
                    "vade": v.strftime("%d.%m.%Y") if pd.notna(v) else "",
                    "vade_iso": v.strftime("%Y-%m-%d") if pd.notna(v) else "",
                    "tur": str(r.get(tur_col, "") or "").strip() if tur_col else "",
                    "banka": str(r.get(banka_col, "") or "").strip() if banka_col else "",
                    "sube": str(r.get(sube_col, "") or "").strip() if sube_col else "",
                    "tutar": round(abs(float(r[tutar_col])), 2),
                    "cek_no": str(r.get(no_col, "") or "").strip() if no_col else "",
                    "ciro": str(r.get(ciro_col, "") or "").strip() if ciro_col else "",
                    "aciklama": str(r.get(ack_col, "") or "").strip() if ack_col else "",
                    "veris_tarihi": (r[ver_col].strftime("%d.%m.%Y") if (ver_col and pd.notna(r.get(ver_col))) else ""),
                    "durum": durum,
                    "yon": yon,
                    "kapali": kapali,
                    "gecikmis": gec,
                    "gun_fark": int((v.date() - today).days) if pd.notna(v) else None,
                })
        except Exception as e:
            log.error(f"{sheet_name}: {e}")
        return rows_out

    cekler_verilen = _parse_cek_sheet("VERİLEN ÇEK LİSTESİ", 2, "out")
    cekler_alinan = _parse_cek_sheet("ALINAN ÇEKLER LİSTESİ", 1, "in")
    data["cekler_verilen"] = clean(cekler_verilen)
    data["cekler_alinan"] = clean(cekler_alinan)

    try:
        bekleyen_v = [c for c in cekler_verilen if not c["kapali"]]
        bekleyen_a = [c for c in cekler_alinan if not c["kapali"]]
        gec_v = [c for c in bekleyen_v if c["gecikmis"]]
        gec_a = [c for c in bekleyen_a if c["gecikmis"]]
        data["kpi"].update({
            "cek_verilen_sayi": len(cekler_verilen),
            "cek_alinan_sayi": len(cekler_alinan),
            "cek_verilen_bekleyen": len(bekleyen_v),
            "cek_alinan_bekleyen": len(bekleyen_a),
            "cek_verilen_bekleyen_tutar": round(sum(c["tutar"] for c in bekleyen_v), 2),
            "cek_alinan_bekleyen_tutar": round(sum(c["tutar"] for c in bekleyen_a), 2),
            "cek_verilen_gecikmis": len(gec_v),
            "cek_alinan_gecikmis": len(gec_a),
        })
    except Exception as e:
        log.error(f"CEK ozet: {e}")

    # ---------- BAKIYE_LOG ----------
    try:
        if "BAKIYE_LOG" in xl.sheet_names:
            bl = xl.parse("BAKIYE_LOG", header=None)
            hdr_row = None
            for i in range(min(len(bl), 10)):
                row = [str(x).upper() for x in bl.iloc[i].tolist() if pd.notna(x)]
                if any("TARİH" in x or "TARIH" in x for x in row) and any("BANKA" in x for x in row):
                    hdr_row = i
                    break
            if hdr_row is not None and len(bl) > hdr_row + 1:
                headers = [str(x).strip() for x in bl.iloc[hdr_row].tolist()]
                bl = bl.iloc[hdr_row + 1:].copy()
                bl.columns = headers[:len(bl.columns)]
                t_col = next((c for c in bl.columns if "TARİH" in str(c).upper() or "TARIH" in str(c).upper()), None)
                b_col = next((c for c in bl.columns if str(c).upper().strip() == "BANKA"), None)
                h_col = next((c for c in bl.columns if "HESAP" in str(c).upper()), None)
                bak_col = next((c for c in bl.columns if "BAKİYE" in str(c).upper() or "BAKIYE" in str(c).upper()), None)
                if t_col and b_col and bak_col:
                    bl[t_col] = pd.to_datetime(bl[t_col], errors="coerce")
                    bl[bak_col] = pd.to_numeric(bl[bak_col], errors="coerce")
                    bl = bl.dropna(subset=[t_col, bak_col])
                    bl = bl.sort_values(t_col).tail(500)
                    data["bakiye_log"] = [
                        {
                            "tarih": r[t_col].strftime("%Y-%m-%d"),
                            "banka": str(r[b_col] or "").strip(),
                            "hesap_turu": str(r[h_col] or "TL").strip().upper().replace("EURO", "EUR") if h_col else "TL",
                            "bakiye": round(float(r[bak_col]), 2),
                        }
                        for _, r in bl.iterrows() if str(r[b_col] or "").strip()
                    ]
    except Exception as e:
        log.error(f"BAKIYE_LOG: {e}")

    # ---------- KUR_GECMISI ----------
    try:
        if "KUR_GECMISI" in xl.sheet_names:
            kur = xl.parse("KUR_GECMISI", header=2)
            kur.columns = [str(c).strip() for c in kur.columns]
            rn = {}
            for c in kur.columns:
                u = c.upper()
                if "TARİH" in u or "TARIH" in u: rn[c] = "tarih"
                elif "EUR" in u: rn[c] = "eur"
                elif "USD" in u: rn[c] = "usd"
            kur = kur.rename(columns=rn)
            if all(c in kur.columns for c in ["tarih", "eur", "usd"]):
                kur["tarih"] = pd.to_datetime(kur["tarih"], errors="coerce")
                kur = kur.dropna(subset=["tarih"])
                kur["eur"] = pd.to_numeric(kur["eur"], errors="coerce")
                kur["usd"] = pd.to_numeric(kur["usd"], errors="coerce")
                kur = kur.dropna(subset=["eur", "usd"]).sort_values("tarih").tail(60)
                data["kurlar"] = [
                    {"tarih": r["tarih"].strftime("%d.%m.%y"),
                     "eur": round(float(r["eur"]), 4),
                     "usd": round(float(r["usd"]), 4)}
                    for _, r in kur.iterrows()
                ]
                if len(kur) > 0:
                    last = kur.iloc[-1]
                    data["kpi"].update({
                        "son_eur": round(float(last["eur"]), 4),
                        "son_usd": round(float(last["usd"]), 4),
                    })
                    if len(kur) > 1:
                        prev = kur.iloc[-2]
                        data["kpi"]["eur_degisim"] = round(float(last["eur"]) - float(prev["eur"]), 4)
                        data["kpi"]["usd_degisim"] = round(float(last["usd"]) - float(prev["usd"]), 4)
    except Exception as e:
        log.error(f"KUR: {e}")

    data["parse_sure"] = round(time.time() - t0, 2)
    log.info(f"Parse OK ({data['parse_sure']}s) -- "
             f"{data['kpi'].get('islem_sayisi', 0)} islem, "
             f"{len(data['bakiyeler'])} banka, "
             f"{len(data['kredi_kartlari'])} kart, "
             f"{len(data.get('hakedisler', []))} hakedis, "
             f"{len(data['cekler_verilen'])} verilen cek, "
             f"{len(data['cekler_alinan'])} alinan cek, "
             f"{len(data['bekleyen'])} bekleyen")
    xl.close()
    return json.loads(json.dumps(data, cls=SafeEncoder, ensure_ascii=False, default=str))


# ---------- WebSocket connection manager ----------
class CM:
    def __init__(self): self.ws = []
    async def connect(self, w): await w.accept(); self.ws.append(w)
    def disconnect(self, w):
        if w in self.ws: self.ws.remove(w)
    async def broadcast(self, d):
        msg = json.dumps(d, cls=SafeEncoder, ensure_ascii=False, default=str)
        dead = []
        for w in self.ws:
            try: await w.send_text(msg)
            except: dead.append(w)
        for w in dead:
            if w in self.ws: self.ws.remove(w)
        if self.ws: log.info(f"  {len(self.ws)} istemciye gonderildi")


mgr = CM()
latest = {}
excel_path = ""
_loop = None


class FH(FileSystemEventHandler):
    def __init__(self): self._t = 0
    def on_modified(self, e):
        if e.is_directory: return
        fn = os.path.basename(e.src_path).lower()
        if fn.startswith("~$") or not fn.endswith((".xlsm", ".xlsx", ".xls")): return
        now = time.time()
        if now - self._t < 3: return
        self._t = now
        log.info("Dosya degisti -> yeniden okunuyor")
        if _loop and _loop.is_running():
            asyncio.run_coroutine_threadsafe(_reload(), _loop)


async def _reload():
    global latest
    await asyncio.sleep(1.5)
    try:
        latest = parse_excel(excel_path)
        latest["tip"] = "guncelleme"
        await mgr.broadcast(latest)
    except Exception as e:
        log.error(f"Reload: {e}")


# ---------- Restart flag watcher ----------
def _restart_watcher():
    flag = Path(__file__).parent / "restart.flag"
    stop = Path(__file__).parent / "stop.flag"
    while True:
        if stop.exists() and sys.platform == "win32":
            try:
                stop.unlink()
            except Exception:
                pass
            log.info("stop.flag algilandi — yerel sunucu kapatiliyor...")
            time.sleep(1)
            sys.exit(0)
        if flag.exists():
            try:
                flag.unlink()
            except Exception:
                pass
            log.info("restart.flag algilandi — yeniden baslatiliyor...")
            time.sleep(1)
            os.execv(sys.executable, [sys.executable] + sys.argv)
        time.sleep(5)


# ---------- mDNS (kasa.local) ----------
def _get_local_ip():
    """Yerel IP adresini tespit et — birden fazla yontem dener."""
    import socket
    # Yontem 1: UDP socket ile default route uzerinden IP al
    for target in ["8.8.8.8", "1.1.1.1", "192.168.1.1"]:
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.settimeout(1)
            s.connect((target, 80))
            ip = s.getsockname()[0]
            s.close()
            if ip and not ip.startswith("127."):
                return ip
        except Exception:
            pass
    # Yontem 2: hostname uzerinden
    try:
        return socket.gethostbyname(socket.gethostname())
    except Exception:
        return "0.0.0.0"

def _mdns_register(port: int):
    try:
        import socket
        from zeroconf import Zeroconf, ServiceInfo
        local_ip = _get_local_ip()
        log.info(f"mDNS icin tespit edilen IP: {local_ip}")
        info = ServiceInfo(
            "_http._tcp.local.",
            "Kasa Dashboard._http._tcp.local.",
            addresses=[socket.inet_aton(local_ip)],
            port=port,
            properties={"path": "/"},
            server="kasa.local.",
        )
        zc = Zeroconf()
        zc.register_service(info)
        log.info(f"mDNS: http://kasa.local:{port} ({local_ip})")
        while True:
            time.sleep(3600)
    except Exception as e:
        log.warning(f"mDNS kurulamadi: {e}")


# ---------- FastAPI ----------
BASE = Path(__file__).parent
security = HTTPBasic()

def verify(credentials: HTTPBasicCredentials = Depends(security)):
    ok_user = secrets.compare_digest(credentials.username, os.environ.get("KASA_USER", "CHANGE-ME-USER"))
    ok_pass = secrets.compare_digest(credentials.password, os.environ.get("KASA_PASS", "CHANGE-ME-PASS"))
    if not (ok_user and ok_pass):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED,
                            detail="Yetkisiz", headers={"WWW-Authenticate": "Basic"})


@asynccontextmanager
async def lifespan(app: FastAPI):
    global _loop
    _loop = asyncio.get_running_loop()
    threading.Thread(target=_restart_watcher, daemon=True).start()
    yield


app = FastAPI(lifespan=lifespan)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


@app.get("/api/data")
async def api_data(credentials: HTTPBasicCredentials = Depends(verify)):
    try:
        return JSONResponse(latest if latest else parse_excel(excel_path))
    except Exception as e:
        return JSONResponse({"hata": str(e), "trace": traceback.format_exc()}, status_code=500)


@app.get("/api/refresh")
async def api_refresh(credentials: HTTPBasicCredentials = Depends(verify)):
    global latest
    try:
        latest = parse_excel(excel_path)
        if _loop:
            asyncio.run_coroutine_threadsafe(mgr.broadcast(latest), _loop)
        return JSONResponse({"durum": "ok", "guncelleme": latest.get("guncelleme", "-")})
    except Exception as e:
        return JSONResponse({"hata": str(e)}, status_code=500)


@app.get("/api/health")
async def health():
    return {"durum": "aktif", "ws": len(mgr.ws),
            "son": latest.get("guncelleme", "-"),
            "dosya": os.path.basename(excel_path)}


@app.get("/api/test")
async def test():
    try:
        d = parse_excel(excel_path)
        t = json.dumps(d, cls=SafeEncoder, ensure_ascii=False, default=str)
        return PlainTextResponse(
            f"OK - {len(t)} byte\n"
            f"KPI: {list(d.get('kpi', {}).keys())}\n"
            f"Odeme turleri: {len(d.get('odeme_turleri', []))}\n"
            f"Aylik: {len(d.get('aylik', []))}\n"
            f"Bakiye: {len(d.get('bakiyeler', []))}\n"
            f"Kredi Kartlari: {len(d.get('kredi_kartlari', []))}\n"
            f"Hakedisler: {len(d.get('hakedisler', []))}\n"
            f"Verilen cek: {len(d.get('cekler_verilen', []))}\n"
            f"Alinan cek: {len(d.get('cekler_alinan', []))}\n"
            f"Bakiye log: {len(d.get('bakiye_log', []))}\n"
            f"Bekleyen: {len(d.get('bekleyen', []))}\n"
            f"Takvim gunu: {len(d.get('takvim', []))}\n"
            f"Islem: {len(d.get('islemler', []))}\n"
            f"Kur: {len(d.get('kurlar', []))}"
        )
    except Exception as e:
        return PlainTextResponse(f"HATA: {e}\n{traceback.format_exc()}", status_code=500)


@app.get("/chart.min.js")
async def chartjs():
    p = BASE / "chart.min.js"
    if p.exists(): return FileResponse(p, media_type="application/javascript")
    return PlainTextResponse("// Chart.js not found", status_code=404)


@app.websocket("/ws")
async def ws_ep(ws: WebSocket):
    await mgr.connect(ws)
    if latest:
        try:
            await ws.send_text(json.dumps({**latest, "tip": "ilk"},
                                          cls=SafeEncoder, ensure_ascii=False, default=str))
        except: pass
    try:
        while True: await ws.receive_text()
    except WebSocketDisconnect:
        mgr.disconnect(ws)


@app.get("/", response_class=HTMLResponse)
async def index(credentials: HTTPBasicCredentials = Depends(verify)):
    p = BASE / "dashboard.html"
    if not p.exists():
        return HTMLResponse("<h1>dashboard.html bulunamadi</h1>", 404)
    return p.read_text(encoding="utf-8")


def main():
    import sys
    # pythonw.exe ile sys.stdout/stderr None olur, uvicorn crash yapar -- log dosyasina yonlendir
    if sys.stdout is None:
        sys.stdout = open(_log_file, "a", encoding="utf-8", buffering=1)
    if sys.stderr is None:
        sys.stderr = sys.stdout

    global excel_path, latest
    pa = argparse.ArgumentParser()
    pa.add_argument("--file", required=True)
    pa.add_argument("--port", type=int, default=8765)
    pa.add_argument("--host", default="0.0.0.0")
    pa.add_argument("--polling", action="store_true")
    pa.add_argument("--user", default=None, help="Basic Auth kullanici - KASA_USER env var ile de verilebilir")
    pa.add_argument("--password", default=None, help="Basic Auth sifresi - KASA_PASS env var ile de verilebilir")
    a = pa.parse_args()
    if a.user: os.environ["KASA_USER"] = a.user
    if a.password: os.environ["KASA_PASS"] = a.password
    excel_path = os.path.abspath(a.file)
    if not os.path.exists(excel_path):
        log.error(f"Dosya yok: {excel_path}")
        raise SystemExit(1)
    log.info("Ilk yukleme...")
    try:
        latest = parse_excel(excel_path)
        latest["tip"] = "ilk"
    except Exception as e:
        log.error(f"Hata: {e}")
        latest = {"hata": str(e), "kpi": {}}
    wd = str(Path(excel_path).parent)
    h = FH()
    ob = PollingObserver(timeout=3)
    ob.schedule(h, wd, recursive=False)
    ob.start()
    threading.Thread(target=_mdns_register, args=(a.port,), daemon=True).start()
    log.info(f"Dashboard: http://localhost:{a.port}")
    try:
        uvicorn.run(app, host=a.host, port=a.port, log_level="info")
    finally:
        ob.stop()
        ob.join()


if __name__ == "__main__":
    main()
