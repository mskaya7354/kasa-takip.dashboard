"""
Kasa Takip - Dashboard Sunucusu
"""
import asyncio, argparse, json, logging, math, os, time, traceback
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import uvicorn
from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, PlainTextResponse, FileResponse
from watchdog.events import FileSystemEventHandler
from watchdog.observers.polling import PollingObserver

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("kasa")

class SafeEncoder(json.JSONEncoder):
    def default(self, o):
        if isinstance(o, float) and (math.isnan(o) or math.isinf(o)): return None
        if hasattr(o, 'isoformat'): return o.isoformat()
        return str(o)

def safe_val(v):
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)): return None
    return v

def clean(records):
    return [{k: safe_val(v) for k, v in r.items()} for r in records]

def parse_excel(path):
    log.info(f"Excel okunuyor: {path}")
    t0 = time.time()
    xl = pd.ExcelFile(path, engine="openpyxl")
    data = {"guncelleme": datetime.now().strftime("%d.%m.%Y %H:%M:%S"), "dosya": os.path.basename(path), "kpi": {}}

    try:
        if "VERI_GIRISI" in xl.sheet_names:
            df = xl.parse("VERI_GIRISI", header=3)
            df.columns = df.columns.str.strip()
            cm = {}
            for c in df.columns:
                u = c.upper()
                if "TUTAR" in u and "TL" in u: cm[c]="tutar_tl"
                elif any(x in u for x in ["ÖDEME TÜRÜ","ODEME TURU"]): cm[c]="odeme_turu"
                elif any(x in u for x in ["ÖDEME ARACI","ODEME ARACI"]): cm[c]="odeme_araci"
                elif "DURUM" in u: cm[c]="durum"
                elif any(x in u for x in ["TARİH","TARIH"]): cm[c]="tarih"
            df = df.rename(columns=cm)
            if "tutar_tl" in df.columns:
                df["tutar_tl"] = pd.to_numeric(df["tutar_tl"], errors="coerce")
                df = df[df["tutar_tl"].notna() & (df["tutar_tl"] != 0)]
                data["kpi"].update({"toplam_tutar":round(float(df["tutar_tl"].sum()),2),"islem_sayisi":len(df),"ortalama_tutar":round(float(df["tutar_tl"].mean()),2),"en_buyuk":round(float(df["tutar_tl"].max()),2),"bekleyen_adet":int(df["durum"].eq("BEKLİYOR").sum()) if "durum" in df.columns else 0})
                if "odeme_turu" in df.columns:
                    g=df.groupby("odeme_turu")["tutar_tl"].agg(["sum","count"]).reset_index();g.columns=["odeme_turu","tutar","adet"];g["tutar"]=g["tutar"].round(2);g["adet"]=g["adet"].astype(int)
                    data["odeme_turleri"]=clean(g.sort_values("tutar",ascending=False).to_dict("records"))
                if "odeme_araci" in df.columns:
                    g=df.groupby("odeme_araci")["tutar_tl"].agg(["sum","count"]).reset_index();g.columns=["odeme_araci","tutar","adet"];g["tutar"]=g["tutar"].round(2);g["adet"]=g["adet"].astype(int)
                    data["odeme_araclari"]=clean(g.sort_values("tutar",ascending=False).head(12).to_dict("records"))
                if "tarih" in df.columns:
                    df["tarih"]=pd.to_datetime(df["tarih"],errors="coerce");df2=df.dropna(subset=["tarih"]).copy()
                    if len(df2)>0:
                        df2["ay_key"]=df2["tarih"].dt.to_period("M");df2["ay"]=df2["tarih"].dt.strftime("%b %y")
                        m=df2.groupby(["ay_key","ay"])["tutar_tl"].sum().reset_index().sort_values("ay_key").tail(18)
                        data["aylik"]=[{"ay":r["ay"],"tutar":round(float(r["tutar_tl"]),2)} for _,r in m.iterrows()]
    except Exception as e: log.error(f"VERI_GIRISI: {e}")

    try:
        if "KASA DURUM" in xl.sheet_names:
            kd=xl.parse("KASA DURUM",header=0);kd.columns=[str(c).strip() for c in kd.columns]
            if "BANKA ADI" in kd.columns and "BAKİYE TL" in kd.columns:
                kd2=kd[["BANKA ADI","BAKİYE TL","KULLANILABİLİR BAKİYE","BLOKE"]].copy();kd2.columns=["banka","bakiye_tl","kullanilabilir","bloke"]
                kd2=kd2[kd2["banka"].notna()]
                for c in ["bakiye_tl","kullanilabilir","bloke"]: kd2[c]=pd.to_numeric(kd2[c],errors="coerce").fillna(0).round(2)
                kd2=kd2[kd2["bakiye_tl"]!=0]
                data["bakiyeler"]=clean(kd2.to_dict("records"))
                data["kpi"].update({"toplam_bakiye":round(float(kd2["bakiye_tl"].sum()),2),"toplam_kullanilabilir":round(float(kd2["kullanilabilir"].sum()),2),"toplam_bloke":round(float(kd2["bloke"].sum()),2)})
    except Exception as e: log.error(f"KASA DURUM: {e}")

    try:
        if "KUR_GECMISI" in xl.sheet_names:
            kur=xl.parse("KUR_GECMISI",header=2);kur.columns=[str(c).strip() for c in kur.columns]
            rn={}
            for c in kur.columns:
                if "TARİH" in c.upper() or "TARIH" in c.upper(): rn[c]="tarih"
                elif "EUR" in c.upper(): rn[c]="eur"
                elif "USD" in c.upper(): rn[c]="usd"
            kur=kur.rename(columns=rn)
            if all(c in kur.columns for c in ["tarih","eur","usd"]):
                kur["tarih"]=pd.to_datetime(kur["tarih"],errors="coerce");kur=kur.dropna(subset=["tarih"])
                kur["eur"]=pd.to_numeric(kur["eur"],errors="coerce");kur["usd"]=pd.to_numeric(kur["usd"],errors="coerce")
                kur=kur.dropna(subset=["eur","usd"]).sort_values("tarih").tail(60)
                data["kurlar"]=[{"tarih":r["tarih"].strftime("%d.%m.%y"),"eur":round(float(r["eur"]),4),"usd":round(float(r["usd"]),4)} for _,r in kur.iterrows()]
                if len(kur)>0: s=kur.iloc[-1]; data["kpi"].update({"son_eur":round(float(s["eur"]),4),"son_usd":round(float(s["usd"]),4)})
    except Exception as e: log.error(f"KUR: {e}")

    data["parse_sure"]=round(time.time()-t0,2)
    log.info(f"Parse OK ({data['parse_sure']}s)")
    return json.loads(json.dumps(data, cls=SafeEncoder, ensure_ascii=False, default=str))

# Connection manager
class CM:
    def __init__(self): self.ws=[]
    async def connect(self,w): await w.accept(); self.ws.append(w)
    def disconnect(self,w):
        if w in self.ws: self.ws.remove(w)
    async def broadcast(self,d):
        msg=json.dumps(d,cls=SafeEncoder,ensure_ascii=False,default=str);dead=[]
        for w in self.ws:
            try: await w.send_text(msg)
            except: dead.append(w)
        for w in dead:
            if w in self.ws: self.ws.remove(w)
        if self.ws: log.info(f"  {len(self.ws)} istemciye gonderildi")

mgr=CM(); latest={}; excel_path=""; _loop=None

class FH(FileSystemEventHandler):
    def __init__(self): self._t=0
    def on_modified(self,e):
        if e.is_directory: return
        fn=os.path.basename(e.src_path).lower()
        if fn.startswith("~$") or not fn.endswith((".xlsm",".xlsx",".xls")): return
        now=time.time()
        if now-self._t<3: return
        self._t=now; log.info(f"Dosya degisti")
        if _loop and _loop.is_running(): asyncio.run_coroutine_threadsafe(_reload(),_loop)

async def _reload():
    global latest
    await asyncio.sleep(1.5)
    try: latest=parse_excel(excel_path); latest["tip"]="guncelleme"; await mgr.broadcast(latest)
    except Exception as e: log.error(f"Reload: {e}")

app=FastAPI()
app.add_middleware(CORSMiddleware,allow_origins=["*"],allow_methods=["*"],allow_headers=["*"])
BASE=Path(__file__).parent

@app.on_event("startup")
async def startup():
    global _loop; _loop=asyncio.get_running_loop()

@app.get("/api/data")
async def api_data():
    try: return JSONResponse(latest if latest else parse_excel(excel_path))
    except Exception as e: return JSONResponse({"hata":str(e)},status_code=500)

@app.get("/api/health")
async def health():
    return {"durum":"aktif","ws":len(mgr.ws),"son":latest.get("guncelleme","—"),"dosya":os.path.basename(excel_path)}

@app.get("/api/test")
async def test():
    try:
        d=parse_excel(excel_path); t=json.dumps(d,cls=SafeEncoder,ensure_ascii=False,default=str)
        return PlainTextResponse(f"OK - {len(t)} byte\nKPI: {list(d.get('kpi',{}).keys())}\nOdeme: {len(d.get('odeme_turleri',[]))}\nAylik: {len(d.get('aylik',[]))}\nBakiye: {len(d.get('bakiyeler',[]))}\nKur: {len(d.get('kurlar',[]))}")
    except Exception as e: return PlainTextResponse(f"HATA: {e}\n{traceback.format_exc()}",status_code=500)

# Chart.js'i yerel dosyadan servis et
@app.get("/chart.min.js")
async def chartjs():
    p=BASE/"chart.min.js"
    if p.exists(): return FileResponse(p, media_type="application/javascript")
    return PlainTextResponse("// Chart.js not found", status_code=404)

@app.websocket("/ws")
async def ws_ep(ws:WebSocket):
    await mgr.connect(ws)
    if latest:
        try: await ws.send_text(json.dumps({**latest,"tip":"ilk"},cls=SafeEncoder,ensure_ascii=False,default=str))
        except: pass
    try:
        while True: await ws.receive_text()
    except WebSocketDisconnect: mgr.disconnect(ws)

@app.get("/",response_class=HTMLResponse)
async def index():
    p=BASE/"dashboard.html"
    if not p.exists(): return HTMLResponse("<h1>dashboard.html yok</h1>",404)
    return p.read_text(encoding="utf-8")

def main():
    global excel_path, latest
    pa=argparse.ArgumentParser();pa.add_argument("--file",required=True);pa.add_argument("--port",type=int,default=8765);pa.add_argument("--host",default="0.0.0.0");pa.add_argument("--polling",action="store_true")
    a=pa.parse_args(); excel_path=os.path.abspath(a.file)
    if not os.path.exists(excel_path): log.error(f"Dosya yok: {excel_path}"); raise SystemExit(1)
    log.info("Ilk yukleme...")
    try: latest=parse_excel(excel_path); latest["tip"]="ilk"; log.info(f"JSON OK: {len(json.dumps(latest,cls=SafeEncoder,default=str))} byte")
    except Exception as e: log.error(f"Hata: {e}"); latest={"hata":str(e),"kpi":{}}
    wd=str(Path(excel_path).parent);h=FH();ob=PollingObserver(timeout=3);ob.schedule(h,wd,recursive=False);ob.start()
    log.info(f"Dashboard: http://localhost:{a.port}")
    try: uvicorn.run(app,host=a.host,port=a.port,log_level="info")
    finally: ob.stop();ob.join()

if __name__=="__main__": main()
