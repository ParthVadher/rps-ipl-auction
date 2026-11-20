# auction.py — RPS IPL Auction Platform (single file, zero CLI during auction)
import os, io, re, csv, json, sqlite3, datetime, threading, socket, webbrowser
from contextlib import contextmanager
from pathlib import Path
from typing import Optional
from fastapi import FastAPI, UploadFile, File, HTTPException, Body
from fastapi.responses import HTMLResponse, StreamingResponse
import uvicorn

# ─── CONFIG (edit team names here if needed) ─────────────────
TEAMS = ["Rising Pune Supergiant","Chennai Super Kings","Mumbai Indians",
         "Royal Challengers Bengaluru","Kolkata Knight Riders","Sunrisers Hyderabad",
         "Delhi Capitals","Rajasthan Royals","Punjab Kings","Gujarat Titans",
         "Lucknow Super Giants","Gujarat Lions"]
CFG = {"purse":120.0,"min_players":18,"max_players":25,
       "min_indian":14,"max_overseas":8,"min_spend_pct":0.75}
# ─────────────────────────────────────────────────────────────

DB = Path(__file__).parent / "auction.db"
LOCK = threading.Lock()
app = FastAPI(title="RPS Auction")

def conn():
    c = sqlite3.connect(DB, timeout=30, isolation_level=None)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA journal_mode=WAL")
    c.execute("PRAGMA foreign_keys=ON")
    return c

@contextmanager
def db():
    c = conn()
    try: yield c
    finally: c.close()

def init():
    with db() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS players(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL, name_lc TEXT UNIQUE NOT NULL,
            role TEXT, category TEXT, batting_position TEXT,
            rating REAL, overseas INTEGER, base_price TEXT,
            bowling_style TEXT,
            source TEXT, ts TEXT);
        CREATE TABLE IF NOT EXISTS buys(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            player_name TEXT NOT NULL, team TEXT NOT NULL,
            price REAL NOT NULL, is_overseas INTEGER,
            ts TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS unsold(
            player_name TEXT UNIQUE NOT NULL,
            ts TEXT NOT NULL);
        """)
        try: c.execute("ALTER TABLE players ADD COLUMN bowling_style TEXT")
        except sqlite3.OperationalError: pass

def now(): return datetime.datetime.now().isoformat(timespec="seconds")

# ─── import parsing ─────────────────────────────────────────
ALIAS = {
    "name": ["name","player","player name","cricketer","full name","player_name"],
    "role": ["role","player role","type","speciality","playing role"],
    "category": ["set","set name","category","auction set","group","batch"],
    "rating": ["rating","ratings","score","overall","rtg","r"],
    "batting_position": ["batting position","position","bat pos","batting pos","order","slot"],
    "bowling_style": ["bowling style","bowling","bowl","style","pace","spin"],
    "overseas": ["overseas","foreign","is overseas","country","nationality","origin"],
    "base_price": ["base price","base","price","reserve","reserve price","bp"],
}
CAT_ORDER = ["MARQUEE","BATTER","WICKETKEEPER","ALL-ROUNDER","ALLROUNDER",
             "BOWLER","CAPPED","UNCAPPED","OTHER"]

def ncol(s): return re.sub(r"[^a-z0-9]","",str(s).lower())
def cat_class(s):
    s = str(s or "").strip().upper()
    if not s: return "OTHER"
    for c in CAT_ORDER:
        if c in s: return c
    return "OTHER"
def pos_order(p):
    if not p: return 9
    s = str(p).lower()
    if "open" in s: return 1
    if "top" in s: return 2
    if "one-down" in s or "1-down" in s: return 3
    if "mid" in s: return 4
    if "finish" in s or "lower" in s: return 5
    m = re.search(r"\d+", s)
    return int(m.group()) if m and 1 <= int(m.group()) <= 11 else 9
def to_float(v):
    try: return float(str(v).replace(",","").strip())
    except: return None

def is_spinner(text):
    if not text: return False
    t = text.lower()
    return any(k in t for k in ["spin", "orthodox", "legbreak", "offbreak", "wrist", "finger", "leg break", "off break"])

def is_pacer(text):
    if not text: return False
    t = text.lower()
    return any(k in t for k in ["fast", "faster", "medium", "pace", "seam", "swing"])

def resolve_headers(row):
    m = {}
    for k in row.keys():
        nk = ncol(k)
        for f, aliases in ALIAS.items():
            if f in m: continue
            if nk in {ncol(a) for a in aliases}:
                m[f] = k
    return m

def rows_from_dicts(rows, source):
    if not rows: return []
    h = resolve_headers(rows[0])
    if "name" not in h: h["name"] = list(rows[0].keys())[0]
    out = []
    for r in rows:
        name = str(r.get(h.get("name",""),"")).strip()
        if not name or not re.search(r"[A-Za-z]", name): continue
        ov_raw = str(r.get(h.get("overseas",""),"")).strip().lower()
        ov = None
        if ov_raw in ("true","1","yes","y","overseas","os","foreign"): ov = 1
        elif ov_raw in ("false","0","no","n","indian","ind","india","domestic"): ov = 0
        if ov is None and ov_raw:
            ov = 0 if ov_raw in ("india","ind","in") else 1
        out.append({
            "name": name, "name_lc": name.lower().strip(),
            "role": str(r.get(h.get("role",""),"")).strip(),
            "category": cat_class(r.get(h.get("category",""),"")),
            "batting_position": str(r.get(h.get("batting_position",""),"")).strip(),
            "bowling_style": str(r.get(h.get("bowling_style",""),"")).strip(),
            "rating": to_float(r.get(h.get("rating",""),"")),
            "overseas": ov,
            "base_price": str(r.get(h.get("base_price",""),"")).strip(),
            "source": source,
        })
    return out

def read_csv_bytes(data):
    txt = data.decode("utf-8-sig", errors="replace")
    return list(csv.DictReader(io.StringIO(txt)))

def read_xlsx_bytes(data):
    import pandas as pd
    df = pd.read_excel(io.BytesIO(data), dtype=str).fillna("")
    return df.to_dict(orient="records")

def ocr_image(data):
    try:
        import pytesseract
        from PIL import Image
        try: pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        except Exception: pass
        img = Image.open(io.BytesIO(data)).convert("L")
        if img.width < 1600:
            s = 1600/img.width
            img = img.resize((int(img.width*s), int(img.height*s)), Image.LANCZOS)
        return pytesseract.image_to_string(img, config="--psm 6")
    except Exception as e:
        raise HTTPException(400, f"OCR unavailable: {e}")

def read_pdf_bytes(data):
    try:
        import fitz
    except ImportError:
        raise HTTPException(400, "pymupdf not installed")
    doc = fitz.open(stream=data, filetype="pdf")
    texts = []
    for page in doc:
        t = page.get_text()
        if len(t.strip()) > 80:
            texts.append(t)
        else:
            pix = page.get_pixmap(dpi=220)
            texts.append(ocr_image(pix.tobytes("png")))
    return texts

def parse_freeform(text):
    rows = []
    for line in text.splitlines():
        line = line.strip()
        if not line or len(line) < 4: continue
        if re.match(r"^(name|player|s\.?no|sr\.?|#)\b", line, re.I) and len(line) < 25: continue
        
        parts = [p.strip() for p in re.split(r"\s{2,}|\t+", line) if p.strip()]
        if len(parts) < 2: parts = line.split(" ")
        if len(parts) < 2: continue
        name = parts[0]
        if not re.search(r"[A-Za-z]{3}", name): continue
        
        rating, role, pos, bowl = None, "", "", ""
        for tok in parts[1:]:
            n = to_float(tok)
            if n is not None and 0 < n <= 100 and rating is None: rating = n; continue
            low = tok.lower()
            if any(k in low for k in ("batter","batsman","bowler","all-round","allround","wicket","keeper")): 
                role = tok; continue
            if any(k in low for k in ("spin", "pace", "fast", "medium", "orthodox")): 
                bowl = tok; continue
            if re.search(r"\b(opener|top|middle|finish|no\.?\s*\d|\d)\b", low): 
                pos = tok
                
        rows.append({"name":name,"name_lc":name.lower(),
                     "role":role,"category":"OTHER","batting_position":pos,
                     "bowling_style":bowl,"rating":rating,"overseas":None,"base_price":"",
                     "source":"ocr"})
    return rows

def import_file(filename, data):
    name = (filename or "").lower()
    try:
        if name.endswith(".csv"): return rows_from_dicts(read_csv_bytes(data), filename)
        if name.endswith((".xlsx",".xls")): return rows_from_dicts(read_xlsx_bytes(data), filename)
        if name.endswith(".pdf"):
            pages = read_pdf_bytes(data)
            rows, raw = [], ""
            for t in pages:
                raw += t + "\n"
                rows.extend(parse_freeform(t))
            for r in rows: r["source"] = filename
            return rows, raw
        if name.endswith((".jpg",".jpeg",".png",".webp",".bmp")):
            raw = ocr_image(data)
            rows = parse_freeform(raw)
            for r in rows: r["source"] = filename
            return rows, raw
        raise HTTPException(400, f"unsupported file: {filename}")
    except HTTPException: raise
    except Exception as e: raise HTTPException(400, f"failed to parse {filename}: {e}")

# ─── state / buys ───────────────────────────────────────────
def team_state():
    with db() as c:
        rows = c.execute("SELECT * FROM buys").fetchall()
    s = {t:{"spent":0.0,"count":0,"indian":0,"overseas":0} for t in TEAMS}
    for r in rows:
        if r["team"] not in s: continue
        s[r["team"]]["spent"] += r["price"]
        s[r["team"]]["count"] += 1
        s[r["team"]]["overseas" if r["is_overseas"] else "indian"] += 1
    for t in s:
        v = s[t]
        v["remaining"] = round(CFG["purse"] - v["spent"], 2)
        v["slots_left"] = CFG["max_players"] - v["count"]
        v["indian_ok"] = v["indian"] + v["slots_left"] >= CFG["min_indian"]
        v["overseas_ok"] = v["overseas"] <= CFG["max_overseas"]
        v["min_spend_ok"] = v["spent"] >= CFG["purse"] * CFG["min_spend_pct"]
    return s

def do_buy(player_name, team, price, is_overseas):
    player_name = player_name.strip()
    if not player_name: raise HTTPException(400, "empty player name")
    if team not in TEAMS: raise HTTPException(400, "unknown team")
    if price <= 0: raise HTTPException(400, "price must be > 0")
    with LOCK, db() as c:
        c.execute("BEGIN IMMEDIATE")
        try:
            existing = c.execute("SELECT team FROM buys WHERE LOWER(player_name)=?", (player_name.lower(),)).fetchone()
            if existing: raise HTTPException(400, f"'{player_name}' is already sold to {existing['team']}!")
            s = team_state()[team]
            if s["remaining"] < price: raise HTTPException(400, f"{team} has only {s['remaining']} Cr left, cannot buy for {price}")
            if s["count"] + 1 > CFG["max_players"]: raise HTTPException(400, f"{team} already has {s['count']}/{CFG['max_players']} players")
            if is_overseas and s["overseas"] + 1 > CFG["max_overseas"]: raise HTTPException(400, f"{team} already has {s['overseas']}/{CFG['max_overseas']} overseas")
            new_ind = s["indian"] + (0 if is_overseas else 1)
            slots_after = CFG["max_players"] - (s["count"] + 1)
            if new_ind + slots_after < CFG["min_indian"]: raise HTTPException(400, f"would leave fewer than {CFG['min_indian']} Indian slots reachable")
            
            c.execute("INSERT INTO buys(player_name,team,price,is_overseas,ts) VALUES(?,?,?,?,?)",
                      (player_name, team, price, 1 if is_overseas else 0, now()))
            c.execute("DELETE FROM unsold WHERE LOWER(player_name)=?", (player_name.lower(),))
            c.execute("COMMIT")
        except HTTPException: c.execute("ROLLBACK"); raise
    return {"ok": True, "state": team_state()[team]}

# ─── API ────────────────────────────────────────────────────
@app.post("/api/import/preview")
async def api_preview(files: list[UploadFile] = File(...)):
    all_rows, errors, raw_texts = [], [], []
    for f in files:
        try:
            res = import_file(f.filename, await f.read())
            if isinstance(res, tuple):
                rows, raw = res
                all_rows.extend(rows)
                if raw: raw_texts.append(f"--- {f.filename} ---\n{raw}")
            else: all_rows.extend(res)
        except HTTPException as e: errors.append(f"{f.filename}: {e.detail}")
    seen = {}
    for r in all_rows:
        k = r["name_lc"]
        if k not in seen or (r.get("rating") or 0) > (seen[k].get("rating") or 0):
            seen[k] = r
    return {"rows": list(seen.values()), "errors": errors, "raw": "\n\n".join(raw_texts)}

@app.post("/api/import/commit")
def api_commit(rows: list[dict] = Body(...)):
    ins = upd = 0
    with db() as c:
        for r in rows:
            nm = str(r.get("name","")).strip()
            if not nm: continue
            lc = nm.lower()
            ex = c.execute("SELECT id FROM players WHERE name_lc=?", (lc,)).fetchone()
            if ex:
                c.execute("""UPDATE players SET role=COALESCE(NULLIF(?,''),role), 
                             category=COALESCE(NULLIF(?,''),category), 
                             batting_position=COALESCE(NULLIF(?,''),batting_position),
                             bowling_style=COALESCE(NULLIF(?,''),bowling_style),
                             rating=COALESCE(?,rating), 
                             overseas=COALESCE(?,overseas), 
                             base_price=COALESCE(NULLIF(?,''),base_price), 
                             source=? WHERE id=?""",
                          (r.get("role",""), r.get("category","OTHER"), r.get("batting_position",""),
                           r.get("bowling_style",""), r.get("rating"), r.get("overseas"), 
                           r.get("base_price",""), r.get("source",""), ex["id"]))
                upd += 1
            else:
                c.execute("""INSERT INTO players(name,name_lc,role,category,batting_position,bowling_style,rating,
                             overseas,base_price,source,ts) VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                          (nm, lc, r.get("role",""), r.get("category","OTHER"),
                           r.get("batting_position",""), r.get("bowling_style",""), r.get("rating"), 
                           r.get("overseas"), r.get("base_price",""), r.get("source",""), now()))
                ins += 1
    return {"inserted": ins, "updated": upd}

@app.post("/api/players/add")
def api_add_player(payload: dict = Body(...)):
    name = payload["name"].strip()
    if not name: raise HTTPException(400, "Name required")
    with LOCK, db() as c:
        c.execute("BEGIN IMMEDIATE")
        try:
            lc = name.lower()
            ex = c.execute("SELECT id FROM players WHERE name_lc=?", (lc,)).fetchone()
            if ex: raise HTTPException(400, "Player already exists in database")
            
            c.execute("""INSERT INTO players(name,name_lc,role,category,batting_position,bowling_style,rating,
                         overseas,base_price,source,ts) VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                      (name, lc, payload.get("role",""), "OTHER",
                       "", "", None, 
                       1 if payload.get("overseas") else 0, payload.get("base_price",""), "manual", now()))
            c.execute("COMMIT")
        except HTTPException: c.execute("ROLLBACK"); raise
    return {"ok": True}

@app.get("/api/players")
def api_players(q: str = "", category: str = "", roles: str = "", sort: str = "category"):
    with db() as c:
        rows = [dict(r) for r in c.execute("SELECT * FROM players").fetchall()]
    
    if q:
        ql = q.lower()
        rows = [r for r in rows if ql in r["name"].lower() or ql in (r["role"] or "").lower()]
    if category: rows = [r for r in rows if r["category"] == category]
        
    if roles:
        role_list = [r.strip().lower() for r in roles.split(",") if r.strip()]
        def match_role(role_str):
            rs = (role_str or "").lower()
            for req in role_list:
                if req == "bat" and ("bat" in rs or "batsman" in rs) and "wk" not in rs and "keep" not in rs: return True
                if req == "bowl" and ("bowl" in rs or "pace" in rs or "spin" in rs): return True
                if req == "all" and ("all" in rs or "ar" in rs or "round" in rs): return True
                if req == "wk" and ("wk" in rs or "keep" in rs or "wicket" in rs): return True
            return False
        rows = [r for r in rows if match_role(r["role"])]

    for r in rows:
        r["pos_order"] = pos_order(r["batting_position"])

    if sort == "rating":
        rows.sort(key=lambda r: (-(r["rating"] or 0), r["name"].lower()))
    elif sort == "name":
        rows.sort(key=lambda r: r["name"].lower())
    elif sort == "position" or sort == "batting_order":
        rows.sort(key=lambda r: (r["pos_order"], -(r["rating"] or 0), r["name"].lower()))
    elif sort == "spinner":
        rows = [r for r in rows if is_spinner(r.get("bowling_style")) or is_spinner(r.get("role"))]
        rows.sort(key=lambda r: (-(r["rating"] or 0), r["name"].lower()))
    elif sort == "pacer":
        rows = [r for r in rows if is_pacer(r.get("bowling_style")) or is_pacer(r.get("role"))]
        rows.sort(key=lambda r: (-(r["rating"] or 0), r["name"].lower()))
    else:
        def cidx(c): return CAT_ORDER.index(c) if c in CAT_ORDER else len(CAT_ORDER)
        rows.sort(key=lambda r: (cidx(r["category"]), r["pos_order"], -(r["rating"] or 0), r["name"].lower()))
    return rows

@app.delete("/api/players/all")
def api_clear_players():
    with db() as c: c.execute("DELETE FROM players")
    return {"ok": True}

@app.post("/api/buy")
def api_buy(payload: dict = Body(...)):
    return do_buy(payload["player"], payload["team"], float(payload["price"]), bool(payload.get("is_overseas", False)))

@app.post("/api/unsold")
def api_unsold(payload: dict = Body(...)):
    player_name = payload["player"].strip()
    if not player_name: raise HTTPException(400, "Empty name")
    with LOCK, db() as c:
        c.execute("BEGIN IMMEDIATE")
        try:
            existing = c.execute("SELECT team FROM buys WHERE LOWER(player_name)=?", (player_name.lower(),)).fetchone()
            if existing: raise HTTPException(400, f"'{player_name}' is already sold to {existing['team']}! Cannot mark unsold.")
            c.execute("INSERT OR REPLACE INTO unsold(player_name, ts) VALUES(?,?)", (player_name, now()))
            c.execute("COMMIT")
        except HTTPException: c.execute("ROLLBACK"); raise
    return {"ok": True}

@app.delete("/api/unsold/{name}")
def api_bring_back(name: str):
    with LOCK, db() as c:
        c.execute("DELETE FROM unsold WHERE LOWER(player_name)=?", (name.lower(),))
    return {"ok": True}

@app.get("/api/unsold")
def api_get_unsold():
    with db() as c:
        return [dict(r) for r in c.execute("SELECT * FROM unsold ORDER BY ts").fetchall()]

@app.post("/api/edit_buy/{buy_id}")
def api_edit_buy(buy_id: int, payload: dict = Body(...)):
    new_price = float(payload["price"])
    new_team = payload["team"]
    
    with LOCK, db() as c:
        c.execute("BEGIN IMMEDIATE")
        try:
            b = c.execute("SELECT * FROM buys WHERE id=?", (buy_id,)).fetchone()
            if not b: raise HTTPException(404, "Purchase not found")
            old_team = b["team"]
            
            if old_team == new_team:
                s = team_state()[old_team]
                diff = new_price - b["price"]
                if s["remaining"] - diff < 0:
                    raise HTTPException(400, f"Not enough purse. Max price is {b['price'] + s['remaining']} Cr")
            else:
                s = team_state()[new_team]
                if s["remaining"] < new_price:
                    raise HTTPException(400, f"{new_team} has only {s['remaining']} Cr left, cannot buy for {new_price}")
                if s["count"] + 1 > CFG["max_players"]:
                    raise HTTPException(400, f"{new_team} already has {s['count']}/{CFG['max_players']} players")
                if b["is_overseas"] and s["overseas"] + 1 > CFG["max_overseas"]:
                    raise HTTPException(400, f"{new_team} already has {s['overseas']}/{CFG['max_overseas']} overseas")
                new_ind = s["indian"] + (0 if b["is_overseas"] else 1)
                slots_after = CFG["max_players"] - (s["count"] + 1)
                if new_ind + slots_after < CFG["min_indian"]:
                    raise HTTPException(400, f"would leave fewer than {CFG['min_indian']} Indian slots reachable for {new_team}")
            
            c.execute("UPDATE buys SET team=?, price=? WHERE id=?", (new_team, new_price, buy_id))
            c.execute("COMMIT")
        except HTTPException:
            c.execute("ROLLBACK"); raise
    return {"ok": True}

@app.post("/api/undo/{buy_id}")
def api_undo(buy_id: int):
    with LOCK, db() as c: c.execute("DELETE FROM buys WHERE id=?", (buy_id,))
    return {"ok": True}

@app.get("/api/state")
def api_state(): return team_state()

@app.get("/api/buys")
def api_buys():
    with db() as c: return [dict(r) for r in c.execute("SELECT * FROM buys ORDER BY id DESC").fetchall()]

@app.get("/api/team_squad/{team_name}")
def api_team_squad(team_name: str):
    with db() as c:
        buys = c.execute("SELECT * FROM buys WHERE team=? ORDER BY id", (team_name,)).fetchall()
        squad = []
        for b in buys:
            p = c.execute("SELECT role, category, rating FROM players WHERE name_lc=?", (b["player_name"].lower(),)).fetchone()
            squad.append({
                "id": b["id"], "name": b["player_name"], "price": b["price"], "is_overseas": b["is_overseas"],
                "role": p["role"] if p and p["role"] else "Unknown", "category": p["category"] if p else "OTHER",
                "rating": p["rating"] if p else None
            })
        return squad

@app.get("/api/export/buys")
def api_export():
    with db() as c: rows = c.execute("SELECT * FROM buys ORDER BY id").fetchall()
    buf = io.StringIO()
    if rows:
        w = csv.DictWriter(buf, fieldnames=rows[0].keys())
        w.writeheader(); w.writerows([dict(r) for r in rows])
    return StreamingResponse(io.BytesIO(buf.getvalue().encode()), media_type="text/csv",
                             headers={"Content-Disposition":"attachment; filename=purchases.csv"})

@app.get("/", response_class=HTMLResponse)
def root(): return HTML

# ─── frontend ───────────────────────────────────────────────
HTML = r"""<!doctype html>
<html><head><meta charset="utf-8"><title>RPS Auction</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
*{box-sizing:border-box}body{margin:0;font-family:system-ui,sans-serif;background:#0b1220;color:#e6ecf7;padding:16px;max-width:1400px;margin:0 auto}
h1{font-size:18px;margin:0 0 12px}h1 span{color:#ff5a5f}
.tabs{display:flex;gap:6px;margin-bottom:14px;flex-wrap:wrap}
.tabs button{background:#0d1527;color:#8b98b3;border:1px solid #1e2a45;padding:8px 14px;border-radius:8px;cursor:pointer;font-size:13px}
.tabs button.active{background:#ff5a5f;color:#fff;border-color:#ff5a5f;font-weight:600}
.card{background:#111a2e;border:1px solid #1e2a45;border-radius:12px;padding:14px;margin-bottom:12px;transition:border-color 0.2s}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{padding:8px;text-align:left;border-bottom:1px solid #1e2a45;vertical-align:middle}
th{color:#8b98b3;font-size:11px;text-transform:uppercase;letter-spacing:.05em}
input,select,button{font:inherit;background:#0d1527;border:1px solid #1e2a45;color:#e6ecf7;padding:8px 12px;border-radius:8px;font-size:13px}
input[type=checkbox]{width:16px;height:16px;padding:0}
button.pri{background:#ff5a5f;border:0;color:#fff;font-weight:600;cursor:pointer}
button.ok{background:#10b981;border:0;color:#fff;font-weight:600;cursor:pointer}
button.warn{background:#f59e0b;border:0;color:#fff;font-weight:600;cursor:pointer}
button.sm{padding:4px 10px;font-size:11px;cursor:pointer;margin-right:4px}
.tag{padding:2px 8px;border-radius:99px;font-size:11px;font-weight:600;display:inline-block;margin-right:4px}
.tag.g{background:rgba(52,211,153,.15);color:#34d399}
.tag.r{background:rgba(239,68,68,.15);color:#ef4444}
.tag.y{background:rgba(245,158,11,.15);color:#f59e0b}
.tag.b{background:rgba(59,130,246,.15);color:#3b82f6}
.err{color:#ef4444;font-size:12px;min-height:16px;margin-top:6px}
.drop{border:2px dashed #1e2a45;border-radius:12px;padding:30px;text-align:center;color:#8b98b3;cursor:pointer}
.drop.hover{border-color:#ff5a5f;background:rgba(255,90,95,.05);color:#e6ecf7}
.hidden{display:none !important}
.big{font-size:22px;font-weight:700}
.muted{font-size:11px;color:#8b98b3}
.cat-hdr{background:#0d1527;font-weight:700;font-size:11px;letter-spacing:.08em;color:#ff5a5f;text-transform:uppercase;padding:6px 8px}
pre{background:#0d1527;padding:10px;border-radius:8px;font-size:12px;overflow:auto;max-height:200px;color:#8b98b3}
.role-btn{border-radius:99px;padding:4px 12px;font-size:11px;font-weight:600;cursor:pointer;border:1px solid #1e2a45;background:#0d1527;color:#8b98b3;transition:all 0.2s}
.role-btn.active{background:#ff5a5f;color:#fff;border-color:#ff5a5f}
.team-card{cursor:pointer}
.team-card:hover{border-color:#ff5a5f}

/* Modal Styles */
.modal-overlay{position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.8);z-index:1000;display:flex;justify-content:center;align-items:center}
.modal-content{width:95%;max-width:1200px;max-height:90vh;overflow-y:auto;background:#0b1220;border:1px solid #1e2a45;border-radius:12px;padding:20px;position:relative}
.modal-content.sm{max-width:400px}
.close-btn{position:absolute;top:15px;right:20px;background:#ff5a5f;border:0;color:#fff;font-weight:bold;font-size:16px;width:30px;height:30px;border-radius:50%;cursor:pointer;z-index:10}
.player-card{background:#111a2e;border-left:4px solid #1e2a45;padding:10px;border-radius:6px;display:flex;justify-content:space-between;align-items:center}
.player-card.overseas{border-left-color:#f59e0b}
.player-card.indian{border-left-color:#10b981}
.player-card.in-xi{background:#1a243b;border:1px solid #3b82f6;box-shadow:0 0 10px rgba(59,130,246,0.2)}
.role-group{margin-bottom:20px}
.role-title{font-size:12px;color:#8b98b3;text-transform:uppercase;letter-spacing:.08em;font-weight:700;margin-bottom:10px;border-bottom:1px solid #1e2a45;padding-bottom:5px}
.xi-badge{background:#3b82f6;color:#fff;font-size:10px;padding:2px 6px;border-radius:4px;margin-right:5px}

/* Draggable XI List */
.xi-item{background:#111a2e;padding:8px;border-radius:6px;display:flex;justify-content:space-between;align-items:center;border-left:3px solid #3b82f6;cursor:grab;}
.xi-item:active{cursor:grabbing}
.xi-item.drag-over{border-top:2px solid #ff5a5f}
</style></head><body>

<h1>RPS <span>AUCTION</span> PLATFORM</h1>

<div class="tabs">
  <button data-tab="auction" class="active">🎯 Auction</button>
  <button data-tab="teams">🏏 Teams</button>
  <button data-tab="players">📋 Players</button>
  <button data-tab="import">📥 Import List</button>
</div>

<!-- AUCTION -->
<div id="tab-auction">
  <div class="card">
    <h3 style="font-size:12px;color:#8b98b3;margin:0 0 10px">RECORD PURCHASE</h3>
    <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
      <input id="bPlayer" list="playerList" placeholder="Player name (type or pick)" style="flex:2;min-width:230px" autocomplete="off" autofocus>
      <datalist id="playerList"></datalist>
      <select id="bTeam" style="min-width:180px"></select>
      <input id="bPrice" type="number" step="0.05" min="0.05" placeholder="Price (Cr)" style="width:120px">
      <label style="display:flex;align-items:center;gap:6px"><input type="checkbox" id="bOS"> Overseas</label>
      <button class="pri" onclick="doBuy()">Buy</button>
      <button class="warn" onclick="markUnsold()">Unsold</button>
    </div>
    <div id="bErr" class="err"></div>
  </div>
  
  <div id="teamCards" class="grid"></div>
  
  <div style="display:grid;grid-template-columns:2fr 1fr;gap:12px;margin-top:14px;align-items:start">
    <!-- Purchases List -->
    <div class="card">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;flex-wrap:wrap;gap:8px">
        <h3 style="font-size:12px;color:#8b98b3;margin:0">PURCHASES</h3>
        <div style="display:flex;gap:8px;align-items:center">
          <input id="searchBuys" type="text" placeholder="Search bought players..." oninput="renderBuys()" style="padding:4px 8px;font-size:12px;width:200px">
          <a href="/api/export/buys" style="color:#ff5a5f;font-size:12px;text-decoration:none">⬇ Download CSV</a>
        </div>
      </div>
      <table id="buysTable"></table>
    </div>

    <!-- Unsold List -->
    <div class="card" style="background:#1a1112;border-color:#3f2122">
      <h3 style="font-size:12px;color:#f59e0b;margin:0 0 10px">UNSOLD (ACCELERATED ROUND)</h3>
      <div id="unsoldList"></div>
    </div>
  </div>
</div>

<!-- TEAMS -->
<div id="tab-teams" class="hidden">
  <div class="card">
    <h3 style="font-size:12px;color:#8b98b3;margin:0 0 10px">Click a team to view their squad</h3>
    <table id="teamsTable"></table>
  </div>
</div>

<!-- PLAYERS -->
<div id="tab-players" class="hidden">
  <div class="card">
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px">
      <input id="pSearch" placeholder="Search name" oninput="loadPlayers()" style="flex:1;min-width:200px">
      <select id="pCategory" onchange="loadPlayers()"><option value="">All categories</option></select>
    </div>
    
    <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:10px;padding:8px 0;border-top:1px solid #1e2a45">
      <span style="font-size:11px;color:#8b98b3;font-weight:600">ROLES:</span>
      <button class="role-btn active" data-role="bat" onclick="toggleRole(this)">Batter</button>
      <button class="role-btn active" data-role="wk" onclick="toggleRole(this)">Wicketkeeper</button>
      <button class="role-btn active" data-role="all" onclick="toggleRole(this)">All-Rounder</button>
      <button class="role-btn active" data-role="bowl" onclick="toggleRole(this)">Bowler</button>
    </div>

    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px;align-items:center">
      <select id="pSort" onchange="loadPlayers()">
        <option value="category">Sort: category → position → rating</option>
        <option value="batting_order">Sort: Batting Order (Top/Mid/Lower)</option>
        <option value="rating">Sort: rating (high first)</option>
        <option value="position">Sort: slot / position</option>
        <option value="spinner">Sort: Spinners</option>
        <option value="pacer">Sort: Pacers</option>
        <option value="name">Sort: name A–Z</option>
      </select>
      <button class="sm" style="background:#3b82f6;border:0;color:#fff" onclick="openAddPlayerModal()">+ Add Player Manually</button>
      <button class="sm" onclick="clearPlayers()">Clear all players</button>
    </div>
    
    <div id="playersCount" class="muted" style="margin-bottom:8px"></div>
    <div id="playersTable" style="max-height:600px;overflow:auto"></div>
  </div>
</div>

<!-- IMPORT -->
<div id="tab-import" class="hidden">
  <div class="card">
    <h3 style="font-size:12px;color:#8b98b3;margin:0 0 10px">UPLOAD PLAYER LIST(S)</h3>
    <div id="drop" class="drop">
      <div style="font-size:16px;margin-bottom:6px">Drop files here, or click to browse</div>
      <div class="muted">CSV · XLSX · PDF · JPG · PNG — you can select multiple files at once, or upload set-by-set during the day</div>
    </div>
    <input id="fileInput" type="file" multiple accept=".csv,.xlsx,.xls,.pdf,.jpg,.jpeg,.png" class="hidden">
    <div id="importErr" class="err"></div>
    <div id="rawTextCard" class="hidden" style="margin-top:10px">
      <div style="font-size:12px;color:#ef4444;font-weight:600;margin-bottom:4px">OCR could not parse this into rows. Here is what it saw. Copy it into Excel and save as CSV if needed.</div>
      <pre id="rawText"></pre>
    </div>
  </div>
  <div id="previewCard" class="card hidden">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;flex-wrap:wrap;gap:8px">
      <div><b id="previewCount"></b> players extracted — <span class="muted">uncheck any rows to skip them</span></div>
      <button class="ok" onclick="commitImport()">✓ Import Selected</button>
    </div>
    <div id="previewTable" style="max-height:520px;overflow:auto"></div>
  </div>
</div>

<!-- SQUAD MODAL -->
<div id="squadModal" class="modal-overlay hidden" onclick="if(event.target === this) closeSquadModal()">
  <div class="modal-content">
    <button class="close-btn" onclick="closeSquadModal()">✕</button>
    <h2 id="modalTeamName" style="margin-top:0;font-size:20px;color:#ff5a5f"></h2>
    
    <div style="display:flex;gap:16px;flex-wrap:wrap">
      <div style="flex:1;min-width:280px;background:#0d1527;border:1px solid #1e2a45;border-radius:8px;padding:12px">
        <h3 style="font-size:14px;margin:0 0 10px;color:#3b82f6">Playing XI Builder</h3>
        <div id="xiStats" style="font-size:11px;margin-bottom:10px;color:#8b98b3"></div>
        <div id="xiList" style="display:flex;flex-direction:column;gap:6px"></div>
      </div>
      <div style="flex:2;min-width:400px">
        <h3 style="font-size:14px;margin:0 0 10px;color:#e6ecf7">Full Squad</h3>
        <div id="modalSquadList"></div>
      </div>
    </div>
  </div>
</div>

<!-- EDIT PURCHASE MODAL -->
<div id="editModal" class="modal-overlay hidden" onclick="if(event.target === this) closeEditModal()">
  <div class="modal-content sm">
    <button class="close-btn" onclick="closeEditModal()">✕</button>
    <h2 style="margin-top:0;font-size:18px;color:#ff5a5f">Edit Purchase</h2>
    <input type="hidden" id="editId">
    <div style="margin-bottom:12px">
      <label style="font-size:12px;color:#8b98b3">Player</label>
      <input id="editPlayerName" readonly style="width:100%;background:#1a243b;cursor:not-allowed">
    </div>
    <div style="margin-bottom:12px">
      <label style="font-size:12px;color:#8b98b3">Team</label>
      <select id="editTeam" style="width:100%"></select>
    </div>
    <div style="margin-bottom:12px">
      <label style="font-size:12px;color:#8b98b3">Price (Cr)</label>
      <input id="editPrice" type="number" step="0.05" style="width:100%">
    </div>
    <div id="editErr" class="err" style="margin-bottom:12px"></div>
    <button class="pri" style="width:100%" onclick="saveEdit()">Save Changes</button>
  </div>
</div>

<!-- ADD PLAYER MODAL -->
<div id="addPlayerModal" class="modal-overlay hidden" onclick="if(event.target === this) closeAddPlayerModal()">
  <div class="modal-content sm">
    <button class="close-btn" onclick="closeAddPlayerModal()">✕</button>
    <h2 style="margin-top:0;font-size:18px;color:#3b82f6">Add Player Manually</h2>
    <div style="margin-bottom:12px">
      <label style="font-size:12px;color:#8b98b3">Player Name</label>
      <input id="addPlayerName" style="width:100%" placeholder="e.g. Virat Kohli">
    </div>
    <div style="margin-bottom:12px">
      <label style="font-size:12px;color:#8b98b3">Role</label>
      <select id="addPlayerRole" style="width:100%">
        <option value="Batter">Batter</option>
        <option value="Bowler">Bowler</option>
        <option value="All-Rounder">All-Rounder</option>
        <option value="Wicketkeeper">Wicketkeeper</option>
      </select>
    </div>
    <div style="margin-bottom:12px;display:flex;align-items:center;gap:8px">
      <input type="checkbox" id="addPlayerOS" style="width:16px;height:16px">
      <label style="font-size:13px;cursor:pointer" for="addPlayerOS">Overseas Player</label>
    </div>
    <div style="margin-bottom:12px">
      <label style="font-size:12px;color:#8b98b3">Base Price (Cr) - Optional</label>
      <input id="addPlayerPrice" type="number" step="0.05" style="width:100%" placeholder="e.g. 2.0">
    </div>
    <div id="addPlayerErr" class="err" style="margin-bottom:12px"></div>
    <button class="pri" style="width:100%" onclick="saveNewPlayer()">Add Player</button>
  </div>
</div>

<script>
const $ = s => document.querySelector(s);
const $$ = s => document.querySelectorAll(s);
let previewRows = [];
let teamsCache = [];
let openTeam = null;
let currentModalTeam = null;
let dragSrcIdx = null;
let dragTeam = null;
let availablePlayers = [];
let allBuys = [];
let allUnsold = [];

let playingXI = {};
try { playingXI = JSON.parse(localStorage.getItem('rps_xi') || '{}'); } catch(e) { playingXI = {}; }

$$('.tabs button').forEach(b => b.onclick = () => {
  $$('.tabs button').forEach(x => x.classList.remove('active'));
  b.classList.add('active');
  ['auction','teams','players','import'].forEach(t => $('#tab-'+t).classList.add('hidden'));
  $('#tab-'+b.dataset.tab).classList.remove('hidden');
  if (b.dataset.tab === 'players') loadPlayers();
  if (b.dataset.tab === 'teams') loadTeams();
});

async function api(path, opts={}) {
  const r = await fetch(path, opts);
  if (!r.ok) {
    let m = await r.text();
    try { m = JSON.parse(m).detail || m } catch(e){}
    throw new Error(m);
  }
  const ct = r.headers.get('content-type')||'';
  return ct.includes('json') ? r.json() : r.text();
}

// ── auction ─────────────────────────────
async function loadAuction() {
  const [state, buys, unsold] = await Promise.all([
    api('/api/state'), 
    api('/api/buys'),
    api('/api/unsold')
  ]);
  teamsCache = Object.keys(state);
  allBuys = buys;
  allUnsold = unsold;

  if ($('#bTeam').options.length !== teamsCache.length) {
    $('#bTeam').innerHTML = teamsCache.map(t => `<option>${t}</option>`).join('');
  }

  $('#teamCards').innerHTML = teamsCache.map(t => {
    const v = state[t];
    return `<div class="card team-card" onclick="openSquadModal('${t}')">
      <div style="font-weight:600;font-size:13px;margin-bottom:4px">${t} <span style="font-size:10px;color:#8b98b3">(Click for Squad)</span></div>
      <div class="big">${v.remaining.toFixed(2)} Cr</div>
      <div class="muted">spent ${v.spent.toFixed(2)} · ${v.count} players · ${v.slots_left} slots left</div>
      <div style="margin-top:8px">
        <span class="tag ${v.indian_ok?'g':'r'}">IND ${v.indian}</span>
        <span class="tag ${v.overseas_ok?'g':'r'}">OS ${v.overseas}</span>
        ${v.min_spend_ok?'<span class="tag g">75% ✓</span>':''}
      </div>
    </div>`;
  }).join('');

  renderBuys();
  renderUnsold();

  const players = await api('/api/players?sort=name');
  const soldNames = new Set(buys.map(b => b.player_name.toLowerCase()));
  const unsoldNames = new Set(unsold.map(u => u.player_name.toLowerCase()));
  
  availablePlayers = players.filter(p => !soldNames.has(p.name.toLowerCase()) && !unsoldNames.has(p.name.toLowerCase()));
  
  $('#playerList').innerHTML = availablePlayers
    .map(p => `<option value="${p.name.replace(/"/g,'&quot;')}" data-os="${p.overseas===1?'1':p.overseas===0?'0':''}">${p.role||''} ${p.category||''} ${p.rating?('· '+p.rating):''}</option>`)
    .join('');
}

function renderBuys() {
  const q = $('#searchBuys').value.toLowerCase().trim();
  const filtered = q ? allBuys.filter(b => b.player_name.toLowerCase().includes(q)) : allBuys;

  $('#buysTable').innerHTML =
    '<tr><th>#</th><th>Player</th><th>Team</th><th>Price</th><th></th><th>Time</th><th>Actions</th></tr>' +
    filtered.map(b => `<tr>
      <td>${b.id}</td>
      <td>${b.player_name}</td>
      <td>${b.team}</td>
      <td><b>${b.price} Cr</b></td>
      <td>${b.is_overseas?'<span class="tag y">OS</span>':'<span class="tag g">IND</span>'}</td>
      <td class="muted">${b.ts.slice(11,16)}</td>
      <td>
        <button class="sm" onclick="openEditModal(${b.id}, '${b.player_name.replace(/'/g,"\\'")}', ${b.price}, '${b.team.replace(/'/g,"\\'")}')">Edit</button>
        <button class="sm" onclick="undoBuy(${b.id})">Undo</button>
      </td>
    </tr>`).join('');
}

function renderUnsold() {
  if (!allUnsold.length) {
    $('#unsoldList').innerHTML = '<div class="muted" style="text-align:center;padding:10px;">No unsold players yet.</div>';
    return;
  }
  
  $('#unsoldList').innerHTML = allUnsold.map(u => `
    <div style="background:#1a1112;border:1px solid #3f2122;padding:8px;border-radius:6px;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center;">
      <div style="font-size:13px;font-weight:600">${u.player_name}</div>
      <div>
        <button class="sm" onclick="quickBuyUnsold('${u.player_name.replace(/'/g,"\\'")}')">Buy ▸</button>
        <button class="sm" style="background:transparent;border:1px solid #8b98b3;color:#8b98b3" onclick="bringBack('${u.player_name.replace(/'/g,"\\'")}')">Bring Back</button>
      </div>
    </div>
  `).join('');
}

async function doBuy() {
  $('#bErr').textContent = '';
  const player = $('#bPlayer').value.trim();
  if (!player) { $('#bErr').textContent = 'Enter a player name'; return; }
  try {
    await api('/api/buy', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({
        player, team: $('#bTeam').value,
        price: parseFloat($('#bPrice').value),
        is_overseas: $('#bOS').checked,
      }),
    });
    $('#bPlayer').value = ''; $('#bPrice').value = ''; $('#bOS').checked = false;
    $('#bPlayer').focus();
    loadAuction();
    if(currentModalTeam) openSquadModal(currentModalTeam);
  } catch (e) {
    $('#bErr').textContent = '✗ ' + e.message;
  }
}

async function markUnsold() {
  $('#bErr').textContent = '';
  const player = $('#bPlayer').value.trim();
  if (!player) { $('#bErr').textContent = 'Enter a player name to mark unsold'; return; }
  if (!confirm(`Mark ${player} as UNSOLD?`)) return;
  
  try {
    await api('/api/unsold', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({ player }),
    });
    $('#bPlayer').value = ''; $('#bPrice').value = ''; $('#bOS').checked = false;
    $('#bPlayer').focus();
    loadAuction();
  } catch (e) {
    $('#bErr').textContent = '✗ ' + e.message;
  }
}

async function bringBack(name) {
  try {
    await api('/api/unsold/' + encodeURIComponent(name), { method: 'DELETE' });
    loadAuction();
  } catch (e) { alert(e.message); }
}

function quickBuyUnsold(name) {
  $('#bPlayer').value = name;
  $('#bPrice').focus();
}

// ── Add Player Modal ─────────────────────
function openAddPlayerModal() {
  $('#addPlayerName').value = '';
  $('#addPlayerRole').value = 'Batter';
  $('#addPlayerOS').checked = false;
  $('#addPlayerPrice').value = '';
  $('#addPlayerErr').textContent = '';
  $('#addPlayerModal').classList.remove('hidden');
}

function closeAddPlayerModal() {
  $('#addPlayerModal').classList.add('hidden');
}

async function saveNewPlayer() {
  const name = $('#addPlayerName').value.trim();
  if (!name) { $('#addPlayerErr').textContent = 'Name required'; return; }
  
  try {
    await api('/api/players/add', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({
        name: name,
        role: $('#addPlayerRole').value,
        overseas: $('#addPlayerOS').checked,
        base_price: $('#addPlayerPrice').value
      })
    });
    closeAddPlayerModal();
    loadPlayers();
  } catch (e) {
    $('#addPlayerErr').textContent = '✗ ' + e.message;
  }
}

// ── Edit Modal ───────────────────────────
function openEditModal(id, playerName, price, team) {
  $('#editId').value = id;
  $('#editPlayerName').value = playerName;
  $('#editPrice').value = price;
  $('#editTeam').innerHTML = teamsCache.map(t => `<option ${t===team?'selected':''}>${t}</option>`).join('');
  $('#editErr').textContent = '';
  $('#editModal').classList.remove('hidden');
}

function closeEditModal() {
  $('#editModal').classList.add('hidden');
}

async function saveEdit() {
  const id = $('#editId').value;
  const team = $('#editTeam').value;
  const price = parseFloat($('#editPrice').value);
  
  try {
    await api('/api/edit_buy/' + id, {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({ team, price })
    });
    closeEditModal();
    loadAuction();
    if(currentModalTeam) openSquadModal(currentModalTeam);
  } catch (e) {
    $('#editErr').textContent = '✗ ' + e.message;
  }
}

async function undoBuy(id) {
  if (!confirm('Undo this purchase?')) return;
  await fetch('/api/undo/' + id, {method:'POST'});
  loadAuction();
  if(currentModalTeam) openSquadModal(currentModalTeam);
}

$('#bPlayer').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') {
    const val = this.value.trim().toLowerCase();
    if (!val) return;
    e.preventDefault(); 
    
    let match = availablePlayers.find(p => p.name.toLowerCase() === val);
    if (!match) {
      const matches = availablePlayers.filter(p => p.name.toLowerCase().includes(val));
      if (matches.length > 0) {
        matches.sort((a, b) => a.name.toLowerCase().indexOf(val) - b.name.toLowerCase().indexOf(val));
        match = matches[0];
      }
    }
    if (match) {
      this.value = match.name;
      $('#bOS').checked = match.overseas === 1;
      $('#bPrice').focus(); 
    } else { $('#bPrice').focus(); }
  }
});

$('#bPrice').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') { e.preventDefault(); doBuy(); }
});

// ── SQUAD MODAL & XI BUILDER ─────────────
async function openSquadModal(teamName) {
  currentModalTeam = teamName;
  $('#modalTeamName').textContent = teamName + " Squad";
  $('#squadModal').classList.remove('hidden');
  $('#modalSquadList').innerHTML = '<div class="muted">Loading squad...</div>';
  
  try {
    const squad = await api('/api/team_squad/' + encodeURIComponent(teamName));
    if (!squad.length) {
      $('#modalSquadList').innerHTML = '<div class="muted" style="padding:20px;text-align:center;">No players purchased yet.</div>';
      renderXI(teamName, []);
      return;
    }
    
    const groups = { 'BATTERS': [], 'WICKETKEEPERS': [], 'ALL-ROUNDERS': [], 'BOWLERS': [], 'OTHER': [] };
    for (const p of squad) {
      let r = (p.role||'').toUpperCase();
      if (r.includes('BAT') && !r.includes('WK')) groups['BATTERS'].push(p);
      else if (r.includes('BOWL')) groups['BOWLERS'].push(p);
      else if (r.includes('ALL') || r === 'AR') groups['ALL-ROUNDERS'].push(p);
      else if (r.includes('WK') || r.includes('KEEP')) groups['WICKETKEEPERS'].push(p);
      else groups['OTHER'].push(p);
    }
    
    let html = '';
    const order = ['BATTERS', 'WICKETKEEPERS', 'ALL-ROUNDERS', 'BOWLERS', 'OTHER'];
    for (const role of order) {
      if (!groups[role].length) continue;
      html += `<div class="role-group">
        <div class="role-title">${role}</div>
        <div style="display:grid;grid-template-columns:repeat(auto-fill, minmax(220px, 1fr));gap:10px">
          ${groups[role].map(p => {
            const typeClass = p.is_overseas ? 'overseas' : 'indian';
            return `<div class="player-card ${typeClass}" id="pc-${p.id}">
              <div>
                <div style="font-weight:600;font-size:13px">${p.name}</div>
                <div style="font-size:11px;color:#8b98b3;margin-top:2px">${p.price} Cr</div>
              </div>
              <div>
                <button class="sm" id="btn-${p.id}" onclick="toggleXI('${teamName}', ${p.id}, '${p.name.replace(/'/g,"\\'")}', ${p.is_overseas})">+ XI</button>
              </div>
            </div>`;
          }).join('')}
        </div>
      </div>`;
    }
    
    $('#modalSquadList').innerHTML = html;
    renderXI(teamName, squad);
  } catch (e) {
    $('#modalSquadList').innerHTML = `<div class="muted" style="color:#ef4444">Error: ${e.message}</div>`;
  }
}

function closeSquadModal() {
  $('#squadModal').classList.add('hidden');
  currentModalTeam = null;
}

function toggleXI(team, playerId, playerName, isOverseas) {
  if (!playingXI[team]) playingXI[team] = [];
  const idx = playingXI[team].findIndex(p => p.id === playerId);
  
  if (idx > -1) { playingXI[team].splice(idx, 1); }
  else {
    if (playingXI[team].length >= 11) return alert("Playing XI is already full (11 players).");
    const currentOS = playingXI[team].filter(p => p.isOverseas).length;
    if (isOverseas && currentOS >= 4) return alert("Cannot add more than 4 Overseas players.");
    playingXI[team].push({ id: playerId, name: playerName, isOverseas: isOverseas });
  }
  localStorage.setItem('rps_xi', JSON.stringify(playingXI));
  openSquadModal(team); 
}

function dragStart(e, team, index) { dragSrcIdx = index; dragTeam = team; e.dataTransfer.effectAllowed = 'move'; }
function dragOver(e) { e.preventDefault(); e.currentTarget.classList.add('drag-over'); }
function dragLeave(e) { e.currentTarget.classList.remove('drag-over'); }
function drop(e, team, targetIdx) {
  e.preventDefault(); e.currentTarget.classList.remove('drag-over');
  if (dragTeam !== team || dragSrcIdx === targetIdx) return;
  const arr = playingXI[team];
  const item = arr.splice(dragSrcIdx, 1)[0];
  arr.splice(targetIdx, 0, item);
  localStorage.setItem('rps_xi', JSON.stringify(playingXI));
  openSquadModal(team);
}

function moveXI(team, index, direction) {
  if (index + direction < 0 || index + direction >= playingXI[team].length) return;
  const arr = playingXI[team];
  const temp = arr[index];
  arr[index] = arr[index + direction];
  arr[index + direction] = temp;
  localStorage.setItem('rps_xi', JSON.stringify(playingXI));
  openSquadModal(team);
}

async function renderXI(team, squad) {
  const xi = playingXI[team] || [];
  const osCount = xi.filter(p => p.isOverseas).length;
  const indCount = xi.length - osCount;
  
  $('#xiStats').innerHTML = `
    <span class="tag b">Selected: ${xi.length} / 11</span>
    <span class="tag ${osCount > 4 ? 'r' : 'g'}">OS: ${osCount} / 4</span>
    <span class="tag g">IND: ${indCount}</span>
  `;
  
  if (xi.length === 0) {
    $('#xiList').innerHTML = '<div class="muted" style="padding:10px;text-align:center;">Click "+ XI" on players to build your Playing 11.</div>';
  } else {
    $('#xiList').innerHTML = xi.map((p, i) => `
      <div class="xi-item" draggable="true" ondragstart="dragStart(event, '${team}', ${i})" ondragover="dragOver(event)" ondragleave="dragLeave(event)" ondrop="drop(event, '${team}', ${i})">
        <div style="display:flex;align-items:center;gap:8px;">
          <span style="color:#8b98b3;font-size:16px;">☰</span>
          <div style="font-size:13px;font-weight:600;">${i+1}. ${p.name}</div>
        </div>
        <div style="display:flex;gap:4px;">
          <button class="sm" onclick="moveXI('${team}', ${i}, -1)">↑</button>
          <button class="sm" onclick="moveXI('${team}', ${i}, 1)">↓</button>
          <button class="sm" style="background:#ef4444;border:0;color:#fff;margin-left:4px" onclick="toggleXI('${team}', ${p.id}, '${p.name.replace(/'/g,"\\'")}', ${p.isOverseas})">✕</button>
        </div>
      </div>
    `).join('');
  }
  
  if (squad) {
    squad.forEach(p => {
      const btn = document.getElementById('btn-' + p.id);
      const card = document.getElementById('pc-' + p.id);
      if (btn && card) {
        if (xi.find(x => x.id === p.id)) {
          btn.textContent = 'Remove'; btn.style.background = '#3b82f6'; btn.style.color = '#fff';
          card.classList.add('in-xi');
        } else {
          btn.textContent = '+ XI'; btn.style.background = ''; btn.style.color = '';
          card.classList.remove('in-xi');
        }
      }
    });
  }
}

// ── teams ───────────────────────────────
async function loadTeams() {
  const s = await api('/api/state');
  let html = '<tr><th>Team</th><th>Purse</th><th>Spent</th><th>Remaining</th><th>Players</th><th>Indian</th><th>Overseas</th><th>75% min</th></tr>';
  for (const t of Object.keys(s)) {
    const v = s[t];
    html += `<tr class="team-row" onclick="openSquadModal('${t}')" style="cursor:pointer">
      <td><b>${t}</b> <span style="font-size:10px;color:#8b98b3">▼</span></td>
      <td>${v.spent + v.remaining} Cr</td>
      <td>${v.spent.toFixed(2)}</td>
      <td><b>${v.remaining.toFixed(2)}</b></td>
      <td>${v.count} / ${v.count + v.slots_left}</td>
      <td>${v.indian}${v.indian_ok?'':' <span class="tag r">low</span>'}</td>
      <td>${v.overseas}${v.overseas_ok?'':' <span class="tag r">over</span>'}</td>
      <td>${v.min_spend_ok?'<span class="tag g">MET</span>':'<span class="tag y">pending</span>'}</td>
    </tr>`;
  }
  $('#teamsTable').innerHTML = html;
}

// ── players ─────────────────────────────
function toggleRole(btn) {
  btn.classList.toggle('active');
  loadPlayers();
}

async function loadPlayers() {
  const q = $('#pSearch').value, cat = $('#pCategory').value, sort = $('#pSort').value;
  const activeRoles = Array.from(document.querySelectorAll('.role-btn.active')).map(b => b.dataset.role);
  const rolesStr = activeRoles.join(',');

  const url = `/api/players?q=${encodeURIComponent(q)}&category=${encodeURIComponent(cat)}&roles=${encodeURIComponent(rolesStr)}&sort=${sort}`;
  const rows = await api(url);
  
  $('#playersCount').textContent = `${rows.length} player${rows.length!==1?'s':''} found`;

  const cats = [...new Set(rows.map(r => r.category))].sort();
  const cur = $('#pCategory').value;
  if ($('#pCategory').options.length <= 1 || cur === '') {
    $('#pCategory').innerHTML = '<option value="">All categories</option>' + cats.map(c => `<option ${c===cur?'selected':''}>${c}</option>`).join('');
  }

  if (!rows.length) { $('#playersTable').innerHTML = '<div class="muted" style="padding:10px">No players match these filters.</div>'; return; }

  let html = '<table><tr><th>#</th><th>Name</th><th>Role</th><th>Bowl</th><th>Pos</th><th>Rating</th><th>Overseas</th><th>Base</th><th></th></tr>';
  let lastCat = null;
  let counter = 1;

  for (const r of rows) {
    let group = r.category;

    if (sort === 'batting_order') {
      const po = r.pos_order;
      if (po <= 3) group = 'Top Order (1-3)';
      else if (po <= 5) group = 'Middle Order (4-5)';
      else if (po <= 8) group = 'Lower Order (6-8)';
      else group = 'Unspecified / Bowlers';
    }

    if (group !== lastCat) {
      html += `<tr><td colspan="9" class="cat-hdr">${group}</td></tr>`;
      lastCat = group;
    }

    const ov = r.overseas === 1 ? '<span class="tag y">OS</span>' : r.overseas === 0 ? '<span class="tag g">IND</span>' : '—';
    html += `<tr>
      <td>${counter++}</td>
      <td><b>${r.name}</b></td>
      <td>${r.role||'—'}</td>
      <td>${r.bowling_style||'—'}</td>
      <td>${r.batting_position||'—'}</td>
      <td>${r.rating??'—'}</td>
      <td>${ov}</td>
      <td>${r.base_price||'—'}</td>
      <td><button class="sm" onclick="quickBuy('${r.name.replace(/'/g,"\\'")}', ${r.overseas??''})">Buy ▸</button></td>
    </tr>`;
  }
  html += '</table>';
  $('#playersTable').innerHTML = html;
}

function quickBuy(name, os) {
  $$('.tabs button').forEach(b => b.classList.remove('active'));
  $('#tab-auction').classList.remove('hidden');
  $('#tab-teams').classList.add('hidden');
  $('#tab-players').classList.add('hidden');
  $('#tab-import').classList.add('hidden');
  document.querySelector('[data-tab=auction]').classList.add('active');
  $('#bPlayer').value = name;
  if (os === 1 || os === 0) $('#bOS').checked = os === 1;
  $('#bPrice').focus();
}

async function clearPlayers() {
  if (!confirm('Delete ALL imported players? This does not delete purchases.')) return;
  await fetch('/api/players/all', {method:'DELETE'});
  loadPlayers();
}

// ── import ──────────────────────────────
const dropzone = $('#drop'), fileInput = $('#fileInput');
dropzone.onclick = () => fileInput.click();
dropzone.ondragover = e => { e.preventDefault(); dropzone.classList.add('hover'); };
dropzone.ondragleave = () => dropzone.classList.remove('hover');
dropzone.ondrop = e => { e.preventDefault(); dropzone.classList.remove('hover'); handleFiles(e.dataTransfer.files); };
fileInput.onchange = () => handleFiles(fileInput.files);

async function handleFiles(files) {
  if (!files.length) return;
  $('#importErr').textContent = '';
  $('#rawTextCard').classList.add('hidden');
  const fd = new FormData();
  for (const f of files) fd.append('files', f);
  dropzone.innerHTML = '<div style="font-size:16px">⏳ Reading ' + files.length + ' file(s)…</div>';
  try {
    const res = await api('/api/import/preview', {method:'POST', body:fd});
    if (res.errors && res.errors.length) $('#importErr').textContent = res.errors.join(' | ');
    if (res.rows.length === 0 && res.raw) {
      $('#rawTextCard').classList.remove('hidden');
      $('#rawText').textContent = res.raw;
    }
    previewRows = res.rows.map(r => ({...r, _checked: true}));
    renderPreview();
  } catch (e) { $('#importErr').textContent = '✗ ' + e.message; }
  dropzone.innerHTML = '<div style="font-size:16px;margin-bottom:6px">Drop more files, or click to browse</div><div class="muted">CSV · XLSX · PDF · JPG · PNG — multiple files supported</div>';
  fileInput.value = '';
}

function renderPreview() {
  if (!previewRows.length) { $('#previewCard').classList.add('hidden'); return; }
  $('#previewCard').classList.remove('hidden');
  $('#previewCount').textContent = previewRows.length;
  let html = '<table><tr><th></th><th>Name</th><th>Role</th><th>Bowl</th><th>Category</th><th>Pos</th><th>Rating</th><th>OV</th><th>Base</th></tr>';
  previewRows.forEach((r, i) => {
    html += `<tr>
      <td><input type="checkbox" ${r._checked?'checked':''} onchange="previewRows[${i}]._checked=this.checked"></td>
      <td>${r.name}</td><td>${r.role||''}</td><td>${r.bowling_style||''}</td><td>${r.category||''}</td><td>${r.batting_position||''}</td>
      <td>${r.rating??''}</td><td>${r.overseas===1?'OS':r.overseas===0?'IND':''}</td>
      <td>${r.base_price||''}</td>
    </tr>`;
  });
  html += '</table>';
  $('#previewTable').innerHTML = html;
}

async function commitImport() {
  const keep = previewRows.filter(r => r._checked);
  if (!keep.length) { alert('Nothing selected'); return; }
  try {
    const res = await api('/api/import/commit', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify(keep),
    });
    alert(`✓ Inserted ${res.inserted} new, updated ${res.updated} existing`);
    previewRows = [];
    renderPreview();
    loadPlayers();
  } catch (e) { alert('✗ ' + e.message); }
}

// ── boot ────────────────────────────────
loadAuction();
loadPlayers();
setInterval(loadAuction, 3000);
</script></body></html>
"""

# ─── launch ─────────────────────────────────────────────────
def free_port(start=8000):
    for p in range(start, start + 20):
        s = socket.socket(); s.settimeout(0.2)
        try:
            s.bind(("0.0.0.0", p)); s.close(); return p
        except OSError: continue
    return start

def lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]; s.close(); return ip
    except Exception: return "127.0.0.1"

if __name__ == "__main__":
    init()
    port = free_port(8000)
    ip = lan_ip()
    print("=" * 60)
    print(f" RPS AUCTION PLATFORM is running")
    print(f" Local:  http://localhost:{port}")
    print(f" LAN:    http://{ip}:{port}   (use this on Laptop 2/3)")
    print(f" Data:   {DB}")
    print("=" * 60)
    print(" Keep this window OPEN. Close it to stop the app.")
    print("=" * 60)
    def open_browser():
        import time; time.sleep(1.2)
        webbrowser.open(f"http://localhost:{port}")
    threading.Thread(target=open_browser, daemon=True).start()
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="warning")