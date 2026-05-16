"""
BTC Ladder Bot — Render-ready
==============================
Run modes:
  python bot.py           → GUI (local)
  python bot.py live      → headless live/dry trader
  python bot.py backtest  → backtest
  python bot.py valrtest  → VALR connection test
  python bot.py premium   → USDC premium check

Environment variables (set in Render dashboard):
  EXCHANGE          VALR or BINANCE  (default: VALR)
  TRADING_PAIR      e.g. BTCUSDC      (default: BTCUSDC)
  VALR_API_KEY      your VALR API key
  VALR_API_SECRET   your VALR API secret
  BINANCE_API_KEY   your Binance API key
  BINANCE_API_SECRET your Binance API secret
  DRY_RUN           true/false       (default: true)
  CHECK_INTERVAL_MIN minutes between checks (default: 5)
  GEN_CAPITAL       capital per generation (default: 1100)
  SLOT_SIZE         capital per slot       (default: 100)
  STATE_DIR         directory for state file (default: /tmp)
  PORT              HTTP health-check port (Render sets this)
  TELEGRAM_TOKEN    optional — bot token for trade alerts
  TELEGRAM_CHAT_ID  optional — chat id for trade alerts
"""

import os, sys, json, time, logging, hashlib, hmac as hmac_lib
import urllib.request, urllib.error, urllib.parse
import base64, threading, signal
from datetime import datetime, timezone
from pathlib import Path
from http.server import HTTPServer, BaseHTTPRequestHandler

# ── Optional imports ───────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    GUI_OK = True
except ImportError:
    GUI_OK = False

try:
    import pandas as pd
    PANDAS_OK = True
except ImportError:
    PANDAS_OK = False

# =============================================================
#  SECTION 1 — CONFIGURATION
# =============================================================
APP_DIR     = Path(os.path.dirname(os.path.abspath(__file__)))
CONFIG_FILE = APP_DIR / "strategy_config.json"

# State file goes in STATE_DIR env var, or /tmp on Linux, or APP_DIR locally
_state_dir  = Path(os.environ.get("STATE_DIR", "/tmp" if sys.platform != "win32" else str(APP_DIR)))
STATE_FILE  = _state_dir / "bot_state.json"
LOG_FILE    = _state_dir / "bot.log"

DEFAULT_PARAMS = {
    "exchange":           os.environ.get("EXCHANGE", "VALR"),
    "pair":               os.environ.get("TRADING_PAIR", "BTCUSDC"),
    "gen_capital":        float(os.environ.get("GEN_CAPITAL", "1100")),
    "slot_size":          float(os.environ.get("SLOT_SIZE", "100")),
    "split_min_size":     30.0,
    "max_generations":    3,
    "buy_drop_pct":       2.0,
    "lookback_hours":     12,
    "dip_sell_pct":       5.0,
    "sell_decay_days":    7,
    "sell_decay_rate":    1.0,
    "sell_floor_pct":     5.0,
    "escalation_start":   5,
    "escalation_step":    2.0,
    "escalation_max_lvl": 3,
    "entry_decay_days":   15,
    "entry_decay_rate":   1.0,
    "new_gen_threshold":  20.0,
    "new_gen_cooldown_d": 30,
    "check_interval_min": int(os.environ.get("CHECK_INTERVAL_MIN", "5")),
    "dry_run":            os.environ.get("DRY_RUN", "true").lower() != "false",
}

PARAM_LABELS = {
    "exchange":           ("Exchange",               "VALR or BINANCE"),
    "pair":               ("Trading pair",           "e.g. BTCUSDC or BTCUSDT"),
    "gen_capital":        ("Capital per generation", "Currency per generation"),
    "slot_size":          ("Slot size",              "Capital per slot"),
    "split_min_size":     ("Min slot size (split)",  "No split below this"),
    "max_generations":    ("Max generations",        "Hard cap on generations"),
    "buy_drop_pct":       ("Buy trigger drop %",     "% drop over lookback window"),
    "lookback_hours":     ("Lookback hours",         "Window for buy trigger"),
    "dip_sell_pct":       ("Sell target %",          "Base sell target per trade"),
    "sell_decay_days":    ("Sell decay start (days)","Days before target decays"),
    "sell_decay_rate":    ("Sell decay rate %/day",  "% reduction per day"),
    "sell_floor_pct":     ("Sell floor %",           "Minimum sell target"),
    "escalation_start":   ("Escalation after N pos", "Start escalating after N open"),
    "escalation_step":    ("Escalation step %",      "Extra % per escalation level"),
    "escalation_max_lvl": ("Max escalation levels",  "Cap on escalation depth"),
    "entry_decay_days":   ("Entry decay start (days)","Days before entry eases"),
    "entry_decay_rate":   ("Entry decay rate %/day", "% ease per day"),
    "new_gen_threshold":  ("New gen trigger %",      "Price below peak to spawn gen"),
    "new_gen_cooldown_d": ("New gen cooldown (days)","Min days between generations"),
    "check_interval_min": ("Check interval (min)",   "How often bot checks price"),
    "dry_run":            ("Dry run mode",           "True=simulate False=live"),
}

def load_strategy_config():
    cfg = dict(DEFAULT_PARAMS)
    if CONFIG_FILE.exists():
        try:
            cfg.update(json.loads(CONFIG_FILE.read_text()))
        except Exception:
            pass
    # Env vars always win over saved file
    for k, v in [
        ("exchange",           os.environ.get("EXCHANGE")),
        ("pair",               os.environ.get("TRADING_PAIR")),
        ("dry_run",            os.environ.get("DRY_RUN")),
        ("check_interval_min", os.environ.get("CHECK_INTERVAL_MIN")),
        ("gen_capital",        os.environ.get("GEN_CAPITAL")),
        ("slot_size",          os.environ.get("SLOT_SIZE")),
    ]:
        if v is not None:
            orig = DEFAULT_PARAMS[k]
            if isinstance(orig, bool):
                cfg[k] = v.lower() != "false"
            elif isinstance(orig, int):
                cfg[k] = int(float(v))
            elif isinstance(orig, float):
                cfg[k] = float(v)
            else:
                cfg[k] = v
    return cfg

def save_strategy_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2))

# =============================================================
#  SECTION 2 — API CREDENTIALS (env vars first, file fallback)
# =============================================================
_API_FILE = APP_DIR / "api_config.enc"

def _machine_key():
    import socket
    seed = socket.gethostname() + str(os.getenv("USERNAME", "user"))
    return hashlib.sha256(seed.encode()).hexdigest()[:32]

def _xor_enc(data):
    key = _machine_key().encode()
    raw = json.dumps(data).encode()
    enc = bytes(b ^ key[i % len(key)] for i, b in enumerate(raw))
    return base64.b64encode(enc)

def _xor_dec(data):
    key = _machine_key().encode()
    raw = base64.b64decode(data)
    dec = bytes(b ^ key[i % len(key)] for i, b in enumerate(raw))
    return json.loads(dec.decode())

def load_api_config():
    """Return API config. Env vars take priority — this is how Render injects secrets."""
    exchange = os.environ.get("EXCHANGE", "VALR")

    # Check env vars for both exchanges
    valr_key    = os.environ.get("VALR_API_KEY", "")
    valr_sec    = os.environ.get("VALR_API_SECRET", "")
    bnb_key     = os.environ.get("BINANCE_API_KEY", "")
    bnb_sec     = os.environ.get("BINANCE_API_SECRET", "")

    if exchange == "BINANCE" and bnb_key:
        return {"exchange": "BINANCE", "api_key": bnb_key, "api_secret": bnb_sec,
                "pair": os.environ.get("TRADING_PAIR", "BTCUSDT")}
    if valr_key:
        return {"exchange": "VALR", "api_key": valr_key, "api_secret": valr_sec,
                "pair": os.environ.get("TRADING_PAIR", "BTCUSDC")}

    # Fallback: encrypted local file (local dev only)
    if _API_FILE.exists():
        try:
            return _xor_dec(_API_FILE.read_bytes())
        except Exception:
            pass

    return {"api_key": "", "api_secret": "", "exchange": exchange,
            "pair": os.environ.get("TRADING_PAIR", "BTCUSDC")}

def save_api_config(cfg):
    _API_FILE.write_bytes(_xor_enc(cfg))

# =============================================================
#  SECTION 3 — EXCHANGE API CLIENTS
# =============================================================
FEE_RATE = 0.001

def valr_sign(secret, ts, verb, path, body=""):
    payload = f"{ts}{verb.upper()}{path}{body}"
    return hmac_lib.new(bytearray(secret, "utf-8"),
                        bytearray(payload, "utf-8"),
                        digestmod=hashlib.sha512).hexdigest()

def valr_request(path, method="GET", body=None, public=False,
                 api_key="", api_secret=""):
    ts       = str(int(time.time() * 1000))
    body_str = json.dumps(body) if body else ""
    headers  = {"Content-Type": "application/json"}
    if not public and api_key:
        headers["X-VALR-API-KEY"]   = api_key
        headers["X-VALR-SIGNATURE"] = valr_sign(api_secret, ts, method, path, body_str)
        headers["X-VALR-TIMESTAMP"] = ts
    data = body_str.encode() if body_str else None
    req  = urllib.request.Request(
        "https://api.valr.com" + path, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            return json.loads(r.read().decode()), r.status
    except urllib.error.HTTPError as e:
        return {"error": e.code, "detail": e.read().decode()[:300]}, e.code
    except Exception as e:
        return {"error": str(e)}, 0

def binance_sign(secret, qs):
    return hmac_lib.new(secret.encode(), qs.encode(),
                        digestmod=hashlib.sha256).hexdigest()

def binance_request(path, method="GET", params=None, public=False,
                    api_key="", api_secret=""):
    p = dict(params or {})
    if not public and api_secret:
        p["timestamp"] = str(int(time.time() * 1000))
        qs  = urllib.parse.urlencode(p)
        qs += "&signature=" + binance_sign(api_secret, qs)
    else:
        qs = urllib.parse.urlencode(p)
    url     = "https://api.binance.com" + path + ("?" + qs if qs else "")
    headers = {"Content-Type": "application/json"}
    if not public and api_key:
        headers["X-MBX-APIKEY"] = api_key
    req = urllib.request.Request(url, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            return json.loads(r.read().decode()), r.status
    except urllib.error.HTTPError as e:
        return {"error": e.code, "detail": e.read().decode()[:300]}, e.code
    except Exception as e:
        return {"error": str(e)}, 0

def parse_pair(pair):
    for q in ["USDT", "USDC", "BUSD", "USDC", "EUR", "GBP", "USD"]:
        if pair.endswith(q):
            return pair[:-len(q)], q
    return pair[:3], pair[3:]

def fmt_price(val, quote):
    sym = {"USDC": "R", "USD": "$", "USDT": "$", "USDC": "$", "EUR": "€", "GBP": "£"}.get(quote, quote + " ")
    if val >= 10000: return f"{sym}{val:,.0f}"
    if val >= 1:     return f"{sym}{val:,.2f}"
    return f"{sym}{val:.6f}"

def get_live_price(exchange, pair, api_key, api_secret):
    try:
        if exchange == "VALR":
            data, _ = valr_request(f"/v1/public/{pair}/marketsummary",
                                   public=True, api_key=api_key, api_secret=api_secret)
            if "lastTradedPrice" in data:
                return {"price": float(data["lastTradedPrice"]),
                        "bid":   float(data.get("bidPrice", 0)),
                        "ask":   float(data.get("askPrice", 0)),
                        "high":  float(data.get("highPrice", 0)),
                        "low":   float(data.get("lowPrice", 0)),
                        "vol":   float(data.get("baseVolume", 0)),
                        "chg":   data.get("changeFromPrevious", "?")}
        elif exchange == "BINANCE":
            data, _ = binance_request("/api/v3/ticker/24hr",
                                      params={"symbol": pair}, public=True,
                                      api_key=api_key, api_secret=api_secret)
            if "lastPrice" in data:
                prev = float(data.get("prevClosePrice", 0))
                last = float(data["lastPrice"])
                chg  = ((last - prev) / prev * 100) if prev else 0
                return {"price": last,
                        "bid":   float(data.get("bidPrice", 0)),
                        "ask":   float(data.get("askPrice", 0)),
                        "high":  float(data.get("highPrice", 0)),
                        "low":   float(data.get("lowPrice", 0)),
                        "vol":   float(data.get("volume", 0)),
                        "chg":   f"{chg:.4f}"}
    except Exception:
        pass
    return None

def get_live_price_with_retry(exchange, pair, api_key, api_secret, retries=4):
    """Fetch price with exponential backoff — critical for cloud stability."""
    delay = 10
    for attempt in range(retries):
        result = get_live_price(exchange, pair, api_key, api_secret)
        if result:
            return result
        if attempt < retries - 1:
            logging.warning(f"Price fetch failed (attempt {attempt+1}/{retries}), retry in {delay}s")
            time.sleep(delay)
            delay *= 2
    return None

def get_account_balance(exchange, pair, api_key, api_secret):
    base, quote = parse_pair(pair)
    try:
        if exchange == "VALR":
            data, _ = valr_request("/v1/account/balances",
                                   api_key=api_key, api_secret=api_secret)
            if isinstance(data, list):
                result = {}
                for b in data:
                    cur = b.get("currency", "")
                    tot = float(b.get("total", 0))
                    av  = float(b.get("available", 0))
                    rsv = float(b.get("reserved", 0))
                    if cur in (base, quote) or tot > 0:
                        result[cur] = {"total": tot, "available": av, "reserved": rsv,
                                       "is_base": cur == base, "is_quote": cur == quote}
                return result
        elif exchange == "BINANCE":
            data, _ = binance_request("/api/v3/account",
                                      api_key=api_key, api_secret=api_secret)
            if "balances" in data:
                result = {}
                for b in data["balances"]:
                    cur  = b.get("asset", "")
                    free = float(b.get("free", 0))
                    lkd  = float(b.get("locked", 0))
                    tot  = free + lkd
                    if cur in (base, quote) or tot > 0:
                        result[cur] = {"total": tot, "available": free, "reserved": lkd,
                                       "is_base": cur == base, "is_quote": cur == quote}
                return result
    except Exception as e:
        return {"_error": str(e)}
    return None

def get_all_pairs(exchange, api_key, api_secret):
    try:
        if exchange == "VALR":
            data, _ = valr_request("/v1/public/pairs", public=True,
                                   api_key=api_key, api_secret=api_secret)
            if isinstance(data, list):
                return sorted(p.get("symbol", "") for p in data
                              if p.get("symbol", "") and p.get("active", True))
        elif exchange == "BINANCE":
            data, _ = binance_request("/api/v3/exchangeInfo", public=True,
                                      api_key=api_key, api_secret=api_secret)
            if "symbols" in data:
                return sorted(s["symbol"] for s in data["symbols"]
                              if s.get("status") == "TRADING")
    except Exception:
        pass
    return []

# =============================================================
#  SECTION 4 — TELEGRAM NOTIFICATIONS (optional)
# =============================================================
_TG_TOKEN   = os.environ.get("TELEGRAM_TOKEN", "")
_TG_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

def _tg_send(msg):
    if not _TG_TOKEN or not _TG_CHAT_ID:
        return
    try:
        url  = f"https://api.telegram.org/bot{_TG_TOKEN}/sendMessage"
        body = json.dumps({"chat_id": _TG_CHAT_ID, "text": msg, "parse_mode": "Markdown"}).encode()
        req  = urllib.request.Request(url, data=body,
                                      headers={"Content-Type": "application/json"}, method="POST")
        urllib.request.urlopen(req, timeout=8)
    except Exception as e:
        logging.warning(f"Telegram send failed: {e}")

# =============================================================
#  SECTION 5 — WEB DASHBOARD (replaces plain health endpoint)
# =============================================================
_bot_status = {"alive": True, "last_price": None, "last_tick": None, "dry_run": True}
_snap       = {"state": None, "price_data": None, "cfg": None}
_snap_lock  = threading.Lock()
_log_lines  = []

class _LogCapture(logging.Handler):
    def emit(self, record):
        _log_lines.append(self.format(record))
        if len(_log_lines) > 300:
            del _log_lines[:-300]

_DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>BTC Ladder Bot</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#0d1117;color:#e6edf3;min-height:100vh}
#topbar{background:#161b22;border-bottom:1px solid #30363d;padding:12px 24px;display:flex;align-items:center;gap:14px;flex-wrap:wrap}
#topbar h1{font-size:17px;font-weight:700;color:#58a6ff;white-space:nowrap}
.badge{padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;letter-spacing:.5px;white-space:nowrap}
.b-ok{background:#1a7f4b33;color:#3fb950;border:1px solid #1a7f4b}
.b-dry{background:#9a500033;color:#f0883e;border:1px solid #9a5000}
.b-err{background:#6e242433;color:#f85149;border:1px solid #6e2424}
#b-mode{cursor:pointer;user-select:none}
#b-mode:hover{filter:brightness(1.25)}
#price-ticker{margin-left:auto;text-align:right;white-space:nowrap}
#price-ticker .pair{color:#8b949e;font-size:11px}
#price-ticker .price{font-size:20px;font-weight:700;color:#e6edf3}
#price-ticker .chg{font-size:12px;margin-left:6px}
#tabs{background:#161b22;border-bottom:1px solid #30363d;padding:0 20px;display:flex;gap:2px}
.tb{padding:10px 18px;border:none;background:none;color:#8b949e;cursor:pointer;font-size:13px;border-bottom:2px solid transparent;font-family:inherit}
.tb.on{color:#58a6ff;border-bottom-color:#58a6ff}
.tab{display:none;padding:20px;max-width:1400px;margin:0 auto}
.tab.on{display:block}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:20px}
.card{background:#161b22;border:1px solid #30363d;border-radius:8px;padding:14px 18px}
.card .lbl{color:#8b949e;font-size:11px;margin-bottom:4px;text-transform:uppercase;letter-spacing:.5px}
.card .val{font-size:21px;font-weight:700}
.green{color:#3fb950}.red{color:#f85149}.blue{color:#58a6ff}.orange{color:#f0883e}
.gen{margin-bottom:18px;border:1px solid #30363d;border-radius:8px;overflow:hidden}
.gen-hdr{background:#21262d;padding:12px 16px;display:flex;flex-wrap:wrap;gap:18px;align-items:center}
.gen-hdr .gtitle{font-weight:700;color:#58a6ff;font-size:14px}
.gstat{font-size:12px;color:#8b949e}.gstat b{color:#e6edf3;font-weight:600}
.slots{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:8px;padding:12px;background:#0d1117}
.slot{border-radius:6px;padding:10px 12px;border:1px solid #30363d}
.slot.idle{background:#161b22}
.slot.dip{background:#0d2818;border-color:#1a7f4b55}
.slot .sh{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px}
.slot .sid{font-size:11px;color:#6e7681;font-weight:700}
.stbadge{font-size:10px;font-weight:700;padding:2px 7px;border-radius:10px}
.stbadge.idle{background:#21262d;color:#6e7681}
.stbadge.dip{background:#1a7f4b33;color:#3fb950;border:1px solid #1a7f4b55}
.slot .sentry{font-size:13px;font-weight:600;color:#e6edf3;margin-bottom:2px}
.slot .stgt{font-size:11px;color:#8b949e;margin-bottom:3px}
.slot .spnl{font-size:12px;font-weight:700}
.slot .swait{font-size:12px;color:#484f58;margin-top:4px}
.tradelog{margin-top:20px}
.tradelog h3{color:#8b949e;font-size:12px;text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px}
.tlog-row{display:grid;grid-template-columns:90px 60px 1fr 1fr 90px;gap:8px;font-size:12px;padding:6px 0;border-bottom:1px solid #21262d;align-items:center}
.tlog-row.hdr{color:#6e7681;font-size:11px;font-weight:700}
.tlog-buy{color:#3fb950}.tlog-sell{color:#f85149}
.fsec{background:#161b22;border:1px solid #30363d;border-radius:8px;padding:18px 20px;margin-bottom:14px}
.fsec h3{color:#58a6ff;font-size:13px;font-weight:700;margin-bottom:14px;padding-bottom:8px;border-bottom:1px solid #30363d}
.fgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:12px}
.ff{display:flex;flex-direction:column;gap:4px}
.ff label{font-size:11px;color:#8b949e;text-transform:uppercase;letter-spacing:.4px}
.ff input,.ff select{background:#0d1117;border:1px solid #30363d;color:#e6edf3;padding:8px 10px;border-radius:6px;font-size:13px;font-family:inherit;width:100%}
.ff input:focus,.ff select:focus{outline:none;border-color:#58a6ff;box-shadow:0 0 0 3px #58a6ff22}
.ff small{color:#6e7681;font-size:11px}
.btnrow{display:flex;gap:10px;flex-wrap:wrap;margin-top:18px}
.btn{padding:9px 18px;border-radius:6px;border:1px solid transparent;cursor:pointer;font-size:13px;font-weight:600;font-family:inherit;transition:opacity .15s}
.btn:hover{opacity:.85}
.btn-g{background:#238636;color:#fff;border-color:#2ea043}
.btn-s{background:#21262d;color:#e6edf3;border-color:#30363d}
.btn-o{background:#9a500033;color:#f0883e;border-color:#9a5000}
#log-pre{background:#0d1117;border:1px solid #30363d;border-radius:8px;padding:14px 16px;font-family:'Cascadia Code',monospace;font-size:12px;max-height:65vh;overflow-y:auto;white-space:pre-wrap;color:#8b949e;line-height:1.5}
#toast{position:fixed;bottom:24px;right:24px;padding:12px 20px;border-radius:8px;font-size:13px;font-weight:600;display:none;z-index:999;animation:fadein .2s}
@keyframes fadein{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
.refresh-info{color:#484f58;font-size:11px;white-space:nowrap}
.no-data{text-align:center;color:#484f58;padding:48px 20px}
</style>
</head>
<body>
<div id="topbar">
  <h1>&#9889; BTC Ladder Bot</h1>
  <span id="b-status" class="badge b-err">LOADING</span>
  <span id="b-mode"   class="badge b-dry" style="display:none" title="Click to toggle Dry Run / Live mode" onclick="toggleMode()"></span>
  <div id="price-ticker">
    <div class="pair" id="t-pair"></div>
    <span class="price" id="t-price">&#8212;</span>
    <span class="chg"   id="t-chg"></span>
  </div>
  <span class="refresh-info" id="t-refresh"></span>
</div>
<div id="tabs">
  <button class="tb on" onclick="showTab('dashboard',this)">Dashboard</button>
  <button class="tb"    onclick="showTab('strategy',this)">Strategy</button>
  <button class="tb"    onclick="showTab('logs',this)">Logs</button>
</div>

<div id="tab-dashboard" class="tab on">
  <div class="cards">
    <div class="card"><div class="lbl">Portfolio Value</div><div class="val blue" id="c-port">&#8212;</div></div>
    <div class="card"><div class="lbl">Invested</div><div class="val" id="c-inv">&#8212;</div></div>
    <div class="card"><div class="lbl">Total P&amp;L</div><div class="val" id="c-pnl">&#8212;</div></div>
    <div class="card"><div class="lbl">Realised Profit</div><div class="val green" id="c-real">&#8212;</div></div>
    <div class="card"><div class="lbl">Open Positions</div><div class="val" id="c-open">&#8212;</div></div>
    <div class="card"><div class="lbl" id="c-qlbl">Quote Balance</div><div class="val blue" id="c-qbal">&#8212;</div><div style="font-size:10px;color:#484f58;margin-top:2px" id="c-qrsv"></div></div>
    <div class="card"><div class="lbl" id="c-blbl">Base Balance</div><div class="val" id="c-bbal">&#8212;</div><div style="font-size:10px;color:#484f58;margin-top:2px" id="c-brsv"></div></div>
  </div>
  <div id="gens"></div>
</div>

<div id="tab-strategy" class="tab">
  <form id="sf" onsubmit="saveStrategy(event)">
    <div class="fsec">
      <h3>Exchange &amp; Pair</h3>
      <div class="fgrid">
        <div class="ff"><label>Exchange</label>
          <select name="exchange" id="f-ex" onchange="loadPairs()">
            <option value="VALR">VALR</option>
            <option value="BINANCE">Binance</option>
          </select>
        </div>
        <div class="ff"><label>Trading Pair</label>
          <select name="pair" id="f-pair"><option>Loading&#8230;</option></select>
        </div>
      </div>
    </div>
    <div class="fsec">
      <h3>Capital Settings</h3>
      <div class="fgrid">
        <div class="ff"><label>Capital per Generation</label><input type="number" name="gen_capital" step="any"><small>Quote currency per generation</small></div>
        <div class="ff"><label>Slot Size</label><input type="number" name="slot_size" step="any"><small>Capital per slot</small></div>
        <div class="ff"><label>Min Slot Size (split)</label><input type="number" name="split_min_size" step="any"><small>No split below this</small></div>
        <div class="ff"><label>Max Generations</label><input type="number" name="max_generations"><small>Hard cap on generations</small></div>
      </div>
    </div>
    <div class="fsec">
      <h3>Buy Logic</h3>
      <div class="fgrid">
        <div class="ff"><label>Buy Trigger Drop %</label><input type="number" name="buy_drop_pct" step="0.1"><small>% drop over lookback window</small></div>
        <div class="ff"><label>Lookback Hours</label><input type="number" name="lookback_hours"><small>Window used for the buy signal</small></div>
      </div>
    </div>
    <div class="fsec">
      <h3>Sell Logic</h3>
      <div class="fgrid">
        <div class="ff"><label>Sell Target %</label><input type="number" name="dip_sell_pct" step="0.1"><small>Base sell target per trade</small></div>
        <div class="ff"><label>Sell Decay Start (days)</label><input type="number" name="sell_decay_days"><small>Days before target decays</small></div>
        <div class="ff"><label>Sell Decay Rate %/day</label><input type="number" name="sell_decay_rate" step="0.1"><small>% reduction per day</small></div>
        <div class="ff"><label>Sell Floor %</label><input type="number" name="sell_floor_pct" step="0.1"><small>Minimum sell target</small></div>
      </div>
    </div>
    <div class="fsec">
      <h3>Escalation</h3>
      <div class="fgrid">
        <div class="ff"><label>Escalation After N Positions</label><input type="number" name="escalation_start"><small>Start escalating after N open slots</small></div>
        <div class="ff"><label>Escalation Step %</label><input type="number" name="escalation_step" step="0.1"><small>Extra % per escalation level</small></div>
        <div class="ff"><label>Max Escalation Levels</label><input type="number" name="escalation_max_lvl"><small>Cap on escalation depth</small></div>
      </div>
    </div>
    <div class="fsec">
      <h3>Entry Decay</h3>
      <div class="fgrid">
        <div class="ff"><label>Entry Decay Start (days)</label><input type="number" name="entry_decay_days"><small>Days before entry threshold eases</small></div>
        <div class="ff"><label>Entry Decay Rate %/day</label><input type="number" name="entry_decay_rate" step="0.1"><small>% ease per day</small></div>
      </div>
    </div>
    <div class="fsec">
      <h3>Generation Spawning</h3>
      <div class="fgrid">
        <div class="ff"><label>New Gen Trigger %</label><input type="number" name="new_gen_threshold" step="0.1"><small>Price below peak to spawn new gen</small></div>
        <div class="ff"><label>New Gen Cooldown (days)</label><input type="number" name="new_gen_cooldown_d"><small>Min days between generations</small></div>
      </div>
    </div>
    <div class="fsec">
      <h3>Bot Settings</h3>
      <div class="fgrid">
        <div class="ff"><label>Check Interval (minutes)</label><input type="number" name="check_interval_min"><small>How often the bot checks the price</small></div>
        <div class="ff"><label>Mode</label>
          <select name="dry_run" id="f-dry">
            <option value="true">Dry Run (paper trading)</option>
            <option value="false">LIVE TRADING (real orders)</option>
          </select>
          <small class="orange">&#9888; Live mode places real orders on your account</small>
        </div>
      </div>
    </div>
    <div class="btnrow">
      <button type="submit" class="btn btn-g">Save Strategy</button>
      <button type="button" class="btn btn-s" onclick="exportStrategy()">Export JSON</button>
      <label class="btn btn-s" style="cursor:pointer">Import JSON
        <input type="file" accept=".json" style="display:none" onchange="importStrategy(this)">
      </label>
    </div>
    <p style="color:#484f58;font-size:11px;margin-top:12px">&#8505; Strategy changes take effect on the next bot restart.</p>
  </form>
</div>

<div id="tab-logs" class="tab">
  <pre id="log-pre">Loading&#8230;</pre>
</div>

<div id="toast"></div>

<script>
var _cfg = null;

function showTab(name, btn) {
  document.querySelectorAll('.tab').forEach(function(e){e.classList.remove('on')});
  document.querySelectorAll('.tb').forEach(function(e){e.classList.remove('on')});
  document.getElementById('tab-'+name).classList.add('on');
  btn.classList.add('on');
  if (name === 'strategy') loadStrategyForm();
  if (name === 'logs') loadLogs();
}

function sym(cfg) {
  var p = (cfg && cfg.pair) || 'BTCUSDC';
  if (p.endsWith('USDC'))  return 'R';
  if (p.endsWith('USDT') || p.endsWith('USDC') || p.endsWith('USD') || p.endsWith('BUSD')) return '$';
  if (p.endsWith('EUR'))  return '\\u20ac';
  if (p.endsWith('GBP'))  return '\\u00a3';
  return '';
}

function fmtN(v, s, dec) {
  if (v === null || v === undefined) return '\\u2014';
  s = s || '';
  dec = (dec === undefined) ? 0 : dec;
  if (Math.abs(v) >= 10000) return s + v.toLocaleString('en',{maximumFractionDigits:0});
  if (Math.abs(v) >= 1)     return s + v.toLocaleString('en',{minimumFractionDigits:dec,maximumFractionDigits:Math.max(dec,2)});
  return s + v.toFixed(6);
}

function pnlColor(v) { return v >= 0 ? 'green' : 'red'; }

async function toggleMode() {
  var isDry = (_cfg && _cfg.dry_run !== undefined) ? _cfg.dry_run : true;
  var goLive = isDry;
  if (goLive && !confirm('Switch to LIVE TRADING?\\nThis will place real orders on your account.\\nAre you sure?')) return;
  if (!goLive && !confirm('Switch to DRY RUN (paper trading)?')) return;
  try {
    await fetch('/api/config', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({dry_run: !isDry})});
    await refresh();
  } catch(e) { alert('Failed to toggle mode: ' + e); }
}

async function refresh() {
  try {
    var r = await fetch('/api/state');
    var d = await r.json();
    _cfg = d.cfg;
    // Top bar
    var sb = document.getElementById('b-status');
    sb.textContent = d.status === 'ok' ? 'ONLINE' : 'DEGRADED';
    sb.className   = 'badge ' + (d.status === 'ok' ? 'b-ok' : 'b-err');
    var mb = document.getElementById('b-mode');
    mb.style.display = '';
    mb.textContent   = d.dry_run ? 'DRY RUN' : 'LIVE';
    mb.className     = 'badge ' + (d.dry_run ? 'b-dry' : 'b-ok');
    document.getElementById('t-pair').textContent = d.cfg ? (d.cfg.pair || '') : '';
    var pd = d.price_data;
    var s  = sym(d.cfg);
    document.getElementById('t-price').textContent = pd ? s + (pd.price||0).toLocaleString('en',{maximumFractionDigits:0}) : (d.last_price ? s+d.last_price.toLocaleString('en',{maximumFractionDigits:0}) : '\\u2014');
    var chg = pd ? parseFloat(pd.chg || 0) : 0;
    var chgEl = document.getElementById('t-chg');
    chgEl.textContent = pd ? ((chg>=0?'+':'')+chg.toFixed(2)+'%') : '';
    chgEl.className   = 'chg ' + (chg >= 0 ? 'green' : 'red');
    document.getElementById('t-refresh').textContent = 'Updated '+new Date().toLocaleTimeString();
    if (document.getElementById('tab-dashboard').classList.contains('on')) {
      updateDashboard(d);
    }
  } catch(e) {
    document.getElementById('b-status').textContent = 'OFFLINE';
    document.getElementById('b-status').className   = 'badge b-err';
  }
}

function updateDashboard(d) {
  var state = d.state;
  var s     = sym(d.cfg);
  var price = d.price_data ? d.price_data.price : (d.last_price || 0);
  if (!state) {
    document.getElementById('gens').innerHTML = '<div class="no-data">Bot not yet started &#8212; waiting for first tick</div>';
    return;
  }
  var totalVal = 0, totalReal = 0, totalOpen = 0;
  var invested = state.total_invested || 0;
  (state.generations || []).forEach(function(gen){
    var btcV = (gen.slots||[]).filter(function(sl){return sl.state==='DIP'}).reduce(function(a,sl){return a+sl.dip_btc*price},0);
    totalVal  += (gen.free_cash||0) + btcV;
    totalReal += (gen.trade_log||[]).filter(function(t){return t.type==='SELL'}).reduce(function(a,t){return a+(t.profit||0)},0);
    totalOpen += (gen.slots||[]).filter(function(sl){return sl.state==='DIP'}).length;
  });
  var pnl = totalVal - invested;
  document.getElementById('c-port').textContent = s+totalVal.toLocaleString('en',{maximumFractionDigits:0});
  document.getElementById('c-inv').textContent  = s+invested.toLocaleString('en',{maximumFractionDigits:0});
  var pe = document.getElementById('c-pnl');
  pe.textContent = (pnl>=0?'+':'')+s+Math.abs(pnl).toLocaleString('en',{maximumFractionDigits:0});
  pe.className   = 'val '+pnlColor(pnl);
  document.getElementById('c-real').textContent = (totalReal>=0?'+':'')+s+Math.abs(totalReal).toLocaleString('en',{minimumFractionDigits:2,maximumFractionDigits:2});
  document.getElementById('c-open').textContent = totalOpen;

  var html = '';
  (state.generations||[]).forEach(function(gen){
    var btcV  = (gen.slots||[]).filter(function(sl){return sl.state==='DIP'}).reduce(function(a,sl){return a+sl.dip_btc*price},0);
    var total = (gen.free_cash||0)+btcV;
    var gReal = (gen.trade_log||[]).filter(function(t){return t.type==='SELL'}).reduce(function(a,t){return a+(t.profit||0)},0);
    var nOpen = (gen.slots||[]).filter(function(sl){return sl.state==='DIP'}).length;
    var nIdle = (gen.slots||[]).filter(function(sl){return sl.state==='IDLE'}).length;
    var pct   = gen.capital ? ((total-gen.capital)/gen.capital*100) : 0;
    var slotsHtml = '';
    (gen.slots||[]).forEach(function(sl){
      if (sl.state === 'DIP') {
        var tp  = sl.sell_order_price || (sl.dip_entry ? sl.dip_entry*(1+sl.sell_target) : 0);
        var unr = sl.dip_btc*price - sl.dip_cost;
        slotsHtml += '<div class="slot dip">' +
          '<div class="sh"><span class="sid">SLOT '+(sl.id+1)+'</span><span class="stbadge dip">DIP</span></div>' +
          '<div class="sentry">'+s+(sl.dip_entry||0).toLocaleString('en',{maximumFractionDigits:0})+'</div>' +
          '<div class="stgt">Target: '+s+tp.toLocaleString('en',{maximumFractionDigits:0})+'</div>' +
          '<div class="spnl '+pnlColor(unr)+'">'+(unr>=0?'+':'')+s+Math.abs(unr).toLocaleString('en',{minimumFractionDigits:2,maximumFractionDigits:2})+'</div>' +
          '</div>';
      } else {
        slotsHtml += '<div class="slot idle">' +
          '<div class="sh"><span class="sid">SLOT '+(sl.id+1)+'</span><span class="stbadge idle">IDLE</span></div>' +
          '<div class="swait">Waiting for signal</div>' +
          '</div>';
      }
    });
    var logs = (gen.trade_log||[]).slice(-8).reverse();
    var logHtml = '';
    if (logs.length) {
      logHtml = '<div class="tradelog"><h3>Recent Trades (Gen '+gen.id+')</h3>' +
        '<div class="tlog-row hdr"><span>Time</span><span>Type</span><span>Price</span><span>Amount</span><span>Profit</span></div>';
      logs.forEach(function(t){
        var ttime = t.time ? t.time.substring(0,16).replace('T',' ') : (t.candle||'');
        logHtml += '<div class="tlog-row">' +
          '<span style="color:#484f58">'+ttime+'</span>' +
          '<span class="tlog-'+t.type.toLowerCase()+'">'+t.type+'</span>' +
          '<span>'+s+(t.price||0).toLocaleString('en',{maximumFractionDigits:0})+'</span>' +
          '<span>'+(t.btc ? t.btc.toFixed(6)+' BTC' : (t.slot_size ? s+(t.slot_size||0).toLocaleString('en',{maximumFractionDigits:2}) : ''))+'</span>' +
          '<span class="'+(t.profit>=0?'green':'red')+'">'+(t.profit!==undefined?(t.profit>=0?'+':'')+s+Math.abs(t.profit||0).toFixed(2):'')+'</span>' +
          '</div>';
      });
      logHtml += '</div>';
    }
    html += '<div class="gen">' +
      '<div class="gen-hdr">' +
        '<span class="gtitle">Generation '+gen.id+'</span>' +
        '<span class="gstat">Capital: <b>'+s+(gen.capital||0).toLocaleString('en',{maximumFractionDigits:0})+'</b></span>' +
        '<span class="gstat">Free Cash: <b>'+s+(gen.free_cash||0).toLocaleString('en',{minimumFractionDigits:2,maximumFractionDigits:2})+'</b></span>' +
        '<span class="gstat">Total: <b>'+s+total.toLocaleString('en',{maximumFractionDigits:0})+'</b></span>' +
        '<span class="gstat">P&amp;L: <b class="'+pnlColor(pct)+'">'+(pct>=0?'+':'')+pct.toFixed(1)+'%</b></span>' +
        '<span class="gstat">Realised: <b class="green">'+(gReal>=0?'+':'')+s+Math.abs(gReal).toFixed(2)+'</b></span>' +
        '<span class="gstat">Open: <b>'+nOpen+'</b> / Idle: <b>'+nIdle+'</b></span>' +
      '</div>' +
      '<div class="slots">'+slotsHtml+'</div>' +
      logHtml +
      '</div>';
  });
  document.getElementById('gens').innerHTML = html;
}

async function loadStrategyForm() {
  try {
    var r = await fetch('/api/config');
    var c = await r.json();
    _cfg  = c;
    var f = document.getElementById('sf');
    Object.keys(c).forEach(function(k){
      var el = f.elements[k];
      if (!el) return;
      el.value = c[k];
    });
    await loadPairs();
  } catch(e) { console.error(e); }
}

async function loadPairs() {
  var ex   = document.getElementById('f-ex').value;
  var psel = document.getElementById('f-pair');
  psel.innerHTML = '<option>Loading&#8230;</option>';
  try {
    var r   = await fetch('/api/pairs?exchange='+encodeURIComponent(ex));
    var d   = await r.json();
    var cur = _cfg ? (_cfg.pair||'') : '';
    psel.innerHTML = (d.pairs||[]).map(function(p){
      return '<option value="'+p+'"'+(p===cur?' selected':'')+'>'+p+'</option>';
    }).join('');
    if (!psel.value && cur)
      psel.innerHTML = '<option value="'+cur+'" selected>'+cur+'</option>' + psel.innerHTML;
  } catch(e) {
    psel.innerHTML = '<option value="">Failed to load pairs</option>';
  }
}

async function saveStrategy(e) {
  e.preventDefault();
  var f = document.getElementById('sf'), data = {};
  Array.from(f.elements).forEach(function(el){ if (el.name) data[el.name] = el.value; });
  try {
    var r = await fetch('/api/config',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
    var d = await r.json();
    toast(d.ok ? '&#10003; Strategy saved!' : 'Error: '+d.error, d.ok ? '#238636' : '#6e2424');
  } catch(ex) { toast('Save failed: '+ex,'#6e2424'); }
}

function exportStrategy() { window.location.href='/api/export'; }

function importStrategy(inp) {
  var file = inp.files[0]; if (!file) return;
  var rd = new FileReader();
  rd.onload = async function(e) {
    try {
      var data = JSON.parse(e.target.result);
      var r    = await fetch('/api/config',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
      var d    = await r.json();
      if (d.ok) { toast('&#10003; Strategy imported!','#238636'); loadStrategyForm(); }
      else toast('Import error: '+d.error,'#6e2424');
    } catch(ex) { toast('Invalid JSON file','#6e2424'); }
  };
  rd.readAsText(file);
  inp.value = '';
}

async function loadLogs() {
  try {
    var r = await fetch('/api/logs');
    var d = await r.json();
    var pre = document.getElementById('log-pre');
    pre.textContent = (d.lines||[]).join('\\n') || 'No logs yet.';
    pre.scrollTop   = pre.scrollHeight;
  } catch(e) { document.getElementById('log-pre').textContent = 'Could not load logs.'; }
}

function toast(msg, color) {
  var t = document.getElementById('toast');
  t.innerHTML   = msg;
  t.style.background = color || '#238636';
  t.style.color = '#fff';
  t.style.display = 'block';
  clearTimeout(t._tid);
  t._tid = setTimeout(function(){ t.style.display='none'; }, 3500);
}

async function refreshBalances() {
  try {
    var r = await fetch('/api/balances');
    var d = await r.json();
    var balances = d.balances || {};
    var s = sym(d.cfg || _cfg);
    if (balances._error) {
      document.getElementById('c-qbal').textContent = 'Error';
      document.getElementById('c-bbal').textContent = 'Error';
      return;
    }
    Object.keys(balances).forEach(function(cur) {
      var b = balances[cur];
      if (b.is_quote) {
        document.getElementById('c-qlbl').textContent = cur + ' Balance';
        document.getElementById('c-qbal').textContent = s + (b.available||0).toLocaleString('en',{minimumFractionDigits:2,maximumFractionDigits:2});
        var rsv = b.reserved || 0;
        document.getElementById('c-qrsv').textContent = rsv > 0 ? 'Reserved: '+s+(rsv).toLocaleString('en',{minimumFractionDigits:2,maximumFractionDigits:2}) : '';
      }
      if (b.is_base) {
        document.getElementById('c-blbl').textContent = cur + ' Balance';
        document.getElementById('c-bbal').textContent = (b.available||0).toFixed(6) + ' ' + cur;
        var rsv = b.reserved || 0;
        document.getElementById('c-brsv').textContent = rsv > 0 ? 'Reserved: '+(rsv).toFixed(6) : '';
      }
    });
  } catch(e) { console.error('Balance refresh failed', e); }
}

refresh();
refreshBalances();
setInterval(refresh, 30000);
setInterval(refreshBalances, 60000);
</script>
</body>
</html>"""

class _DashboardHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        path = self.path.split("?")[0]
        if path == "/":
            body = _DASHBOARD_HTML.encode()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif path == "/api/status":
            self._json({
                "status":     "ok" if _bot_status["alive"] else "degraded",
                "last_price": _bot_status["last_price"],
                "last_tick":  _bot_status["last_tick"],
                "dry_run":    _bot_status["dry_run"],
                "time":       datetime.now(timezone.utc).isoformat(),
            })
        elif path == "/api/state":
            with _snap_lock:
                data = {
                    "status":     "ok" if _bot_status["alive"] else "degraded",
                    "last_price": _bot_status["last_price"],
                    "last_tick":  _bot_status["last_tick"],
                    "dry_run":    _bot_status["dry_run"],
                    "time":       datetime.now(timezone.utc).isoformat(),
                    "state":      _snap["state"],
                    "price_data": _snap["price_data"],
                    "cfg":        _snap["cfg"],
                }
            self._json(data)
        elif path == "/api/config":
            self._json(load_strategy_config())
        elif path == "/api/pairs":
            qs       = urllib.parse.parse_qs(self.path.split("?", 1)[-1])
            exchange = (qs.get("exchange", [None])[0]
                        or (_snap.get("cfg") or load_strategy_config()).get("exchange", "VALR"))
            pairs    = get_all_pairs(exchange, "", "")
            self._json({"pairs": pairs, "exchange": exchange})
        elif path == "/api/balances":
            api_cfg  = load_api_config()
            cfg_now  = (_snap.get("cfg") or load_strategy_config())
            exchange = cfg_now.get("exchange", "VALR")
            pair     = cfg_now.get("pair", "BTCUSDC")
            api_key  = api_cfg.get("api_key", "")
            api_sec  = api_cfg.get("api_secret", "")
            balances = get_account_balance(exchange, pair, api_key, api_sec)
            self._json({"balances": balances or {}, "exchange": exchange, "pair": pair})
        elif path == "/api/logs":
            self._json({"lines": list(_log_lines[-200:])})
        elif path == "/api/export":
            cfg  = load_strategy_config()
            body = json.dumps(cfg, indent=2).encode()
            name = f"btc_ladder_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Disposition", f'attachment; filename="{name}"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == "/api/config":
            try:
                length  = int(self.headers.get("Content-Length", 0))
                payload = json.loads(self.rfile.read(length).decode())
                cfg     = load_strategy_config()
                for k, v in payload.items():
                    if k not in DEFAULT_PARAMS:
                        continue
                    orig = DEFAULT_PARAMS[k]
                    if isinstance(orig, bool):
                        cfg[k] = str(v).lower() not in ("false", "0", "")
                    elif isinstance(orig, int):
                        cfg[k] = int(float(v))
                    elif isinstance(orig, float):
                        cfg[k] = float(v)
                    else:
                        cfg[k] = str(v)
                save_strategy_config(cfg)
                with _snap_lock:
                    _snap["cfg"] = cfg
                self._json({"ok": True})
            except Exception as e:
                self._json({"ok": False, "error": str(e)})
        else:
            self.send_response(404)
            self.end_headers()

    def _json(self, data):
        body = json.dumps(data, default=str).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args):
        pass

def start_health_server():
    port = int(os.environ.get("PORT", 8080))
    try:
        srv = HTTPServer(("0.0.0.0", port), _DashboardHandler)
        t   = threading.Thread(target=srv.serve_forever, daemon=True)
        t.start()
        logging.info(f"Dashboard listening on :{port}")
        _start_keep_alive(port)
    except Exception as e:
        logging.warning(f"Dashboard server could not start: {e}")

def _start_keep_alive(port):
    """Ping own /api/status every 9 min to prevent Render free-tier sleep."""
    def _ping():
        url = f"http://127.0.0.1:{port}/api/status"
        while True:
            time.sleep(540)  # 9 minutes
            try:
                urllib.request.urlopen(url, timeout=10)
            except Exception:
                pass
    threading.Thread(target=_ping, daemon=True, name="keep-alive").start()

# =============================================================
#  SECTION 6 — LIVE TRADER
# =============================================================
def make_slot(idx):
    return {"id": idx, "active": idx == 0, "state": "IDLE",
            "sell_target": 0.05, "dip_btc": 0.0, "dip_entry": None,
            "dip_cost": 0.0, "dip_open_time": None, "dip_orig_target": 0.05,
            "sell_order_id": None, "sell_order_price": None,
            "peak_price": None, "dip_buys": 0, "dip_sells": 0}

def make_generation(gen_id, capital, start_time):
    n = max(1, int(capital // 100))
    return {"id": gen_id, "capital": capital, "free_cash": capital,
            "slot_size": capital / n, "n_slots": n,
            "slots": [make_slot(i) for i in range(n)],
            "last_dip_buy_price": None, "last_dip_buy_time": None,
            "last_dip_activity": None, "start_time": start_time, "trade_log": []}

def bot_place_limit_sell(btc_qty, sell_price, slot_id, gen_id,
                          dry_run, exchange, symbol, api_key, api_secret):
    if dry_run:
        oid = f"DRY_{gen_id}_{slot_id}_{int(time.time())}"
        logging.info(f"[DRY] LIMIT SELL Gen{gen_id} Sl{slot_id+1} "
                     f"{btc_qty:.6f} @ {sell_price:,.2f} id={oid}")
        return oid
    if exchange == "VALR":
        ts   = int(time.time())
        body = {"side": "SELL", "quantity": f"{btc_qty:.8f}", "price": f"{sell_price:.8f}",
                "pair": symbol, "timeInForce": "GTC",
                "customerOrderId": f"LADDER-{gen_id}-{slot_id}-{ts}"[:50]}
        data, status = valr_request("/v1/orders/limit", method="POST", body=body,
                                    api_key=api_key, api_secret=api_secret)
        logging.info(f"VALR LIMIT SELL status={status} data={data}")
        if status in (200, 202):
            oid = str(data.get("id", data.get("orderId", "")))
            logging.info(f"LIMIT SELL PLACED Gen{gen_id} Sl{slot_id+1} "
                         f"{btc_qty:.8f} @ {sell_price:.8f} id={oid}")
            return oid
        if status == 400 and isinstance(data, dict) and data.get("code") == -12005:
            sell_base_cur = symbol
            for q in ["USDT", "USDC", "BUSD", "USDC", "EUR", "GBP", "USD"]:
                if symbol.endswith(q):
                    sell_base_cur = symbol[:-len(q)]
                    break
            data2, status2 = valr_request(
                f"/v1/simple/{symbol}/order",
                method="POST", body={
                    "payInCurrency": sell_base_cur,
                    "payAmount":     f"{btc_qty:.8f}",
                    "side":          "SELL",
                }, api_key=api_key, api_secret=api_secret)
            logging.info(f"VALR SIMPLE SELL status={status2} data={data2}")
            if status2 in (200, 202):
                oid = str(data2.get("id", data2.get("orderId", f"SIMPLE-{ts}")))
                return oid
    elif exchange == "BINANCE":
        params = {"symbol": symbol, "side": "SELL", "type": "LIMIT",
                  "timeInForce": "GTC", "quantity": f"{btc_qty:.6f}",
                  "price": f"{sell_price:.2f}"}
        data, status = binance_request("/api/v3/order", method="POST",
                                       params=params, api_key=api_key,
                                       api_secret=api_secret)
        if status == 200 and "orderId" in data:
            return str(data["orderId"])
    logging.error(f"Limit sell FAILED Gen{gen_id} Sl{slot_id+1}")
    return None

def bot_market_buy(quote_amount, slot_id, gen_id,
                   dry_run, exchange, symbol, api_key, api_secret):
    if dry_run:
        price = get_live_price(exchange, symbol, api_key, api_secret)
        p   = price["price"] if price else 1
        btc = (quote_amount * (1 - FEE_RATE)) / p
        logging.info(f"[DRY] BUY Gen{gen_id} Sl{slot_id+1} "
                     f"{quote_amount:.2f} -> {btc:.6f} BTC @ {p:,.2f}")
        return btc, p

    if exchange == "VALR":
        cur = get_live_price("VALR", symbol, api_key, api_secret)
        ask = float(cur.get("ask") or cur.get("price") or 0) if cur else 0
        if ask <= 0:
            return 0, "VALR error: could not fetch price"
        base_qty = round((quote_amount * (1 - FEE_RATE)) / ask, 8)
        ts = int(time.time())
        data, status = valr_request("/v1/orders/market", method="POST", body={
            "side":            "BUY",
            "quoteAmount":     f"{quote_amount:.8f}",
            "pair":            symbol,
            "customerOrderId": f"BUY-{gen_id}-{slot_id}-{ts}"[:50],
        }, api_key=api_key, api_secret=api_secret)
        logging.info(f"VALR MARKET BUY status={status} body={data}")

        if status == 400 and isinstance(data, dict) and data.get("code") == -12005:
            simple_body = {"side": "BUY"}
            for q in ["USDT", "USDC", "BUSD", "USDC", "EUR", "GBP", "USD"]:
                if symbol.endswith(q):
                    simple_body["payInCurrency"] = q
                    break
            simple_body["payAmount"] = f"{quote_amount:.8f}"
            data, status = valr_request(f"/v1/simple/{symbol}/order",
                                        method="POST", body=simple_body,
                                        api_key=api_key, api_secret=api_secret)
            logging.info(f"VALR SIMPLE BUY status={status} data={data}")
            if status not in (200, 202):
                for tif in ["IOC", "FOK", "GTC"]:
                    data, status = valr_request("/v1/orders/limit", method="POST", body={
                        "side": "BUY", "quantity": f"{base_qty:.8f}",
                        "price": f"{ask:.8f}", "pair": symbol,
                        "postOnly": False, "timeInForce": tif,
                        "customerOrderId": f"L{tif}-{gen_id}-{slot_id}-{ts}"[:50],
                    }, api_key=api_key, api_secret=api_secret)
                    logging.info(f"VALR LIMIT {tif} status={status} data={data}")
                    if status in (200, 202):
                        break
                    if isinstance(data, dict) and data.get("code") != -12005:
                        break

        if status not in (200, 202):
            detail = (data.get("message", "") or data.get("detail", "") or str(data))[:300]
            logging.error(f"VALR BUY FAILED status={status}: {detail}")
            return 0, f"VALR error {status}: {detail}"

        time.sleep(2)
        oid = data.get("id", "") or data.get("orderId", "")
        if oid:
            filled, _ = valr_request(f"/v1/orders/{symbol}/orderid/{oid}",
                                     api_key=api_key, api_secret=api_secret)
            qty   = float(filled.get("executedQuantity") or
                          filled.get("originalQuantity") or base_qty)
            total = float(filled.get("executedCost") or 0)
            avg   = (total / qty) if (qty > 0 and total > 0) else ask
            if qty > 0:
                return qty, avg
        return base_qty, ask

    elif exchange == "BINANCE":
        params = {"symbol": symbol, "side": "BUY", "type": "MARKET",
                  "quoteOrderQty": f"{quote_amount:.2f}"}
        data, status = binance_request("/api/v3/order", method="POST",
                                       params=params, api_key=api_key,
                                       api_secret=api_secret)
        if status == 200:
            qty  = float(data.get("executedQty", 0))
            cost = float(data.get("cummulativeQuoteQty", quote_amount))
            avg  = cost / qty if qty else 0
            return qty, avg
        return 0, f"Binance error {status}: {data.get('msg', str(data))}"

    return 0, "Unsupported exchange"

def bot_cancel_order(order_id, slot_id, gen_id,
                     dry_run, exchange, symbol, api_key, api_secret):
    if dry_run or not order_id or str(order_id).startswith("DRY_"):
        return True
    if exchange == "VALR":
        _, status = valr_request("/v1/orders/order", method="DELETE",
                                 body={"orderId": order_id, "pair": symbol},
                                 api_key=api_key, api_secret=api_secret)
        return status in (200, 202, 204)
    elif exchange == "BINANCE":
        _, status = binance_request("/api/v3/order", method="DELETE",
                                    params={"symbol": symbol, "orderId": order_id},
                                    api_key=api_key, api_secret=api_secret)
        return status == 200
    return False

def bot_order_status(order_id, dry_run, exchange, symbol, api_key, api_secret):
    if dry_run:
        return "OPEN"
    if exchange == "VALR":
        data, _ = valr_request(f"/v1/orders/{symbol}/orderid/{order_id}",
                               api_key=api_key, api_secret=api_secret)
        s = data.get("orderStatusType", "").upper()
        if s in ("FILLED", "EXECUTED"):              return "FILLED"
        if s in ("CANCELLED", "FAILED", "EXPIRED"):  return "CANCELLED"
        return "OPEN"
    elif exchange == "BINANCE":
        data, _ = binance_request("/api/v3/order",
                                  params={"symbol": symbol, "orderId": order_id},
                                  api_key=api_key, api_secret=api_secret)
        s = data.get("status", "").upper()
        if s == "FILLED":                                    return "FILLED"
        if s in ("CANCELED", "EXPIRED", "REJECTED"):         return "CANCELLED"
        return "OPEN"
    return "UNKNOWN"

def replace_sell_order(slot, gen_id, dry_run, exchange, symbol, api_key, api_secret):
    if slot.get("sell_order_id"):
        bot_cancel_order(slot["sell_order_id"], slot["id"], gen_id,
                         dry_run, exchange, symbol, api_key, api_secret)
        slot["sell_order_id"]    = None
        slot["sell_order_price"] = None
    if slot["dip_btc"] > 0 and slot["dip_entry"]:
        tp  = slot["dip_entry"] * (1 + slot["sell_target"])
        oid = bot_place_limit_sell(slot["dip_btc"], tp, slot["id"], gen_id,
                                   dry_run, exchange, symbol, api_key, api_secret)
        slot["sell_order_id"]    = oid
        slot["sell_order_price"] = tp

def record_sell(gen, slot, fill_price, now_str):
    proceeds = slot["dip_btc"] * fill_price * (1 - FEE_RATE)
    profit   = proceeds - slot["dip_cost"]
    gen["free_cash"] += proceeds
    gen["slot_size"]  = gen["free_cash"] / gen["n_slots"]
    gen["trade_log"].append({
        "time": now_str, "type": "SELL", "slot": slot["id"],
        "price": fill_price, "btc": slot["dip_btc"],
        "proceeds": proceeds, "profit": profit,
        "peak": slot.get("peak_price", fill_price),
        "slot_size": gen["slot_size"]
    })
    logging.info(f"Gen{gen['id']} Sl{slot['id']+1} SELL {slot['dip_btc']:.6f} "
                 f"@ {fill_price:,.2f} profit={profit:+.2f}")
    _tg_send(f"*SELL* Gen{gen['id']} Slot{slot['id']+1}\n"
             f"Price: {fill_price:,.2f}  Profit: {profit:+.2f}")
    remaining = [s for s in gen["slots"] if s["state"] == "DIP" and s["id"] != slot["id"]]
    slot.update(state="IDLE", dip_btc=0.0, dip_entry=None, dip_cost=0.0,
                dip_open_time=None, dip_orig_target=0.05, sell_target=0.05,
                sell_order_id=None, sell_order_price=None, peak_price=None)
    slot["dip_sells"] += 1
    if not remaining:
        gen["last_dip_buy_price"] = None
        gen["last_dip_buy_time"]  = None

def process_generation(gen, current_price, ref_price, now_str, cfg, api_key, api_secret):
    slots     = gen["slots"]
    free_cash = gen["free_cash"]
    slot_size = gen["slot_size"]
    gen_id    = gen["id"]
    exchange  = cfg.get("exchange", "VALR")
    symbol    = cfg.get("pair", "BTCUSDC")
    dry_run   = cfg.get("dry_run", True)

    BUY_DROP        = cfg.get("buy_drop_pct", 2.0) / 100
    SELL_BASE       = cfg.get("dip_sell_pct", 5.0) / 100
    SELL_DECAY_DAYS = cfg.get("sell_decay_days", 7)
    SELL_DECAY_RATE = cfg.get("sell_decay_rate", 1.0) / 100
    SELL_FLOOR      = cfg.get("sell_floor_pct", 5.0) / 100
    ESC_START       = cfg.get("escalation_start", 5)
    ESC_STEP        = cfg.get("escalation_step", 2.0) / 100
    ESC_MAX         = cfg.get("escalation_max_lvl", 3)
    ENTRY_DECAY_D   = cfg.get("entry_decay_days", 15)
    SPLIT_MIN       = cfg.get("split_min_size", 30.0)

    for slot in slots:
        if slot["state"] == "DIP" and slot["dip_entry"]:
            pk = slot.get("peak_price") or slot["dip_entry"]
            if current_price > pk:
                slot["peak_price"] = current_price

    n_idle = sum(1 for s in slots if s["state"] == "IDLE")
    if n_idle == 1:
        new_sz = free_cash / (len(slots) + 1)
        if new_sz >= SPLIT_MIN:
            gen["n_slots"] += 1
            ns = make_slot(len(slots))
            ns["active"] = True
            slots.append(ns)
            slot_size = free_cash / gen["n_slots"]
            gen["slot_size"] = slot_size
            for s in slots:
                if s["state"] == "DIP" and s["sell_order_id"]:
                    replace_sell_order(s, gen_id, dry_run, exchange, symbol, api_key, api_secret)
            logging.info(f"Gen{gen_id} SPLIT -> {gen['n_slots']} slots slot_size={slot_size:.2f}")

    for s in range(1, len(slots)):
        if not slots[s]["active"] and slots[s-1]["dip_buys"] > 0:
            slots[s]["active"] = True

    for slot in slots:
        if slot["state"] == "DIP" and slot.get("dip_open_time"):
            try:
                ot        = datetime.fromisoformat(slot["dip_open_time"])
                days_held = (datetime.now(timezone.utc) - ot).total_seconds() / 86400
            except Exception:
                days_held = 0
            if days_held > SELL_DECAY_DAYS:
                days_past  = days_held - SELL_DECAY_DAYS
                reduction  = int(days_past) * SELL_DECAY_RATE
                new_target = max(SELL_FLOOR, round(slot["dip_orig_target"] - reduction, 4))
                if new_target != slot["sell_target"]:
                    slot["sell_target"] = new_target
                    replace_sell_order(slot, gen_id, dry_run, exchange, symbol, api_key, api_secret)

    for slot in slots:
        if slot["state"] != "DIP":
            continue
        oid = slot.get("sell_order_id")
        if dry_run:
            tp = slot.get("sell_order_price") or (
                slot["dip_entry"] * (1 + slot["sell_target"]) if slot["dip_entry"] else None)
            if tp and current_price >= tp:
                record_sell(gen, slot, current_price, now_str)
                free_cash = gen["free_cash"]
                slot_size = gen["slot_size"]
        elif oid:
            status = bot_order_status(oid, dry_run, exchange, symbol, api_key, api_secret)
            if status == "FILLED":
                fill_price = slot.get("sell_order_price", current_price)
                record_sell(gen, slot, fill_price, now_str)
                free_cash = gen["free_cash"]
                slot_size = gen["slot_size"]
            elif status == "CANCELLED":
                replace_sell_order(slot, gen_id, dry_run, exchange, symbol, api_key, api_secret)

    n_dip = sum(1 for s in slots if s["state"] == "DIP")
    if n_dip < ESC_START:
        buy_signal = current_price <= ref_price * (1 - BUY_DROP)
    else:
        ref = gen.get("last_dip_buy_price")
        if ref:
            lvl       = min(ESC_MAX, n_dip - ESC_START + 1)
            drop_need = BUY_DROP + ESC_STEP * lvl
            if gen.get("last_dip_buy_time"):
                try:
                    lbt = datetime.fromisoformat(gen["last_dip_buy_time"])
                    ds  = (datetime.now(timezone.utc) - lbt).total_seconds() / 86400
                    if ds > ENTRY_DECAY_D:
                        drop_need = max(BUY_DROP, drop_need - int(ds - ENTRY_DECAY_D) * 0.01)
                except Exception:
                    pass
            buy_signal = current_price <= ref * (1 - drop_need)
        else:
            buy_signal = False

    if buy_signal and free_cash >= slot_size:
        for slot in slots:
            if slot["active"] and slot["state"] == "IDLE":
                btc, avg = bot_market_buy(slot_size, slot["id"], gen_id,
                                          dry_run, exchange, symbol, api_key, api_secret)
                if btc > 0:
                    orig = min(0.12, SELL_BASE + n_dip * 0.01)
                    tp   = avg * (1 + orig)
                    oid  = bot_place_limit_sell(btc, tp, slot["id"], gen_id,
                                                dry_run, exchange, symbol, api_key, api_secret)
                    slot.update(state="DIP", dip_btc=btc, dip_entry=avg,
                                dip_cost=slot_size, dip_open_time=now_str,
                                dip_orig_target=orig, sell_target=orig,
                                sell_order_id=oid, sell_order_price=tp, peak_price=avg)
                    free_cash -= slot_size
                    gen["last_dip_buy_price"] = avg
                    gen["last_dip_buy_time"]  = now_str
                    gen["last_dip_activity"]  = now_str
                    slot["dip_buys"] += 1
                    gen["trade_log"].append({
                        "time": now_str, "type": "BUY", "slot": slot["id"],
                        "price": avg, "btc": btc, "cost": slot_size,
                        "sell_target": orig, "target_price": tp, "order_id": oid
                    })
                    logging.info(f"Gen{gen_id} Sl{slot['id']+1} BUY {btc:.6f} "
                                 f"@ {avg:,.2f} sell_order @ {tp:,.2f}")
                    _tg_send(f"*BUY* Gen{gen_id} Slot{slot['id']+1}\n"
                             f"Price: {avg:,.2f}  Qty: {btc:.6f}  Target: {tp:,.2f}")
                break

    gen["free_cash"] = free_cash
    gen["slot_size"] = slot_size

def print_bot_dashboard(state, current_price, cfg):
    _, quote = parse_pair(cfg.get("pair", "BTCUSDC"))
    sym  = {"USD": "$", "USDT": "$", "USDC": "$", "EUR": "€", "GBP": "£"}.get(quote, quote + " ")
    mode = "[DRY RUN]" if cfg.get("dry_run", True) else "[LIVE]"
    now  = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    print(f"\n{'='*76}")
    print(f"  {cfg.get('pair','BTC')} Ladder  {mode}  |  {now}")
    print(f"  Current price: {sym}{current_price:,.2f}")
    print(f"{'='*76}")
    total_val = total_real = 0
    for gen in state["generations"]:
        slots   = gen["slots"]
        cash    = gen["free_cash"]
        btc_val = sum(s["dip_btc"] * current_price for s in slots if s["state"] == "DIP")
        total   = cash + btc_val
        capital = gen["capital"]
        real    = sum(t.get("profit", 0) for t in gen["trade_log"] if t.get("type") == "SELL")
        n_open  = sum(1 for s in slots if s["state"] == "DIP")
        n_idle  = sum(1 for s in slots if s["state"] == "IDLE")
        pnl_pct = (total - capital) / capital * 100 if capital else 0
        total_val  += total
        total_real += real
        print(f"\n  GEN {gen['id']}  capital={sym}{capital:,.0f}  "
              f"cash={sym}{cash:,.2f}  total={sym}{total:,.2f}  "
              f"P&L={pnl_pct:+.1f}%  open={n_open}  idle={n_idle}")
    invested  = state.get("total_invested", 0)
    grand     = total_val - invested
    print(f"\n{'─'*76}")
    print(f"  TOTAL  portfolio={sym}{total_val:,.2f}  invested={sym}{invested:,.2f}  "
          f"P&L={sym}{grand:+,.2f} ({grand/invested*100:+.1f}%)  realised={sym}{total_real:+,.2f}")
    print(f"{'='*76}")

def run_live_bot():
    # Set up root logger first so _LogCapture captures every log from startup onward
    root_log = logging.getLogger()
    root_log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s  %(levelname)s  %(message)s")
    sh  = logging.StreamHandler()
    sh.setFormatter(fmt)
    root_log.addHandler(sh)
    if LOG_FILE.parent.exists():
        try:
            fh = logging.FileHandler(LOG_FILE)
            fh.setFormatter(fmt)
            root_log.addHandler(fh)
        except Exception:
            pass
    # Register _LogCapture before any logs so the web dashboard sees everything
    cap = _LogCapture()
    cap.setFormatter(fmt)
    root_log.addHandler(cap)

    cfg      = load_strategy_config()
    api_cfg  = load_api_config()
    api_key  = api_cfg.get("api_key", "")
    api_sec  = api_cfg.get("api_secret", "")
    exchange = cfg.get("exchange", "VALR")
    pair     = cfg.get("pair", "BTCUSDC")
    interval = cfg.get("check_interval_min", 5) * 60
    dry_run  = cfg.get("dry_run", True)

    _bot_status["dry_run"] = dry_run
    with _snap_lock:
        _snap["cfg"] = cfg
    start_health_server()

    logging.info(f"Bot starting  exchange={exchange}  pair={pair}  dry_run={dry_run}")
    if not api_key:
        logging.warning("No API key found — set VALR_API_KEY / BINANCE_API_KEY env var")

    # Graceful shutdown
    _stop = threading.Event()
    def _handle_sigterm(sig, frame):
        logging.info("SIGTERM received — shutting down after current tick")
        _stop.set()
    signal.signal(signal.SIGTERM, _handle_sigterm)
    signal.signal(signal.SIGINT,  _handle_sigterm)

    state = None
    if STATE_FILE.exists():
        try:
            state = json.loads(STATE_FILE.read_text())
            logging.info(f"State loaded — {len(state['generations'])} generation(s)")
        except Exception:
            pass
    if state is None:
        now = datetime.now(timezone.utc).isoformat()
        state = {
            "generations":    [make_generation(1, cfg["gen_capital"], now),
                               make_generation(2, cfg["gen_capital"], now)],
            "total_invested": cfg["gen_capital"] * 2,
            "created_at":     now,
        }
        STATE_FILE.write_text(json.dumps(state, indent=2))
        logging.info("Fresh state created")

    price_history = []
    lkb = cfg.get("lookback_hours", 12)
    consecutive_failures = 0

    while not _stop.is_set():
        try:
            # Reload config each tick so dashboard changes (dry_run, pair, interval…) take effect immediately
            new_cfg = load_strategy_config()
            if new_cfg.get("dry_run") != cfg.get("dry_run"):
                logging.info(f"Mode changed: dry_run {cfg['dry_run']} -> {new_cfg['dry_run']}  "
                             f"({'DRY RUN' if new_cfg['dry_run'] else 'LIVE TRADING'})")
            cfg      = new_cfg
            exchange = cfg.get("exchange", "VALR")
            pair     = cfg.get("pair", "BTCUSDC")
            interval = cfg.get("check_interval_min", 5) * 60
            lkb      = cfg.get("lookback_hours", 12)
            _bot_status["dry_run"] = cfg.get("dry_run", True)
            with _snap_lock:
                _snap["cfg"] = cfg

            now_str    = datetime.now(timezone.utc).isoformat()
            price_data = get_live_price_with_retry(exchange, pair, api_key, api_sec)
            if not price_data:
                consecutive_failures += 1
                wait = min(300, 60 * consecutive_failures)
                logging.warning(f"Price fetch failed ({consecutive_failures}x) — waiting {wait}s")
                _stop.wait(wait)
                continue

            consecutive_failures = 0
            current_price = price_data["price"]
            _bot_status["last_price"] = current_price
            _bot_status["last_tick"]  = now_str
            _bot_status["alive"]      = True

            price_history.append(current_price)
            if len(price_history) > lkb * 2:
                price_history = price_history[-lkb*2:]
            ref_price = price_history[-lkb-1] if len(price_history) > lkb else current_price

            for gen in state["generations"]:
                process_generation(gen, current_price, ref_price, now_str,
                                   cfg, api_key, api_sec)

            if len(state["generations"]) < cfg.get("max_generations", 3):
                all_stuck = all(
                    all(s["state"] == "DIP" for s in g["slots"] if s["active"])
                    and any(s["dip_entry"] for s in g["slots"])
                    and current_price < max((s["dip_entry"] or 0) for s in g["slots"])
                      * (1 - cfg.get("new_gen_threshold", 20) / 100)
                    for g in state["generations"]
                )
                if all_stuck:
                    now_s = datetime.now(timezone.utc).isoformat()
                    ng    = make_generation(len(state["generations"]) + 1,
                                           cfg["gen_capital"], now_s)
                    state["generations"].append(ng)
                    logging.info(f"GEN {ng['id']} SPAWNED @ {current_price:,.0f}")
                    _tg_send(f"*GEN {ng['id']} SPAWNED* @ {current_price:,.2f}")

            # Always derive total_invested from actual generation capitals (never a running counter)
            state["total_invested"] = sum(g["capital"] for g in state["generations"])

            print_bot_dashboard(state, current_price, cfg)
            STATE_FILE.write_text(json.dumps(state, indent=2, default=str))
            with _snap_lock:
                _snap["state"]      = json.loads(json.dumps(state, default=str))
                _snap["price_data"] = price_data
                _snap["cfg"]        = cfg

        except Exception as e:
            logging.error(f"Loop error: {e}", exc_info=True)

        _stop.wait(interval)

    logging.info("Bot stopped — saving state")
    STATE_FILE.write_text(json.dumps(state, indent=2, default=str))

# =============================================================
#  SECTION 7 — BACKTEST ENGINE
# =============================================================
def run_backtest():
    CSV_HOURLY = APP_DIR / "btc_hourly.csv"
    CSV_DAILY  = APP_DIR / "btc_daily.csv"
    print("Loading BTC data for backtest...")
    df = None
    for csv in (CSV_HOURLY, CSV_DAILY):
        if csv.exists():
            try:
                if PANDAS_OK:
                    d = pd.read_csv(csv, skiprows=1)
                    d.columns = [c.lower().strip() for c in d.columns]
                    cc = next((c for c in d.columns if c in ("close", "price")), None)
                    dc = next((c for c in d.columns if any(k in c for k in ("date","time","stamp"))), None)
                    if cc and dc:
                        d = d[[dc, cc]].copy()
                        d.columns = ["timestamp", "close"]
                        d["timestamp"] = pd.to_datetime(d["timestamp"], format="mixed")
                        d["close"]     = pd.to_numeric(d["close"], errors="coerce")
                        d = d.dropna().sort_values("timestamp").reset_index(drop=True)
                        df = d
                        print(f"  Loaded {len(df):,} candles from {csv.name}")
                        break
            except Exception as e:
                print(f"  Failed ({csv.name}): {e}")

    if df is None:
        print("  No local CSV — trying CryptoCompare (free, no key needed)...")
        try:
            end_ts   = int(time.time())
            start_ts = end_ts - (5 * 365 * 24 * 3600)
            all_rows, to_ts = [], end_ts
            while to_ts > start_ts:
                url = (f"https://min-api.cryptocompare.com/data/v2/histohour"
                       f"?fsym=BTC&tsym=USD&limit=2000&toTs={to_ts}&aggregate=1")
                req = urllib.request.Request(url, timeout=20)
                with urllib.request.urlopen(req) as r:
                    data = json.loads(r.read().decode())
                candles = [c for c in data["Data"]["Data"] if c["time"] >= start_ts]
                if not candles:
                    break
                all_rows = candles + all_rows
                to_ts    = data["Data"]["Data"][0]["time"] - 1
                time.sleep(0.3)
            if PANDAS_OK and all_rows:
                df = pd.DataFrame(all_rows)
                df["timestamp"] = pd.to_datetime(df["time"], unit="s")
                df["close"]     = df["close"].astype(float)
                df = (df[["timestamp","close"]]
                      .drop_duplicates("timestamp")
                      .sort_values("timestamp")
                      .reset_index(drop=True))
                print(f"  Loaded {len(df):,} candles from CryptoCompare")
        except Exception as e:
            print(f"  CryptoCompare failed: {e}")

    if df is None:
        print("  ERROR: No data. Place btc_hourly.csv or btc_daily.csv next to bot.py")
        return

    prices = df["close"].values
    cfg    = load_strategy_config()
    LKBK      = cfg.get("lookback_hours", 12)
    BUY_DROP  = cfg.get("buy_drop_pct", 2.0) / 100
    SELL_BASE = cfg.get("dip_sell_pct", 5.0) / 100
    SELL_FLOOR= cfg.get("sell_floor_pct", 5.0) / 100
    SELL_DDAY = cfg.get("sell_decay_days", 7)
    SELL_DRAT = cfg.get("sell_decay_rate", 1.0) / 100
    ESC_START = cfg.get("escalation_start", 5)
    ESC_STEP  = cfg.get("escalation_step", 2.0) / 100
    ESC_MAX   = cfg.get("escalation_max_lvl", 3)
    GEN_CAP   = cfg.get("gen_capital", 1100.0)
    SPLIT_MIN = cfg.get("split_min_size", 30.0)
    MAX_GENS  = cfg.get("max_generations", 3)
    NEW_GEN_T = cfg.get("new_gen_threshold", 20.0) / 100
    CDPD      = 24

    def bs_make_gen(gid, cap, ci):
        n = max(1, int(cap // 100))
        return {"id": gid, "capital": cap, "free_cash": cap, "slot_size": cap/n, "n_slots": n,
                "slots": [{"active": i==0, "state": "IDLE", "sell_target": SELL_BASE,
                            "dip_btc": 0.0, "dip_entry": None, "dip_cost": 0.0,
                            "dip_open_candle": None, "dip_orig_target": SELL_BASE,
                            "peak_price": None, "buys": 0, "sells": 0}
                           for i in range(n)],
                "last_buy_price": None, "last_buy_candle": None,
                "start_candle": ci, "trade_log": []}

    gens = [bs_make_gen(1, GEN_CAP, LKBK), bs_make_gen(2, GEN_CAP, LKBK)]
    total_invested  = GEN_CAP * 2
    last_gen_candle = LKBK
    print(f"Running backtest on {len(prices):,} candles...")

    for i in range(LKBK, len(prices)):
        p_ref = prices[i - LKBK]
        p_t   = prices[i]
        for gen in gens:
            slots     = gen["slots"]
            free_cash = gen["free_cash"]
            slot_size = gen["slot_size"]
            for s in slots:
                if s["state"] == "DIP" and s["dip_entry"]:
                    pk = s.get("peak_price") or s["dip_entry"]
                    if p_t > pk: s["peak_price"] = p_t
            n_idle = sum(1 for s in slots if s["state"] == "IDLE")
            if n_idle == 1:
                new_sz = free_cash / (len(slots) + 1)
                if new_sz >= SPLIT_MIN:
                    gen["n_slots"] += 1
                    slots.append({"active": True, "state": "IDLE", "sell_target": SELL_BASE,
                                  "dip_btc": 0.0, "dip_entry": None, "dip_cost": 0.0,
                                  "dip_open_candle": None, "dip_orig_target": SELL_BASE,
                                  "peak_price": None, "buys": 0, "sells": 0})
                    slot_size = free_cash / gen["n_slots"]
                    gen["slot_size"] = slot_size
            for s in range(1, len(slots)):
                if not slots[s]["active"] and slots[s-1].get("buys", 0) > 0:
                    slots[s]["active"] = True
            for s in slots:
                if s["state"] == "DIP" and s.get("dip_open_candle") is not None:
                    dh = (i - s["dip_open_candle"]) / CDPD
                    if dh > SELL_DDAY:
                        r  = int(dh - SELL_DDAY) * SELL_DRAT
                        nt = max(SELL_FLOOR, round(s["dip_orig_target"] - r, 4))
                        s["sell_target"] = nt
            for s in slots:
                if s["state"] == "DIP" and s["dip_entry"]:
                    if p_t >= s["dip_entry"] * (1 + s["sell_target"]):
                        proc   = s["dip_btc"] * p_t
                        profit = proc - s["dip_cost"]
                        free_cash += proc
                        slot_size  = free_cash / gen["n_slots"]
                        gen["trade_log"].append({"type": "SELL", "candle": i, "price": p_t,
                                                 "profit": profit, "peak": s.get("peak_price", p_t),
                                                 "slot_size": slot_size})
                        s.update(state="IDLE", dip_btc=0.0, dip_entry=None, dip_cost=0.0,
                                 dip_open_candle=None, dip_orig_target=SELL_BASE,
                                 sell_target=SELL_BASE, peak_price=None)
                        s["sells"] = s.get("sells", 0) + 1
            n_dip = sum(1 for s in slots if s["state"] == "DIP")
            if n_dip < ESC_START:
                buy_ok = p_t <= p_ref * (1 - BUY_DROP)
            else:
                ref = gen.get("last_buy_price")
                if ref:
                    lvl = min(ESC_MAX, n_dip - ESC_START + 1)
                    dn  = BUY_DROP + ESC_STEP * lvl
                    if gen.get("last_buy_candle") is not None:
                        ds = (i - gen["last_buy_candle"]) / CDPD
                        if ds > 15:
                            dn = max(BUY_DROP, dn - int(ds - 15) * 0.01)
                    buy_ok = p_t <= ref * (1 - dn)
                else:
                    buy_ok = False
            if buy_ok and free_cash >= slot_size:
                for s in slots:
                    if s["active"] and s["state"] == "IDLE":
                        orig = min(0.12, SELL_BASE + n_dip * 0.01)
                        s.update(state="DIP", dip_btc=slot_size/p_t, dip_entry=p_t,
                                 dip_cost=slot_size, dip_open_candle=i,
                                 dip_orig_target=orig, sell_target=orig, peak_price=p_t)
                        s["buys"] = s.get("buys", 0) + 1
                        free_cash -= slot_size
                        gen["last_buy_price"]  = p_t
                        gen["last_buy_candle"] = i
                        gen["trade_log"].append({"type": "BUY", "candle": i,
                                                 "price": p_t, "slot_size": slot_size})
                        break
            gen["free_cash"] = free_cash
            gen["slot_size"] = slot_size

        if len(gens) < MAX_GENS and (i - last_gen_candle) >= CDPD * 30:
            all_stuck = all(
                all(s["state"] == "DIP" for s in g["slots"] if s["active"])
                and any(s["dip_entry"] for s in g["slots"])
                and p_t < max((s["dip_entry"] or 0) for s in g["slots"]) * (1 - NEW_GEN_T)
                for g in gens
            )
            if all_stuck:
                gens.append(bs_make_gen(len(gens) + 1, GEN_CAP, i))
                total_invested += GEN_CAP
                last_gen_candle = i
                print(f"  Gen {len(gens)} spawned at candle {i} price={p_t:,.0f}")

    final = prices[-1]; start = prices[LKBK]
    bh    = (total_invested / start) * final
    print(f"\n{'='*65}\n  BACKTEST RESULTS\n{'='*65}")
    print(f"  Period  : {df['timestamp'].iloc[LKBK].date()} -> {df['timestamp'].iloc[-1].date()}")
    print(f"  Candles : {len(prices)-LKBK:,}  Invested: {total_invested:,.2f}  Buy&Hold: {bh:,.2f} ({(bh/total_invested-1)*100:+.1f}%)")
    grand = grand_real = 0
    for gen in gens:
        slots   = gen["slots"]
        btc_val = sum(s["dip_btc"] * final for s in slots if s["state"] == "DIP")
        total   = gen["free_cash"] + btc_val
        real    = sum(t.get("profit", 0) for t in gen["trade_log"] if t["type"] == "SELL")
        n_sells = sum(1 for t in gen["trade_log"] if t["type"] == "SELL")
        pnl_pct = (total - gen["capital"]) / gen["capital"] * 100
        grand += total; grand_real += real
        print(f"  Gen {gen['id']}  capital={gen['capital']:,.0f}  total={total:,.2f} ({pnl_pct:+.1f}%)  "
              f"realised={real:,.2f}  trades={n_sells}")
    grand_pnl = grand - total_invested
    print(f"\n  GRAND TOTAL : {grand:,.2f}  P&L={grand_pnl:+,.2f} ({grand_pnl/total_invested*100:+.1f}%)")
    print(f"{'='*65}")

# =============================================================
#  SECTION 8 — USDC PREMIUM & VALR TEST
# =============================================================
def run_premium_check():
    print("\n  Fetching prices...")
    try:
        req = urllib.request.Request(
            "https://api.coingecko.com/api/v3/simple/price?ids=bitcoin&vs_currencies=usd,usdc",
            headers={"User-Agent": "Mozilla/5.0", "Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=10) as r:
            cg = json.loads(r.read().decode())
        btc_usd = cg["bitcoin"]["usd"]
        btc_USDC_cg = cg["bitcoin"]["usdc"]
    except Exception as e:
        print(f"  CoinGecko failed: {e}"); return
    api = load_api_config()
    vp  = get_live_price("VALR", "BTCUSDC", api.get("api_key",""), api.get("api_secret",""))
    usd_USDC = btc_USDC_cg / btc_usd
    print(f"\n{'='*55}\n  USDC PREMIUM MONITOR\n{'='*55}")
    print(f"  BTC/USD  (global)   : ${btc_usd:>12,.2f}")
    if vp:
        prem = (vp["price"] - btc_USDC_cg) / btc_USDC_cg * 100
        print(f"  BTC/USDC  (VALR)     : R{vp['price']:>12,.2f}")
        print(f"  SA Premium          : {prem:>+11.2f}%")
       
    print(f"{'='*55}\n")

def run_valr_test():
    cfg = load_api_config()
    key, sec = cfg.get("api_key",""), cfg.get("api_secret","")
    if not key:
        print("  No API key — set VALR_API_KEY env var or configure via GUI"); return
    print("\n  Testing VALR connection (read-only)...")
    t, _ = valr_request("/v1/public/time", public=True)
    print(f"  Server time : {t.get('time','error')}")
    tk, _ = valr_request("/v1/public/BTCUSDC/marketsummary", public=True)
    if "lastTradedPrice" in tk:
        print(f"  BTC/USDC : R{float(tk['lastTradedPrice']):,.2f}  "
              f"bid=R{float(tk.get('bidPrice',0)):,.0f}  ask=R{float(tk.get('askPrice',0)):,.0f}")
    bals, _ = valr_request("/v1/account/balances", api_key=key, api_secret=sec)
    if isinstance(bals, list):
        print("  Balances:")
        for b in bals:
            tot = float(b.get("total", 0))
            if tot > 0 or b.get("currency") in ("BTC","USDC"):
                print(f"    {b.get('currency','?'):6s}  total={tot:.6f}  "
                      f"avail={float(b.get('available',0)):.6f}")
    else:
        print(f"  Balance error: {bals}")
    orders, _ = valr_request("/v1/orders/open", api_key=key, api_secret=sec)
    if isinstance(orders, list):
        print(f"  Open orders : {len(orders)}")
    print("  Test complete — no orders placed.")

# =============================================================
#  SECTION 9 — GUI (local use only)
# =============================================================
def run_gui():
    if not GUI_OK:
        print("tkinter not available — run with 'python bot.py live' for headless mode")
        return

    class TradingApp(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("BTC Ladder Trader")
            self.geometry("1280x900")
            self.minsize(1100, 720)
            self.configure(bg="#F5F5F5")
            self.params      = load_strategy_config()
            self.api_cfg     = load_api_config()
            self.bot_running = False
            self.bot_thread  = None
            self.price_data  = {}
            self.balances    = {}
            self.bot_state   = {}
            self.param_vars  = {}
            self._setup_styles()
            self._build_ui()
            self._start_refresh()

        @staticmethod
        def _parse_pair(pair):
            for q in ["USDT","USDC","BUSD","EUR","GBP","USD"]:
                if pair.endswith(q): return pair[:-len(q)], q
            return pair[:3], pair[3:]

        @staticmethod
        def _fmt(val, quote):
            sym = {"USD":"$","USDT":"$","USDC":"$","EUR":"€","GBP":"£"}.get(quote, quote+" ")
            if val >= 10000: return f"{sym}{val:,.0f}"
            if val >= 1:     return f"{sym}{val:,.2f}"
            return f"{sym}{val:.6f}"

        def _setup_styles(self):
            s = ttk.Style(self); s.theme_use("clam")
            s.configure("TNotebook",     background="#F5F5F5", borderwidth=0)
            s.configure("TNotebook.Tab", padding=[14,7], font=("Segoe UI",10))
            s.configure("TFrame",        background="#F5F5F5")
            s.configure("TLabel",        background="#F5F5F5", font=("Segoe UI",10))
            s.configure("TButton",       font=("Segoe UI",10), padding=[10,5])
            s.configure("TEntry",        font=("Segoe UI",10), padding=4)

        def _build_ui(self):
            top = tk.Frame(self, bg="#1D9E75", height=52)
            top.pack(fill="x"); top.pack_propagate(False)
            tk.Label(top, text="  BTC Ladder Trader", bg="#1D9E75", fg="white",
                     font=("Segoe UI",13,"bold")).pack(side="left", padx=8, pady=10)
            self.status_lbl = tk.Label(top, text="  STOPPED  ", bg="#CC3333", fg="white",
                                       font=("Segoe UI",9,"bold"), padx=8, pady=2)
            self.status_lbl.pack(side="right", padx=12, pady=10)
            self.mode_lbl = tk.Label(top, text="DRY RUN", bg="#BA7517", fg="white",
                                     font=("Segoe UI",9,"bold"), padx=8, pady=2)
            self.mode_lbl.pack(side="right", padx=4, pady=10)
            nb = ttk.Notebook(self)
            nb.pack(fill="both", expand=True, padx=8, pady=8)
            self._build_dashboard(nb)
            self._build_positions(nb)
            self._build_import_tab(nb)
            self._build_settings(nb)
            self._build_api_tab(nb)
            self._build_log(nb)

        def _build_dashboard(self, nb):
            tab = ttk.Frame(nb); nb.add(tab, text="  Dashboard  ")
            ctrl = tk.Frame(tab, bg="#F5F5F5"); ctrl.pack(fill="x", padx=12, pady=10)
            self.start_btn = tk.Button(ctrl, text="▶  Start Bot",
                command=self._toggle_bot, bg="#1D9E75", fg="white",
                font=("Segoe UI",11,"bold"), padx=14, pady=7, relief="flat", cursor="hand2")
            self.start_btn.pack(side="left", padx=(0,8))
            tk.Button(ctrl, text="↻  Refresh", command=self._manual_refresh,
                bg="#378ADD", fg="white", font=("Segoe UI",10),
                padx=10, pady=7, relief="flat", cursor="hand2").pack(side="left", padx=(0,8))
            self.dry_btn = tk.Button(ctrl, text="MODE: DRY RUN",
                command=self._toggle_dry_run, bg="#BA7517", fg="white",
                font=("Segoe UI",10,"bold"), padx=10, pady=7, relief="flat", cursor="hand2")
            self.dry_btn.pack(side="left")
            self.upd_lbl = tk.Label(ctrl, text="Last update: never",
                bg="#F5F5F5", font=("Segoe UI",9), fg="#999999")
            self.upd_lbl.pack(side="right")

            cf = tk.Frame(tab, bg="#F5F5F5"); cf.pack(fill="x", padx=12, pady=4)
            self.metric_cards = {}
            metrics = [("price","Price","Loading..."), ("change","24h Change","—"),
                       ("portfolio","Portfolio","—"), ("realised","Realised P&L","—"),
                       ("openpos","Open Pos","—"), ("gens","Generations","—")]
            for i, (key, lbl, dflt) in enumerate(metrics):
                card = tk.Frame(cf, bg="white", highlightbackground="#E0E0E0", highlightthickness=1)
                card.grid(row=0, column=i, padx=4, pady=4, sticky="nsew")
                cf.columnconfigure(i, weight=1)
                hdr = tk.Label(card, text=lbl, bg="white", font=("Segoe UI",9), fg="#666666")
                hdr.pack(pady=(8,2))
                vl  = tk.Label(card, text=dflt, bg="white", font=("Segoe UI",13,"bold"), fg="#222222")
                vl.pack(pady=(0,4))
                sl  = tk.Label(card, text="", bg="white", font=("Segoe UI",8), fg="#999999")
                sl.pack(pady=(0,6))
                self.metric_cards[key] = vl
                if key == "price":
                    self.price_hdr_lbl = hdr
                    self.spread_lbl    = sl

            bf = tk.LabelFrame(tab, text="  Account Balances  ", bg="#F5F5F5",
                               font=("Segoe UI",10,"bold"), fg="#444444",
                               relief="flat", highlightbackground="#E0E0E0", highlightthickness=1)
            bf.pack(fill="x", padx=12, pady=4)
            tk.Button(bf, text="↻", command=lambda: threading.Thread(
                target=self._fetch_balances, daemon=True).start(),
                font=("Segoe UI",10), relief="flat", bg="#378ADD", fg="white",
                cursor="hand2", padx=6).pack(side="right", padx=8, pady=4)
            self.bal_q = tk.Label(bf, text="Quote: —", bg="#F5F5F5", font=("Segoe UI",10), fg="#1D9E75")
            self.bal_q.pack(anchor="w", padx=12, pady=(6,2))
            self.bal_b = tk.Label(bf, text="Base: —",  bg="#F5F5F5", font=("Segoe UI",10), fg="#BA7517")
            self.bal_b.pack(anchor="w", padx=12, pady=2)
            self.bal_t = tk.Label(bf, text="Total: —", bg="#F5F5F5", font=("Segoe UI",10,"bold"), fg="#444444")
            self.bal_t.pack(anchor="w", padx=12, pady=(2,6))

            gf = tk.LabelFrame(tab, text="  Generation Summary  ", bg="#F5F5F5",
                               font=("Segoe UI",10,"bold"), fg="#444444",
                               relief="flat", highlightbackground="#E0E0E0", highlightthickness=1)
            gf.pack(fill="both", expand=True, padx=12, pady=(4,12))
            cols = ("Gen","Started","Capital","Cash","Open/Total","Idle","Slot Size",
                    "Realised","Unrealised","Total","P&L%","Trades")
            self.gen_tree = ttk.Treeview(gf, columns=cols, show="headings", height=5)
            for c, w in zip(cols, [60,90,90,90,80,50,80,90,90,90,70,60]):
                self.gen_tree.heading(c, text=c)
                self.gen_tree.column(c, width=w, anchor="center")
            self.gen_tree.tag_configure("profit", background="#F0FFF0")
            self.gen_tree.tag_configure("loss",   background="#FFF0F0")
            vsb = ttk.Scrollbar(gf, orient="vertical", command=self.gen_tree.yview)
            self.gen_tree.configure(yscrollcommand=vsb.set)
            vsb.pack(side="right", fill="y")
            self.gen_tree.pack(fill="both", expand=True, padx=4, pady=4)

        def _build_positions(self, nb):
            tab = ttk.Frame(nb); nb.add(tab, text="  Positions  ")
            fb = tk.Frame(tab, bg="#F5F5F5"); fb.pack(fill="x", padx=12, pady=8)
            tk.Label(fb, text="Show:", bg="#F5F5F5", font=("Segoe UI",10)).pack(side="left")
            self.pos_filter = ttk.Combobox(fb, values=["All","Open","Idle"], width=14, state="readonly")
            self.pos_filter.set("All"); self.pos_filter.pack(side="left", padx=8)
            self.pos_filter.bind("<<ComboboxSelected>>", lambda e: self._refresh_positions())
            cols = ("Gen","Sl","State","Entry","Target","Target%","Peak","Peak%",
                    "Current","Curr%","Hold val","Unreal","Days","Order")
            self.pos_tree = ttk.Treeview(tab, columns=cols, show="headings", height=20)
            for c, w in zip(cols, [40,35,60,100,100,70,100,70,100,70,90,90,60,90]):
                self.pos_tree.heading(c, text=c); self.pos_tree.column(c, width=w, anchor="center")
            self.pos_tree.tag_configure("dip",   background="#FFF8F0")
            self.pos_tree.tag_configure("idle",  background="#F8FFF8")
            self.pos_tree.tag_configure("close", background="#E8FFE8")
            vsb = ttk.Scrollbar(tab, orient="vertical", command=self.pos_tree.yview)
            self.pos_tree.configure(yscrollcommand=vsb.set)
            vsb.pack(side="right", fill="y")
            self.pos_tree.pack(fill="both", expand=True, padx=12, pady=(0,12))

        def _build_import_tab(self, nb):
            tab = ttk.Frame(nb); nb.add(tab, text="  Import / Enter Trades  ")

            # ── button bar ─────────────────────────────────────────────
            btn_bar = tk.Frame(tab, bg="#F5F5F5")
            btn_bar.pack(fill="x", padx=12, pady=8)
            bkw = dict(font=("Segoe UI",10), padx=10, pady=6, relief="flat", cursor="hand2")
            tk.Button(btn_bar, text="Import CSV / Excel", command=self._import_trades_file,
                bg="#378ADD", fg="white", **bkw).pack(side="left", padx=(0,6))
            tk.Button(btn_bar, text="Delete Selected", command=self._delete_import_row,
                bg="#CC3333", fg="white", **bkw).pack(side="left", padx=(0,6))
            tk.Button(btn_bar, text="Clear All", command=self._clear_import_rows,
                bg="#888888", fg="white", **bkw).pack(side="left", padx=(0,6))
            tk.Button(btn_bar, text="Save to Bot State", command=self._apply_imports_to_state,
                bg="#1D9E75", fg="white", font=("Segoe UI",10,"bold"),
                padx=10, pady=6, relief="flat", cursor="hand2").pack(side="right")

            tk.Label(tab,
                text="CSV/Excel columns (header row required):  gen_id  |  entry_price  |  btc_qty  |  "
                     "cost (optional)  |  sell_target_pct (optional)  |  open_time (optional)",
                bg="#F5F5F5", font=("Segoe UI",8), fg="#888888",
                justify="left").pack(anchor="w", padx=14, pady=(0,4))

            # ── trade table ────────────────────────────────────────────
            tf = tk.Frame(tab, bg="#F5F5F5")
            tf.pack(fill="both", expand=True, padx=12, pady=(0,8))
            imp_cols = ("Gen", "Entry Price", "BTC Qty", "Cost", "Sell Target %", "Open Time")
            self.imp_tree = ttk.Treeview(tf, columns=imp_cols, show="headings",
                                         height=12, selectmode="extended")
            for c, w in zip(imp_cols, [50, 120, 110, 110, 110, 180]):
                self.imp_tree.heading(c, text=c)
                self.imp_tree.column(c, width=w, anchor="center")
            vsb_imp = ttk.Scrollbar(tf, orient="vertical", command=self.imp_tree.yview)
            self.imp_tree.configure(yscrollcommand=vsb_imp.set)
            vsb_imp.pack(side="right", fill="y")
            self.imp_tree.pack(fill="both", expand=True)

            # ── manual entry form ──────────────────────────────────────
            mf = tk.LabelFrame(tab, text="  Manual Entry  ", bg="#F5F5F5",
                font=("Segoe UI",10,"bold"), fg="#1D9E75", relief="flat",
                highlightbackground="#D0D0D0", highlightthickness=1, padx=8, pady=8)
            mf.pack(fill="x", padx=12, pady=(0,12))

            default_tgt = str(self.params.get("dip_sell_pct", 5.0))
            self._imp_gen_var    = tk.StringVar(value="1")
            self._imp_price_var  = tk.StringVar(value="")
            self._imp_qty_var    = tk.StringVar(value="")
            self._imp_cost_var   = tk.StringVar(value="")
            self._imp_target_var = tk.StringVar(value=default_tgt)
            self._imp_time_var   = tk.StringVar(
                value=datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S+00:00"))

            row0 = tk.Frame(mf, bg="#F5F5F5"); row0.pack(fill="x", pady=4)
            for lbl, var, w in [
                ("Gen ID",          self._imp_gen_var,    5),
                ("Entry Price",     self._imp_price_var,  14),
                ("BTC Qty",         self._imp_qty_var,    14),
                ("Cost (auto-calc)",self._imp_cost_var,   14),
                ("Sell Target %",   self._imp_target_var, 8),
                ("Open Time (UTC)", self._imp_time_var,   22),
            ]:
                tk.Label(row0, text=lbl+":", bg="#F5F5F5",
                    font=("Segoe UI",9)).pack(side="left", padx=(10,2))
                ttk.Entry(row0, textvariable=var, width=w).pack(side="left", padx=(0,4))

            tk.Button(row0, text="+ Add Row", command=self._add_manual_trade,
                bg="#1D9E75", fg="white", font=("Segoe UI",10,"bold"),
                padx=10, pady=4, relief="flat", cursor="hand2").pack(side="left", padx=(14,0))

        def _import_trades_file(self):
            import csv as _csv
            path = filedialog.askopenfilename(
                title="Import Trades",
                filetypes=[("CSV files","*.csv"),
                           ("Excel files","*.xlsx *.xls"),
                           ("All files","*.*")])
            if not path:
                return
            rows = []
            try:
                if path.lower().endswith((".xlsx",".xls")):
                    if not EXCEL_OK:
                        messagebox.showerror("Error",
                            "openpyxl not installed.\nRun:  pip install openpyxl")
                        return
                    wb  = openpyxl.load_workbook(path, data_only=True)
                    ws  = wb.active
                    hdr = [str(c.value).strip().lower() if c.value else ""
                           for c in next(ws.iter_rows(min_row=1, max_row=1))]
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        rows.append(dict(zip(hdr, row)))
                else:
                    with open(path, newline="", encoding="utf-8-sig") as f:
                        rows = [{k.strip().lower(): v
                                 for k, v in r.items()}
                                for r in _csv.DictReader(f)]
            except Exception as exc:
                messagebox.showerror("Import Error", str(exc))
                return

            default_tgt = self.params.get("dip_sell_pct", 5.0)
            added = 0
            for r in rows:
                try:
                    gen_id = int(float(r.get("gen_id") or r.get("gen") or 1))
                    entry  = float(r.get("entry_price") or r.get("entry") or
                                   r.get("price") or 0)
                    qty    = float(r.get("btc_qty") or r.get("btc_quantity") or
                                   r.get("qty") or 0)
                    if not entry or not qty:
                        continue
                    raw_cost = r.get("cost")
                    cost = float(raw_cost) if raw_cost not in (None,"","0") \
                           else round(entry * qty, 8)
                    tgt  = float(r.get("sell_target_pct") or
                                 r.get("sell_target") or default_tgt)
                    otime = str(r.get("open_time") or
                                datetime.now(timezone.utc).isoformat())
                    self.imp_tree.insert("", "end",
                        values=(gen_id, f"{entry:,.2f}", f"{qty:.8f}",
                                f"{cost:,.2f}", f"{tgt:.2f}", otime))
                    added += 1
                except Exception:
                    continue
            from pathlib import Path as _P
            self._log(f"Imported {added} row(s) from {_P(path).name}")

        def _add_manual_trade(self):
            try:
                gen_id = int(float(self._imp_gen_var.get() or 1))
                entry  = float(self._imp_price_var.get())
                qty    = float(self._imp_qty_var.get())
                tgt    = float(self._imp_target_var.get())
                cost_s = self._imp_cost_var.get().strip()
                cost   = float(cost_s) if cost_s else round(entry * qty, 8)
                otime  = self._imp_time_var.get().strip() or \
                         datetime.now(timezone.utc).isoformat()
            except ValueError as exc:
                messagebox.showerror("Invalid Input", str(exc))
                return
            if not entry or not qty:
                messagebox.showerror("Error","Entry Price and BTC Qty are required")
                return
            self.imp_tree.insert("", "end",
                values=(gen_id, f"{entry:,.2f}", f"{qty:.8f}",
                        f"{cost:,.2f}", f"{tgt:.2f}", otime))
            self._imp_price_var.set("")
            self._imp_qty_var.set("")
            self._imp_cost_var.set("")

        def _delete_import_row(self):
            for item in self.imp_tree.selection():
                self.imp_tree.delete(item)

        def _clear_import_rows(self):
            if messagebox.askyesno("Clear All","Remove all rows from the import list?"):
                for item in self.imp_tree.get_children():
                    self.imp_tree.delete(item)

        def _apply_imports_to_state(self):
            rows = self.imp_tree.get_children()
            if not rows:
                messagebox.showinfo("Nothing to do","No trade rows in the list.")
                return

            # ── load or initialise state ───────────────────────────────
            state = None
            if STATE_FILE.exists():
                try:
                    state = json.loads(STATE_FILE.read_text())
                except Exception:
                    pass
            if state is None:
                now = datetime.now(timezone.utc).isoformat()
                state = {"generations": [], "total_invested": 0.0, "created_at": now}

            trades = []
            for item in rows:
                v = self.imp_tree.item(item)["values"]
                trades.append({
                    "gen_id": int(v[0]),
                    "entry":  float(str(v[1]).replace(",","")),
                    "qty":    float(str(v[2]).replace(",","")),
                    "cost":   float(str(v[3]).replace(",","")),
                    "tgt":    float(str(v[4])) / 100.0,
                    "otime":  str(v[5]),
                })

            total_added = 0
            for t in trades:
                gid = t["gen_id"]
                gen = next((g for g in state["generations"] if g["id"] == gid), None)
                if gen is None:
                    gen = make_generation(gid, self.params.get("gen_capital", 1100),
                                          datetime.now(timezone.utc).isoformat())
                    gen["id"] = gid
                    state["generations"].append(gen)

                slot = next((s for s in gen["slots"] if s["state"] == "IDLE"), None)
                if slot is None:
                    slot = make_slot(len(gen["slots"]))
                    gen["slots"].append(slot)
                    gen["n_slots"] = len(gen["slots"])

                slot.update({
                    "state":           "DIP",
                    "active":          True,
                    "dip_btc":         t["qty"],
                    "dip_entry":       t["entry"],
                    "dip_cost":        t["cost"],
                    "sell_target":     t["tgt"],
                    "dip_orig_target": t["tgt"],
                    "peak_price":      t["entry"],
                    "dip_open_time":   t["otime"],
                })
                gen["free_cash"] = max(0.0, gen.get("free_cash", 0.0) - t["cost"])
                state["total_invested"] = state.get("total_invested", 0.0) + t["cost"]
                total_added += 1

            STATE_FILE.write_text(json.dumps(state, indent=2))
            self.bot_state = state
            self._refresh_positions()
            for item in self.imp_tree.get_children():
                self.imp_tree.delete(item)
            self._log(f"Saved {total_added} trade(s) to bot state → {STATE_FILE}")
            messagebox.showinfo("Saved",
                f"{total_added} trade(s) written to bot state.\n"
                "Check the Positions tab to confirm.")

        def _build_settings(self, nb):
            tab = ttk.Frame(nb); nb.add(tab, text="  Strategy Settings  ")
            tb = tk.Frame(tab, bg="#F5F5F5"); tb.pack(fill="x", padx=12, pady=8)
            bkw = dict(font=("Segoe UI",10), padx=10, pady=6, relief="flat", cursor="hand2")
            tk.Button(tb, text="Save",     command=self._save_settings,
                bg="#1D9E75", fg="white", **bkw).pack(side="left", padx=(0,6))
            tk.Button(tb, text="Defaults", command=self._reset_defaults,
                bg="#666666", fg="white", **bkw).pack(side="left", padx=(0,6))

            canvas = tk.Canvas(tab, bg="#F5F5F5", highlightthickness=0)
            sb = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
            canvas.configure(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y"); canvas.pack(fill="both", expand=True, padx=12, pady=4)
            inner = tk.Frame(canvas, bg="#F5F5F5")
            cw = canvas.create_window((0,0), window=inner, anchor="nw")
            inner.bind("<Configure>", lambda e: (canvas.configure(scrollregion=canvas.bbox("all")),
                                                 canvas.itemconfig(cw, width=canvas.winfo_width())))
            canvas.bind("<Configure>", lambda e: canvas.itemconfig(cw, width=canvas.winfo_width()))
            def _mw(e):
                if e.delta: canvas.yview_scroll(int(-1*(e.delta/120)), "units")
                elif e.num==4: canvas.yview_scroll(-1, "units")
                elif e.num==5: canvas.yview_scroll(1,  "units")
            canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _mw))
            canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

            groups = {
                "Capital & Sizing":  ["gen_capital","slot_size","split_min_size","max_generations"],
                "Buy Trigger":       ["exchange","pair","buy_drop_pct","lookback_hours"],
                "Sell Targets":      ["dip_sell_pct","sell_decay_days","sell_decay_rate","sell_floor_pct"],
                "Escalation":        ["escalation_start","escalation_step","escalation_max_lvl",
                                      "entry_decay_days","entry_decay_rate"],
                "Generation Spawn":  ["new_gen_threshold","new_gen_cooldown_d"],
                "Bot Control":       ["check_interval_min","dry_run"],
            }
            self.param_vars = {}
            for grp, keys in groups.items():
                gf = tk.LabelFrame(inner, text=f"  {grp}  ", bg="#F5F5F5",
                    font=("Segoe UI",10,"bold"), fg="#1D9E75", relief="flat",
                    highlightbackground="#D0D0D0", highlightthickness=1, padx=8, pady=4)
                gf.pack(fill="x", padx=4, pady=6)
                for key in keys:
                    lbl, desc = PARAM_LABELS.get(key, (key, ""))
                    row = tk.Frame(gf, bg="#F5F5F5"); row.pack(fill="x", pady=2)
                    tk.Label(row, text=lbl+":", bg="#F5F5F5", font=("Segoe UI",10),
                        width=26, anchor="e").pack(side="left", padx=(4,8))
                    val = self.params.get(key, "")
                    if isinstance(val, bool):
                        var = tk.BooleanVar(value=val)
                        ttk.Checkbutton(row, variable=var, text="Enabled").pack(side="left")
                    elif key == "pair":
                        var = tk.StringVar(value=str(val))
                        self.settings_pair_combo = ttk.Combobox(row, textvariable=var,
                            values=[str(val)], width=16, state="readonly")
                        self.settings_pair_combo.pack(side="left")
                    elif key == "exchange":
                        var = tk.StringVar(value=str(val))
                        ttk.Combobox(row, textvariable=var, values=["VALR","BINANCE"],
                            width=16, state="readonly").pack(side="left")
                    else:
                        var = tk.StringVar(value=str(val))
                        ttk.Entry(row, textvariable=var, width=16).pack(side="left")
                    self.param_vars[key] = var
                    tk.Label(row, text=desc, bg="#F5F5F5", font=("Segoe UI",9),
                        fg="#888888").pack(side="left", padx=8)

        def _build_api_tab(self, nb):
            tab = ttk.Frame(nb); nb.add(tab, text="  API & Exchange  ")
            fr = tk.Frame(tab, bg="#F5F5F5"); fr.pack(padx=40, pady=24, fill="x")
            tk.Label(fr, text="Exchange API Configuration", bg="#F5F5F5",
                font=("Segoe UI",13,"bold")).pack(anchor="w", pady=(0,4))
            tk.Label(fr,
                text="On Render: set VALR_API_KEY / VALR_API_SECRET as environment secrets.\n"
                     "Locally: enter below and click Save (stored encrypted on this machine only).",
                bg="#F5F5F5", font=("Segoe UI",9), fg="#CC3333").pack(anchor="w", pady=(0,14))

            self.exchange_var  = tk.StringVar(value=self.api_cfg.get("exchange","VALR"))
            self.apikey_var    = tk.StringVar(value=self.api_cfg.get("api_key",""))
            self.apisecret_var = tk.StringVar(value=self.api_cfg.get("api_secret",""))

            for lbl, attr, secret in [("Exchange","exchange_var",False),
                                       ("API Key","apikey_var",False),
                                       ("API Secret","apisecret_var",True)]:
                row = tk.Frame(fr, bg="#F5F5F5"); row.pack(fill="x", pady=5)
                tk.Label(row, text=lbl+":", bg="#F5F5F5", font=("Segoe UI",10),
                    width=12, anchor="e").pack(side="left", padx=(0,8))
                var = getattr(self, attr)
                if lbl == "Exchange":
                    ttk.Combobox(row, textvariable=var, values=["VALR","BINANCE"],
                        width=20, state="readonly").pack(side="left")
                else:
                    e = ttk.Entry(row, textvariable=var, show="*" if secret else "", width=56)
                    e.pack(side="left")
                    if secret:
                        tk.Button(row, text="show/hide",
                            command=lambda en=e: en.config(show="" if en.cget("show")=="*" else "*"),
                            font=("Segoe UI",8), relief="flat", cursor="hand2").pack(side="left", padx=6)

            pr = tk.Frame(fr, bg="#F5F5F5"); pr.pack(fill="x", pady=6)
            tk.Label(pr, text="Trading Pair:", bg="#F5F5F5", font=("Segoe UI",10),
                width=12, anchor="e").pack(side="left", padx=(0,8))
            self.pair_combo = ttk.Combobox(pr, values=[self.params.get("pair","BTCUSDC")],
                width=20, state="readonly")
            self.pair_combo.set(self.params.get("pair","BTCUSDC"))
            self.pair_combo.pack(side="left")
            self.pair_combo.bind("<<ComboboxSelected>>", self._on_pair_selected)
            tk.Button(pr, text="↻ Fetch pairs",
                command=lambda: threading.Thread(target=self._fetch_pairs, daemon=True).start(),
                font=("Segoe UI",9), relief="flat", bg="#378ADD", fg="white",
                cursor="hand2", padx=8, pady=4).pack(side="left", padx=8)

            br = tk.Frame(fr, bg="#F5F5F5"); br.pack(fill="x", pady=14)
            tk.Button(br, text="Save Credentials", command=self._save_api,
                bg="#1D9E75", fg="white", font=("Segoe UI",10,"bold"),
                padx=12, pady=7, relief="flat", cursor="hand2").pack(side="left", padx=(0,8))
            tk.Button(br, text="Test Connection", command=self._test_connection,
                bg="#378ADD", fg="white", font=("Segoe UI",10),
                padx=12, pady=7, relief="flat", cursor="hand2").pack(side="left")
            self.conn_lbl = tk.Label(fr, text="", bg="#F5F5F5", font=("Segoe UI",10))
            self.conn_lbl.pack(anchor="w", pady=6)

        def _build_log(self, nb):
            tab = ttk.Frame(nb); nb.add(tab, text="  Log  ")
            ctrl = tk.Frame(tab, bg="#F5F5F5"); ctrl.pack(fill="x", padx=12, pady=6)
            tk.Button(ctrl, text="Clear", command=self._clear_log,
                font=("Segoe UI",9), relief="flat", bg="#666666", fg="white",
                cursor="hand2", padx=8, pady=4).pack(side="left")
            self.log_box = scrolledtext.ScrolledText(tab, wrap=tk.WORD,
                font=("Courier New",9), bg="#1A1A2E", fg="#00FF88",
                insertbackground="#00FF88", state="disabled")
            self.log_box.pack(fill="both", expand=True, padx=12, pady=(0,12))

        # ── Bot control ────────────────────────────────────────
        def _toggle_bot(self):
            if self.bot_running:
                self.bot_running = False
                self.start_btn.configure(text="▶  Start Bot", bg="#1D9E75")
                self.status_lbl.configure(text="  STOPPED  ", bg="#CC3333")
                self._log("Bot stopping (will finish current tick)...")
            else:
                self.params  = load_strategy_config()
                self.api_cfg = load_api_config()
                if not self.api_cfg.get("api_key") and not self.params.get("dry_run", True):
                    messagebox.showerror("No API Key", "API key required for live trading.\n"
                                         "Set in API & Exchange tab or set VALR_API_KEY env var.")
                    return
                self.bot_running = True
                self.start_btn.configure(text="⏹  Stop Bot", bg="#CC3333")
                self.status_lbl.configure(text="  RUNNING  ", bg="#1D9E75")
                self.bot_thread = threading.Thread(target=self._bot_loop, daemon=True)
                self.bot_thread.start()
                self._log(f"Bot started — {self.params.get('pair')} "
                          f"{'[DRY RUN]' if self.params.get('dry_run') else '[LIVE]'}")

        def _bot_loop(self):
            cfg     = self.params
            api_cfg = self.api_cfg
            api_key = api_cfg.get("api_key","")
            api_sec = api_cfg.get("api_secret","")
            exchange= cfg.get("exchange","VALR")
            pair    = cfg.get("pair","BTCUSDC")
            interval= cfg.get("check_interval_min",5) * 60

            state = None
            if STATE_FILE.exists():
                try: state = json.loads(STATE_FILE.read_text())
                except Exception: pass
            if state is None:
                now   = datetime.now(timezone.utc).isoformat()
                state = {"generations": [make_generation(1, cfg["gen_capital"], now),
                                         make_generation(2, cfg["gen_capital"], now)],
                         "total_invested": cfg["gen_capital"] * 2, "created_at": now}

            price_history = []
            lkb = cfg.get("lookback_hours", 12)

            while self.bot_running:
                try:
                    now_str    = datetime.now(timezone.utc).isoformat()
                    price_data = get_live_price_with_retry(exchange, pair, api_key, api_sec)
                    if not price_data:
                        self._log("Price fetch failed — retrying in 60s")
                        time.sleep(60); continue

                    current_price = price_data["price"]
                    price_history.append(current_price)
                    if len(price_history) > lkb * 2:
                        price_history = price_history[-lkb*2:]
                    ref_price = price_history[-lkb-1] if len(price_history) > lkb else current_price

                    for gen in state["generations"]:
                        process_generation(gen, current_price, ref_price, now_str,
                                           cfg, api_key, api_sec)

                    self.bot_state  = state
                    self.price_data = price_data
                    STATE_FILE.write_text(json.dumps(state, indent=2, default=str))
                    self.after(0, self._refresh_dashboard)
                    _, quote = parse_pair(pair)
                    self._log(f"Tick: {self._fmt(current_price, quote)}")
                except Exception as e:
                    self._log(f"Error: {e}")
                time.sleep(interval)

        def _toggle_dry_run(self):
            self.params["dry_run"] = not self.params.get("dry_run", True)
            save_strategy_config(self.params)
            dr = self.params["dry_run"]
            self.dry_btn.configure(text=f"MODE: {'DRY RUN' if dr else 'LIVE  '}", bg="#BA7517" if dr else "#CC3333")
            self.mode_lbl.configure(text="DRY RUN" if dr else "LIVE", bg="#BA7517" if dr else "#CC3333")

        def _manual_refresh(self):
            threading.Thread(target=self._do_refresh, daemon=True).start()

        def _do_refresh(self):
            self.params  = load_strategy_config()
            self.api_cfg = load_api_config()
            key  = self.api_cfg.get("api_key","")
            sec  = self.api_cfg.get("api_secret","")
            exch = self.params.get("exchange","VALR")
            pair = self.params.get("pair","BTCUSDC")
            pd   = get_live_price(exch, pair, key, sec)
            if pd:
                self.price_data = pd
            if STATE_FILE.exists():
                try: self.bot_state = json.loads(STATE_FILE.read_text())
                except Exception: pass
            threading.Thread(target=self._fetch_balances, daemon=True).start()
            self.after(0, self._refresh_dashboard)

        def _start_refresh(self):
            self._do_refresh()
            self.after(30000, self._start_refresh)

        def _fetch_balances(self):
            key  = self.api_cfg.get("api_key","")
            sec  = self.api_cfg.get("api_secret","")
            exch = self.params.get("exchange","VALR")
            pair = self.params.get("pair","BTCUSDC")
            bals = get_account_balance(exch, pair, key, sec)
            if bals and "_error" not in bals:
                self.balances = bals
                self.after(0, self._update_balance_display)

        def _update_balance_display(self):
            _, quote = self._parse_pair(self.params.get("pair","BTCUSDC"))
            for cur, info in self.balances.items():
                sym = {"USDC":"R","USD":"$","USDT":"$","USDC":"$","EUR":"€","GBP":"£"}.get(cur, "")
                if info.get("is_quote"):
                    self.bal_q.configure(text=f"{cur}: {sym}{info['total']:,.2f}  (avail: {sym}{info['available']:,.2f})")
                elif info.get("is_base"):
                    self.bal_b.configure(text=f"{cur}: {info['total']:.6f}  (avail: {info['available']:.6f})")

        def _fetch_pairs(self):
            exch = self.exchange_var.get()
            key  = self.apikey_var.get().strip()
            sec  = self.apisecret_var.get().strip()
            pairs = get_all_pairs(exch, key, sec)
            if pairs:
                self.after(0, self._update_pair_selector, pairs)
                self._log(f"Fetched {len(pairs)} pairs")

        def _update_pair_selector(self, pairs):
            if hasattr(self, "pair_combo"):
                self.pair_combo["values"] = pairs
            if hasattr(self, "settings_pair_combo"):
                self.settings_pair_combo["values"] = pairs

        def _on_pair_selected(self, _=None):
            pair = self.pair_combo.get()
            self.params["pair"] = pair
            if hasattr(self, "settings_pair_combo"):
                self.settings_pair_combo.set(pair)
            if "pair" in self.param_vars:
                self.param_vars["pair"].set(pair)

        def _refresh_dashboard(self):
            self.upd_lbl.configure(text=f"Last update: {datetime.now().strftime('%H:%M:%S')}")
            pd = self.price_data
            if pd:
                _, quote = self._parse_pair(self.params.get("pair","BTCUSDC"))
                self.price_hdr_lbl.configure(text=f"Price ({self.params.get('pair','')})")
                self.metric_cards["price"].configure(text=self._fmt(pd["price"], quote))
                chg = pd.get("chg","?")
                try:
                    c = float(chg)
                    self.metric_cards["change"].configure(
                        text=f"{c:+.2f}%", fg="#1D9E75" if c >= 0 else "#CC3333")
                except Exception:
                    self.metric_cards["change"].configure(text=str(chg), fg="#444444")
                if pd.get("bid") and pd.get("ask"):
                    sp = pd["ask"] - pd["bid"]
                    self.spread_lbl.configure(
                        text=f"bid {self._fmt(pd['bid'],quote)}  ask {self._fmt(pd['ask'],quote)}  spread {self._fmt(sp,quote)}")

            state = self.bot_state
            if not state:
                self._show_strategy_preview(self.params.get("pair","BTCUSDC"))
                return

            for item in self.gen_tree.get_children():
                self.gen_tree.delete(item)
            price   = self.price_data.get("price", 0)
            gens    = state.get("generations", [])
            _, quote = self._parse_pair(self.params.get("pair","BTCUSDC"))
            sym     = {"USDC":"R","USD":"$","USDT":"$","USDC":"$","EUR":"€","GBP":"£"}.get(quote, quote+" ")
            total_val = total_real = total_open = total_idle = 0

            for gen in gens:
                slots   = gen.get("slots",[])
                capital = gen.get("capital", 0)
                cash    = gen.get("free_cash", 0)
                n_slots = gen.get("n_slots", 0)
                slot_sz = gen.get("slot_size", 0)
                n_open  = sum(1 for s in slots if s.get("state") == "DIP")
                n_idle  = sum(1 for s in slots if s.get("state") == "IDLE")
                btc_val = sum(s.get("dip_btc",0) * price for s in slots if s.get("state") == "DIP")
                open_cost   = sum(s.get("dip_cost",0) for s in slots if s.get("state") == "DIP")
                unrealised  = btc_val - open_cost
                tot         = cash + btc_val
                realised    = sum(t.get("profit",0) for t in gen.get("trade_log",[]) if t.get("type") == "SELL")
                n_trades    = sum(1 for t in gen.get("trade_log",[]) if t.get("type") == "SELL")
                pnl         = (tot - capital) / capital * 100 if capital else 0
                start       = str(gen.get("start_time","?"))[:10]
                total_val  += tot; total_real += realised; total_open += n_open; total_idle += n_idle
                tag = "profit" if pnl > 0 else ("loss" if pnl < 0 else "")
                self.gen_tree.insert("","end", tags=(tag,), values=(
                    f"Gen {gen.get('id',1)}", start, f"{sym}{capital:,.2f}", f"{sym}{cash:,.2f}",
                    f"{n_open} / {n_slots}", n_idle, f"{sym}{slot_sz:,.4f}",
                    f"{sym}{realised:,.4f}", f"{sym}{unrealised:+,.4f}",
                    f"{sym}{tot:,.4f}", f"{pnl:+.2f}%", n_trades))

            invested  = state.get("total_invested", 0)
            grand_pnl = total_val - invested
            self.metric_cards["portfolio"].configure(text=f"{sym}{total_val:,.0f}")
            self.metric_cards["realised"].configure(
                text=f"{sym}{total_real:+,.2f}", fg="#1D9E75" if total_real >= 0 else "#CC3333")
            self.metric_cards["openpos"].configure(text=str(total_open))
            self.metric_cards["gens"].configure(text=str(len(gens)))
            self._refresh_positions()

        def _show_strategy_preview(self, pair):
            for item in self.gen_tree.get_children(): self.gen_tree.delete(item)
            _, quote = self._parse_pair(pair)
            sym = {"USDC":"R","USD":"$","USDT":"$","USDC":"$","EUR":"€","GBP":"£"}.get(quote, quote+" ")
            def _pv(key, default):
                if key in self.param_vars:
                    try:
                        v = self.param_vars[key].get()
                        orig = self.params.get(key, default)
                        if isinstance(orig, bool):   return bool(v)
                        elif isinstance(orig, int):   return int(float(str(v)))
                        elif isinstance(orig, float): return float(v)
                        return str(v)
                    except Exception: pass
                return self.params.get(key, default)
            cap     = _pv("gen_capital", 1100.0)
            slot_sz = _pv("slot_size",   100.0)
            n_slots = max(1, int(cap // slot_sz)) if slot_sz > 0 else 11
            slot_sz = cap / n_slots
            for g in range(1, 3):
                self.gen_tree.insert("","end", values=(
                    f"Gen {g}", "not started yet",
                    f"{sym}{cap:,.2f}", f"{sym}{cap:,.2f}",
                    f"0 / {n_slots}", n_slots, f"{sym}{slot_sz:,.2f}",
                    f"{sym}0.00", f"{sym}0.00", f"{sym}{cap:,.2f}", "0.0%", "0"))

        def _refresh_positions(self):
            for item in self.pos_tree.get_children(): self.pos_tree.delete(item)
            state = self.bot_state; price = self.price_data.get("price",0)
            filt  = self.pos_filter.get()
            _, quote = self._parse_pair(self.params.get("pair","BTCUSDC"))
            sym  = {"USDC":"R","USD":"$","USDT":"$","USDC":"$","EUR":"€","GBP":"£"}.get(quote, quote+" ")
            for gen in state.get("generations",[]):
                gid = gen.get("id",1)
                for s in gen.get("slots",[]):
                    st = s.get("state","IDLE")
                    if filt=="Open" and st!="DIP":  continue
                    if filt=="Idle" and st!="IDLE": continue
                    if st == "DIP":
                        entry  = s.get("dip_entry",0)
                        tpct   = s.get("sell_target",0.05)
                        tp     = s.get("sell_order_price") or (entry*(1+tpct) if entry else 0)
                        peak   = s.get("peak_price") or entry
                        hold   = s.get("dip_btc",0) * price
                        unreal = hold - s.get("dip_cost",0)
                        cpct   = (price-entry)/entry*100 if entry else 0
                        pkpct  = (peak-entry)/entry*100  if entry else 0
                        days   = 0
                        if s.get("dip_open_time"):
                            try:
                                ot = datetime.fromisoformat(s["dip_open_time"])
                                days = (datetime.now(timezone.utc)-ot).days
                            except Exception: pass
                        oid  = s.get("sell_order_id","")
                        ostr = f"ORD:{oid[-6:]}" if oid and not str(oid).startswith("DRY_") \
                               else ("DRY" if oid else "NONE")
                        tag  = "close" if tp and (tp-price)/price < 0.005 else "dip"
                        vals = (gid, s.get("id",0)+1, "DIP",
                                f"{sym}{entry:,.0f}" if entry else "—",
                                f"{sym}{tp:,.0f}" if tp else "—", f"{tpct*100:.0f}%",
                                f"{sym}{peak:,.0f}" if peak else "—", f"{pkpct:+.1f}%",
                                f"{sym}{price:,.0f}" if price else "—", f"{cpct:+.1f}%",
                                f"{sym}{hold:,.2f}", f"{sym}{unreal:+,.2f}", f"{days}d", ostr)
                    else:
                        tag  = "idle"
                        vals = (gid, s.get("id",0)+1, "IDLE",
                                "—","—","—","—","—","—","—","—","—","—","—")
                    self.pos_tree.insert("","end", values=vals, tags=(tag,))

        def _save_settings(self):
            changed = []
            for key, var in self.param_vars.items():
                val  = var.get()
                orig = self.params.get(key)
                try:
                    if isinstance(orig, bool):   new_val = bool(val)
                    elif isinstance(orig, int):   new_val = int(float(str(val)))
                    elif isinstance(orig, float): new_val = float(val)
                    else:                         new_val = str(val)
                    if new_val != orig:
                        changed.append(f"{key}: {orig} -> {new_val}")
                    self.params[key] = new_val
                except Exception: pass
            save_strategy_config(self.params)
            self._log(f"Settings saved ({len(changed)} changes)")
            messagebox.showinfo("Saved", f"Strategy settings saved.\n{len(changed)} parameter(s) changed.")

        def _reset_defaults(self):
            if messagebox.askyesno("Reset","Reset all parameters to defaults?"):
                self.params = dict(DEFAULT_PARAMS)
                for key, var in self.param_vars.items():
                    var.set(str(self.params.get(key,"")))

        def _save_api(self):
            cfg = {"exchange": self.exchange_var.get(),
                   "api_key":  self.apikey_var.get().strip(),
                   "api_secret": self.apisecret_var.get().strip(),
                   "pair": self.pair_combo.get() if hasattr(self,"pair_combo") else self.params.get("pair","BTCUSDC")}
            save_api_config(cfg); self.api_cfg = cfg
            self.params["pair"] = cfg["pair"]
            self._log("API credentials saved (encrypted locally)")
            messagebox.showinfo("Saved","Credentials saved securely.")

        def _test_connection(self):
            self.conn_lbl.configure(text="Testing...", fg="#BA7517")
            self.update_idletasks()
            def do_test():
                exchange = self.exchange_var.get()
                key      = self.apikey_var.get().strip()
                secret   = self.apisecret_var.get().strip()
                pair     = self.pair_combo.get() if hasattr(self,"pair_combo") else self.params.get("pair","BTCUSDC")
                price    = get_live_price(exchange, pair, key, secret)
                if price:
                    _, quote = self._parse_pair(pair)
                    msg = f"Connected — {pair}: {self._fmt(price['price'],quote)}"
                    self.after(0, self.conn_lbl.configure, {"text":msg,"fg":"#1D9E75"})
                    self._log(f"Connection OK — {pair} = {price['price']:,.2f}")
                    threading.Thread(target=self._fetch_pairs, daemon=True).start()
                else:
                    self.after(0, self.conn_lbl.configure,
                        {"text":"Connection failed — check key / internet","fg":"#CC3333"})
            threading.Thread(target=do_test, daemon=True).start()

        def _log(self, msg):
            def _do():
                self.log_box.configure(state="normal")
                ts = datetime.now().strftime("%H:%M:%S")
                self.log_box.insert("end", f"{ts}  {msg}\n")
                self.log_box.see("end")
                self.log_box.configure(state="disabled")
            self.after(0, _do)

        def _clear_log(self):
            self.log_box.configure(state="normal")
            self.log_box.delete("1.0","end")
            self.log_box.configure(state="disabled")

    TradingApp().mainloop()

# =============================================================
#  ENTRY POINT
# =============================================================
if __name__ == "__main__":
    # On Render: RUN_MODE env var or no display → default to live
    _default = os.environ.get("RUN_MODE",
                              "gui" if (GUI_OK and not os.environ.get("RENDER")) else "live")
    mode = (sys.argv[1].lower() if len(sys.argv) > 1 else _default)

    if mode == "backtest":
        run_backtest()
    elif mode == "live":
        run_live_bot()
    elif mode == "valrtest":
        run_valr_test()
    elif mode == "premium":
        run_premium_check()
    else:
        run_gui()
