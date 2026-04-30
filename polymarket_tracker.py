"""
Polymarket Wallet Tracker & Market Predictor
Uses public Data API (no auth needed).

Requirements:
    pip install requests pandas openpyxl python-dotenv schedule

Railway setup:
    - Mount a persistent volume at /data
    - Set OUTPUT_DIR=/data in environment variables
"""

import os
import time
import requests
import pandas as pd
from datetime import datetime, timezone
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Config ────────────────────────────────────────────────────────────────────

WALLETS = [
    "0x53e55bc7cb3d67ad177c023ce891ad076a9d6177",
    "0xc6587b11a2209e46dfe3928b31c5514a8e33b784",
]

OUTPUT_DIR = os.environ.get("OUTPUT_DIR", "/data")
EXCEL_PATH = os.path.join(OUTPUT_DIR, "polymarket_tracker.xlsx")
os.makedirs(OUTPUT_DIR, exist_ok=True)

DATA_API  = "https://data-api.polymarket.com"
GAMMA_API = "https://gamma-api.polymarket.com"
HEADERS   = {"Accept": "application/json"}

# ── Fetch ─────────────────────────────────────────────────────────────────────

def get_json(url, params=None, retries=3):
    for attempt in range(retries):
        try:
            r = requests.get(url, params=params, headers=HEADERS, timeout=20)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            print(f"    [retry {attempt+1}] {url} — {e}")
            time.sleep(2 ** attempt)
    return None


def fetch_trades(wallet: str) -> list:
    """
    Fetch all trades for a wallet from the public Data API.
    Endpoint: GET /trades?user=<wallet>&limit=500&offset=N
    Response fields: proxyWallet, side, price, size, outcome, timestamp,
                     conditionId, transactionHash, market (slug)
    """
    trades = []
    offset = 0
    limit  = 500

    while True:
        data = get_json(f"{DATA_API}/trades", params={
            "user":   wallet.lower(),
            "limit":  limit,
            "offset": offset,
        })
        if not data or not isinstance(data, list):
            break
        trades.extend(data)
        print(f"    Fetched {len(trades)} trades so far...")
        if len(data) < limit:
            break
        offset += limit
        time.sleep(0.4)

    return trades


def fetch_market_info(condition_id: str) -> dict:
    """Fetch market metadata — question, end date, outcomes."""
    if not condition_id:
        return {}
    data = get_json(f"{GAMMA_API}/markets", params={"conditionId": condition_id})
    if isinstance(data, list) and data:
        return data[0]
    return {}


# ── Parse & Enrich ────────────────────────────────────────────────────────────

def parse_trades(raw: list, wallet: str) -> pd.DataFrame:
    rows = []
    for t in raw:
        try:
            # Data API timestamps can be seconds or milliseconds
            ts = int(t.get("timestamp", 0) or 0)
            if ts > 9_999_999_999:
                ts = ts // 1000
            dt = datetime.fromtimestamp(ts, tz=timezone.utc) if ts else None

            price = float(t.get("price", 0) or 0)
            size  = float(t.get("size",  0) or 0)

            rows.append({
                "wallet":       wallet,
                "timestamp":    dt,
                "side":         (t.get("side") or "").upper(),
                "outcome":      t.get("outcome") or t.get("name") or "",
                "price":        price,
                "implied_prob": round(price * 100, 1),
                "size_usdc":    size,
                "value_usdc":   round(price * size, 2),
                "condition_id": t.get("conditionId") or t.get("market") or "",
                "tx_hash":      t.get("transactionHash") or "",
                "market_name":  "",   # filled below
            })
        except Exception:
            continue

    df = pd.DataFrame(rows)
    if not df.empty:
        df.sort_values("timestamp", ascending=False, inplace=True)
    return df


def enrich_market_names(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    unique_cids = [c for c in df["condition_id"].unique() if c]
    print(f"    Fetching names for {len(unique_cids)} markets...")
    cid_map = {}
    for cid in unique_cids:
        info = fetch_market_info(cid)
        cid_map[cid] = info.get("question") or info.get("title") or cid[:20]
        time.sleep(0.2)
    df["market_name"] = df["condition_id"].map(cid_map).fillna("Unknown")
    return df


# ── Stats & Predictions ───────────────────────────────────────────────────────

def wallet_stats(df: pd.DataFrame) -> dict:
    if df.empty:
        return {"note": "No trades found"}
    buys  = df[df["side"] == "BUY"]
    sells = df[df["side"] == "SELL"]
    return {
        "total_trades":    len(df),
        "buy_volume_usdc": round(buys["size_usdc"].sum(), 2),
        "sell_volume_usdc":round(sells["size_usdc"].sum(), 2),
        "unique_markets":  df["condition_id"].nunique(),
        "avg_buy_price":   round(buys["price"].mean(), 4) if not buys.empty else 0,
        "avg_sell_price":  round(sells["price"].mean(), 4) if not sells.empty else 0,
        "first_trade":     df["timestamp"].min(),
        "last_trade":      df["timestamp"].max(),
    }


def predict_next_markets(df: pd.DataFrame, top_n=10) -> pd.DataFrame:
    if df.empty or "market_name" not in df.columns:
        return pd.DataFrame()

    now = datetime.now(tz=timezone.utc)
    df  = df.copy()
    df["days_ago"] = (now - df["timestamp"]).dt.total_seconds() / 86400
    df["weight"]   = df["days_ago"].apply(lambda d: 0.95 ** d)

    grp = df.groupby(["condition_id", "market_name"])

    def avg_buy_price(g):
        buys = g[g["side"] == "BUY"]["price"]
        return buys.mean() if not buys.empty else None

    def net_usdc(g):
        b = g[g["side"] == "BUY"]["size_usdc"].sum()
        s = g[g["side"] == "SELL"]["size_usdc"].sum()
        return b - s

    scores = grp.apply(lambda g: pd.Series({
        "total_weight":    g["weight"].sum(),
        "trade_count":     len(g),
        "avg_buy_price":   avg_buy_price(g),
        "net_usdc":        net_usdc(g),
        "last_trade":      g["timestamp"].max(),
        "yes_pct":         (g["outcome"] == "YES").mean(),
    })).reset_index()

    # Focus on markets they've exited (likely to re-enter similar ones)
    closed = scores[scores["net_usdc"] <= 5].copy()
    closed["score"] = closed["total_weight"] * 0.6 + closed["trade_count"] * 0.4
    closed = closed.sort_values("score", ascending=False).head(top_n)

    closed["predicted_entry"] = closed["avg_buy_price"].apply(
        lambda p: f"{p:.2f}  ({p*100:.0f}¢)" if pd.notna(p) else "N/A"
    )
    closed["likely_outcome"] = (closed["yes_pct"] >= 0.5).map({True: "YES", False: "NO"})

    return closed[[
        "market_name", "condition_id", "trade_count",
        "last_trade", "predicted_entry", "likely_outcome", "score"
    ]].rename(columns={
        "market_name":    "Market",
        "condition_id":   "Condition ID",
        "trade_count":    "Historical Trades",
        "last_trade":     "Last Seen",
        "predicted_entry":"Predicted Entry (odds)",
        "likely_outcome": "Likely Outcome",
        "score":          "Confidence Score",
    })


# ── Excel ─────────────────────────────────────────────────────────────────────

NAVY  = "1F3864"
GOLD  = "FFD700"
GREEN = "E2EFDA"
RED   = "FCE4D6"
GRAY  = "F5F5F5"

def hdr(cell, bg=NAVY, fg="FFFFFF"):
    cell.font      = Font(bold=True, color=fg, size=10)
    cell.fill      = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

def data_cell(cell, val, row_idx, side=None):
    thin = Side(style="thin", color="CCCCCC")
    if side == "BUY":
        fill = GREEN
    elif side == "SELL":
        fill = RED
    else:
        fill = GRAY if row_idx % 2 == 0 else "FFFFFF"
    if isinstance(val, datetime):
        val = val.strftime("%Y-%m-%d %H:%M UTC")
    cell.value     = val
    cell.font      = Font(size=9)
    cell.fill      = PatternFill("solid", start_color=fill)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border    = Border(left=thin, right=thin, top=thin, bottom=thin)

def write_df(ws, df, start_row=1):
    for ci, col in enumerate(df.columns, 1):
        c = ws.cell(row=start_row, column=ci, value=col)
        hdr(c)
        ws.column_dimensions[get_column_letter(ci)].width = max(16, len(str(col)) + 4)
    for ri, row in enumerate(df.itertuples(index=False), start_row + 1):
        side = None
        if "side" in df.columns:
            side = getattr(row, "side", None)
        for ci, val in enumerate(row, 1):
            data_cell(ws.cell(row=ri, column=ci), val, ri, side=side)
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    ws.auto_filter.ref = ws.dimensions


def build_excel(all_data: dict):
    wb = Workbook()
    wb.remove(wb.active)

    # ── Summary ───────────────────────────────────────────────────────────────
    ws = wb.create_sheet("Summary")
    ws["A1"] = "Polymarket Wallet Intelligence"
    ws["A1"].font = Font(bold=True, size=16, color=NAVY)
    ws["A2"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M UTC')}"
    ws["A2"].font = Font(italic=True, size=9, color="888888")

    row = 4
    for wallet, d in all_data.items():
        ws.cell(row=row, column=1, value=f"Wallet: {wallet}").font = Font(bold=True, size=11, color=NAVY)
        row += 1
        for k, v in d["stats"].items():
            ws.cell(row=row, column=1, value=k.replace("_", " ").title()).font = Font(bold=True, size=9)
            val = v.strftime("%Y-%m-%d %H:%M UTC") if isinstance(v, datetime) else v
            ws.cell(row=row, column=2, value=val).font = Font(size=9)
            row += 1
        row += 2

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 30

    # ── Per wallet sheets ─────────────────────────────────────────────────────
    for wallet, d in all_data.items():
        short = wallet[:8]
        df    = d["trades"]

        # Trades sheet
        ws_t = wb.create_sheet(f"Trades {short}")
        if not df.empty:
            cols = ["timestamp", "market_name", "side", "outcome",
                    "price", "implied_prob", "size_usdc", "value_usdc",
                    "condition_id", "tx_hash"]
            cols = [c for c in cols if c in df.columns]
            write_df(ws_t, df[cols])
        else:
            ws_t["A1"] = "No trades found for this wallet."

        # Predictions sheet
        preds = d["predictions"]
        ws_p  = wb.create_sheet(f"Predict {short}")
        ws_p["A1"] = f"Next Market Predictions — {wallet}"
        ws_p["A1"].font = Font(bold=True, size=12, color=NAVY)
        ws_p["A2"] = "Confidence Score = recency-weighted trade frequency. Higher = more likely next entry."
        ws_p["A2"].font = Font(italic=True, size=8, color="666666")
        ws_p.row_dimensions[1].height = 18

        if not preds.empty:
            write_df(ws_p, preds, start_row=4)
            # Gold highlight for top prediction
            for ci in range(1, len(preds.columns) + 1):
                c = ws_p.cell(row=5, column=ci)
                c.fill = PatternFill("solid", start_color=GOLD)
                c.font = Font(bold=True, size=9)
        else:
            ws_p["A4"] = "Not enough trade history to generate predictions."

    wb.save(EXCEL_PATH)
    print(f"\n✅ Excel saved → {EXCEL_PATH}")


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    print(f"\n{'='*60}")
    print(f"  Polymarket Tracker — {datetime.now().strftime('%Y-%m-%d %H:%M')} UTC")
    print(f"{'='*60}\n")

    all_data = {}

    for wallet in WALLETS:
        print(f"📍 Wallet: {wallet}")
        print("  Fetching trades from Data API...")

        raw = fetch_trades(wallet)
        print(f"  → {len(raw)} raw trades")

        df = parse_trades(raw, wallet)
        df = enrich_market_names(df)

        stats = wallet_stats(df)
        preds = predict_next_markets(df, top_n=10)

        all_data[wallet] = {"trades": df, "stats": stats, "predictions": preds}

        print(f"  Stats: {stats.get('total_trades', 0)} trades across "
              f"{stats.get('unique_markets', 0)} markets")

        if not preds.empty:
            print(f"\n  🔮 Top predictions:")
            for _, r in preds.head(3).iterrows():
                print(f"     • {str(r['Market'])[:55]}")
                print(f"       Entry: {r['Predicted Entry (odds)']}  |  Outcome: {r['Likely Outcome']}")
        print()

    build_excel(all_data)


if __name__ == "__main__":
    import schedule

    mode = os.environ.get("RUN_MODE", "once")
    if mode == "scheduled":
        run()
        schedule.every(6).hours.do(run)
        print("\n⏰ Scheduled every 6 hours. Ctrl+C to stop.")
        while True:
            schedule.run_pending()
            time.sleep(60)
    else:
        run()
