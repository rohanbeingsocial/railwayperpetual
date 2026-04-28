"""
Polymarket Copy Trade P&L Tracker
Railway deployment — runs 24/7
"""

import requests, time, os, json, base64
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# CONFIG
POLL_INTERVAL       = int(os.getenv('POLL_INTERVAL', 60))
LEADERBOARD_REFRESH = int(os.getenv('LEADERBOARD_REFRESH', 30))
TOP_N               = int(os.getenv('TOP_N', 20))
MIN_POSITION_USD    = float(os.getenv('MIN_POSITION_USD', 10))
PAPER_BALANCE       = float(os.getenv('PAPER_BALANCE', 1000.0))
BET_SIZE_PCT        = float(os.getenv('BET_SIZE_PCT', 0.05))
EXCEL_FILE          = os.getenv('EXCEL_FILE', 'copy_trade_pnl.xlsx')

DATA_API  = 'https://data-api.polymarket.com/v1'
CLOB_API  = 'https://clob.polymarket.com'
CATEGORIES = ['POLITICS','ECONOMICS','FINANCE','CULTURE','TECH','SPORTS','CRYPTO']

GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
GITHUB_REPO  = os.getenv("GITHUB_REPO")

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0','Accept': 'application/json'})

def log(msg):
    print(f'[{datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")}] {msg}', flush=True)

# ================= GITHUB UPLOAD =================
def upload_to_github(file_path, commit_msg="update excel"):
    if not GITHUB_TOKEN or not GITHUB_REPO:
        return
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{file_path}"
    try:
        with open(file_path, "rb") as f:
            content = base64.b64encode(f.read()).decode()

        headers = {"Authorization": f"token {GITHUB_TOKEN}"}

        res = requests.get(url, headers=headers)
        sha = res.json().get("sha") if res.status_code == 200 else None

        data = {"message": commit_msg, "content": content}
        if sha:
            data["sha"] = sha

        requests.put(url, headers=headers, json=data)
        log("☁️ Uploaded to GitHub")

    except Exception as e:
        log(f"GitHub error: {e}")

# ================= API =================
def fetch_top_wallets(top_n=TOP_N):
    wallet_map = {}
    for cat in CATEGORIES:
        try:
            res = session.get(f'{DATA_API}/leaderboard',
                params={'timePeriod':'MONTH','limit':top_n,'orderBy':'PNL','category':cat})
            if res.status_code != 200:
                continue
            for t in res.json():
                w = t.get('proxyWallet','')
                if not w.startswith('0x'): continue
                wallet_map[w] = {
                    'username': t.get('userName') or w[:10],
                    'wallet': w,
                    'category': cat
                }
        except:
            pass
    return wallet_map

def get_activity(wallet):
    try:
        return session.get(f'{DATA_API}/activity',
            params={'user':wallet,'limit':20}).json()
    except:
        return []

def get_current_price(token_id):
    try:
        r = session.get(f'{CLOB_API}/midpoint',params={'token_id':token_id})
        return float(r.json().get('mid',0))
    except:
        return None

def parse_trade(item):
    try:
        price = float(item.get('price',0))
        size  = float(item.get('size',0))
        return {
            'market': item.get('title','Unknown'),
            'outcome': item.get('outcome','?'),
            'side': item.get('side','?'),
            'price': price,
            'size': size,
            'usd_value': price*size,
            'tx_hash': item.get('transactionHash',''),
            'asset': item.get('asset',''),
            'condition_id': item.get('conditionId',''),
            'outcome_index': item.get('outcomeIndex',0)
        }
    except:
        return None

def pos_key(wallet, cid, idx):
    return f"{wallet[:6]}_{cid[:6]}_{idx}"

# ================= EXCEL =================
def init_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Trades"
    ws.append(["Time","Wallet","Market","Side","Price","Size"])
    wb.save(EXCEL_FILE)

def update_excel(row):
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append(row)
    wb.save(EXCEL_FILE)

# ================= MAIN =================
def main():
    log("🚀 Starting...")

    init_excel()
    seen = {}
    wallets = fetch_top_wallets()

    cycle = 0

    while True:
        cycle += 1

        for w, info in wallets.items():
            acts = get_activity(w)
            for item in acts:
                tx = item.get('transactionHash')
                if not tx or tx in seen.get(w,set()):
                    continue

                seen.setdefault(w,set()).add(tx)
                trade = parse_trade(item)

                if not trade or trade['usd_value'] < MIN_POSITION_USD:
                    continue

                log(f"{info['username']} {trade['side']} {trade['market']}")

                update_excel([
                    datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S'),
                    w,
                    trade['market'],
                    trade['side'],
                    trade['price'],
                    trade['size']
                ])

        # upload every 5 cycles
        if cycle % 5 == 0:
            upload_to_github(EXCEL_FILE)

        time.sleep(POLL_INTERVAL)

if __name__ == "__main__":
    main()
