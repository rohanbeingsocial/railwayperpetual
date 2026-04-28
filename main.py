"""
Polymarket Copy Trade P&L Tracker
Railway deployment — runs 24/7 with Bulletproof State Persistence
"""
import requests, time, os, json, base64
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════════════════
# CONFIG — edit these or set as Railway environment variables
# ═══════════════════════════════════════════════════════════
POLL_INTERVAL       = int(os.getenv('POLL_INTERVAL', 60))
LEADERBOARD_REFRESH = int(os.getenv('LEADERBOARD_REFRESH', 30))
TOP_N               = int(os.getenv('TOP_N', 100))
MIN_POSITION_USD    = float(os.getenv('MIN_POSITION_USD', 10))
PAPER_BALANCE       = float(os.getenv('PAPER_BALANCE', 1000.0))
BET_SIZE_PCT        = float(os.getenv('BET_SIZE_PCT', 0.05))
EXCEL_FILE          = os.getenv('EXCEL_FILE', 'copy_trade_pnl.xlsx')
STATE_FILE          = 'tracker_state.json'

DATA_API  = 'https://data-api.polymarket.com/v1'
CLOB_API  = 'https://clob.polymarket.com'
CATEGORIES = ['POLITICS', 'ECONOMICS', 'FINANCE', 'CULTURE', 'TECH', 'SPORTS', 'CRYPTO']

# GitHub config
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
GITHUB_REPO  = os.getenv("GITHUB_REPO")

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'application/json',
})

def log(msg):
    print(f'[{datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")}] {msg}', flush=True)

# ═══════════════════════════════════════════════════════════
# STATE RECOVERY & GITHUB UPLOAD
# ═══════════════════════════════════════════════════════════

def save_state(paper_balance, open_positions, closed_trades):
    state = {
        'paper_balance': paper_balance,
        'open_positions': open_positions,
        'closed_trades': closed_trades
    }
    try:
        with open(STATE_FILE, 'w') as f:
            json.dump(state, f)
    except Exception as e:
        log(f'⚠️ Failed to save state locally: {e}')

def load_state():
    if GITHUB_TOKEN and GITHUB_REPO:
        log("🌐 Checking GitHub for previous memory state...")
        # Add cache-buster to prevent GitHub from serving a stale file
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{STATE_FILE}?t={int(time.time())}"
        headers = {
            "Authorization": f"token {GITHUB_TOKEN}",
            "Accept": "application/vnd.github.v3+json"
        }
        try:
            res = requests.get(url, headers=headers)
            if res.status_code == 200:
                content = base64.b64decode(res.json().get("content", "")).decode('utf-8')
                state = json.loads(content)
                # Save a local copy immediately
                with open(STATE_FILE, 'w') as f:
                    json.dump(state, f)
                log(f"☁️ Successfully downloaded memory state! Restored {len(state.get('open_positions',{}))} open, {len(state.get('closed_trades',[]))} closed trades.")
                return state
            elif res.status_code == 404:
                log("ℹ️ No previous memory found on GitHub (404). Starting completely fresh.")
                return None
            else:
                # If we get a 403 Rate Limit or 500 error, DO NOT wipe the state.
                log(f"🚨 CRITICAL: GitHub API returned {res.status_code}: {res.text}")
                log("🚨 Halting to prevent overwriting your Excel data. Retrying in 60s...")
                time.sleep(60)
                return load_state() # Recursive retry
        except Exception as e:
            log(f"🚨 CRITICAL: Network error while fetching state: {e}")
            log("🚨 Retrying in 60s...")
            time.sleep(60)
            return load_state()

    # Local fallback only if GitHub env variables are missing
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r') as f:
                state = json.load(f)
            log(f"💾 Memory recovered locally! Restored {len(state.get('open_positions',{}))} open, {len(state.get('closed_trades',[]))} closed trades.")
            return state
        except Exception as e:
            log(f"⚠️ Could not read local state file: {e}")
    return None

def upload_to_github(file_path, commit_msg="Hourly Sync"):
    if not GITHUB_TOKEN or not GITHUB_REPO:
        return

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{file_path}"
    try:
        with open(file_path, "rb") as f:
            content = base64.b64encode(f.read()).decode()

        headers = {
            "Authorization": f"token {GITHUB_TOKEN}",
            "Accept": "application/vnd.github.v3+json"
        }

        res = requests.get(url, headers=headers)
        sha = res.json().get("sha") if res.status_code == 200 else None

        data = {
            "message": commit_msg,
            "content": content,
            "branch": "main"
        }
        if sha:
            data["sha"] = sha

        res = requests.put(url, headers=headers, json=data)

        if res.status_code in [200, 201]:
            log(f"☁️ {file_path} successfully synced to GitHub")
        else:
            log(f"❌ GitHub upload failed for {file_path}: {res.text}")
    except Exception as e:
        log(f"❌ Upload error for {file_path}: {e}")

# ═══════════════════════════════════════════════════════════
# LEADERBOARD
# ═══════════════════════════════════════════════════════════

def fetch_top_wallets(top_n=TOP_N):
    wallet_map = {}
    for cat in CATEGORIES:
        try:
            res = session.get(
                f'{DATA_API}/leaderboard',
                params={'timePeriod': 'MONTH', 'limit': top_n, 'orderBy': 'PNL', 'category': cat},
                timeout=15
            )
            if res.status_code != 200:
                continue
            traders = res.json()
            if not isinstance(traders, list):
                continue
            for t in traders:
                w = t.get('proxyWallet', '')
                if not w or not w.startswith('0x'):
                    continue
                pnl = float(t.get('pnl', 0))
                if w not in wallet_map:
                    wallet_map[w] = {
                        'username':   t.get('userName') or w[:10],
                        'wallet':     w,
                        'categories': [cat],
                        'pnl':        pnl,
                    }
                else:
                    if cat not in wallet_map[w]['categories']:
                        wallet_map[w]['categories'].append(cat)
                    wallet_map[w]['pnl'] = max(wallet_map[w]['pnl'], pnl)
            time.sleep(0.3)
        except Exception as e:
            log(f'⚠️ Leaderboard error ({cat}): {e}')

    for w in wallet_map:
        wallet_map[w]['category'] = '+'.join(wallet_map[w]['categories'])
    return wallet_map

def build_watch_list(live_wallets, open_positions, seen_tx):
    combined = dict(live_wallets)
    locked_wallets = {pos['wallet'] for pos in open_positions.values()}
    for wallet in locked_wallets:
        if wallet not in combined:
            pos = next(p for p in open_positions.values() if p['wallet'] == wallet)
            combined[wallet] = {
                'username': pos['username'],
                'wallet':   wallet,
                'category': pos['category'],
                'pnl':      0,
                'locked':   True,
            }
    for wallet in combined:
        seen_tx.setdefault(wallet, set())
    return combined

# ═══════════════════════════════════════════════════════════
# API HELPERS
# ═══════════════════════════════════════════════════════════

def get_activity(wallet, limit=20):
    try:
        res = session.get(f'{DATA_API}/activity',
