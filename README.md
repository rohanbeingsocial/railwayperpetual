# Polymarket Wallet Tracking & Slippage Analysis

## Overview

This project is a real-time Polymarket wallet tracking system built to evaluate whether copying top-performing wallets can be a profitable strategy in practice.

The idea is based on a simple assumption:

> If all wallet activity on Polymarket is public, we should be able to track top wallets, replicate their trades, and achieve similar profits.

The system tracks 100+ wallets, captures their entry and exit activity, and logs trades into Excel for analysis.

---

## Hypothesis

Top wallets ("smart money") consistently generate profits.

Therefore:

- Copying their trades should also be profitable  
- Even with slippage, profits should still exist (though reduced)

---

## Core Problem: Slippage

Copy trading introduces unavoidable delays:

1. Delay in detecting wallet activity  
2. Delay in executing trades  
3. Price movement during that time  

This results in real-world slippage.

---

## Experiment Design

To simulate realistic conditions, the system does NOT use wallet prices.

### Entry Logic

- Wallet buys → signal received  
- We enter at the current market price  
- Not the wallet’s entry price  

### Exit Logic

- Wallet sells → signal received  
- We exit at the current market price  

> This ensures results reflect actual executable trades.

---

## System Architecture

**Backend:** Python  
**Deployment:** Railway (persistent runtime)  
**Storage:** Excel  

### Tracking Scope

- 100+ wallets  
- Multiple markets within selected categories  
- Timestamped entry and exit events  

### Key Files

- `main_copywallets.py` → Core logic  
- `copytradesfinal.xlsx` → Final results  
- `tracker_state.json` → State persistence  

---

## Results

After running the system:

- Wallets combined profit: **~$50,000**  
- Simulated copy trading profit: **~$0.47**

---

## Key Observation

In most cases:

> Our entry price ≈ exit price

Meaning:

- By the time we enter, the price has already moved  
- By the time we exit, the opportunity is gone  

---

## Validation

- Script reviewed multiple times  
- Logic verified  
- A few trades behaved correctly  
- Majority showed the same pattern  

> Conclusion: The system is functioning correctly — this is not a bug.

---

## Insight

During analysis:

- Many top wallets appear to be **market makers**
- They:
  - Enter early  
  - Provide liquidity  
  - Capture spreads  

So:

> By the time trades are visible publicly, the edge is already gone.

---

## Final Conclusion

Copy trading top Polymarket wallets is not viable.

### Reasons

- Slippage removes profit  
- Public signals are delayed  
- Market makers operate differently from retail traders  

---

## Future Direction

A better strategy would be:

- Identify non-market-maker wallets  
- Focus on:
  - Consistent directional traders  
  - High conviction positions  
- Track a curated set instead of top-volume wallets  

---

## Deployment

Railway Deployment:

https://brilliant-ambition-production.up.railway.app

---

## Data

Results are stored in:

`copytradesfinal.xlsx`

---

## Key Takeaway

> Transparency ≠ Profitability

Just because data is public does not mean it is exploitable.
