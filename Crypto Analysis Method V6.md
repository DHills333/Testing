# CRYPTO ANALYSIS METHOD (V6)

*Version 6, 7 October 2026. Replaces the Lean Evaluation Guide (V5.2) and the Lean Evaluation Procedure (V5.2): the rules are unchanged, now written once. What it produces for each coin: one score, one rating, one tier and a one-page report. No buy or sell calls. This is an analytical framework, not financial advice; ratings do not predict price.*

---

## 1. The rules that never change

1. **Retrieve, then score.** Every grade comes from a number or a yes/no from a named source. If it cannot be retrieved, leave the line **blank**. Never guess and never give a "neutral" grade.
2. **Blank is not zero.** A blank line is left out of the score and lowers coverage. "Not applicable" (NA) lines are removed from the denominator.
3. **Date everything.** Every fact is written as *value | source | date*. Depth, sentiment and prices are one-day snapshots.
4. **Say plainly what could not be found, and why.** A gap listed honestly is better than a filled gap.
5. Do not use MCAP/TVL as a metric (user decision). Do not use the old V4.6 form.
6. **Grades:** A = 100% of a line's points, B = 75%, C = 50%, D = 25%, E = 0%.
7. **Score** = earned points ÷ scored points × 100. **Rating** = score ÷ 10. **Coverage** = scored points ÷ applicable points.
8. **Tiers:** Exceptional 90–100 | Strong 75–89.9 | Average/Decent 60–74.9 | Weak below 60. **Below 70% coverage no tier is printed** (the score is provisional).
9. Every ruling the user makes is written on the **Decisions** sheet, so it is never asked twice.

---

## 2. The files

| File | What it is |
|---|---|
| `Crypto Analysis V<version>.xlsx` | **The workbook, the single source of truth.** Sheets: **Grades** (the live scorer, where grades are entered), **Evidence** (value \| source \| date for every coin and line), **Coin notes** (what it does, biggest risk, open point, caps applied), **Usage method** (line 6 workings), **Decisions**, **Change log**, **Settings** (points, cut-offs, group rules, tier thresholds, market-share basket). |
| `Crypto Analysis Method V<version>.md` | This document. |
| `Crypto Analysis How I Use It V<version>.docx` | The user's own quick guide. |
| `crypto_analysis_summary.py` | Builds the one-page summary of all coins from the workbook. |
| `crypto_analysis_report.py` | Builds a coin's report from the workbook. |

**Naming:** every change gets a new version number. A major change (new features; new, moved or removed sheets or columns; changed calculations or rules) raises the first number (V6 → V7). A minor change (a coin added or refreshed, a grade changed, wording, small fixes) raises the decimal (V6 → V6.1, or V6.0.1 for something tiny). Never overwrite an earlier version. Update the version line and version log on the workbook's **Read me** sheet. Keep only the newest workbook in the Project so the scripts can never pick the wrong one (they take the highest version number).

---

## 3. The fifteen lines (100 points)

The **Settings** sheet holds the same table; formulas read their points from it.

| # | Line | Pts | Cut-offs | Source |
|---|---|---|---|---|
| 1 | Supply dilution (MC/FDV) | 10 | A ≥0.80, B 0.60–0.79, C 0.40–0.59, D 0.20–0.39, E <0.20 | CoinGecko, TradingView |
| 2 | New supply in next 90 days (unlocks + issuance, % of circulating) | 5 | A ≤1%, B >1–3, C >3–5, D >5–10, E >10 | Tokenomist, Tokenomics.com, DropsTab |
| 3 | Insider share (team, investors, foundation) | 5 | A ≤15% or fair launch, B >15–25, C >25–40, D >40–50, E >50 | Tokenomist, Tokenomics.com |
| 4 | Holder value capture (fees burned, bought back or paid out) | 5 | A ≥50%, B 25–49, C 10–24, D 1–9, E 0; capped at C if fees < $1M a month | DefiLlama |
| 5 | US exit liquidity: 2% depth on Coinbase, Kraken, Binance.US ÷ $100K | 10 | A ≥10×, B 5–<10, C 3–<5, D 1–<3, E <1 | Exchange APIs |
| 6 | Usage trend (90-day) | 10 | A ≥+10%, B 0–10, C −10–0, D −30 to −10, E < −30 (30-day cut-offs if no 90-day: +5/0/−5/−15) | Section 5 |
| 7 | Social sentiment | 5 | Mean of two steps, a half rounds down: sentiment 85/80/75/70 and Galaxy Score 70/60/50/40 | TradingView (CoinGecko votes only if flagged) |
| 8 | Incident record, last 24 months | 5 | A none found, B ≤$1M remediated, C $1–10M remediated, D >$10M remediated, E unremediated | DefiLlama hacks, news |
| 9 | US regulatory status | 10 | A ETF-approved or commodity, no action; B regulated issuer or clear; C unclear, no action; D restricted; E active enforcement | Coinbase tags, SEC/CFTC |
| 10 | Admin control | 5 | A immutable or no admin path; B timelocked multisig or DAO; C multisig without timelock or disclosed freeze; D single key or foundation; E unverified mint/upgrade | Docs, DefiLlama |
| 11 | Purpose: documented use beyond trading in a live product | 10 | Yes = A, No = E, blank if no source | DefiLlama fees/TVL, CoinGecko, project docs |
| 12 | Backers: total raised from outside investors | 5 | A ≥$100M, B $25–<100M, C $5–<25M, D >$0–<5M; NA where no raise recorded | DropsTab, Tokenomist |
| 13 | Adoption: production use or holding by a government, central bank or major company | 5 | Yes = A, No = E (searched, none found), blank if not researched | News, SEC filings |
| 14 | Size: market cap | 5 | A ≥$100B, B $10–<100B, C $1–<10B, D $0.1–<1B, E <$0.1B | CoinGecko |
| 15 | Market-share trend: 90-day change in dominance vs a top-20 basket | 5 | Same cut-offs as line 6; NA for USDT, USDC, PAXG | CoinGecko history |

### What to retrieve, and what to watch for

| # | How | Watch for |
|---|---|---|
| 1 | Market cap ÷ FDV from CoinGecko and TradingView. | If the sources differ because one uses total and one max supply, use TradingView (FDV on max supply) and write both in the evidence (AVAX: 0.62 vs 0.94). Note a value within 0.01 of a cut-off. |
| 2 | Unlocks plus issuance as % of circulating. | No forward schedule: use realized 90-day circulating-supply growth (CoinGecko history), labelled "proxy". Conflicting boilerplate ("schedule ended") is not evidence. |
| 3 | Team + investors + foundation as % of supply. | Write down what counts (NIGHT: Foundation only, 35%; HBAR: investors, team and council but not treasury pools). Fair launch = A. |
| 4 | DefiLlama holders revenue ÷ fees, 30 days (`dailyHoldersRevenue`, `dailyFees`). | Capped at C under $1M fees a month. Under 1% but above 0 = E (HBAR). |
| 5 | Bid-side depth within 2% of mid on Coinbase, Kraken and Binance.US (USD and USDT pairs), summed, ÷ $100,000. | One-day snapshot; date it. Request the deepest book each venue allows (Kraken: 500 levels). If the 2% mark is not reached, record the total as a lower bound. |
| 6 | Section 5. | Lower of two measures; † if within 5 points of a cut-off. |
| 7 | TradingView `sentiment` and `galaxyscore`. | If either is missing, leave blank. CoinGecko votes may stand in only if flagged. |
| 8 | DefiLlama `/hacks` for the last 24 months, matched by **name and chain**, then checked against news. | Name matches give false positives ("Bitcoin Mission" is not Bitcoin). Score the coin's **own direct loss**. Losses at third-party wallets, exchanges, bridges or apps on a chain are listed in the evidence for context but not graded. |
| 9 | ETFs, commodity or security treatment, SEC/CFTC/DOJ action. | A general SEC framework statement does not lift a coin above C. Say "no enforcement found" only after searching. |
| 10 | Docs and explorers: keys, upgrades, freezes, timelocks. | Blank if the project's own documents do not say. |
| 11 | Yes or No with a named source. | DOGE, SHIB and PEPE are No (user decision). |
| 12 | DropsTab, Tokenomist raise totals. | Company-level rounds count (user decision). A predecessor token's raise (SKY from MKR) is an assumption: flag it. |
| 13 | Documented production use or holding. | Pilots, MoUs, announcements, ETFs, small-company treasury buys and node operators do not count. |
| 14 | CoinGecko market cap. | Scored for all coins, including stablecoins. |
| 15 | Section 6. | |

---

## 4. Groups

| Group | Examples | Lines that are NA |
|---|---|---|
| Chain / protocol | ETH, SOL, AAVE, CAKE | none |
| Money coin | BTC, LTC, XMR, ZEC | L4 |
| Memecoin | DOGE, SHIB, PEPE | L4; **tier held at Average/Decent automatically** |
| Stablecoin / RWA | USDC, USDT, PAXG | L1, L2, L3, L4, L12, L15 |

A fair launch or a coin with no recorded raise also has **L12 = NA**. NA is typed in the grade cell; the **Group check** column on Grades says "Group NA missing" if a line the group requires as NA is not NA.

---

## 5. Usage trend (line 6)

Final grade = the **lower** of:
- **Price-free measure:** stablecoin supply on the chain (stablecoins.llama.fi), 90-day change; for apps and anything without a stablecoin series, 30-day average fees (USD) compared with the same average 90 days earlier.
- **Native-token measure:** the same activity priced in the coin's own token: TVL ÷ token price for chains and apps (chain TVL from DefiLlama's chain series, which excludes double counting and liquid staking; app TVL with liquid-staking and receipt tokens removed); fees ÷ token price for money coins, ONDO and RENDER.

Mark "†" (Borderline = yes on the Usage method sheet) if either value is within 5 points of a cut-off. UNI and RAY have no complete token breakdown (unadjusted result). DOGE has no fee series (earlier grade kept). No usage series: PAXG, PEPE, SAND, FET, QNT, NIGHT, RED, POWR (line blank).

---

## 6. Market-share trend (line 15)

1. 90-day change in the coin's market cap (CoinGecko daily history).
2. Relative change = (1 + coin change) ÷ (1 + basket change) − 1, where the basket is the top-20 market-cap list on the **Settings** sheet.
3. Grade with the line 6 cut-offs.
4. The basket's 90-day change was +33.5% on 2026-10-05 (a later re-computation gave +32.7%). **Recompute it for any other date** and write the figure and date in the evidence.

---

## 7. Caps and flags (only if the fact was retrieved)

Entered in the **Cap** column on Grades:
- **blank:** none. Memecoins need no entry (automatic).
- **3, a flag:** the tier drops at most one step below what the score gives, and never below Average/Decent from the flag alone. Flags: active SEC/CFTC/DOJ enforcement; material incident in the last 90 days, remediated, loss over $1M (the 90 days run from remediation); team-controlled mint or upgrade key with no timelock; top-10 non-exchange wallets over 60%; stablecoin depeg over 5% for over 24 hours; privacy coin banned or delisted. Remove the flag when it no longer applies (for example, 90 days after remediation) and log it.
- **4, Cap 4 (Weak):** unresolved exploit, unbacked minting, or stablecoin/RWA reserves shown not to be independently verified.

---

## 8. Evaluating a coin, step by step

1. **Universe.** Confirm the coin trades spot on Coinbase (including its in-app DEX), Kraken or Binance.US. If not, stop. Record the venues.
2. **Group.** Pick the group (section 4).
3. **Retrieve all 15 lines** (section 3), each as value | source | date. For every blank line, write what was tried and why it failed.
4. **Show the user** the proposed grades, every judgment call and every blank **before changing any file**. Wait for approval.
5. **Enter it in the workbook:**
   - **Grades:** one row per coin. New coin: copy the last row down (formulas come with it), then fill ticker, name, group, cap, the 15 grades, Last evaluated and CoinGecko ID in the yellow cells.
   - **Evidence:** one row per line (15 rows).
   - **Coin notes:** what it does, biggest risk, open point, caps applied.
   - **Usage method:** one row (both measures, both grades, final, Borderline). The † marks come from this sheet.
   - **Decisions:** any new ruling.
   - **Change log:** one row per change (date, coin, line, old grade, new grade, value, source).
   - **Read me:** version line and version log.
6. **Recalculate** (open in Excel, or have Claude recalculate with LibreOffice) and check: no errors, Group check = OK for every coin.
7. **Read the result:** coverage first (under 70% = not rated); fragility (within about 2 points of a tier line, or coverage close to 70%); what rests on judgment calls.
8. **Report:** run `crypto_analysis_report.py TICKER`. Reports are always regenerated, never edited by hand. New findings go in the **Change log**, never in a report.
9. Save the workbook under the next version number (section 2).

**Updating a coin later:** change the grade on Grades, update its Evidence row and Last evaluated date, add a Change log row, recalculate, regenerate the report and, if wanted, the summary.

**One-page summary:** `python3 crypto_analysis_summary.py [WORKBOOK] [--out FILE.docx] [--date YYYY-MM-DD]`. It retrieves no data and changes no file. With no workbook named, it uses the highest-version `Crypto Analysis V*.xlsx` and warns if there is more than one. Always state which workbook was used.

### Sources that worked (as of 2026-10-05)

| Need | Source |
|---|---|
| Spot listings and order books | api.exchange.coinbase.com, api.kraken.com, api.binance.us |
| Market cap, FDV, supply, price history | api.coingecko.com (IDs in the CoinGecko ID column on Grades) |
| Sentiment %, Galaxy Score, FDV on max supply | TradingView connector (`sentiment`, `galaxyscore`, `market_cap_diluted_calc`) |
| TVL, fees, holders revenue, hacks | api.llama.fi (`/protocols`, `/protocol/{slug}`, `/v2/historicalChainTvl/{Chain}`, `/overview/fees/{chain}`, `/summary/fees/{protocol}`, `/hacks`) |
| Stablecoin supply on a chain | stablecoins.llama.fi (`/stablecoincharts/{Chain}`) |
| Backers and vesting | dropstab.com (`/coins/{slug}/fundraising`, `/vesting`), tokenomist.ai, Tokenomics.com |
| ETFs, regulation, adoption, incidents | web search, SEC and exchange notices, project docs |

**Do not rely on:** cryptorank.io (blocked), news.dropstab.com (blocked), DefiLlama paid endpoints (`/raises`, `/oracles`, `/emissions`), DexScreener for exit liquidity (look-alike tokens). If a source is blocked, say so and move on.

---

## 9. Worked example

AVAX, 2026-10-05. Grades: L1 B, L2 B, L3 C, L4 B, L5 A, L6 D, L7 C, L8 A, L9 A, L10 blank, L11 A, L12 A, L13 A, L14 C, L15 A.
- Earned = 7.5 + 3.75 + 2.5 + 3.75 + 10 + 2.5 + 2.5 + 5 + 10 + 10 + 5 + 5 + 2.5 + 5 = **75.0**
- Scored = 100 − 5 (L10 blank) = **95**; applicable = 100
- Score = 75 ÷ 95 × 100 = **78.9**; rating 7.9; coverage 95%; tier **Strong**

---

## 10. Checklist

- [ ] Trades spot on Coinbase, Kraken or Binance.US
- [ ] Group chosen; NA lines set (Group check = OK)
- [ ] 15 lines retrieved or left blank with reason; each as value | source | date
- [ ] Judgment calls shown to the user and written on Decisions
- [ ] Hacks matched by name **and** chain, then checked against news
- [ ] Usage: both measures, lower taken, † if borderline
- [ ] Market-share basket recomputed for today's date
- [ ] Caps and flags only for retrieved facts
- [ ] Recalculated, no errors; coverage ≥ 70% or "Not rated"
- [ ] Fragility noted
- [ ] Evidence, Coin notes, Usage method, Change log, Read me updated; report regenerated; new version saved

---

## 11. Known limits (state these when sharing a result)

- Depth, sentiment and prices are one-day snapshots.
- Hack records are matched by name and chain, so smaller incidents can be missed (re-checked against DefiLlama on 2026-10-05).
- Many adoption calls rest on crypto media or the project's own announcements; line 13 is blank for many coins.
- Where no forward unlock schedule exists, line 2 uses realized supply growth (a proxy).
- Market-share trend separates little in a rising market and includes newly unlocked supply. The basket includes stablecoins (USDT, USDC, USDS), which dampen its change.
- Native-token TVL is a worst case for chains (it treats all locked value as the native token).
- A blank line is never scored, so a coin with many blanks can score higher than it would with full data. Always read coverage next to the score; the report shows the range the score would fall in if the blanks were filled.
- Differentiation and competitors are not scored (judgement).
- Ratings describe the quality of retrievable facts, **not** future price.

---

## 12. Starting a new chat: the ready prompt

> **"Using the Crypto Analysis V6 method only, evaluate [COIN]. Use web search and your tools to retrieve every line yourself; do not answer from memory. Confirm it trades spot on Coinbase, Kraken or Binance.US, pick the group, then show me all 15 lines with value, source, date and proposed grade, every judgment call, and every line you could not retrieve (with what you tried). Do not change any file until I approve. Then add the coin to the workbook, save it as the next version, and run the report script."**

Useful follow-ups: "Re-check Line 8 for [COIN] against DefiLlama hacks." "What would change [COIN]'s tier if Line X moved one grade?" "Regenerate the report for [COIN]." "Summary."
