# Semiconductor Where-to-Play Map for an Industrial-Automation Company

Chips were an $800bn market in 2025, and governments are paying for new fabs. This model asks **which parts of the value chain, in which countries, an automation company should go after**: controls, safety, gas detection, building automation and software, the things Honeywell Technologies actually sells. It sizes the value chain with 2025 data, then scores 11 segment-by-country plays on attractiveness vs right to win, with a deep-dive on India's 12 approved projects.

Built by Pavitra Seth (BITS Pilani) as preparation for Honeywell PS-II. It uses public data only and is not affiliated with Honeywell.

**Files**
- `semiconductor_where_to_play.xlsx`: tabs README, Capabilities, Segments, Geographies, Plays, Matrix (bubble chart), India, Sensitivity, Sources. It has 259 live formulas and no errors.
- `build.py`: rebuilds the workbook (`python build.py`).

## The value chain in 2025
| Segment | 2025 size | Growth | Source |
|---|---|---|---|
| Chips (context) | $795.6bn | +26.2% | WSTS |
| Equipment (tools) | $135.1bn | +15% | SEMI |
| Wafer fab materials | $45.8bn | +5.4% | SEMI |
| Packaging materials | $27.4bn | +9.3% | SEMI |
| Fab building and facility systems | ~$45bn | follows tool spend | Derived: tools are 70–80% of fab cost |
| **of which automation, controls and safety** | **~$4.5bn** | | 10% of facility spend (assumption) |

**First point:** an automation company is not competing for the $135bn tool market. Its realistic pool is about **$4.5bn a year**, worth winning but small. Honeywell's materials business went with the Solstice spin-off in October 2025, so materials is no longer its game either.

## Results
| Rank | Play | Attractiveness | Right to win | Verdict |
|---|---|---|---|---|
| 1 | Fab facility automation & safety: **US** | 4.10 | 4.25 | Invest |
| 2 | Fab facility automation & safety: **Taiwan** | 4.55 | 3.80 | Invest |
| 3 | Fab facility automation & safety: **India** | 3.65 | 4.20 | Invest |
| 4 | Fab facility automation & safety: Japan | 3.70 | 3.80 | Invest |
| 5 | ATMP / OSAT facility automation: India | 3.65 | 3.80 | Invest |
| 6 | Fab facility automation & safety: Korea | 3.85 | 3.40 | Partner |
| 8 | Fab facility automation & safety: China | 2.80 | 3.25 | Avoid |
| 10–11 | Materials (Taiwan), equipment (Korea) | 3.5–4.3 | 1.0–1.4 | Partner or buy only |

The top three stay in the top four under all three weighting scenarios (Sensitivity tab).

## India deep-dive
- **The projects:** 12 approved projects with ₹1.64 lakh crore of investment. One silicon fab (Tata–PSMC, Dholera, ₹91,526 Cr), two compound-semiconductor fabs, and nine packaging (ATMP/OSAT) units.
- **The pool:** applying facility shares (25% for fabs, 35% for packaging) and a 10% automation share gives an **automation and safety pool of about ₹4,775 Cr (~$540m)** over the 2025–28 build.
- **Where it sits:** 75% of the pool is in Gujarat.
- **First three customers:** the Tata–PSMC fab, the Micron Sanand ATMP and the Tata Assam OSAT.

## Three takeaways
1. **Play the building, not the tools.** Cleanroom controls, toxic-gas detection, safety systems and utility automation are where Honeywell's portfolio fits fabs. That pool grows with fab construction (+15% in 2025).
2. **The US and Taiwan are the scale plays; India is the share play.** Taiwan's tool spending nearly doubled in 2025 (+90%), and the US has $30.9bn of CHIPS funding across 40 fab projects, though several have slipped. India's pool is small in dollars, but it is greenfield, has few incumbents, and Honeywell already has a listed local arm.
3. **Avoid China and Europe for now, for different reasons.** China is the biggest market (36% of tool spending) but carries export-control and localisation risk. Europe's tool spending fell 41% in 2025 as projects were delayed.

## Method
- **Segments** are sized from SEMI and WSTS releases. The facility pool = equipment ÷ equipment share × facility share.
- **Geographies:** the automation pool is split by each region's share of 2025 equipment billings.
- **Plays:**
  - Attractiveness = pool size 30%, growth 25%, policy tailwind 15%, access 15%, profit pool 15%.
  - Right to win = portfolio fit 40%, customer access 25%, installed base 20%, low competition 15%.
  - Scores are 1–5 with an evidence note each. The verdict is "Invest" when both scores are at least 3.5.
- **Robustness:** the Sensitivity tab re-ranks under growth-heavy and safety-first weights.

## Limitations
- The equipment share of fab cost (75%), the automation share of facility spend (10%) and the packaging facility share (35%) are assumptions.
- The India pool covers a multi-year build, while the other pools are annual.
- Play scores are analyst judgment.
- India project details come from a compiled list; check them against the India Semiconductor Mission site.
- Check Honeywell's fab-relevant products in its latest 10-K.

## Sources
WSTS, SEMI (equipment and materials), Quartz / Construction Physics (fab cost split), Congressional Research Service R49031 (CHIPS), IANS and Adda247 (India projects), the India Semiconductor Mission, and Honeywell results and spin-off releases. Full links are on the Sources tab.
