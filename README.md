# Plan Year Outlook (LF): a concept report for Angle Health's Business portal

**Live demo:** https://mohith26.github.io/angle-health-plan-year-outlook/

Unofficial concept demo by Mohith Gajjela. Not affiliated with Angle Health. All data is synthetic.

## What it is

A new report that would sit next to *Funding Summary (LF)*, *Large Claimants* and the *Top 20 Medications* reports in the Reporting tab of the Business portal (the view a broker or employer sees for a group). It answers the question an employer on a level-funded plan actually asks mid-year: **"Are we going to get a surplus back, and how sure are you?"**

The existing Funding Summary groups paid claims by incurred month and notes that "recent months' claims may increase as more claims are received." That caveat is the whole problem: in the most recent months only a third to two thirds of the claims have arrived, so a group that looks $28K under its claims fund in October can be tracking over it once the missing claims land. The outlook estimates those claims, projects the rest of the year and gives a calibrated probability.

For each group the page shows:

- chance of a claims fund surplus, expected surplus at settlement, projected plan-year claims and claims not yet received
- a plain-English callout that reconciles paid-to-date with the projection, and how a naive run-rate would have done
- claims by incurred month (paid / not yet received / projected) against the monthly benefit expense allocation
- the calibrated distribution of the fund balance at settlement, with the 80% range and break-even
- a build-up from paid claims to projected claims, and the large-claimant table with each member's room left under the specific deductible
- a monthly detail table in the same format as the existing reports
- how the outlook is calibrated, including held-out coverage by month of the plan year

A slider rewinds the outlook to any earlier month of the plan year (everything is recomputed from the claims that had been paid by then), and the group selector switches between five groups at different points in their plan years, including one in run-out.

## How it works

`src/outlook.js` is the whole model (no dependencies). It runs in Node for tests and in the browser for the page.

1. **Development (chain-ladder).** Volume-weighted age-to-age factors on retained paid dollars across a book of completed group-years give, for an incurred month seen at lag *d*, the share of its contract-ultimate (12 incurred / 24 paid) already paid.
2. **Credibility (Buhlmann-Straub).** The group's completed experience ratio (paid ÷ expected-reported) is blended with its underwriting. *k* = EPV/VHM is fitted on the book; the weight *Z* = w/(w+k) uses reported member-months, so immature months count only partially.
3. **Outstanding claims (Bornhuetter-Ferguson exposure + Monte Carlo).** Unreported exposure (Σ(1 − completion)) and the remaining months are simulated per member as compound Poisson claims with a spliced severity (lognormal body below $10K, Pareto tail above it, Hill estimator), with the group's risk level drawn from its credibility posterior. Each member's simulated claims are capped at their remaining room under the $30K specific deductible, so a member who has already met it adds nothing to the fund.
4. **Calibration (split conformal + PIT recalibration).** For every as-of month, the simulator's probability integral transform of the settled amount is computed on a calibration fold of completed group-years. The conformal quantile of |2u − 1| sets which simulation percentiles to report as the 80% range (a finite-sample coverage guarantee under exchangeability), and the empirical CDF of the calibration PITs recalibrates the surplus probability.

`scripts/build.mjs` generates a synthetic book of 1,200 group-years (`src/synth.js`), fits on 400, calibrates on 400 and evaluates on 400 held-out group-years, then writes `data/model.json` and five current groups to `data/groups.json`. The synthetic truth is deliberately richer than the model (within-year risk drift, seasonality, separate Rx and medical lags, multi-month catastrophic episodes with a heavier tail), so the simulator alone is overconfident.

Held-out results (from `data/model.json`, pooled over as-of months 2 to 13):

| | Result |
|---|---|
| Coverage of the simulator's 10th-90th percentile | 74.6% |
| Coverage of the calibrated 80% range | 80.4% |
| Mean absolute error of projected claims, % of fund | 12.4% (outlook) vs 24.3% (paid-claims run-rate) |
| Brier score for "surplus at settlement" | 0.120 raw vs 0.120 recalibrated (no gain; reported as-is) |

## Run it

```
node scripts/build.mjs        # regenerate data (about 35 s)
node --test test/             # 10 tests
open index.html               # or any static server
```

The page is static HTML/CSS/JS with Chart.js from a CDN. `dist/outlook.bundle.js` and `data/*.js` are classic-script copies written by the build script so the page also works from `file://`.

## Tests

`test/outlook.test.mjs` covers the random generators, reproducibility and paid-order specific stop-loss in the synthetic book, chain-ladder monotonicity and agreement with the book's empirical lag, Buhlmann-Straub behaviour (Z grows with exposure, the estimate lies between experience and prior), additivity of the decomposition, that a member at the deductible adds no outstanding retained claims, convergence to the settled amount, split-conformal recovering 80% coverage for a deliberately overconfident forecaster, recalibration monotonicity, and that the committed data only contains payments made by the as-of date.

## Honest limitations

- All data is synthetic. Real lag patterns, severity, plan design (deductibles, Rx carve-outs, capitation) and membership changes would need Angle's data; the structure is what transfers.
- Membership is held constant within the plan year, and the claims fund allocation is spread evenly over 12 months.
- One pooled development pattern; in practice it should be segmented (network, Rx vs medical, claim type) and refreshed as payment speed changes.
- The conformal guarantee is marginal over group-years and assumes the calibration years are exchangeable with the current one; a shift in the book (new networks, a trend spike) breaks that until the calibration fold catches up. Coverage is reported per month of the plan year rather than per group size.
- Surplus is the claims fund balance at settlement. The refund actually paid depends on each group's contract (sharing percentage, conditions) and is not modelled.
- Aggregate stop-loss is assumed to attach at 100% of the claims fund, and the specific deductible is $30K for every member.

## Novelty note

I checked Angle's public site, broker resources, Q4 2026 product update webinar, the Feb 2026 Homebase fireside chat and the public employer-portal bundle. The Business portal today offers Funding Summary (LF), Large Claimants and Top 20 Medications reports plus Explo dashboards (total and monthly claims by member, diagnosis summary, stop-loss eligible members). None of them estimate claims not yet received or forecast the year-end surplus; renewal analysis is prepared by Angle's team about 75 days before renewal. I did not find a level-funded carrier that publishes a calibrated in-portal surplus forecast; several provide interim claims reports and preliminary balances.
