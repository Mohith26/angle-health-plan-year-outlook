// generated from src/outlook.js by scripts/build.mjs
(function(){
// Plan Year Outlook: core model for a level-funded claims fund.
//
// Given a group's claims paid so far in its plan year, estimate where the claims
// fund will land at settlement (12 months incurred / 24 months paid):
//   1. book-wide chain-ladder development pattern -> how complete each incurred month is
//   2. Buhlmann-Straub credibility -> blend the group's completed experience with underwriting
//   3. Bornhuetter-Ferguson style outstanding exposure + Monte Carlo of member-level claims,
//      capped at each member's remaining specific stop-loss deductible
//   4. split-conformal calibration of the predictive quantile levels on completed group-years,
//      so the 80% range holds 80% of outcomes even when the simulator is misspecified
//
// Pure ES module, no dependencies. Runs in the browser and in Node (node --test).

const ISL_DEFAULT = 30000;      // specific stop-loss deductible per member per plan year
const LARGE_THRESHOLD = 10000;  // severity splice point (also Angle's "Large Claimants" cutoff)
const PLAN_MONTHS = 12;
const PAID_MONTHS = 24;         // 12/24 contract

// ---------------------------------------------------------------- random numbers
function rng(seed) {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
function normal(r) {
  let u = 0;
  while (u === 0) u = r();
  return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * r());
}
function gamma(r, shape, scale) {
  if (shape < 1) return gamma(r, shape + 1, scale) * Math.pow(r() || 1e-12, 1 / shape);
  const d = shape - 1 / 3, c = 1 / Math.sqrt(9 * d);
  for (;;) {
    let x, v;
    do { x = normal(r); v = 1 + c * x; } while (v <= 0);
    v = v * v * v;
    const u = r();
    if (u < 1 - 0.0331 * x * x * x * x) return d * v * scale;
    if (Math.log(u) < 0.5 * x * x + d * (1 - v + Math.log(v))) return d * v * scale;
  }
}
function poisson(r, lam) {
  if (!(lam > 0)) return 0;
  if (lam > 40) return Math.max(0, Math.round(lam + Math.sqrt(lam) * normal(r)));
  const L = Math.exp(-lam);
  let k = 0, p = 1;
  do { k++; p *= r(); } while (p > L);
  return k - 1;
}
function normCdf(x) {
  // Abramowitz-Stegun 7.1.26 via erf
  const s = x < 0 ? -1 : 1, z = Math.abs(x) / Math.SQRT2;
  const t = 1 / (1 + 0.3275911 * z);
  const y = 1 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.exp(-z * z);
  return 0.5 * (1 + s * y);
}

// ---------------------------------------------------------------- helpers
function quantileSorted(xs, q) {
  if (!xs.length) return NaN;
  const pos = Math.min(Math.max(q, 0), 1) * (xs.length - 1);
  const lo = Math.floor(pos), hi = Math.ceil(pos);
  return xs[lo] + (xs[hi] - xs[lo]) * (pos - lo);
}
const sum = (a) => a.reduce((s, x) => s + x, 0);

// Observed state of a group-year at "as-of" a = number of paid months observed (1..24).
// payments: [memberIdx, incurredMonth, paidMonth, gross, retained] with months 0-based from plan start.
function observe(group, asOf) {
  const nInc = Math.min(asOf, PLAN_MONTHS);
  const paidRet = new Array(PLAN_MONTHS).fill(0);
  const paidGross = new Array(PLAN_MONTHS).fill(0);
  const memberRet = new Float64Array(group.members.length);
  const memberGross = new Float64Array(group.members.length);
  for (const [j, inc, pay, gross, ret] of group.payments) {
    if (pay > asOf - 1 || inc > nInc - 1) continue;
    paidRet[inc] += ret; paidGross[inc] += gross;
    memberRet[j] += ret; memberGross[j] += gross;
  }
  return { asOf, nInc, paidRet, paidGross, memberRet, memberGross };
}

function ultimateRetained(group) {
  let u = 0;
  for (const p of group.payments) if (p[2] <= PAID_MONTHS - 1) u += p[4];
  return u;
}

// ---------------------------------------------------------------- fitting on the book
// 1. chain-ladder age-to-age factors on retained paid dollars (volume weighted across the book)
function fitDevelopment(book) {
  const maxLag = PAID_MONTHS - 1;
  // cumulative paid C[g][i][k]
  const num = new Array(maxLag).fill(0), den = new Array(maxLag).fill(0);
  for (const g of book) {
    const cell = Array.from({ length: PLAN_MONTHS }, () => new Array(PAID_MONTHS).fill(0));
    for (const [, inc, pay, , ret] of g.payments) if (pay <= maxLag) cell[inc][pay - inc] += ret;
    for (let i = 0; i < PLAN_MONTHS; i++) {
      let cum = 0;
      const cumArr = cell[i].map((x) => (cum += x));
      const lastLag = maxLag - i; // contract window for this incurred month
      for (let k = 0; k < lastLag; k++) { num[k] += cumArr[k + 1]; den[k] += cumArr[k]; }
    }
  }
  const ata = num.map((n, k) => (den[k] > 0 ? n / den[k] : 1));
  return { ata };
}
// fraction of the contract-ultimate already paid for an incurred month seen at lag d with final lag L
function completion(dev, d, L) {
  if (d >= L) return 1;
  let f = 1;
  for (let k = Math.max(d, 0); k < L; k++) f *= dev.ata[k];
  return d < 0 ? 0 : 1 / f;
}

// 2. Buhlmann-Straub structural parameters on the ratio actual retained / expected retained
function fitCredibility(book) {
  let W = 0, WX = 0, epvNum = 0, epvDen = 0, sumW2 = 0;
  const rows = [];
  for (const g of book) {
    const ult = new Array(PLAN_MONTHS).fill(0);
    for (const [, inc, pay, , ret] of g.payments) if (pay <= PAID_MONTHS - 1) ult[inc] += ret;
    const eMonth = g.fund / PLAN_MONTHS;
    const wm = g.members.length; // member-months per month
    const xm = ult.map((u) => u / eMonth);
    const wg = wm * PLAN_MONTHS;
    const xg = sum(ult) / g.fund;
    for (const x of xm) epvNum += wm * (x - xg) ** 2;
    epvDen += PLAN_MONTHS - 1;
    rows.push([wg, xg]);
    W += wg; WX += wg * xg; sumW2 += wg * wg;
  }
  const m = WX / W;
  const epv = epvNum / epvDen;
  let between = 0;
  for (const [wg, xg] of rows) between += wg * (xg - m) ** 2;
  const vhm = Math.max((between - (rows.length - 1) * epv) / (W - sumW2 / W), 1e-4);
  return { m, epv, vhm, k: epv / vhm };
}

// 3. spliced severity: lognormal body (truncated at T) + Pareto tail above T (Hill estimator)
function fitSeverity(book, T = LARGE_THRESHOLD) {
  let nBody = 0, sL = 0, sL2 = 0, nTail = 0, sLogTail = 0, total = 0, n = 0;
  for (const g of book) for (const p of g.payments) {
    const x = p[3];
    if (!(x > 0)) continue;
    n++; total += x;
    if (x > T) { nTail++; sLogTail += Math.log(x / T); }
    else { nBody++; const l = Math.log(x); sL += l; sL2 += l * l; }
  }
  const mu = sL / nBody, sigma = Math.sqrt(Math.max(sL2 / nBody - mu * mu, 1e-6));
  const alpha = Math.max(nTail / sLogTail, 1.05);
  const pi = nTail / n;
  const lt = Math.log(T);
  const bodyMean = Math.exp(mu + sigma * sigma / 2) * normCdf((lt - mu - sigma * sigma) / sigma) / normCdf((lt - mu) / sigma);
  const tailMean = T * alpha / (alpha - 1);
  const cap = 2e6;
  return { mu, sigma, alpha, pi, T, cap, bodyMean, tailMean, mean: (1 - pi) * bodyMean + pi * tailMean, empiricalMean: total / n, nClaims: n };
}
function drawSeverity(r, sev) {
  if (r() < sev.pi) return Math.min(sev.T * Math.pow(r() || 1e-12, -1 / sev.alpha), sev.cap);
  for (let tries = 0; tries < 50; tries++) {
    const x = Math.exp(sev.mu + sev.sigma * normal(r));
    if (x <= sev.T) return x;
  }
  return sev.T;
}

function fitBook(book) {
  return { dev: fitDevelopment(book), cred: fitCredibility(book), sev: fitSeverity(book), nGroupYears: book.length };
}

// ---------------------------------------------------------------- forecast
function forecast(group, asOf, fit, opts = {}) {
  const sims = opts.sims ?? 2000;
  const r = rng(opts.seed ?? 12345);
  const isl = group.isl ?? ISL_DEFAULT;
  const obs = observe(group, asOf);
  const F = group.fund, eMonth = F / PLAN_MONTHS, nMem = group.members.length;
  const { dev, cred, sev } = fit;

  // completion of each observed incurred month
  const comp = [];
  for (let i = 0; i < obs.nInc; i++) comp.push(completion(dev, asOf - 1 - i, PAID_MONTHS - 1 - i));
  const paidToDate = sum(obs.paidRet);
  const reportedExposure = sum(comp);                 // months of expected claims that should be in by now
  const X = reportedExposure > 0 ? paidToDate / (reportedExposure * eMonth) : cred.m;
  const w = reportedExposure * nMem;                  // credibility weight in member-months
  const Z = w / (w + cred.k);
  const Rhat = Z * X + (1 - Z) * cred.m;
  const postVar = Math.max((1 - Z) * cred.vhm, 1e-6);
  const gShape = (Rhat * Rhat) / postVar, gScale = postVar / Rhat;

  const outstandingInc = sum(comp.map((c) => 1 - c)); // IBNR exposure in months
  const futureMonths = Math.max(0, PLAN_MONTHS - obs.nInc);

  // deterministic Bornhuetter-Ferguson view (for the table)
  const months = [];
  for (let i = 0; i < PLAN_MONTHS; i++) {
    if (i < obs.nInc) months.push({ i, paid: obs.paidRet[i], paidGross: obs.paidGross[i], completion: comp[i], ibnr: (1 - comp[i]) * Rhat * eMonth, projected: 0, fund: eMonth });
    else months.push({ i, paid: 0, paidGross: 0, completion: 0, ibnr: 0, projected: Rhat * eMonth, fund: eMonth });
  }

  const lam = group.members.map((m) => m.grossPmpm / sev.mean); // claims per member-month at R = 1
  const head = group.members.map((_, j) => Math.max(0, isl - obs.memberRet[j]));
  const U = new Float64Array(sims);
  let sIbnr = 0, sFut = 0, sReimb = 0;
  for (let s = 0; s < sims; s++) {
    const R = gamma(r, gShape, gScale);
    let ibnrRet = 0, futRet = 0, grossOut = 0;
    for (let j = 0; j < nMem; j++) {
      let room = head[j];
      const nI = poisson(r, lam[j] * R * outstandingInc);
      let gI = 0;
      for (let c = 0; c < nI; c++) gI += drawSeverity(r, sev);
      const nF = poisson(r, lam[j] * R * futureMonths);
      let gF = 0;
      for (let c = 0; c < nF; c++) gF += drawSeverity(r, sev);
      const rI = Math.min(gI, room); room -= rI;
      const rF = Math.min(gF, room);
      ibnrRet += rI; futRet += rF; grossOut += gI + gF;
    }
    sIbnr += ibnrRet; sFut += futRet; sReimb += grossOut - ibnrRet - futRet;
    U[s] = paidToDate + ibnrRet + futRet;
  }
  const sorted = Array.from(U).sort((a, b) => a - b);
  const mean = sum(sorted) / sims;
  const raw = {
    p10: quantileSorted(sorted, 0.1), p50: quantileSorted(sorted, 0.5), p90: quantileSorted(sorted, 0.9),
    pSurplus: cdfSorted(sorted, F),
    expSurplus: sum(sorted.map((u) => Math.max(0, F - u))) / sims,
  };
  return {
    groupId: group.groupId, asOf, fund: F, nMembers: nMem, paidToDate, paidGrossToDate: sum(obs.paidGross),
    reimbursedToDate: sum(obs.paidGross) - paidToDate,
    X, Z, Rhat, w, outstandingInc, futureMonths,
    ibnrMean: sIbnr / sims, futureMean: sFut / sims, islExpected: sReimb / sims,
    mean, samples: sorted, raw, months,
    memberRet: Array.from(obs.memberRet), memberGross: Array.from(obs.memberGross),
    naive: obs.nInc > 0 ? paidToDate * PLAN_MONTHS / obs.nInc : 0,
  };
}

function cdfSorted(sorted, x) {
  // P(U <= x) with linear interpolation between order statistics (randomization-free PIT)
  const n = sorted.length;
  if (x < sorted[0]) return 0.5 / (n + 1);
  if (x >= sorted[n - 1]) return 1 - 0.5 / (n + 1);
  let lo = 0, hi = n - 1;
  while (hi - lo > 1) { const mid = (lo + hi) >> 1; if (sorted[mid] <= x) lo = mid; else hi = mid; }
  const frac = sorted[hi] > sorted[lo] ? (x - sorted[lo]) / (sorted[hi] - sorted[lo]) : 0;
  return (lo + 1 + frac) / (n + 1);
}

// ---------------------------------------------------------------- conformal calibration
// For each as-of month: PIT values u = Fhat(U_true) on a calibration fold.
// 80% interval: conformal quantile of s = |2u - 1| gives the quantile levels to report.
// Probabilities: recalibrated through the empirical CDF of calibration PITs.
function conformalLevel(pits, coverage = 0.8) {
  const s = pits.map((u) => Math.abs(2 * u - 1)).sort((a, b) => a - b);
  const n = s.length;
  const idx = Math.min(n - 1, Math.ceil((n + 1) * coverage) - 1);
  return s[idx];
}
function makeCalibrator(pits, coverage = 0.8) {
  const sortedPits = [...pits].sort((a, b) => a - b);
  const q = conformalLevel(pits, coverage);
  return { n: pits.length, q, coverage, pits: sortedPits.map((x) => +x.toFixed(5)) };
}
function recalibrateProb(cal, u) {
  // G(u) = P(PIT <= u) estimated on calibration fold (smoothed)
  const p = cal.pits, n = p.length;
  let lo = 0, hi = n;
  while (lo < hi) { const mid = (lo + hi) >> 1; if (p[mid] <= u) lo = mid + 1; else hi = mid; }
  return (lo + 0.5) / (n + 1);
}
function applyCalibration(fc, cal) {
  const s = fc.samples, F = fc.fund;
  const lo = (1 - cal.q) / 2, hi = (1 + cal.q) / 2;
  const range = { lo: quantileSorted(s, lo), hi: quantileSorted(s, hi), levels: [lo, hi] };
  const pSurplus = recalibrateProb(cal, cdfSorted(s, F));
  // calibrated predictive draws: U' = Fhat^{-1}(G^{-1}(v)) on an even grid of v
  const M = s.length, draws = new Float64Array(M);
  for (let k = 0; k < M; k++) {
    const v = (k + 0.5) / M;
    const u = quantileSorted(cal.pits, v);
    draws[k] = quantileSorted(s, u);
  }
  let es = 0, mean = 0;
  for (const u of draws) { es += Math.max(0, F - u); mean += u; }
  return {
    range, pSurplus, expSurplus: es / M, mean: mean / M, draws,
    balanceRange: { lo: F - range.hi, hi: F - range.lo },
  };
}

// ---------------------------------------------------------------- evaluation helpers
function evaluateFold(results) {
  // results: [{pit, inRaw, inCal, pRaw, pCal, surplus (0/1), err, naiveErr}]
  const n = results.length;
  const mean = (f) => sum(results.map(f)) / n;
  const bins = [0, 0.2, 0.4, 0.6, 0.8, 1.0001];
  const rel = (key) => bins.slice(0, -1).map((b, k) => {
    const rs = results.filter((x) => x[key] >= b && x[key] < bins[k + 1]);
    return { lo: b, hi: Math.min(bins[k + 1], 1), n: rs.length, pred: rs.length ? sum(rs.map((x) => x[key])) / rs.length : null, obs: rs.length ? sum(rs.map((x) => x.surplus)) / rs.length : null };
  });
  return {
    n,
    coverageRaw: mean((x) => x.inRaw), coverageCal: mean((x) => x.inCal),
    brierRaw: mean((x) => (x.pRaw - x.surplus) ** 2), brierCal: mean((x) => (x.pCal - x.surplus) ** 2),
    maeModel: mean((x) => Math.abs(x.err)), maeNaive: mean((x) => Math.abs(x.naiveErr)),
    biasModel: mean((x) => x.err), biasNaive: mean((x) => x.naiveErr),
    reliabilityRaw: rel('pRaw'), reliabilityCal: rel('pCal'),
  };
}

window.Outlook = {ISL_DEFAULT, LARGE_THRESHOLD, PLAN_MONTHS, PAID_MONTHS, rng, normal, gamma, poisson, normCdf, quantileSorted, sum, observe, ultimateRetained, fitDevelopment, completion, fitCredibility, fitSeverity, drawSeverity, fitBook, forecast, cdfSorted, conformalLevel, makeCalibrator, recalibrateProb, applyCalibration, evaluateFold};
})();
