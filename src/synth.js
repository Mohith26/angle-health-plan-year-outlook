// Synthetic level-funded book: the "truth" process the outlook model never sees directly.
// Deliberately richer than the model (within-year risk drift, seasonality, separate Rx and
// medical lags, multi-month catastrophic episodes with a Pareto tail) so calibration matters.
import { rng, normal, gamma, poisson, ISL_DEFAULT, PLAN_MONTHS, PAID_MONTHS } from './outlook.js';

const MED_LAG = [0.28, 0.35, 0.15, 0.08, 0.05, 0.03, 0.02, 0.015, 0.01, 0.007, 0.004, 0.004, 0.003, 0.002, 0.002, 0.001, 0.001, 0.001];
const CAT_LAG = [0.08, 0.27, 0.24, 0.15, 0.09, 0.06, 0.04, 0.025, 0.015, 0.01, 0.005, 0.005, 0.005, 0.005];
function drawLag(r, dist) {
  let u = r() * dist.reduce((s, x) => s + x, 0);
  for (let k = 0; k < dist.length; k++) { u -= dist[k]; if (u <= 0) return k; }
  return dist.length - 1;
}
const lnorm = (r, med, sig) => med * Math.exp(sig * normal(r));

function ageFactor(age) {
  if (age < 18) return 0.42 + 0.004 * age;
  return 0.55 + 1.55 * Math.pow((age - 18) / 46, 2);
}

// per-claim truth parameters
export const TRUTH = {
  freq: 0.95,               // attritional claims per member-month at factor 1
  rxShare: 0.45, rxMed: 52, rxSig: 1.15,
  medMed: 230, medSig: 1.35,
  catProb: 0.00075,         // catastrophic episode starts per member-month at factor 1
  catXm: 22000, catAlpha: 1.55, catCap: 1.8e6,
  drift: 0.04,              // monthly random-walk sd of the group's log risk
  season: 0.07,
};

function makeMembers(r, nEmp) {
  const members = [];
  for (let e = 0; e < nEmp; e++) {
    const age = 22 + Math.floor(r() * 42);
    members.push({ rel: 'EE', age });
    if (r() < 0.38) members.push({ rel: 'SP', age: Math.max(21, Math.min(64, age + Math.round(normal(r) * 3))) });
    if (age < 55 && r() < 0.42) {
      const kids = 1 + poisson(r, 0.8);
      for (let k = 0; k < kids; k++) members.push({ rel: 'CH', age: Math.floor(r() * 18) });
    }
  }
  return members.map((m) => ({ ...m, a: ageFactor(m.age) }));
}

let BASE = null;
// expected gross and retained cost per member-month at factor 1, by simulation of the truth (cached)
export function baseRates() {
  if (BASE) return BASE;
  const r = rng(777);
  const N = 60000;
  let gross = 0, ret = 0;
  for (let n = 0; n < N; n++) {
    let yr = 0;
    for (let m = 0; m < PLAN_MONTHS; m++) {
      const k = poisson(r, TRUTH.freq);
      for (let c = 0; c < k; c++) yr += r() < TRUTH.rxShare ? lnorm(r, TRUTH.rxMed, TRUTH.rxSig) : lnorm(r, TRUTH.medMed, TRUTH.medSig);
      if (r() < TRUTH.catProb) yr += Math.min(TRUTH.catXm * Math.pow(r() || 1e-12, -1 / TRUTH.catAlpha), TRUTH.catCap);
    }
    gross += yr; ret += Math.min(yr, ISL_DEFAULT);
  }
  BASE = { gross: gross / (N * PLAN_MONTHS), retained: ret / (N * PLAN_MONTHS) };
  return BASE;
}

// One group-year. Returns the record the portal would hold: members with manual (underwritten)
// expected PMPM, the claims fund, and every claim payment with specific stop-loss applied in paid order.
export function generateGroupYear(seed, opts = {}) {
  const r = rng(seed);
  const base = baseRates();
  const nEmp = opts.nEmp ?? Math.max(5, Math.min(140, Math.round(Math.exp(Math.log(18) + 0.65 * normal(r)))));
  const members = makeMembers(r, nEmp);
  const theta = opts.theta ?? gamma(r, 9, 1 / 9);                 // true group risk (cv 0.33)
  const u = Math.pow(theta, 0.4) * Math.exp(0.08 * normal(r));     // underwriting sees part of it
  const startMonth = opts.startMonth ?? 0;                         // calendar month of plan start (0 = Jan)
  const claims = [];                                               // [j, inc, pay, gross]
  let logDrift = 0;
  const forced = opts.forceCat || [];                              // [{j, month, total}]
  for (let m = 0; m < PLAN_MONTHS; m++) {
    logDrift += TRUTH.drift * normal(r);
    const season = 1 + TRUTH.season * Math.cos(2 * Math.PI * ((startMonth + m) % 12) / 12);
    const th = theta * Math.exp(logDrift) * season;
    members.forEach((mem, j) => {
      const k = poisson(r, TRUTH.freq * mem.a * th);
      for (let c = 0; c < k; c++) {
        const rx = r() < TRUTH.rxShare;
        const amt = rx ? lnorm(r, TRUTH.rxMed, TRUTH.rxSig) : lnorm(r, TRUTH.medMed, TRUTH.medSig);
        const lag = rx ? (r() < 0.93 ? 0 : 1) : drawLag(r, MED_LAG);
        claims.push([j, m, m + lag, amt]);
      }
      if (r() < TRUTH.catProb * mem.a * th) addEpisode(r, claims, j, m, Math.min(TRUTH.catXm * Math.pow(r() || 1e-12, -1 / TRUTH.catAlpha), TRUTH.catCap));
    });
    for (const f of forced) if (f.month === m) addEpisode(r, claims, f.j, m, f.total, f.span);
  }
  // specific stop-loss applied in paid order (ties: incurred order)
  claims.sort((x, y) => x[2] - y[2] || x[1] - y[1]);
  const cum = new Float64Array(members.length);
  const payments = [];
  for (const [j, inc, pay, gross] of claims) {
    if (pay > PAID_MONTHS - 1) continue; // outside the 12/24 window
    const ret = Math.min(gross, Math.max(0, ISL_DEFAULT - cum[j]));
    cum[j] += ret;
    payments.push([j, inc, pay, +gross.toFixed(2), +ret.toFixed(2)]);
  }
  const grossPmpm = members.map((m) => base.gross * m.a * u);
  const retPmpm = members.map((m) => base.retained * m.a * u);
  const fund = retPmpm.reduce((s, x) => s + x, 0) * PLAN_MONTHS;
  return {
    seed, theta, u, startMonth, isl: ISL_DEFAULT,
    members: members.map((m, j) => ({ rel: m.rel, age: m.age, grossPmpm: +grossPmpm[j].toFixed(2), retPmpm: +retPmpm[j].toFixed(2) })),
    fund: +fund.toFixed(2), payments,
  };
}
function addEpisode(r, claims, j, m, total, spanOpt) {
  const span = spanOpt ?? 1 + Math.floor(r() * 4);
  const months = [];
  for (let s = 0; s < span && m + s < PLAN_MONTHS; s++) months.push(m + s);
  const w = months.map(() => 0.5 + r());
  const ws = w.reduce((a, b) => a + b, 0);
  months.forEach((mm, k) => {
    const piece = total * w[k] / ws;
    const nParts = 1 + Math.floor(r() * 3);
    for (let p = 0; p < nParts; p++) claims.push([j, mm, mm + drawLag(r, CAT_LAG), piece / nParts]);
  });
}

export function generateBook(n, seed0) {
  const book = [];
  for (let g = 0; g < n; g++) book.push(generateGroupYear(seed0 + g * 7919, { startMonth: g % 12 }));
  return book;
}
