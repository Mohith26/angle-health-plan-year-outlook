// Builds every data file the demo reads. Deterministic (seeded); run: node scripts/build.mjs
//   data/model.json  - book-fitted development, credibility and severity parameters,
//                      per-as-of-month conformal calibrators, held-out test metrics
//   data/groups.json - five current level-funded groups (claims paid through Oct 31, 2026)
import { writeFileSync, mkdirSync, readFileSync } from 'node:fs';
import { fitBook, forecast, ultimateRetained, cdfSorted, makeCalibrator, recalibrateProb, applyCalibration, evaluateFold, quantileSorted } from '../src/outlook.js';
import { generateBook, generateGroupYear } from '../src/synth.js';

const SEED = 20261110;
const N_FOLD = 400;
const AS_OF = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13];
const CAL_SIMS = 300;

const t0 = Date.now();
const book = generateBook(3 * N_FOLD, SEED);
const foldFit = book.slice(0, N_FOLD), foldCal = book.slice(N_FOLD, 2 * N_FOLD), foldTest = book.slice(2 * N_FOLD);
const fit = fitBook(foldFit);
console.log('fit', JSON.stringify(fit.cred), 'sev mean', fit.sev.mean.toFixed(1), `${Date.now() - t0} ms`);

function scoreFold(fold, tag) {
  const out = {};
  for (const a of AS_OF) out[a] = [];
  fold.forEach((g, gi) => {
    const truth = ultimateRetained(g);
    for (const a of AS_OF) {
      const fc = forecast(g, a, fit, { sims: CAL_SIMS, seed: SEED + gi * 31 + a });
      out[a].push({ pit: cdfSorted(fc.samples, truth), fc, truth, fund: g.fund });
    }
  });
  console.log(tag, 'scored', `${Date.now() - t0} ms`);
  return out;
}

const calScores = scoreFold(foldCal, 'calibration');
const calibration = {};
for (const a of AS_OF) calibration[a] = makeCalibrator(calScores[a].map((x) => x.pit), 0.8);

const testScores = scoreFold(foldTest, 'test');
const byAsOf = {}, pooled = [];
for (const a of AS_OF) {
  const rows = testScores[a].map(({ fc, truth, fund }) => {
    const cal = applyCalibration(fc, calibration[a]);
    const r = {
      inRaw: +(truth >= fc.raw.p10 && truth <= fc.raw.p90),
      inCal: +(truth >= cal.range.lo && truth <= cal.range.hi),
      pRaw: fc.raw.pSurplus, pCal: cal.pSurplus,
      surplus: +(truth < fund),
      err: (fc.mean - truth) / fund, naiveErr: (fc.naive - truth) / fund,
    };
    pooled.push(r);
    return r;
  });
  const ev = evaluateFold(rows);
  delete ev.reliabilityRaw; delete ev.reliabilityCal;
  byAsOf[a] = ev;
}
const pooledEval = evaluateFold(pooled);

// ------------------------------------------------------------ current groups (the broker's book)
// claims paid through Oct 31, 2026; asOf = paid months observed in the plan year
const CURRENT = [
  { name: 'Juniper Ridge Dental Partners', groupId: 'UT482113', state: 'UT', start: '2026-01-01', asOf: 10, nEmp: 34, theta: 0.86, seed: 4101,
    forceCat: [{ j: 11, month: 3, total: 61500, span: 3 }] },
  { name: 'Copperline Freight & Logistics', groupId: 'TX306574', state: 'TX', start: '2026-04-01', asOf: 7, nEmp: 58, theta: 1.18, seed: 4219 },
  { name: 'Bluebird Montessori School', groupId: 'GA719240', state: 'GA', start: '2026-07-01', asOf: 4, nEmp: 21, theta: 0.97, seed: 4333 },
  { name: 'Harlan & Pike Architects', groupId: 'OH551087', state: 'OH', start: '2025-10-01', asOf: 13, nEmp: 16, theta: 0.92, seed: 4447 },
  { name: 'Saltwater Marine Supply', groupId: 'FL640932', state: 'FL', start: '2026-02-01', asOf: 9, nEmp: 27, theta: 1.05, seed: 4561 },
];
const groups = CURRENT.map((c) => {
  const startMonth = Number(c.start.slice(5, 7)) - 1;
  const gy = generateGroupYear(c.seed, { nEmp: c.nEmp, theta: c.theta, startMonth, forceCat: c.forceCat });
  return {
    name: c.name, groupId: c.groupId, state: c.state, planStart: c.start, asOfMax: c.asOf,
    isl: gy.isl, fund: gy.fund, members: gy.members,
    // only what the portal could know on Nov 10, 2026: payments made through Oct 31, 2026
    payments: gy.payments.filter((p) => p[2] <= c.asOf - 1),
  };
});

mkdirSync(new URL('../data/', import.meta.url), { recursive: true });
mkdirSync(new URL('../dist/', import.meta.url), { recursive: true });
const model = {
  generated: 'node scripts/build.mjs (seed ' + SEED + ')',
  folds: { fit: N_FOLD, calibration: N_FOLD, test: N_FOLD }, calSims: CAL_SIMS, asOf: AS_OF,
  fit: { dev: { ata: fit.dev.ata.map((x) => +x.toFixed(6)) }, cred: fit.cred, sev: fit.sev, nGroupYears: fit.nGroupYears },
  calibration,
  test: { byAsOf, pooled: pooledEval },
};
writeFileSync(new URL('../data/model.json', import.meta.url), JSON.stringify(model));
writeFileSync(new URL('../data/groups.json', import.meta.url), JSON.stringify({ asOfDate: '2026-10-31', groups }));
// classic-script copies for the static page (works from file:// and GitHub Pages, no build tooling)
const core = readFileSync(new URL('../src/outlook.js', import.meta.url), 'utf8').replace(/^export /gm, '');
const names = [...core.matchAll(/^(?:function|const) ([A-Za-z_][A-Za-z0-9_]*)/gm)].map((m) => m[1]);
writeFileSync(new URL('../dist/outlook.bundle.js', import.meta.url), '// generated from src/outlook.js by scripts/build.mjs\n(function(){\n' + core + '\nwindow.Outlook = {' + names.join(', ') + '};\n})();\n');
writeFileSync(new URL('../data/model.js', import.meta.url), 'window.MODEL = ' + JSON.stringify(model) + ';\n');
writeFileSync(new URL('../data/groups.js', import.meta.url), 'window.GROUPS = ' + JSON.stringify({ asOfDate: '2026-10-31', groups }) + ';\n');
console.log('coverage raw', pooledEval.coverageRaw.toFixed(3), 'cal', pooledEval.coverageCal.toFixed(3),
  'brier raw', pooledEval.brierRaw.toFixed(4), 'cal', pooledEval.brierCal.toFixed(4),
  'MAE model', pooledEval.maeModel.toFixed(3), 'naive', pooledEval.maeNaive.toFixed(3));
for (const a of AS_OF) console.log(a, 'q', calibration[a].q.toFixed(3), 'covRaw', byAsOf[a].coverageRaw.toFixed(3), 'covCal', byAsOf[a].coverageCal.toFixed(3), 'mae', byAsOf[a].maeModel.toFixed(3), 'naive', byAsOf[a].maeNaive.toFixed(3), 'bias', byAsOf[a].biasModel.toFixed(3), byAsOf[a].biasNaive.toFixed(3));
console.log('done', `${Date.now() - t0} ms`);
