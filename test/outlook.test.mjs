import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import {
  rng, poisson, gamma, normal, completion, fitBook, forecast, observe, ultimateRetained,
  conformalLevel, makeCalibrator, recalibrateProb, applyCalibration, cdfSorted, quantileSorted, ISL_DEFAULT,
} from '../src/outlook.js';
import { generateBook, generateGroupYear } from '../src/synth.js';

const book = generateBook(160, 555);
const fit = fitBook(book);

test('random generators are deterministic and have the right means', () => {
  const a = rng(1), b = rng(1);
  for (let i = 0; i < 5; i++) assert.equal(a(), b());
  const r = rng(2);
  let sp = 0, sg = 0, sn = 0;
  const N = 40000;
  for (let i = 0; i < N; i++) { sp += poisson(r, 3.2); sg += gamma(r, 4, 0.5); sn += normal(r); }
  assert.ok(Math.abs(sp / N - 3.2) < 0.05);
  assert.ok(Math.abs(sg / N - 2.0) < 0.03);
  assert.ok(Math.abs(sn / N) < 0.02);
});

test('synthetic group-years are reproducible and apply the specific deductible in paid order', () => {
  const g1 = generateGroupYear(42), g2 = generateGroupYear(42);
  assert.deepEqual(g1.payments.slice(0, 20), g2.payments.slice(0, 20));
  const g = generateGroupYear(7, { forceCat: [{ j: 0, month: 2, total: 90000, span: 2 }] });
  const ret = new Map();
  for (const [j, , , gross, r] of g.payments) { assert.ok(r <= gross + 1e-9); ret.set(j, (ret.get(j) || 0) + r); }
  for (const v of ret.values()) assert.ok(v <= ISL_DEFAULT + 0.05);
  assert.ok(Math.abs(ret.get(0) - ISL_DEFAULT) < 0.05, 'forced large claimant capped exactly at the deductible');
});

test('chain-ladder completion is monotone, ends at 1, and recovers a realistic lag', () => {
  let prev = 0;
  for (let d = 0; d <= 23; d++) { const c = completion(fit.dev, d, 23); assert.ok(c >= prev - 1e-12); prev = c; }
  assert.equal(completion(fit.dev, 23, 23), 1);
  const c0 = completion(fit.dev, 0, 23);
  assert.ok(c0 > 0.25 && c0 < 0.5, `month-of-incurral completion ${c0}`);
  // empirical check on the book: share of contract-ultimate paid in the incurral month
  let p0 = 0, ult = 0;
  for (const g of book) for (const [, inc, pay, , r] of g.payments) { if (pay === inc) p0 += r; ult += r; }
  assert.ok(Math.abs(p0 / ult - c0) < 0.04);
});

test('Buhlmann-Straub parameters are positive and credibility grows with exposure', () => {
  assert.ok(fit.cred.epv > 0 && fit.cred.vhm > 0 && fit.cred.k > 0);
  const g = book[3];
  const z = [3, 6, 9, 12].map((a) => forecast(g, a, fit, { sims: 50, seed: 1 }).Z);
  for (let i = 1; i < z.length; i++) assert.ok(z[i] > z[i - 1]);
  const f = forecast(g, 6, fit, { sims: 50, seed: 1 });
  const lo = Math.min(f.X, fit.cred.m), hi = Math.max(f.X, fit.cred.m);
  assert.ok(f.Rhat >= lo - 1e-12 && f.Rhat <= hi + 1e-12, 'credibility estimate lies between experience and prior');
});

test('forecast decomposition is additive and bounded below by paid claims', () => {
  const g = book[10];
  const f = forecast(g, 7, fit, { sims: 800, seed: 3 });
  assert.ok(Math.abs(f.mean - (f.paidToDate + f.ibnrMean + f.futureMean)) < 1e-6 * f.fund);
  assert.ok(f.samples[0] >= f.paidToDate - 1e-9);
  for (let i = 1; i < f.samples.length; i++) assert.ok(f.samples[i] >= f.samples[i - 1]);
  const obs = observe(g, 7);
  assert.equal(obs.nInc, 7);
});

test('a member already at the specific deductible adds nothing to outstanding claims', () => {
  const members = [{ rel: 'EE', age: 50, grossPmpm: 600, retPmpm: 520 }, { rel: 'EE', age: 50, grossPmpm: 600, retPmpm: 520 }];
  const fund = 2 * 520 * 12;
  // same paid dollars by month in both groups, so experience, credibility and R-hat are identical
  const capped = { members, fund, isl: ISL_DEFAULT, payments: [[0, 0, 1, 90000, 30000]] };
  const split = { members, fund, isl: ISL_DEFAULT, payments: [[0, 0, 1, 15000, 15000], [1, 0, 1, 15000, 15000]] };
  const fa = forecast(capped, 4, fit, { sims: 3000, seed: 5 });
  const fb = forecast(split, 4, fit, { sims: 3000, seed: 5 });
  assert.equal(fa.Rhat, fb.Rhat);
  assert.ok(Math.abs(fa.memberRet[0] - ISL_DEFAULT) < 1e-9);
  assert.ok(fa.ibnrMean + fa.futureMean < fb.ibnrMean + fb.futureMean, 'capped member has no room left');
  assert.ok(fa.islExpected > fb.islExpected, 'more of the outstanding claims are reimbursed by stop-loss');
  assert.equal(fa.reimbursedToDate, 60000);
});

test('as the year completes, the forecast converges on the settled amount', () => {
  const g = book[20];
  const truth = ultimateRetained(g);
  const e3 = Math.abs(forecast(g, 3, fit, { sims: 400, seed: 2 }).mean - truth);
  const e18 = Math.abs(forecast(g, 18, fit, { sims: 400, seed: 2 }).mean - truth);
  assert.ok(e18 < e3 || e18 < 0.02 * g.fund);
  const f24 = forecast(g, 24, fit, { sims: 200, seed: 2 });
  assert.ok(Math.abs(f24.mean - truth) < 1e-6 * g.fund, 'fully developed year equals its paid claims');
});

test('split-conformal level recovers 80% coverage for a miscalibrated forecaster', () => {
  const r = rng(9);
  // forecaster reports N(0, 0.7^2) when truth is N(0, 1): overconfident
  const pit = (y) => { const z = y / 0.7; return 0.5 * (1 + Math.tanh(0.7978845608 * (z + 0.044715 * z * z * z))); };
  const cal = Array.from({ length: 800 }, () => pit(normal(r)));
  const q = conformalLevel(cal, 0.8);
  assert.ok(q > 0.8, 'overconfident model needs wider quantile levels');
  let inRaw = 0, inCal = 0;
  const M = 6000;
  for (let i = 0; i < M; i++) { const u = pit(normal(r)); inRaw += Math.abs(2 * u - 1) <= 0.8; inCal += Math.abs(2 * u - 1) <= q; }
  assert.ok(inRaw / M < 0.7);
  assert.ok(Math.abs(inCal / M - 0.8) < 0.03, `calibrated coverage ${inCal / M}`);
  const uni = Array.from({ length: 2000 }, () => r());
  assert.ok(Math.abs(conformalLevel(uni, 0.8) - 0.8) < 0.04);
});

test('probability recalibration is monotone and applyCalibration is internally consistent', () => {
  const calib = makeCalibrator(Array.from({ length: 300 }, (_, i) => (i + 0.5) / 300), 0.8);
  let prev = 0;
  for (let u = 0; u <= 1; u += 0.05) { const p = recalibrateProb(calib, u); assert.ok(p >= prev && p > 0 && p < 1); prev = p; }
  const f = forecast(book[5], 8, fit, { sims: 1000, seed: 4 });
  const c = applyCalibration(f, calib);
  assert.ok(c.range.lo <= c.range.hi);
  assert.ok(Math.abs(c.balanceRange.lo - (f.fund - c.range.hi)) < 1e-9);
  // uniform PITs => recalibration is (nearly) the identity
  assert.ok(Math.abs(c.pSurplus - cdfSorted(f.samples, f.fund)) < 0.02);
  assert.ok(Math.abs(quantileSorted(f.samples, 0.5) - f.raw.p50) < 1e-9);
});

test('committed data files match the model and only contain information available on the as-of date', () => {
  const model = JSON.parse(readFileSync(new URL('../data/model.json', import.meta.url)));
  const groups = JSON.parse(readFileSync(new URL('../data/groups.json', import.meta.url))).groups;
  assert.ok(model.test.pooled.coverageCal > 0.76 && model.test.pooled.coverageCal < 0.84);
  assert.ok(model.test.pooled.coverageRaw < model.test.pooled.coverageCal);
  assert.ok(model.test.pooled.maeModel < model.test.pooled.maeNaive);
  for (const g of groups) for (const p of g.payments) assert.ok(p[2] <= g.asOfMax - 1);
  for (const a of model.asOf) assert.ok(model.calibration[a].n === model.folds.calibration);
});
