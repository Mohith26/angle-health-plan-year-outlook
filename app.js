/* Plan Year Outlook (LF): UI for the Angle Health Business portal. Every number below is
   computed in the browser from data/groups.js by dist/outlook.bundle.js (= src/outlook.js)
   using the book parameters and conformal calibrators in data/model.js. */
(function () {
  const O = window.Outlook, MODEL = window.MODEL, GROUPS = window.GROUPS.groups;
  const SIMS = 4000;
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const C = { primary: '#6d39d9', p200: '#cdb8f4', p100: '#e6dbfa', p050: '#f3eefc', dark: '#343a40', muted: '#6c757d', grid: '#eef0f4', gray: '#c9ccd3', green: '#2f9e5b' };

  const $ = (s) => document.querySelector(s);
  const fmt$ = (x) => (x < 0 ? '-' : '') + '$' + Math.round(Math.abs(x)).toLocaleString('en-US');
  const fmtK = (x) => (x < 0 ? '-' : '') + '$' + (Math.abs(x) >= 1e6 ? (Math.abs(x) / 1e6).toFixed(2) + 'M' : (Math.abs(x) / 1e3).toFixed(1) + 'K');
  const pct = (x, d = 0) => (100 * x).toFixed(d) + '%';
  const monthLabel = (planStart, i, long) => {
    const y = +planStart.slice(0, 4), m = +planStart.slice(5, 7) - 1 + i;
    const yy = y + Math.floor(m / 12), mm = ((m % 12) + 12) % 12;
    return long ? `${MONTHS[mm]} ${yy}` : `${MONTHS[mm]} '${String(yy).slice(2)}`;
  };
  const lastDay = (planStart, i) => {
    const y = +planStart.slice(0, 4), m = +planStart.slice(5, 7) - 1 + i;
    const d = new Date(Date.UTC(y, m + 1, 0));
    return `${String(d.getUTCMonth() + 1).padStart(2, '0')}/${String(d.getUTCDate()).padStart(2, '0')}/${d.getUTCFullYear()}`;
  };
  const planPeriod = (g) => {
    const y = +g.planStart.slice(0, 4), m = +g.planStart.slice(5, 7);
    return `${String(m).padStart(2, '0')}/01/${y} - ${lastDay(g.planStart, 11)}`;
  };

  const qs = new URLSearchParams(location.search);
  const gi0 = Math.min(GROUPS.length - 1, Math.max(0, +(qs.get('g') || 0)));
  const a0 = Math.min(GROUPS[gi0].asOfMax, Math.max(2, +(qs.get('m') || GROUPS[gi0].asOfMax)));
  let state = { gi: gi0, asOf: a0, view: 'report', choice: 'outlook' };
  const charts = {};

  // ------------------------------------------------------------------ chart defaults (portal look)
  Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
  Chart.defaults.font.size = 11.5;
  Chart.defaults.color = C.muted;
  Chart.defaults.plugins.legend.display = false;
  Chart.defaults.plugins.tooltip.backgroundColor = '#fff';
  Chart.defaults.plugins.tooltip.borderColor = '#e7eaf3';
  Chart.defaults.plugins.tooltip.borderWidth = 1;
  Chart.defaults.plugins.tooltip.titleColor = '#212529';
  Chart.defaults.plugins.tooltip.bodyColor = '#212529';
  Chart.defaults.plugins.tooltip.padding = 9;
  Chart.defaults.plugins.tooltip.cornerRadius = 6;
  Chart.defaults.plugins.tooltip.boxPadding = 4;
  Chart.defaults.animation = false;

  // ------------------------------------------------------------------ compute
  function compute() {
    const g = GROUPS[state.gi], a = state.asOf;
    const fc = O.forecast(g, a, MODEL.fit, { sims: SIMS, seed: 9000 + state.gi * 100 + a });
    const cal = O.applyCalibration(fc, MODEL.calibration[a]);
    const nInc = Math.min(a, 12);
    // allocate simulated means to months (BF weights) so the table, waterfall and KPIs add up exactly
    const ibnrW = fc.months.map((m, i) => (i < nInc ? (1 - m.completion) : 0));
    const futW = fc.months.map((m, i) => (i >= nInc ? 1 : 0));
    const sIb = ibnrW.reduce((s, x) => s + x, 0) || 1, sFu = futW.reduce((s, x) => s + x, 0) || 1;
    const rows = fc.months.map((m, i) => {
      const ibnr = fc.ibnrMean * ibnrW[i] / sIb, proj = fc.futureMean * futW[i] / sFu;
      return { ...m, ibnr, proj, total: m.paid + ibnr + proj, label: monthLabel(g.planStart, i, true), observed: i < nInc };
    });
    return { g, a, fc, cal, rows, nInc, test: MODEL.test.byAsOf[a], calib: MODEL.calibration[a] };
  }

  // ------------------------------------------------------------------ render: report view
  function render() {
    const R = compute();
    const { g, a, fc, cal, rows, nInc, test } = R;
    const F = fc.fund, projected = fc.mean, fundToDate = F / 12 * nInc;
    const paidThrough = monthLabel(g.planStart, a - 1, true);

    $('#rpt-title').textContent = `${g.name} - Plan Year Outlook (LF)`;
    $('#created').textContent = `Group ID ${g.groupId} · Claims paid through ${lastDay(g.planStart, a - 1)} · Created on: 11/10/2026`;
    $('#asof-label').textContent = `Claims paid through ${paidThrough}`;
    $('#asof-month').textContent = a <= 12 ? `Month ${a} of 12` : `Run-out month ${a - 12}`;
    const sl = $('#asof');
    sl.max = g.asOfMax; sl.value = a;
    sl.style.setProperty('--pct', `${100 * (a - 2) / Math.max(1, g.asOfMax - 2)}%`);
    $('#tick-lo').textContent = monthLabel(g.planStart, 1);
    $('#tick-hi').textContent = monthLabel(g.planStart, g.asOfMax - 1);
    $('#period').innerHTML = `<option>${planPeriod(g)}</option>`;

    // stats
    $('#k-prob').textContent = pct(cal.pSurplus);
    $('#k-prob-bar').style.width = pct(cal.pSurplus, 1);
    $('#k-prob-foot').innerHTML = `Calibrated on ${MODEL.folds.calibration} completed group-years. Model alone said ${pct(fc.raw.pSurplus)}.`;
    $('#k-surplus').textContent = fmt$(cal.expSurplus);
    $('#k-surplus').classList.toggle('green', cal.expSurplus >= 0.005 * F);
    $('#k-surplus-foot').innerHTML = `Fund balance, 80% range: ${fmtK(cal.balanceRange.lo)} to ${fmtK(cal.balanceRange.hi)}. Shortfalls are covered by aggregate stop-loss.`;
    $('#k-proj').innerHTML = `${fmt$(projected)}`;
    $('#k-proj-foot').innerHTML = `${pct(projected / F)} of the ${fmt$(F)} claims fund. 80% range ${fmtK(cal.range.lo)} to ${fmtK(cal.range.hi)}.`;
    $('#k-ibnr').textContent = fmt$(fc.ibnrMean);
    $('#k-ibnr-foot').innerHTML = `Incurred ${monthLabel(g.planStart, 0)} to ${monthLabel(g.planStart, nInc - 1)}, not yet paid. Estimated from Angle's claims development.`;

    // callout
    const gap = fundToDate - fc.paidToDate;
    const paidLine = gap >= 0
      ? `Paid claims are running ${fmt$(gap)} under the claims fund allocated for ${monthLabel(g.planStart, 0)} to ${monthLabel(g.planStart, nInc - 1)}`
      : `Paid claims are running ${fmt$(-gap)} over the claims fund allocated for ${monthLabel(g.planStart, 0)} to ${monthLabel(g.planStart, nInc - 1)}`;
    const restLine = fc.futureMonths > 0 ? ` and ${fc.futureMonths} month${fc.futureMonths > 1 ? 's' : ''} still to come` : '';
    $('#callout-t').textContent = projected <= F
      ? `On track to finish ${fmt$(F - projected)} under the claims fund`
      : `Tracking ${fmt$(projected - F)} over the claims fund`;
    $('#callout-d').innerHTML = `${paidLine}, but about ${fmt$(fc.ibnrMean)} incurred in those months has not been paid yet${restLine}. Including both, the plan year is projected at ${pct(projected / F)} of the fund, so there is a ${pct(cal.pSurplus)} chance of a surplus at settlement. A straight run-rate of paid claims would say ${fmt$(fc.naive)} (${pct(fc.naive / F)}); on ${test.n} held-out group-years at this point in the year, that shortcut missed by ${pct(test.maeNaive)} of the fund on average, the outlook by ${pct(test.maeModel)}.`;

    renderMonthly(R);
    renderDist(R);
    renderWaterfall(R);
    renderLarge(R);
    renderTable(R);
    renderModel(R);
  }

  function renderMonthly({ rows, fc }) {
    const labels = rows.map((r) => r.label.split(' ')[0]);
    const data = {
      labels,
      datasets: [
        { type: 'bar', label: 'Paid', data: rows.map((r) => r.paid), backgroundColor: C.primary, stack: 's', borderRadius: 2, maxBarThickness: 34 },
        { type: 'bar', label: 'Not yet received (est.)', data: rows.map((r) => r.ibnr), backgroundColor: C.p200, stack: 's', borderRadius: 2, maxBarThickness: 34 },
        { type: 'bar', label: 'Projected', data: rows.map((r) => r.proj), backgroundColor: C.p100, stack: 's', borderRadius: 2, maxBarThickness: 34 },
        { type: 'line', label: 'Claims fund allocation', data: rows.map((r) => r.fund), borderColor: C.dark, borderWidth: 1.6, borderDash: [5, 4], pointRadius: 0, tension: 0, stack: 'line' },
      ],
    };
    const opts = {
      responsive: true, maintainAspectRatio: false,
      scales: {
        x: { stacked: true, grid: { display: false }, border: { color: '#dfe2e8' } },
        y: { stacked: true, grid: { color: C.grid }, border: { display: false }, ticks: { callback: (v) => fmtK(v), maxTicksLimit: 6 } },
      },
      plugins: { tooltip: { mode: 'index', intersect: false, callbacks: { label: (c) => ` ${c.dataset.label}: ${fmt$(c.raw)}`, footer: (items) => { const t = items.filter((i) => i.dataset.type === 'bar').reduce((s, i) => s + i.raw, 0); return `Total: ${fmt$(t)}`; } } } },
    };
    if (charts.monthly) { charts.monthly.data = data; charts.monthly.update(); }
    else charts.monthly = new Chart($('#c-monthly'), { data, options: opts });
  }

  function renderDist({ fc, cal }) {
    const F = fc.fund;
    const bal = Array.from(cal.draws, (u) => F - u);
    const lo = O.quantileSorted([...bal].sort((x, y) => x - y), 0.005), hi = O.quantileSorted([...bal].sort((x, y) => x - y), 0.995);
    const nb = 34, w = (hi - lo) / nb || 1;
    const counts = new Array(nb).fill(0);
    for (const b of bal) { const k = Math.min(nb - 1, Math.max(0, Math.floor((b - lo) / w))); counts[k]++; }
    const centers = counts.map((_, k) => lo + (k + 0.5) * w);
    const rLo = cal.balanceRange.lo, rHi = cal.balanceRange.hi;
    const colors = centers.map((c) => (c >= 0 ? C.primary : C.gray));
    const data = { labels: centers, datasets: [{ data: counts.map((c) => c / bal.length), backgroundColor: colors, borderRadius: 1.5, categoryPercentage: 1, barPercentage: 0.9 }] };
    const band = {
      id: 'band',
      beforeDatasetsDraw(chart) {
        const { ctx, chartArea: ca, scales: { x } } = chart;
        const px = (v) => x.getPixelForValue((v - lo) / w - 0.5);
        ctx.save();
        ctx.fillStyle = 'rgba(109,57,217,0.06)';
        const a = Math.max(ca.left, px(rLo)), b = Math.min(ca.right, px(rHi));
        ctx.fillRect(a, ca.top, b - a, ca.bottom - ca.top);
        const z = px(0);
        if (z > ca.left && z < ca.right) {
          ctx.strokeStyle = C.dark; ctx.setLineDash([4, 3]); ctx.lineWidth = 1.2;
          ctx.beginPath(); ctx.moveTo(z, ca.top); ctx.lineTo(z, ca.bottom); ctx.stroke();
          ctx.setLineDash([]); ctx.fillStyle = C.dark; ctx.font = '11px ' + Chart.defaults.font.family;
          const right = z > ca.right - 70;
          ctx.textAlign = right ? 'right' : 'left';
          ctx.fillText('Break-even', right ? z - 4 : z + 4, ca.top + 10);
        }
        ctx.fillStyle = C.primary; ctx.textAlign = 'center'; ctx.font = '11px ' + Chart.defaults.font.family;
        ctx.fillText('80% range', (a + b) / 2, ca.top + 24);
        ctx.restore();
      },
    };
    const opts = {
      responsive: true, maintainAspectRatio: false,
      scales: {
        x: { grid: { display: false }, border: { color: '#dfe2e8' }, ticks: { autoSkip: true, maxTicksLimit: 7, maxRotation: 0, callback: function (v) { return fmtK(this.getLabelForValue(v)); } } },
        y: { grid: { color: C.grid }, border: { display: false }, ticks: { callback: (v) => pct(v), maxTicksLimit: 5 } },
      },
      plugins: { tooltip: { callbacks: { title: (it) => { const c = +it[0].label; return `Fund balance ${fmtK(c - w / 2)} to ${fmtK(c + w / 2)}`; }, label: (c) => ` ${pct(c.raw, 1)} of outcomes${+c.label < 0 ? ' (aggregate stop-loss pays)' : ''}` } } },
    };
    if (charts.dist) charts.dist.destroy();
    charts.dist = new Chart($('#c-dist'), { type: 'bar', data, options: opts, plugins: [band] });
    $('#dist-sub').innerHTML = `Claims fund minus projected claims at settlement (12 incurred / 24 paid). ${pct(cal.pSurplus)} of calibrated outcomes end in surplus.`;
  }

  function renderWaterfall({ fc }) {
    const F = fc.fund, proj = fc.mean;
    const maxV = Math.max(F, proj) * 1.06;
    const steps = [
      { lab: 'Paid to date', sub: 'Net of specific stop-loss', v: fc.paidToDate, col: C.primary },
      { lab: 'Not yet received', sub: 'Incurred, not yet paid (est.)', v: fc.ibnrMean, col: C.p200 },
      { lab: 'Remaining months', sub: fc.futureMonths ? `${fc.futureMonths} month${fc.futureMonths > 1 ? 's' : ''} of projected claims` : 'Plan year complete', v: fc.futureMean, col: C.p100 },
    ];
    const fundX = F / maxV * 100;
    const fundLine = `<div style="position:absolute;left:${fundX}%;top:-7px;bottom:-7px;border-left:1.6px dashed ${C.dark}"></div>`;
    let level = 0, html = '';
    for (const s of steps) {
      const left = level / maxV * 100, width = s.v / maxV * 100;
      html += `<div class="wf-row"><div class="lab">${s.lab}<small>${s.sub}</small></div><div class="wf-track">
        <div class="wf-bar" style="left:${left}%;width:${Math.max(width, 0.4)}%;background:${s.col}"></div>${fundLine}</div>
        <div class="amt">${fmt$(s.v)}</div></div>`;
      level += s.v;
    }
    const bal = F - proj, pos = bal >= 0;
    html += `<div class="wf-row total"><div class="lab">Projected plan-year claims<small>vs. ${fmt$(F)} claims fund (dashed line)</small></div><div class="wf-track">
      <div class="wf-bar" style="left:0;width:${proj / maxV * 100}%;background:linear-gradient(90deg,${C.primary} 0 ${fc.paidToDate / proj * 100}%,${C.p200} 0 ${(fc.paidToDate + fc.ibnrMean) / proj * 100}%,${C.p100} 0)"></div>${fundLine}</div>
      <div class="amt">${fmt$(proj)}</div></div>
      <div class="wf-row" style="min-height:34px"><div class="lab" style="font-weight:500">Projected fund balance<small>${pos ? 'Surplus at settlement' : 'Shortfall, paid by aggregate stop-loss'}</small></div><div></div>
      <div class="amt" style="font-weight:600;color:${pos ? C.green : '#c23040'}">${fmt$(bal)}</div></div>`;
    $('#wf').innerHTML = html;
    $('#wf-note').innerHTML = `Point estimate: mean of ${SIMS.toLocaleString()} simulated plan years. Specific stop-loss is expected to reimburse a further ${fmt$(fc.islExpected)} on claims still to come, on top of the ${fmt$(fc.reimbursedToDate)} reimbursed so far.`;
  }

  function renderLarge({ g, fc }) {
    const isl = g.isl;
    const list = fc.memberGross.map((gross, j) => ({ j, gross, ret: fc.memberRet[j], rel: g.members[j].rel }))
      .filter((m) => m.gross > 10000).sort((x, y) => y.gross - x.gross);
    const relName = { EE: 'Employee', SP: 'Spouse', CH: 'Child' };
    $('#large-count').textContent = list.length ? `${list.length} member${list.length > 1 ? 's' : ''} over $10K in claims` : 'No members over $10K in claims';
    $('#large-body').innerHTML = list.length ? list.map((m) => {
      const room = Math.max(0, isl - m.ret), reached = room <= 0.5;
      return `<tr><td>${g.groupId.slice(0, 2)}-${String(1000 + m.j * 37 % 9000).padStart(4, '0')} <span style="color:#8c9196;font-size:12px">${relName[m.rel]}</span></td>
        <td class="num">${fmt$(m.gross)}</td><td class="num">${fmt$(m.ret)}</td>
        <td class="num" style="white-space:nowrap"><span class="pct-bar"><i style="width:${Math.min(100, m.ret / isl * 100)}%"></i></span>${reached ? fmt$(0) : fmt$(room)}</td>
        <td>${reached ? '<span class="pill ok">Stop-loss pays</span>' : '<span class="pill purple">Under deductible</span>'}</td></tr>`;
    }).join('') : `<tr><td colspan="5" style="color:#8c9196">No large claimants this plan year.</td></tr>`;
    $('#large-note').innerHTML = `Specific deductible ${fmt$(isl)} per member per plan year. Amounts above it are reimbursed by specific stop-loss and never reduce the claims fund; the outlook simulates each member's remaining room.`;
  }

  function renderTable({ rows, fc }) {
    let tp = 0, ti = 0, tj = 0, tt = 0, tf = 0;
    const body = rows.map((r) => {
      tp += r.paid; ti += r.ibnr; tj += r.proj; tt += r.total; tf += r.fund;
      const diff = r.fund - r.total;
      return `<tr class="${r.observed ? '' : 'future'}"><td>${r.label}</td>
        <td class="num">${r.observed ? fmt$(r.paid) : '-'}</td>
        <td>${r.observed ? `<span class="pct-bar"><i style="width:${r.completion * 100}%"></i></span>${pct(r.completion)}` : '-'}</td>
        <td class="num">${r.observed ? fmt$(r.ibnr) : '-'}</td>
        <td class="num">${r.observed ? '-' : fmt$(r.proj)}</td>
        <td class="num">${fmt$(r.total)}</td>
        <td class="num">${fmt$(r.fund)}</td>
        <td class="num" style="color:${diff >= 0 ? C.green : '#c23040'}">${fmt$(diff)}</td></tr>`;
    }).join('');
    $('#tbl-body').innerHTML = body + `<tr class="tot"><td>Totals</td><td class="num">${fmt$(tp)}</td><td></td><td class="num">${fmt$(ti)}</td><td class="num">${fmt$(tj)}</td><td class="num">${fmt$(tt)}</td><td class="num">${fmt$(tf)}</td><td class="num" style="color:${tf - tt >= 0 ? C.green : '#c23040'}">${fmt$(tf - tt)}</td></tr>`;
  }

  function renderModel({ fc, a, calib, test }) {
    const dev = MODEL.fit.dev;
    const c1 = O.completion(dev, 0, 23), c2 = O.completion(dev, 1, 23), c4 = O.completion(dev, 3, 23);
    const lo = (1 - calib.q) / 2, hi = (1 + calib.q) / 2;
    $('#m-z').textContent = pct(fc.Z);
    $('#m-z-sub').textContent = `Buhlmann-Straub, k = ${Math.round(MODEL.fit.cred.k)} member-months; ${Math.round(fc.w).toLocaleString()} reported member-months so far`;
    $('#m-r').textContent = `${fc.Rhat.toFixed(2)}x`;
    $('#m-r-sub').textContent = `Own completed experience ${fc.X.toFixed(2)}x, blended with underwriting (book mean ${MODEL.fit.cred.m.toFixed(2)}x)`;
    $('#m-dev').textContent = `${pct(c1)} / ${pct(c2)} / ${pct(c4)}`;
    $('#m-dev-sub').textContent = `Chain-ladder on ${MODEL.fit.nGroupYears} completed group-years (fit fold), retained dollars`;
    $('#m-cal').textContent = `${pct(lo, 1)} to ${pct(hi, 1)} pctl`;
    $('#m-cal-sub').textContent = `Split-conformal on ${calib.n} completed group-years at month ${a}; the raw 10th-90th would be too narrow`;
    $('#m-cov').innerHTML = `${pct(test.coverageRaw, 1)} <span style="color:#9aa0a6">→</span> ${pct(test.coverageCal, 1)}`;
    $('#m-cov-sub').textContent = `${test.n} held-out group-years at this month (target 80%)`;
    const ks = MODEL.asOf;
    const data = {
      labels: ks.map((k) => (k <= 12 ? `M${k}` : `R${k - 12}`)),
      datasets: [
        { label: 'Simulation 10th-90th', data: ks.map((k) => MODEL.test.byAsOf[k].coverageRaw), backgroundColor: C.gray, borderRadius: 2, maxBarThickness: 12 },
        { label: 'Calibrated 80% range', data: ks.map((k) => MODEL.test.byAsOf[k].coverageCal), backgroundColor: ks.map((k) => (k === a ? C.primary : C.p200)), borderRadius: 2, maxBarThickness: 12 },
      ],
    };
    const target = { id: 'target', afterDatasetsDraw(ch) { const { ctx, chartArea: ca, scales: { y } } = ch; const yy = y.getPixelForValue(0.8); ctx.save(); ctx.strokeStyle = C.dark; ctx.setLineDash([4, 3]); ctx.beginPath(); ctx.moveTo(ca.left, yy); ctx.lineTo(ca.right, yy); ctx.stroke(); ctx.fillStyle = C.dark; ctx.font = '10.5px ' + Chart.defaults.font.family; ctx.textAlign = 'right'; ctx.fillText('80% target', ca.right, yy - 5); ctx.restore(); } };
    const opts = { responsive: true, maintainAspectRatio: false, scales: { x: { grid: { display: false }, border: { color: '#dfe2e8' } }, y: { min: 0.6, max: 0.9, grid: { color: C.grid }, border: { display: false }, ticks: { callback: (v) => pct(v), stepSize: 0.1 } } }, plugins: { tooltip: { callbacks: { label: (c) => ` ${c.dataset.label}: ${pct(c.raw, 1)} covered` } } } };
    if (charts.cov) charts.cov.destroy();
    charts.cov = new Chart($('#c-cov'), { type: 'bar', data, options: opts, plugins: [target] });
    $('#m-pooled').innerHTML = `Pooled over all months: ${pct(MODEL.test.pooled.coverageRaw, 1)} raw vs ${pct(MODEL.test.pooled.coverageCal, 1)} calibrated. Projection error ${pct(MODEL.test.pooled.maeModel, 1)} of fund vs ${pct(MODEL.test.pooled.maeNaive, 1)} for a paid-claims run-rate.`;
  }

  // ------------------------------------------------------------------ views + events
  function showView(v) {
    state.view = v;
    $('#view-report').classList.toggle('show', v === 'report');
    $('#view-chooser').classList.toggle('show', v === 'chooser');
    $('#h1').textContent = v === 'report' ? 'Plan Year Outlook' : 'View a Report';
    $('#crumb-last').textContent = v === 'report' ? 'Plan Year Outlook' : 'Reporting';
    $('#crumb-mid').style.display = v === 'report' ? '' : 'none';
    window.scrollTo(0, 0);
  }
  function toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(toast.h); toast.h = setTimeout(() => t.classList.remove('show'), 2200); }

  function init() {
    const gs = $('#group');
    gs.innerHTML = GROUPS.map((g, i) => `<option value="${i}">${g.name} (${g.groupId})</option>`).join('');
    gs.value = String(state.gi);
    gs.addEventListener('change', () => { state.gi = +gs.value; state.asOf = GROUPS[state.gi].asOfMax; render(); });
    $('#asof').addEventListener('input', (e) => { state.asOf = +e.target.value; render(); });
    $('#back').addEventListener('click', (e) => { e.preventDefault(); showView('chooser'); });
    $('#crumb-rep').addEventListener('click', (e) => { e.preventDefault(); showView('chooser'); });
    $('#download').addEventListener('click', () => window.print());
    document.querySelectorAll('.rcard').forEach((c) => c.addEventListener('click', () => {
      document.querySelectorAll('.rcard').forEach((x) => x.classList.remove('sel'));
      c.classList.add('sel'); state.choice = c.dataset.r;
    }));
    $('#view-btn').addEventListener('click', () => {
      if (state.choice === 'outlook') showView('report');
      else toast('This concept demo only implements Plan Year Outlook (LF).');
    });
    $('#dl-all').addEventListener('click', () => toast('Download All is part of the existing portal, not this demo.'));
    document.querySelectorAll('.nav-item').forEach((n) => n.addEventListener('click', () => { if (n.dataset.nav !== 'reporting') toast(`${n.textContent} is part of the existing portal, not this demo.`); else showView('chooser'); }));
    $('#group-list').addEventListener('click', () => toast('Back To Group List is part of the existing broker portal.'));
    $('#chooser-period').innerHTML = `<option>${planPeriod(GROUPS[state.gi])}</option>`;
    render();
    showView(location.hash === '#reports' ? 'chooser' : 'report');
  }
  init();
})();
