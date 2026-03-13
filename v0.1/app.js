// ===== Stock Analysis Dashboard — app.js =====
// Parses Screener.in-format Excel, runs analysis, renders charts & report

(function () {
  'use strict';

  // ── DOM refs ──────────────────────────────────────────
  const $ = (id) => document.getElementById(id);
  const uploadScreen = $('uploadScreen');
  const dashboard    = $('dashboard');
  const dropZone     = $('dropZone');
  const fileInput    = $('fileInput');
  const loading      = $('loadingOverlay');

  // ── Chart instances (so we can destroy on re-upload) ──
  let charts = [];

  // ── Event wiring ──────────────────────────────────────
  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', () => { if (fileInput.files.length) handleFile(fileInput.files[0]); });
  $('backBtn').addEventListener('click', resetToUpload);

  // ── Main handler ──────────────────────────────────────
  function handleFile(file) {
    if (!file.name.match(/\.xlsx?$/i)) { alert('Please upload an .xlsx file'); return; }
    loading.classList.add('visible');
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        const data = parseWorkbook(wb);
        const analysis = analyzeData(data);
        renderDashboard(data, analysis);
        uploadScreen.style.display = 'none';
        dashboard.classList.add('visible');
      } catch (err) {
        console.error(err);
        alert('Error parsing file: ' + err.message);
      }
      loading.classList.remove('visible');
    };
    reader.readAsArrayBuffer(file);
  }

  function resetToUpload() {
    dashboard.classList.remove('visible');
    uploadScreen.style.display = '';
    fileInput.value = '';
    charts.forEach(c => c.destroy());
    charts = [];
  }

  // ════════════════════════════════════════════════════════
  //  PARSER — extracts structured data from "Data Sheet"
  // ════════════════════════════════════════════════════════
  function parseWorkbook(wb) {
    // Find "Data Sheet" or fall back to first sheet
    let sheetName = wb.SheetNames.find(n => /data\s*sheet/i.test(n)) || wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

    // Build a map: label → row values (column 0 = label, columns 1..N = values)
    const labelMap = {};
    raw.forEach((row, idx) => {
      if (row[0] && typeof row[0] === 'string') {
        const key = row[0].trim();
        if (!labelMap[key]) labelMap[key] = [];
        labelMap[key].push({ idx, values: row.slice(1) });
      }
    });

    // Helper: get first matching row's values (cleaned)
    const getRow = (label) => {
      const entries = labelMap[label];
      if (!entries || !entries.length) return [];
      return entries[0].values.map(v => (v === null || v === '' || v === undefined) ? null : v);
    };

    // Helper: get the values for a label that may appear multiple times (get specific occurrence by section)
    const getRowInSection = (label, afterLabel) => {
      const entries = labelMap[label];
      if (!entries || !entries.length) return [];
      const afterEntries = labelMap[afterLabel];
      if (!afterEntries || !afterEntries.length) return entries[0].values.map(v => v ?? null);
      const afterIdx = afterEntries[0].idx;
      const match = entries.find(e => e.idx > afterIdx);
      return match ? match.values.map(v => v ?? null) : entries[0].values.map(v => v ?? null);
    };

    // ── Detect report dates for annual section ──
    const annualDatesRaw = getRow('Report Date');
    const annualDates = annualDatesRaw.map(d => formatDate(d));

    // ── Annual P&L ──
    const annualSales      = toNumbers(getRow('Sales'));
    const annualExpenses   = toNumbers(getRow('Other Expenses'));
    const annualNetProfit  = toNumbers(getRow('Net profit'));
    const annualPBT        = toNumbers(getRow('Profit before tax'));
    const annualTax        = toNumbers(getRow('Tax'));
    const annualOtherIncome = toNumbers(getRow('Other Income'));
    const annualInterest   = toNumbers(getRow('Interest'));
    const annualDepreciation = toNumbers(getRow('Depreciation'));
    const annualDividend   = toNumbers(getRow('Dividend Amount'));
    const annualOPM        = annualSales.map((s, i) => {
      if (!s || annualNetProfit[i] == null) return null;
      const opProfit = s - (toNumbers(getRow('Expenses'))[i] || 0);
      return s ? ((opProfit / s) * 100) : null;
    });

    // ── Detect Quarters section ──
    const quartersIdx = (labelMap['Quarters'] || [])[0]?.idx;
    let qDatesRaw = [], qSales = [], qNetProfit = [], qExpenses = [], qOPM = [], qInterest = [], qPBT = [], qOpProfit = [];
    if (quartersIdx !== undefined) {
      // Find the rows after 'Quarters' label
      const qReportDate = getRowInSection('Report Date', 'Quarters');
      qDatesRaw = qReportDate.map(d => formatDate(d));
      qSales     = toNumbers(getRowInSection('Sales', 'Quarters'));
      qExpenses  = toNumbers(getRowInSection('Expenses', 'Quarters'));
      qNetProfit = toNumbers(getRowInSection('Net profit', 'Quarters'));
      qInterest  = toNumbers(getRowInSection('Interest', 'Quarters'));
      qPBT       = toNumbers(getRowInSection('Profit before tax', 'Quarters'));
      qOpProfit  = toNumbers(getRowInSection('Operating Profit', 'Quarters'));
      qOPM = qSales.map((s, i) => {
        if (!s || qOpProfit[i] == null) return null;
        return (qOpProfit[i] / s) * 100;
      });
    }

    // ── Balance Sheet ──
    const bsDatesRaw     = getRowInSection('Report Date', 'BALANCE SHEET');
    const bsDates        = bsDatesRaw.map(d => formatDate(d));
    const bsEquity       = toNumbers(getRowInSection('Equity Share Capital', 'BALANCE SHEET'));
    const bsReserves     = toNumbers(getRowInSection('Reserves', 'BALANCE SHEET'));
    const bsBorrowings   = toNumbers(getRowInSection('Borrowings', 'BALANCE SHEET'));
    const bsOtherLiab    = toNumbers(getRowInSection('Other Liabilities', 'BALANCE SHEET'));
    const bsTotal        = toNumbers(getRowInSection('Total', 'BALANCE SHEET'));
    const bsNetBlock     = toNumbers(getRowInSection('Net Block', 'BALANCE SHEET'));
    const bsInvestments  = toNumbers(getRowInSection('Investments', 'BALANCE SHEET'));
    const bsOtherAssets  = toNumbers(getRowInSection('Other Assets', 'BALANCE SHEET'));
    const bsCash         = toNumbers(getRowInSection('Cash & Bank', 'BALANCE SHEET'));

    // ── Cash Flow ──
    const cfOperating  = toNumbers(getRowInSection('Cash from Operating Activity', 'CASH FLOW:'));
    const cfInvesting  = toNumbers(getRowInSection('Cash from Investing Activity', 'CASH FLOW:'));
    const cfFinancing  = toNumbers(getRowInSection('Cash from Financing Activity', 'CASH FLOW:'));
    const cfNet        = toNumbers(getRowInSection('Net Cash Flow', 'CASH FLOW:'));

    // ── Meta ──
    const companyName = getRow('COMPANY NAME').find(v => v && typeof v === 'string') || 'Unknown Company';
    const currentPrice = toNumbers(getRow('Current Price')).find(v => v !== null) || 0;
    const marketCap = toNumbers(getRow('Market Capitalization')).find(v => v !== null) || 0;
    const faceValue = toNumbers(getRow('Face Value')).find(v => v !== null) || toNumbers(getRow('Face value')).find(v => v !== null) || 0;
    const numShares = toNumbers(getRow('Number of shares')).find(v => v !== null) || 0;

    return {
      companyName, currentPrice, marketCap, faceValue, numShares,
      annual: { dates: annualDates, sales: annualSales, netProfit: annualNetProfit, pbt: annualPBT, tax: annualTax, otherIncome: annualOtherIncome, interest: annualInterest, depreciation: annualDepreciation, expenses: toNumbers(getRow('Other Expenses')), totalExpenses: toNumbers(getRow('Expenses')), opm: annualOPM, dividend: annualDividend },
      quarterly: { dates: qDatesRaw, sales: qSales, netProfit: qNetProfit, expenses: qExpenses, opm: qOPM, interest: qInterest, pbt: qPBT, opProfit: qOpProfit },
      balance: { dates: bsDates, equity: bsEquity, reserves: bsReserves, borrowings: bsBorrowings, otherLiab: bsOtherLiab, total: bsTotal, netBlock: bsNetBlock, investments: bsInvestments, otherAssets: bsOtherAssets, cash: bsCash },
      cashflow: { operating: cfOperating, investing: cfInvesting, financing: cfFinancing, net: cfNet }
    };
  }

  // ════════════════════════════════════════════════════════
  //  ANALYSIS ENGINE — scores each metric
  // ════════════════════════════════════════════════════════
  function analyzeData(d) {
    const findings = [];
    let score = 0; // positive = good, negative = bad

    // ── Helpers ──
    const lastN = (arr, n) => arr.filter(v => v !== null).slice(-n);
    const pctChange = (oldV, newV) => oldV ? ((newV - oldV) / Math.abs(oldV)) * 100 : 0;
    const cagr = (start, end, years) => (start > 0 && end > 0 && years > 0) ? (Math.pow(end / start, 1 / years) - 1) * 100 : null;

    // ── 1. Revenue Growth ──
    const salesClean = d.annual.sales.filter(v => v !== null);
    let revenueGrowth = null;
    let revenueSignal = 'amber';
    if (salesClean.length >= 2) {
      const recent = salesClean.slice(-2);
      revenueGrowth = pctChange(recent[0], recent[1]);
      if (salesClean.length >= 3) {
        revenueGrowth = cagr(salesClean[0], salesClean[salesClean.length - 1], salesClean.length - 1) || revenueGrowth;
      }
      if (revenueGrowth > 10) { revenueSignal = 'green'; score += 2; findings.push({ signal: 'green', text: `Revenue has been growing at a CAGR of ${revenueGrowth.toFixed(1)}%, which is healthy.` }); }
      else if (revenueGrowth > 0) { revenueSignal = 'amber'; score += 1; findings.push({ signal: 'amber', text: `Revenue growth is modest at ${revenueGrowth.toFixed(1)}% CAGR.` }); }
      else { revenueSignal = 'red'; score -= 2; findings.push({ signal: 'red', text: `Revenue is declining at ${revenueGrowth.toFixed(1)}% CAGR — a fundamental concern.` }); }
    }

    // ── 2. Net Profit Trend ──
    const npClean = d.annual.netProfit.filter(v => v !== null);
    let npSignal = 'amber';
    let latestNP = npClean.length ? npClean[npClean.length - 1] : 0;
    let prevNP = npClean.length >= 2 ? npClean[npClean.length - 2] : null;
    let npChange = prevNP !== null ? pctChange(prevNP, latestNP) : 0;
    if (latestNP > 0 && npChange > 0) { npSignal = 'green'; score += 2; findings.push({ signal: 'green', text: `Net Profit is positive (₹${fmt(latestNP)} Cr) and growing ${npChange > 0 ? '+' : ''}${npChange.toFixed(1)}% YoY.` }); }
    else if (latestNP > 0) { npSignal = 'amber'; score += 0; findings.push({ signal: 'amber', text: `Net Profit is positive (₹${fmt(latestNP)} Cr) but has declined ${npChange.toFixed(1)}% YoY.` }); }
    else { npSignal = 'red'; score -= 3; findings.push({ signal: 'red', text: `Net Profit is in LOSS territory at ₹${fmt(latestNP)} Cr — a major red flag.` }); }

    // ── 3. Quarterly trend ──
    const qNP = d.quarterly.netProfit.filter(v => v !== null);
    let qSignal = 'amber';
    if (qNP.length >= 2) {
      const lastQ = qNP.slice(-4);
      const negativeQs = lastQ.filter(v => v < 0).length;
      const latestQ = lastQ[lastQ.length - 1];
      if (negativeQs >= 3) { qSignal = 'red'; score -= 3; findings.push({ signal: 'red', text: `${negativeQs} of the last 4 quarters show NET LOSSES — situation is deteriorating rapidly.` }); }
      else if (negativeQs >= 1) { qSignal = 'amber'; score -= 1; findings.push({ signal: 'amber', text: `${negativeQs} of the last 4 quarters had losses. Inconsistent profitability is concerning.` }); }
      else if (latestQ > 0) { qSignal = 'green'; score += 1; findings.push({ signal: 'green', text: `All recent quarters are profitable — healthy quarterly momentum.` }); }
    }

    // ── 4. Borrowings vs Reserves ──
    const borr = d.balance.borrowings.filter(v => v !== null);
    const resv = d.balance.reserves.filter(v => v !== null);
    let debtSignal = 'amber';
    if (borr.length >= 2) {
      const borrGrowth = pctChange(borr[borr.length - 2], borr[borr.length - 1]);
      const resvGrowth = resv.length >= 2 ? pctChange(resv[resv.length - 2], resv[resv.length - 1]) : 0;
      const dToE = (borr[borr.length - 1] || 0) / Math.max((d.balance.equity.filter(v => v !== null).slice(-1)[0] || 1) + (resv[resv.length - 1] || 0), 1);
      if (borrGrowth > 20 && borrGrowth > resvGrowth) { debtSignal = 'red'; score -= 2; findings.push({ signal: 'red', text: `Borrowings surging +${borrGrowth.toFixed(0)}% while reserves grew only ${resvGrowth.toFixed(0)}%. Debt-to-Equity is ${dToE.toFixed(1)}x.` }); }
      else if (dToE > 3) { debtSignal = 'red'; score -= 1; findings.push({ signal: 'red', text: `High Debt-to-Equity ratio of ${dToE.toFixed(1)}x — leverage risk is elevated.` }); }
      else if (dToE < 1) { debtSignal = 'green'; score += 1; findings.push({ signal: 'green', text: `Low Debt-to-Equity of ${dToE.toFixed(1)}x — balance sheet is strong.` }); }
      else { findings.push({ signal: 'amber', text: `Debt-to-Equity is ${dToE.toFixed(1)}x and borrowings grew ${borrGrowth.toFixed(0)}%.` }); }
    }

    // ── 5. OPM ──
    const opmClean = d.annual.opm.filter(v => v !== null);
    let opmSignal = 'amber';
    let latestOPM = opmClean.length ? opmClean[opmClean.length - 1] : null;
    if (latestOPM !== null) {
      if (latestOPM > 15) { opmSignal = 'green'; score += 1; findings.push({ signal: 'green', text: `Operating Profit Margin is a healthy ${latestOPM.toFixed(1)}%.` }); }
      else if (latestOPM > 0) { opmSignal = 'amber'; findings.push({ signal: 'amber', text: `OPM is thin at ${latestOPM.toFixed(1)}% — limited margin of safety.` }); }
      else { opmSignal = 'red'; score -= 2; findings.push({ signal: 'red', text: `OPM is NEGATIVE at ${latestOPM.toFixed(1)}% — core operations are loss-making.` }); }
    }

    // ── 6. Cash Flow ──
    const cfOp = d.cashflow.operating.filter(v => v !== null);
    let cfSignal = 'amber';
    if (cfOp.length >= 1) {
      const latestCF = cfOp[cfOp.length - 1];
      if (latestCF > 0) { cfSignal = 'green'; score += 1; findings.push({ signal: 'green', text: `Operating Cash Flow is positive at ₹${fmt(latestCF)} Cr.` }); }
      else { cfSignal = 'red'; score -= 1; findings.push({ signal: 'red', text: `Operating Cash Flow is negative (₹${fmt(latestCF)} Cr) — the business is burning cash.` }); }
    }

    // ── 7. Expense spike ──
    const expClean = d.annual.expenses.filter(v => v !== null);
    if (expClean.length >= 2) {
      const expGrowth = pctChange(expClean[expClean.length - 2], expClean[expClean.length - 1]);
      const revGrowth = salesClean.length >= 2 ? pctChange(salesClean[salesClean.length - 2], salesClean[salesClean.length - 1]) : 0;
      if (expGrowth > revGrowth + 20) {
        score -= 2;
        findings.push({ signal: 'red', text: `Expenses spiked +${expGrowth.toFixed(0)}% vs revenue growth of ${revGrowth.toFixed(0)}% — likely provisions or NPA write-offs.` });
      }
    }

    // ── 8. Valuation quick check ──
    if (d.currentPrice && d.balance.equity.length && d.balance.reserves.length) {
      const latestEquity = d.balance.equity.filter(v => v !== null).slice(-1)[0] || 0;
      const latestReserves = resv.length ? resv[resv.length - 1] : 0;
      const latestShares = d.numShares || 1;
      const bookValue = ((latestEquity + latestReserves)) / (latestShares) * 10; // approx BV per share
      const pb = d.currentPrice / Math.max(bookValue, 0.01);
      if (pb < 1 && npSignal === 'red') {
        findings.push({ signal: 'amber', text: `P/B ratio is ${pb.toFixed(2)}x — looks cheap but may be a "value trap" given deteriorating profits.` });
      } else if (pb < 1) {
        findings.push({ signal: 'green', text: `Trading below book value (P/B = ${pb.toFixed(2)}x) — potentially undervalued.` });
        score += 1;
      } else if (pb > 5) {
        findings.push({ signal: 'amber', text: `P/B ratio is ${pb.toFixed(2)}x — valuation is rich, needs strong growth to justify.` });
      }
    }

    // ── Verdict ──
    let verdict, verdictClass;
    if (score >= 4) { verdict = 'Strong Buy'; verdictClass = 'buy'; }
    else if (score >= 2) { verdict = 'Buy'; verdictClass = 'buy'; }
    else if (score >= 0) { verdict = 'Hold'; verdictClass = 'hold'; }
    else if (score >= -3) { verdict = 'Avoid'; verdictClass = 'avoid'; }
    else { verdict = 'High Risk — Avoid'; verdictClass = 'avoid'; }

    return {
      verdict, verdictClass, score, findings,
      indicators: [
        { label: 'Revenue Growth', value: revenueGrowth !== null ? `${revenueGrowth.toFixed(1)}%` : 'N/A', detail: 'CAGR', signal: revenueSignal },
        { label: 'Net Profit', value: `₹${fmt(latestNP)} Cr`, detail: prevNP !== null ? `${npChange >= 0 ? '+' : ''}${npChange.toFixed(1)}% YoY` : '', signal: npSignal },
        { label: 'Quarterly Trend', value: qNP.length ? `₹${fmt(qNP[qNP.length-1])} Cr` : 'N/A', detail: 'Latest Quarter', signal: qSignal },
        { label: 'Operating Margin', value: latestOPM !== null ? `${latestOPM.toFixed(1)}%` : 'N/A', detail: 'OPM', signal: opmSignal },
        { label: 'Debt Health', value: borr.length ? `₹${fmt(borr[borr.length-1])} Cr` : 'N/A', detail: 'Borrowings', signal: debtSignal },
        { label: 'Cash Flow', value: cfOp.length ? `₹${fmt(cfOp[cfOp.length-1])} Cr` : 'N/A', detail: 'Operating CF', signal: cfSignal },
      ]
    };
  }

  // ════════════════════════════════════════════════════════
  //  RENDERER
  // ════════════════════════════════════════════════════════
  function renderDashboard(data, analysis) {
    // Destroy old charts
    charts.forEach(c => c.destroy());
    charts = [];

    // ── Header ──
    $('companyName').textContent = data.companyName;
    $('metaPrice').textContent = data.currentPrice ? `Price: ₹${data.currentPrice}` : '';
    $('metaMcap').textContent = data.marketCap ? `MCap: ₹${fmt(data.marketCap)} Cr` : '';
    $('metaFV').textContent = data.faceValue ? `FV: ₹${data.faceValue}` : '';

    // ── Verdict ──
    const vb = $('verdictBanner');
    vb.className = 'verdict-banner ' + analysis.verdictClass;
    $('verdictBadge').textContent = analysis.verdict;
    $('verdictText').textContent = analysis.findings.length
      ? analysis.findings[0].text
      : 'Analysis complete.';

    // ── Indicator Cards ──
    const grid = $('indicatorsGrid');
    grid.innerHTML = '';
    analysis.indicators.forEach(ind => {
      const card = document.createElement('div');
      card.className = 'indicator-card';
      card.innerHTML = `
        <div class="signal-dot ${ind.signal}"></div>
        <div class="card-label">${ind.label}</div>
        <div class="card-value ${ind.signal}">${ind.value}</div>
        <div class="card-detail">${ind.detail}</div>
      `;
      grid.appendChild(card);
    });

    // ── Charts ──
    const chartDefaults = {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: { labels: { color: '#94a3b8', font: { family: 'Inter', size: 11 } } },
        tooltip: { backgroundColor: 'rgba(17,24,39,0.95)', titleColor: '#f1f5f9', bodyColor: '#94a3b8', borderColor: 'rgba(255,255,255,0.1)', borderWidth: 1, cornerRadius: 8, padding: 12 }
      },
      scales: {
        x: { ticks: { color: '#64748b', font: { size: 10 } }, grid: { color: 'rgba(255,255,255,0.04)' } },
        y: { ticks: { color: '#64748b', font: { size: 10 } }, grid: { color: 'rgba(255,255,255,0.06)' } }
      }
    };

    // 1. Annual Revenue & Net Profit
    const annLabels = cleanLabels(data.annual.dates);
    charts.push(new Chart($('chartAnnual'), {
      type: 'line',
      data: {
        labels: annLabels,
        datasets: [
          { label: 'Revenue', data: cleanData(data.annual.sales, annLabels), borderColor: '#3b82f6', backgroundColor: 'rgba(59,130,246,0.1)', fill: true, tension: 0.35, pointRadius: 4, pointHoverRadius: 7 },
          { label: 'Net Profit', data: cleanData(data.annual.netProfit, annLabels), borderColor: '#22c55e', backgroundColor: 'rgba(34,197,94,0.1)', fill: true, tension: 0.35, pointRadius: 4, pointHoverRadius: 7 }
        ]
      },
      options: chartDefaults
    }));

    // 2. Quarterly Net Profit (bar)
    const qLabels = cleanLabels(data.quarterly.dates);
    const qProfitData = cleanData(data.quarterly.netProfit, qLabels);
    charts.push(new Chart($('chartQuarterly'), {
      type: 'bar',
      data: {
        labels: qLabels,
        datasets: [{
          label: 'Net Profit',
          data: qProfitData,
          backgroundColor: qProfitData.map(v => v >= 0 ? 'rgba(34,197,94,0.7)' : 'rgba(239,68,68,0.7)'),
          borderColor: qProfitData.map(v => v >= 0 ? '#22c55e' : '#ef4444'),
          borderWidth: 1,
          borderRadius: 4
        }]
      },
      options: chartDefaults
    }));

    // 3. Balance Sheet (stacked bar)
    const bsLabels = cleanLabels(data.balance.dates);
    charts.push(new Chart($('chartBalance'), {
      type: 'bar',
      data: {
        labels: bsLabels,
        datasets: [
          { label: 'Equity + Reserves', data: bsLabels.map((_, i) => (data.balance.equity[i] || 0) + (data.balance.reserves[i] || 0)), backgroundColor: 'rgba(34,197,94,0.65)', borderRadius: 2 },
          { label: 'Borrowings', data: cleanData(data.balance.borrowings, bsLabels), backgroundColor: 'rgba(239,68,68,0.55)', borderRadius: 2 },
          { label: 'Other Liabilities', data: cleanData(data.balance.otherLiab, bsLabels), backgroundColor: 'rgba(245,158,11,0.45)', borderRadius: 2 }
        ]
      },
      options: {
        ...chartDefaults,
        scales: {
          ...chartDefaults.scales,
          x: { ...chartDefaults.scales.x, stacked: true },
          y: { ...chartDefaults.scales.y, stacked: true }
        }
      }
    }));

    // 4. Cash Flow (grouped bar)
    const cfLabels = cleanLabels(data.balance.dates); // CF usually matches BS dates
    charts.push(new Chart($('chartCashflow'), {
      type: 'bar',
      data: {
        labels: cfLabels,
        datasets: [
          { label: 'Operating', data: cleanData(data.cashflow.operating, cfLabels), backgroundColor: 'rgba(59,130,246,0.7)', borderRadius: 3 },
          { label: 'Investing', data: cleanData(data.cashflow.investing, cfLabels), backgroundColor: 'rgba(168,85,247,0.6)', borderRadius: 3 },
          { label: 'Financing', data: cleanData(data.cashflow.financing, cfLabels), backgroundColor: 'rgba(245,158,11,0.6)', borderRadius: 3 }
        ]
      },
      options: chartDefaults
    }));

    // ── Key Metrics Table ──
    renderMetricsTable(data);

    // ── Report ──
    const rl = $('reportList');
    rl.innerHTML = '';
    analysis.findings.forEach(f => {
      const li = document.createElement('li');
      li.innerHTML = `<span class="bullet ${f.signal}"></span><span>${f.text}</span>`;
      rl.appendChild(li);
    });
  }

  // ── Metrics Table ──
  function renderMetricsTable(data) {
    const headerRow = $('metricsHeader');
    const tbody = $('metricsBody');
    headerRow.innerHTML = '';
    tbody.innerHTML = '';

    // Pick last 2 annual dates for comparison
    const dates = data.annual.dates.filter(d => d);
    const n = data.annual.sales.length;
    if (n < 2) return;

    const prevI = n - 2;
    const currI = n - 1;
    const prevLabel = dates[prevI] || `Year ${prevI + 1}`;
    const currLabel = dates[currI] || `Year ${currI + 1}`;

    headerRow.innerHTML = `<th>Metric</th><th>${prevLabel}</th><th>${currLabel}</th><th>Change</th>`;

    const metrics = [
      { name: 'Sales (Revenue)', prev: data.annual.sales[prevI], curr: data.annual.sales[currI] },
      { name: 'Net Profit', prev: data.annual.netProfit[prevI], curr: data.annual.netProfit[currI] },
      { name: 'Interest', prev: data.annual.interest[prevI], curr: data.annual.interest[currI] },
      { name: 'Borrowings', prev: data.balance.borrowings[prevI], curr: data.balance.borrowings[currI] },
      { name: 'Reserves', prev: data.balance.reserves[prevI], curr: data.balance.reserves[currI] },
      { name: 'Other Expenses', prev: data.annual.expenses[prevI], curr: data.annual.expenses[currI] },
    ];

    metrics.forEach(m => {
      if (m.prev == null && m.curr == null) return;
      const change = (m.prev && m.curr) ? ((m.curr - m.prev) / Math.abs(m.prev) * 100) : null;
      const changeClass = change === null ? '' : change > 5 ? (m.name === 'Net Profit' || m.name === 'Reserves' || m.name === 'Sales (Revenue)' ? 'positive' : 'negative') : change < -5 ? (m.name === 'Net Profit' || m.name === 'Reserves' || m.name === 'Sales (Revenue)' ? 'negative' : 'positive') : 'neutral';
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${m.name}</td>
        <td>${m.prev != null ? '₹' + fmt(m.prev) + ' Cr' : '—'}</td>
        <td>${m.curr != null ? '₹' + fmt(m.curr) + ' Cr' : '—'}</td>
        <td class="${changeClass}">${change !== null ? (change >= 0 ? '+' : '') + change.toFixed(1) + '%' : '—'}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  // ════════════════════════════════════════════════════════
  //  UTILITIES
  // ════════════════════════════════════════════════════════
  function toNumbers(arr) {
    return arr.map(v => {
      if (v === null || v === undefined || v === '') return null;
      const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, ''));
      return isNaN(n) ? null : n;
    });
  }

  function formatDate(v) {
    if (!v) return null;
    if (v instanceof Date) {
      const y = v.getFullYear();
      const m = String(v.getMonth() + 1).padStart(2, '0');
      return `${y}-${m}`;
    }
    const s = String(v);
    // Try to extract YYYY or YYYY-MM from string
    const match = s.match(/(\d{4})[-\/]?(\d{2})?/);
    if (match) return match[2] ? `${match[1]}-${match[2]}` : match[1];
    return s.substring(0, 10);
  }

  function cleanLabels(arr) {
    return arr.filter(v => v !== null && v !== undefined && v !== '');
  }

  function cleanData(arr, labels) {
    // Return data aligned to clean labels length
    const clean = arr.filter(v => v !== null);
    // If arrays mismatch, slice from end to match labels
    if (clean.length >= labels.length) return clean.slice(clean.length - labels.length);
    // Pad with null at start
    const padded = new Array(labels.length - clean.length).fill(null).concat(clean);
    return padded;
  }

  function fmt(n) {
    if (n === null || n === undefined) return '—';
    const abs = Math.abs(n);
    if (abs >= 100) return n.toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return n.toFixed(2);
  }

})();
