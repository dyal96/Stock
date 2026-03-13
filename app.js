// ===== Stock Analysis Dashboard v2 — app.js =====
(function () {
  'use strict';

  // ── DOM refs ──
  const $ = (id) => document.getElementById(id);
  const uploadScreen = $('uploadScreen');
  const dashboard = $('dashboard');
  const dropZone = $('dropZone');
  const fileInput = $('fileInput');
  const loading = $('loadingOverlay');

  let charts = [];
  let sparkCharts = [];
  let currentData = null;
  let currentAnalysis = null;
  let compareStocks = []; // { data, analysis }

  // ═══════════════════════════════════════════════
  //  HISTORY (localStorage)
  // ═══════════════════════════════════════════════
  const HISTORY_KEY = 'stockAnalyzerHistory';

  function loadHistory() {
    try { return JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]'); } catch { return []; }
  }
  function saveHistory(entry) {
    const history = loadHistory();
    // Keep max 10, replace if same company
    const idx = history.findIndex(h => h.name === entry.name);
    if (idx >= 0) history.splice(idx, 1);
    history.unshift(entry);
    if (history.length > 10) history.pop();
    localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
  }
  function deleteHistory(name) {
    const history = loadHistory().filter(h => h.name !== name);
    localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
    renderHistory();
  }

  function renderHistory() {
    const panel = $('historyPanel');
    const history = loadHistory();
    if (!history.length) { panel.innerHTML = ''; return; }
    panel.innerHTML = `<h3>📁 Recent Analyses</h3><ul class="history-list">${history.map(h => `
      <li class="history-item" data-name="${esc(h.name)}">
        <span class="hi-name">${esc(h.name)}</span>
        <span class="hi-meta">
          <span class="hi-verdict ${h.verdictClass}">${esc(h.verdict)}</span>
          <span>${h.date}</span>
          <button class="hi-delete" data-del="${esc(h.name)}" title="Remove">✕</button>
        </span>
      </li>`).join('')}</ul>`;
    panel.querySelectorAll('.hi-delete').forEach(btn => {
      btn.addEventListener('click', (e) => { e.stopPropagation(); deleteHistory(btn.dataset.del); });
    });
    // Clicking a history item doesn't re-upload (no stored data), just shows info
  }

  // ═══════════════════════════════════════════════
  //  EVENT WIRING
  // ═══════════════════════════════════════════════
  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
  });
  fileInput.addEventListener('change', () => { if (fileInput.files.length) handleFiles(fileInput.files); });

  $('backBtn').addEventListener('click', resetToUpload);
  $('themeToggle').addEventListener('click', toggleTheme);
  $('exportPdf').addEventListener('click', exportPDF);
  $('printBtn').addEventListener('click', () => window.print());

  // Compare
  $('compareToggle').addEventListener('click', () => {
    const bar = $('compareBar');
    const sec = $('compareSection');
    bar.classList.toggle('hidden');
    sec.classList.toggle('hidden');
  });
  $('compareAddBtn').addEventListener('click', () => $('compareFileInput').click());
  $('compareFileInput').addEventListener('change', () => {
    const f = $('compareFileInput').files[0];
    if (f) addCompareStock(f);
  });

  // Init history
  renderHistory();

  // ═══════════════════════════════════════════════
  //  THEME TOGGLE
  // ═══════════════════════════════════════════════
  function toggleTheme() {
    const html = document.documentElement;
    const current = html.getAttribute('data-theme');
    const next = current === 'dark' ? 'light' : 'dark';
    html.setAttribute('data-theme', next);
    $('themeToggle').textContent = next === 'dark' ? '🌙' : '☀️';
    localStorage.setItem('stockTheme', next);
  }
  // Restore theme
  const savedTheme = localStorage.getItem('stockTheme');
  if (savedTheme) {
    document.documentElement.setAttribute('data-theme', savedTheme);
    $('themeToggle').textContent = savedTheme === 'dark' ? '🌙' : '☀️';
  }

  // ═══════════════════════════════════════════════
  //  FILE HANDLER
  // ═══════════════════════════════════════════════
  function handleFiles(files) {
    const fileArr = Array.from(files).filter(f => /\.xlsx?$/i.test(f.name));
    if (!fileArr.length) { alert('Please upload .xlsx files'); return; }
    loading.classList.add('visible');

    // Process first file as main
    readExcel(fileArr[0], (data, analysis) => {
      currentData = data;
      currentAnalysis = analysis;
      compareStocks = [{ data, analysis }];
      renderDashboard(data, analysis);
      uploadScreen.style.display = 'none';
      dashboard.classList.add('visible');
      showToolbarBtns(true);

      // Save to history
      saveHistory({
        name: data.companyName,
        verdict: analysis.verdict,
        verdictClass: analysis.verdictClass,
        date: new Date().toLocaleDateString('en-IN')
      });

      // If multi files, add rest as compare
      for (let i = 1; i < fileArr.length; i++) {
        addCompareStock(fileArr[i]);
      }
      loading.classList.remove('visible');
    });
  }

  function addCompareStock(file) {
    readExcel(file, (data, analysis) => {
      compareStocks.push({ data, analysis });
      saveHistory({
        name: data.companyName,
        verdict: analysis.verdict,
        verdictClass: analysis.verdictClass,
        date: new Date().toLocaleDateString('en-IN')
      });
      renderCompare();
      $('compareBar').classList.remove('hidden');
      $('compareSection').classList.remove('hidden');
    });
  }

  function readExcel(file, cb) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        const data = parseWorkbook(wb);
        const analysis = analyzeData(data);
        cb(data, analysis);
      } catch (err) {
        console.error(err);
        alert('Error parsing ' + file.name + ': ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function resetToUpload() {
    dashboard.classList.remove('visible');
    uploadScreen.style.display = '';
    fileInput.value = '';
    charts.forEach(c => c.destroy());
    charts = [];
    sparkCharts.forEach(c => c.destroy());
    sparkCharts = [];
    compareStocks = [];
    currentData = null;
    currentAnalysis = null;
    showToolbarBtns(false);
    renderHistory();
  }

  function showToolbarBtns(show) {
    ['backBtn', 'compareToggle', 'exportPdf', 'printBtn'].forEach(id => {
      $(id).style.display = show ? '' : 'none';
    });
  }

  // ═══════════════════════════════════════════════
  //  PDF EXPORT
  // ═══════════════════════════════════════════════
  async function exportPDF() {
    loading.classList.add('visible');
    try {
      const el = $('dashboard');
      const canvas = await html2canvas(el, {
        backgroundColor: getComputedStyle(document.body).getPropertyValue('--bg-primary').trim() || '#0a0e17',
        scale: 1.5,
        useCORS: true,
        logging: false
      });
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: 'portrait', unit: 'px', format: [canvas.width, canvas.height] });
      pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 0, 0, canvas.width, canvas.height);
      pdf.save(`${currentData?.companyName || 'analysis'}_report.pdf`);
    } catch (err) {
      console.error(err);
      alert('PDF export failed: ' + err.message);
    }
    loading.classList.remove('visible');
  }

  // ════════════════════════════════════════════════
  //  PARSER
  // ════════════════════════════════════════════════
  function parseWorkbook(wb) {
    let sheetName = wb.SheetNames.find(n => /data\s*sheet/i.test(n)) || wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

    const labelMap = {};
    raw.forEach((row, idx) => {
      if (row[0] && typeof row[0] === 'string') {
        const key = row[0].trim();
        if (!labelMap[key]) labelMap[key] = [];
        labelMap[key].push({ idx, values: row.slice(1) });
      }
    });

    const getRow = (label) => {
      const entries = labelMap[label];
      if (!entries || !entries.length) return [];
      return entries[0].values.map(v => v ?? null);
    };

    const getRowInSection = (label, afterLabel) => {
      const entries = labelMap[label];
      if (!entries || !entries.length) return [];
      const afterEntries = labelMap[afterLabel];
      if (!afterEntries || !afterEntries.length) return entries[0].values.map(v => v ?? null);
      const afterIdx = afterEntries[0].idx;
      const match = entries.find(e => e.idx > afterIdx);
      return match ? match.values.map(v => v ?? null) : entries[0].values.map(v => v ?? null);
    };

    const annualDates = getRow('Report Date').map(d => formatDate(d));
    const annualSales = toNumbers(getRow('Sales'));
    const annualNetProfit = toNumbers(getRow('Net profit'));
    const annualPBT = toNumbers(getRow('Profit before tax'));
    const annualTax = toNumbers(getRow('Tax'));
    const annualOtherIncome = toNumbers(getRow('Other Income'));
    const annualInterest = toNumbers(getRow('Interest'));
    const annualDepreciation = toNumbers(getRow('Depreciation'));
    const annualDividend = toNumbers(getRow('Dividend Amount'));
    const annualTotalExp = toNumbers(getRow('Expenses'));
    const annualOtherExp = toNumbers(getRow('Other Expenses'));
    const annualOPM = annualSales.map((s, i) => {
      if (!s) return null;
      const exp = annualTotalExp[i] || 0;
      return ((s - exp) / s) * 100;
    });

    // Quarters
    const quartersIdx = (labelMap['Quarters'] || [])[0]?.idx;
    let qDates = [], qSales = [], qNetProfit = [], qExpenses = [], qOPM = [], qInterest = [], qPBT = [], qOpProfit = [];
    if (quartersIdx !== undefined) {
      qDates = getRowInSection('Report Date', 'Quarters').map(d => formatDate(d));
      qSales = toNumbers(getRowInSection('Sales', 'Quarters'));
      qExpenses = toNumbers(getRowInSection('Expenses', 'Quarters'));
      qNetProfit = toNumbers(getRowInSection('Net profit', 'Quarters'));
      qInterest = toNumbers(getRowInSection('Interest', 'Quarters'));
      qPBT = toNumbers(getRowInSection('Profit before tax', 'Quarters'));
      qOpProfit = toNumbers(getRowInSection('Operating Profit', 'Quarters'));
      qOPM = qSales.map((s, i) => (s && qOpProfit[i] != null) ? (qOpProfit[i] / s) * 100 : null);
    }

    // Balance Sheet
    const bsDates = getRowInSection('Report Date', 'BALANCE SHEET').map(d => formatDate(d));
    const bsEquity = toNumbers(getRowInSection('Equity Share Capital', 'BALANCE SHEET'));
    const bsReserves = toNumbers(getRowInSection('Reserves', 'BALANCE SHEET'));
    const bsBorrowings = toNumbers(getRowInSection('Borrowings', 'BALANCE SHEET'));
    const bsOtherLiab = toNumbers(getRowInSection('Other Liabilities', 'BALANCE SHEET'));
    const bsTotal = toNumbers(getRowInSection('Total', 'BALANCE SHEET'));
    const bsNetBlock = toNumbers(getRowInSection('Net Block', 'BALANCE SHEET'));
    const bsInvestments = toNumbers(getRowInSection('Investments', 'BALANCE SHEET'));
    const bsOtherAssets = toNumbers(getRowInSection('Other Assets', 'BALANCE SHEET'));
    const bsCash = toNumbers(getRowInSection('Cash & Bank', 'BALANCE SHEET'));

    // Cash Flow
    const cfOperating = toNumbers(getRowInSection('Cash from Operating Activity', 'CASH FLOW:'));
    const cfInvesting = toNumbers(getRowInSection('Cash from Investing Activity', 'CASH FLOW:'));
    const cfFinancing = toNumbers(getRowInSection('Cash from Financing Activity', 'CASH FLOW:'));
    const cfNet = toNumbers(getRowInSection('Net Cash Flow', 'CASH FLOW:'));

    // Meta
    const companyName = getRow('COMPANY NAME').find(v => v && typeof v === 'string') || 'Unknown Company';
    const currentPrice = toNumbers(getRow('Current Price')).find(v => v !== null) || 0;
    const marketCap = toNumbers(getRow('Market Capitalization')).find(v => v !== null) || 0;
    const faceValue = toNumbers(getRow('Face Value')).find(v => v !== null) || toNumbers(getRow('Face value')).find(v => v !== null) || 0;
    const numShares = toNumbers(getRow('Number of shares')).find(v => v !== null) || 0;

    return {
      companyName, currentPrice, marketCap, faceValue, numShares,
      annual: { dates: annualDates, sales: annualSales, netProfit: annualNetProfit, pbt: annualPBT, tax: annualTax, otherIncome: annualOtherIncome, interest: annualInterest, depreciation: annualDepreciation, expenses: annualOtherExp, totalExpenses: annualTotalExp, opm: annualOPM, dividend: annualDividend },
      quarterly: { dates: qDates, sales: qSales, netProfit: qNetProfit, expenses: qExpenses, opm: qOPM, interest: qInterest, pbt: qPBT, opProfit: qOpProfit },
      balance: { dates: bsDates, equity: bsEquity, reserves: bsReserves, borrowings: bsBorrowings, otherLiab: bsOtherLiab, total: bsTotal, netBlock: bsNetBlock, investments: bsInvestments, otherAssets: bsOtherAssets, cash: bsCash },
      cashflow: { operating: cfOperating, investing: cfInvesting, financing: cfFinancing, net: cfNet }
    };
  }

  // ════════════════════════════════════════════════
  //  ANALYSIS ENGINE
  // ════════════════════════════════════════════════
  function analyzeData(d) {
    const findings = [];
    const scoreParts = []; // { label, points, signal }
    let score = 0;

    const lastN = (arr, n) => arr.filter(v => v !== null).slice(-n);
    const pctChange = (o, n) => o ? ((n - o) / Math.abs(o)) * 100 : 0;
    const cagrFn = (s, e, y) => (s > 0 && e > 0 && y > 0) ? (Math.pow(e / s, 1 / y) - 1) * 100 : null;

    // 1. Revenue Growth
    const salesClean = d.annual.sales.filter(v => v !== null);
    let revenueGrowth = null, revenueSignal = 'amber';
    if (salesClean.length >= 2) {
      revenueGrowth = salesClean.length >= 3
        ? cagrFn(salesClean[0], salesClean[salesClean.length - 1], salesClean.length - 1)
        : pctChange(salesClean[salesClean.length - 2], salesClean[salesClean.length - 1]);
      if (revenueGrowth === null) revenueGrowth = pctChange(salesClean[salesClean.length - 2], salesClean[salesClean.length - 1]);
      if (revenueGrowth > 10) { revenueSignal = 'green'; score += 2; scoreParts.push({ label: 'Revenue ↑', points: +2, signal: 'green' }); findings.push({ signal: 'green', text: `Revenue growing at ${revenueGrowth.toFixed(1)}% CAGR — healthy top-line growth.` }); }
      else if (revenueGrowth > 0) { revenueSignal = 'amber'; score += 1; scoreParts.push({ label: 'Revenue ~', points: +1, signal: 'amber' }); findings.push({ signal: 'amber', text: `Revenue growth is modest at ${revenueGrowth.toFixed(1)}% CAGR.` }); }
      else { revenueSignal = 'red'; score -= 2; scoreParts.push({ label: 'Revenue ↓', points: -2, signal: 'red' }); findings.push({ signal: 'red', text: `Revenue declining at ${revenueGrowth.toFixed(1)}% CAGR — fundamental concern.` }); }
    }

    // 2. Net Profit
    const npClean = d.annual.netProfit.filter(v => v !== null);
    let npSignal = 'amber', latestNP = npClean[npClean.length - 1] || 0;
    let prevNP = npClean.length >= 2 ? npClean[npClean.length - 2] : null;
    let npChange = prevNP !== null ? pctChange(prevNP, latestNP) : 0;
    if (latestNP > 0 && npChange > 0) { npSignal = 'green'; score += 2; scoreParts.push({ label: 'Profit ↑', points: +2, signal: 'green' }); findings.push({ signal: 'green', text: `Net Profit ₹${fmt(latestNP)} Cr, growing ${npChange > 0 ? '+' : ''}${npChange.toFixed(1)}% YoY.` }); }
    else if (latestNP > 0) { npSignal = 'amber'; scoreParts.push({ label: 'Profit ~', points: 0, signal: 'amber' }); findings.push({ signal: 'amber', text: `Net Profit ₹${fmt(latestNP)} Cr but declined ${npChange.toFixed(1)}% YoY.` }); }
    else { npSignal = 'red'; score -= 3; scoreParts.push({ label: 'Profit ↓', points: -3, signal: 'red' }); findings.push({ signal: 'red', text: `Net Profit in LOSS at ₹${fmt(latestNP)} Cr — major red flag.` }); }

    // 3. Quarterly trend
    const qNP = d.quarterly.netProfit.filter(v => v !== null);
    let qSignal = 'amber';
    if (qNP.length >= 2) {
      const lastQ = qNP.slice(-4);
      const negQs = lastQ.filter(v => v < 0).length;
      if (negQs >= 3) { qSignal = 'red'; score -= 3; scoreParts.push({ label: 'Quarters ✕', points: -3, signal: 'red' }); findings.push({ signal: 'red', text: `${negQs}/4 recent quarters show NET LOSSES — deteriorating rapidly.` }); }
      else if (negQs >= 1) { qSignal = 'amber'; score -= 1; scoreParts.push({ label: 'Quarters ~', points: -1, signal: 'amber' }); findings.push({ signal: 'amber', text: `${negQs}/4 recent quarters had losses — inconsistent.` }); }
      else { qSignal = 'green'; score += 1; scoreParts.push({ label: 'Quarters ✓', points: +1, signal: 'green' }); findings.push({ signal: 'green', text: `All recent quarters are profitable — healthy momentum.` }); }
    }

    // 4. Borrowings / Debt
    const borr = d.balance.borrowings.filter(v => v !== null);
    const resv = d.balance.reserves.filter(v => v !== null);
    let debtSignal = 'amber';
    let dToE = null;
    if (borr.length >= 2) {
      const bGrowth = pctChange(borr[borr.length - 2], borr[borr.length - 1]);
      const rGrowth = resv.length >= 2 ? pctChange(resv[resv.length - 2], resv[resv.length - 1]) : 0;
      const eq = d.balance.equity.filter(v => v !== null).slice(-1)[0] || 0;
      const netWorth = Math.max(eq + (resv[resv.length - 1] || 0), 1);
      dToE = borr[borr.length - 1] / netWorth;
      if (bGrowth > 20 && bGrowth > rGrowth) { debtSignal = 'red'; score -= 2; scoreParts.push({ label: 'Debt ↑', points: -2, signal: 'red' }); findings.push({ signal: 'red', text: `Borrowings surging +${bGrowth.toFixed(0)}% vs reserves +${rGrowth.toFixed(0)}%. D/E = ${dToE.toFixed(1)}x.` }); }
      else if (dToE > 3) { debtSignal = 'red'; score -= 1; scoreParts.push({ label: 'Debt high', points: -1, signal: 'red' }); findings.push({ signal: 'red', text: `High D/E of ${dToE.toFixed(1)}x — leverage risk elevated.` }); }
      else if (dToE < 1) { debtSignal = 'green'; score += 1; scoreParts.push({ label: 'Low Debt', points: +1, signal: 'green' }); findings.push({ signal: 'green', text: `Low D/E of ${dToE.toFixed(1)}x — strong balance sheet.` }); }
      else { scoreParts.push({ label: 'Debt ok', points: 0, signal: 'amber' }); findings.push({ signal: 'amber', text: `D/E is ${dToE.toFixed(1)}x. Borrowings grew ${bGrowth.toFixed(0)}%.` }); }
    }

    // 5. OPM
    const opmClean = d.annual.opm.filter(v => v !== null);
    let opmSignal = 'amber', latestOPM = opmClean[opmClean.length - 1] ?? null;
    if (latestOPM !== null) {
      if (latestOPM > 15) { opmSignal = 'green'; score += 1; scoreParts.push({ label: 'OPM ✓', points: +1, signal: 'green' }); findings.push({ signal: 'green', text: `OPM healthy at ${latestOPM.toFixed(1)}%.` }); }
      else if (latestOPM > 0) { opmSignal = 'amber'; scoreParts.push({ label: 'OPM thin', points: 0, signal: 'amber' }); findings.push({ signal: 'amber', text: `OPM thin at ${latestOPM.toFixed(1)}%.` }); }
      else { opmSignal = 'red'; score -= 2; scoreParts.push({ label: 'OPM ✕', points: -2, signal: 'red' }); findings.push({ signal: 'red', text: `OPM NEGATIVE at ${latestOPM.toFixed(1)}% — core ops loss-making.` }); }
    }

    // 6. Cash Flow
    const cfOp = d.cashflow.operating.filter(v => v !== null);
    let cfSignal = 'amber';
    if (cfOp.length >= 1) {
      const latestCF = cfOp[cfOp.length - 1];
      if (latestCF > 0) { cfSignal = 'green'; score += 1; scoreParts.push({ label: 'CF ✓', points: +1, signal: 'green' }); findings.push({ signal: 'green', text: `Operating Cash Flow positive ₹${fmt(latestCF)} Cr.` }); }
      else { cfSignal = 'red'; score -= 1; scoreParts.push({ label: 'CF ✕', points: -1, signal: 'red' }); findings.push({ signal: 'red', text: `Operating Cash Flow negative ₹${fmt(latestCF)} Cr — burning cash.` }); }
    }

    // 7. Expense spike
    const expClean = d.annual.expenses.filter(v => v !== null);
    if (expClean.length >= 2) {
      const expG = pctChange(expClean[expClean.length - 2], expClean[expClean.length - 1]);
      const revG = salesClean.length >= 2 ? pctChange(salesClean[salesClean.length - 2], salesClean[salesClean.length - 1]) : 0;
      if (expG > revG + 20) {
        score -= 2;
        scoreParts.push({ label: 'Exp spike', points: -2, signal: 'red' });
        findings.push({ signal: 'red', text: `Expenses spiked +${expG.toFixed(0)}% vs revenue +${revG.toFixed(0)}% — likely provisions/NPA write-offs.` });
      }
    }

    // 8. Valuation P/B
    let pb = null;
    if (d.currentPrice && d.numShares) {
      const eq = d.balance.equity.filter(v => v !== null).slice(-1)[0] || 0;
      const rs = resv.length ? resv[resv.length - 1] : 0;
      const bv = (eq + rs) / Math.max(d.numShares, 0.01);
      pb = d.currentPrice / Math.max(bv, 0.01);
      if (pb < 1 && npSignal === 'red') { findings.push({ signal: 'amber', text: `P/B = ${pb.toFixed(2)}x — looks cheap but may be a value trap.` }); }
      else if (pb < 1) { findings.push({ signal: 'green', text: `Below book value (P/B = ${pb.toFixed(2)}x) — potentially undervalued.` }); score += 1; }
      else if (pb > 5) { findings.push({ signal: 'amber', text: `P/B = ${pb.toFixed(2)}x — valuation is rich.` }); }
    }

    // 9. Banking-specific: detect if Interest is a major revenue component
    const isBanking = salesClean.length > 0 && d.annual.interest.filter(v => v !== null).length > 0 &&
      (d.annual.interest.filter(v => v !== null).slice(-1)[0] || 0) / (salesClean[salesClean.length - 1] || 1) > 0.2;
    if (isBanking) {
      findings.push({ signal: 'amber', text: `🏦 Banking stock detected. Interest expenses are ${((d.annual.interest.filter(v => v !== null).slice(-1)[0] || 0) / salesClean[salesClean.length - 1] * 100).toFixed(0)}% of revenue — asset quality is critical.` });
      // NIM approximation
      const latestInterestIncome = salesClean[salesClean.length - 1] || 0;
      const latestInterestExpense = d.annual.interest.filter(v => v !== null).slice(-1)[0] || 0;
      const totalAssets = d.balance.total.filter(v => v !== null).slice(-1)[0] || 1;
      const nim = ((latestInterestIncome - latestInterestExpense) / totalAssets * 100);
      if (nim > 0) findings.push({ signal: nim > 3 ? 'green' : 'amber', text: `Net Interest Margin (NIM) ≈ ${nim.toFixed(2)}%.` });
    }

    // Compute extra ratios
    const ratios = computeRatios(d, dToE, pb);

    // Normalize score to 0-10 for gauge
    const maxScore = 10;
    const gaugeScore = Math.max(0, Math.min(maxScore, Math.round((score + 8) / 16 * 10)));

    // Verdict
    let verdict, verdictClass;
    if (score >= 4) { verdict = 'Strong Buy'; verdictClass = 'buy'; }
    else if (score >= 2) { verdict = 'Buy'; verdictClass = 'buy'; }
    else if (score >= 0) { verdict = 'Hold'; verdictClass = 'hold'; }
    else if (score >= -3) { verdict = 'Avoid'; verdictClass = 'avoid'; }
    else { verdict = 'High Risk — Avoid'; verdictClass = 'avoid'; }

    return {
      verdict, verdictClass, score, gaugeScore, findings, scoreParts, ratios, isBanking,
      indicators: [
        { label: 'Revenue Growth', value: revenueGrowth !== null ? `${revenueGrowth.toFixed(1)}%` : 'N/A', detail: 'CAGR', signal: revenueSignal, sparkData: salesClean, tooltip: 'Compound Annual Growth Rate of revenue over available years.' },
        { label: 'Net Profit', value: `₹${fmt(latestNP)} Cr`, detail: prevNP !== null ? `${npChange >= 0 ? '+' : ''}${npChange.toFixed(1)}% YoY` : '', signal: npSignal, sparkData: npClean, tooltip: 'Annual net profit and year-over-year change.' },
        { label: 'Quarterly Trend', value: qNP.length ? `₹${fmt(qNP[qNP.length - 1])} Cr` : 'N/A', detail: 'Latest Quarter', signal: qSignal, sparkData: qNP.slice(-8), tooltip: 'Net profit of the most recent quarter. Red if losses in 3+ of last 4 quarters.' },
        { label: 'Operating Margin', value: latestOPM !== null ? `${latestOPM.toFixed(1)}%` : 'N/A', detail: 'OPM', signal: opmSignal, sparkData: opmClean, tooltip: 'Operating Profit Margin. Green above 15%, red if negative.' },
        { label: 'Debt Health', value: borr.length ? `₹${fmt(borr[borr.length - 1])} Cr` : 'N/A', detail: dToE !== null ? `D/E: ${dToE.toFixed(1)}x` : 'Borrowings', signal: debtSignal, sparkData: borr, tooltip: 'Total borrowings and Debt-to-Equity ratio. Low D/E is better.' },
        { label: 'Cash Flow', value: cfOp.length ? `₹${fmt(cfOp[cfOp.length - 1])} Cr` : 'N/A', detail: 'Operating CF', signal: cfSignal, sparkData: cfOp, tooltip: 'Operating Cash Flow. Positive means the business generates cash.' },
      ]
    };
  }

  // ── Compute Extra Ratios ──
  function computeRatios(d, dToE, pb) {
    const ratios = [];
    const npClean = d.annual.netProfit.filter(v => v !== null);
    const eq = d.balance.equity.filter(v => v !== null).slice(-1)[0] || 0;
    const resv = d.balance.reserves.filter(v => v !== null).slice(-1)[0] || 0;
    const netWorth = eq + resv;
    const totalAssets = d.balance.total.filter(v => v !== null).slice(-1)[0] || 0;
    const latestNP = npClean.length ? npClean[npClean.length - 1] : 0;
    const shares = d.numShares || 1;

    // EPS — both NP and shares are in Cr, so division gives ₹/share directly
    const eps = latestNP / shares;
    ratios.push({ label: 'EPS', value: `₹${eps.toFixed(2)}`, signal: eps > 0 ? 'green' : 'red' });

    // P/E
    if (d.currentPrice && d.currentPrice > 0 && eps > 0) {
      const pe = d.currentPrice / eps;
      ratios.push({ label: 'P/E Ratio', value: pe.toFixed(1), signal: pe < 15 ? 'green' : pe < 30 ? 'amber' : 'red' });
    } else {
      ratios.push({ label: 'P/E Ratio', value: 'N/A', signal: 'amber' });
    }

    // P/B
    if (pb !== null) ratios.push({ label: 'P/B Ratio', value: pb.toFixed(2), signal: pb < 1 ? 'green' : pb < 3 ? 'amber' : 'red' });

    // ROE
    if (netWorth > 0) {
      const roe = (latestNP / netWorth) * 100;
      ratios.push({ label: 'ROE', value: `${roe.toFixed(1)}%`, signal: roe > 15 ? 'green' : roe > 0 ? 'amber' : 'red' });
    }

    // ROA
    if (totalAssets > 0) {
      const roa = (latestNP / totalAssets) * 100;
      ratios.push({ label: 'ROA', value: `${roa.toFixed(2)}%`, signal: roa > 1 ? 'green' : roa > 0 ? 'amber' : 'red' });
    }

    // D/E
    if (dToE !== null) ratios.push({ label: 'D/E Ratio', value: dToE.toFixed(2), signal: dToE < 1 ? 'green' : dToE < 3 ? 'amber' : 'red' });

    // Dividend Yield
    const div = d.annual.dividend.filter(v => v !== null).slice(-1)[0] || 0;
    if (d.currentPrice && d.marketCap && div > 0) {
      const divYield = (div / d.marketCap) * 100;
      ratios.push({ label: 'Div. Yield', value: `${divYield.toFixed(2)}%`, signal: divYield > 1 ? 'green' : 'amber' });
    }

    return ratios;
  }

  // ════════════════════════════════════════════════
  //  RENDERER
  // ════════════════════════════════════════════════
  function renderDashboard(data, analysis) {
    charts.forEach(c => c.destroy()); charts = [];
    sparkCharts.forEach(c => c.destroy()); sparkCharts = [];

    // Header
    $('companyName').textContent = data.companyName;
    $('metaPrice').textContent = data.currentPrice ? `Price: ₹${data.currentPrice}` : '';
    $('metaMcap').textContent = data.marketCap ? `MCap: ₹${fmt(data.marketCap)} Cr` : '';
    $('metaFV').textContent = data.faceValue ? `FV: ₹${data.faceValue}` : '';

    // Gauge
    renderGauge(analysis.gaugeScore);

    // Verdict
    const vb = $('verdictBadge');
    vb.className = 'verdict-badge-large ' + analysis.verdictClass;
    vb.textContent = analysis.verdict;
    $('verdictText').textContent = analysis.findings.length ? analysis.findings[0].text : 'Analysis complete.';

    // Score breakdown tags
    const sb = $('scoreBreakdown');
    sb.innerHTML = '';
    analysis.scoreParts.forEach(p => {
      const tag = document.createElement('span');
      tag.className = 'score-tag ' + p.signal;
      tag.textContent = `${p.label} (${p.points > 0 ? '+' : ''}${p.points})`;
      sb.appendChild(tag);
    });

    // Indicator Cards with sparklines + tooltips
    renderIndicators(analysis.indicators);

    // Ratios row
    renderRatios(analysis.ratios);

    // Charts
    renderCharts(data);

    // Metrics table
    renderMetricsTable(data);

    // Report
    const rl = $('reportList');
    rl.innerHTML = '';
    analysis.findings.forEach(f => {
      const li = document.createElement('li');
      li.innerHTML = `<span class="bullet ${f.signal}"></span><span>${f.text}</span>`;
      rl.appendChild(li);
    });

    // Animate numbers
    animateCounters();

    // Render compare if multiple
    if (compareStocks.length > 1) renderCompare();
  }

  // ── Score Gauge ──
  function renderGauge(score) {
    const arc = document.getElementById('gaugeFill');
    // Arc length: semicircle with r=75
    const totalLen = Math.PI * 75; // ~235.6
    arc.setAttribute('stroke-dasharray', totalLen);
    const offset = totalLen - (score / 10) * totalLen;
    // Initial state
    arc.setAttribute('stroke-dashoffset', totalLen);
    // Color
    let color = score >= 7 ? '#22c55e' : score >= 4 ? '#f59e0b' : '#ef4444';
    arc.setAttribute('stroke', color);
    // Animate
    requestAnimationFrame(() => {
      setTimeout(() => { arc.setAttribute('stroke-dashoffset', offset); }, 100);
    });
    // Score text with counter animation
    const el = $('gaugeScoreText');
    el.textContent = '0';
    el.style.color = color;
    animateValue(el, 0, score, 1200);
  }

  // ── Indicator Cards ──
  function renderIndicators(indicators) {
    const grid = $('indicatorsGrid');
    grid.innerHTML = '';
    indicators.forEach((ind, i) => {
      const card = document.createElement('div');
      card.className = 'indicator-card';
      card.innerHTML = `
        <div class="signal-dot ${ind.signal}"></div>
        <div class="tooltip-text">${ind.tooltip}</div>
        <div class="card-label">${ind.label}</div>
        <div class="card-value ${ind.signal} counter-target" data-value="${ind.value}">${ind.value}</div>
        <div class="card-detail">${ind.detail}</div>
        <div class="sparkline-wrap"><canvas id="spark${i}"></canvas></div>
      `;
      grid.appendChild(card);
    });

    // Render sparklines
    indicators.forEach((ind, i) => {
      if (ind.sparkData && ind.sparkData.length >= 2) {
        const ctx = document.getElementById(`spark${i}`);
        const isPositive = ind.sparkData[ind.sparkData.length - 1] >= (ind.sparkData[ind.sparkData.length - 2] || 0);
        const color = ind.signal === 'green' ? '#22c55e' : ind.signal === 'red' ? '#ef4444' : '#f59e0b';
        const c = new Chart(ctx, {
          type: 'line',
          data: {
            labels: ind.sparkData.map((_, idx) => idx),
            datasets: [{
              data: ind.sparkData,
              borderColor: color,
              borderWidth: 1.5,
              fill: true,
              backgroundColor: color.replace(')', ',0.1)').replace('rgb', 'rgba'),
              pointRadius: 0,
              tension: 0.4
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false }, tooltip: { enabled: false } },
            scales: { x: { display: false }, y: { display: false } },
            animation: { duration: 800 }
          }
        });
        sparkCharts.push(c);
      }
    });
  }

  // ── Ratios Row ──
  function renderRatios(ratios) {
    const grid = $('ratiosGrid');
    grid.innerHTML = '';
    ratios.forEach(r => {
      const card = document.createElement('div');
      card.className = 'ratio-card';
      card.innerHTML = `
        <div class="ratio-label">${r.label}</div>
        <div class="ratio-value ${r.signal}">${r.value}</div>
      `;
      grid.appendChild(card);
    });
  }

  // ── Charts ──
  function renderCharts(data) {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
    const tickColor = isDark ? '#64748b' : '#94a3b8';

    const defaults = {
      responsive: true, maintainAspectRatio: true,
      plugins: {
        legend: { labels: { color: tickColor, font: { family: 'Inter', size: 11 } } },
        tooltip: { backgroundColor: isDark ? 'rgba(17,24,39,0.95)' : 'rgba(255,255,255,0.95)', titleColor: isDark ? '#f1f5f9' : '#1e293b', bodyColor: tickColor, borderColor: 'rgba(128,128,128,0.2)', borderWidth: 1, cornerRadius: 8, padding: 12 }
      },
      scales: {
        x: { ticks: { color: tickColor, font: { size: 10 } }, grid: { color: gridColor } },
        y: { ticks: { color: tickColor, font: { size: 10 } }, grid: { color: gridColor } }
      }
    };

    // 1. Annual Revenue & NP
    const aLabels = cleanLabels(data.annual.dates);
    charts.push(new Chart($('chartAnnual'), {
      type: 'line',
      data: {
        labels: aLabels,
        datasets: [
          { label: 'Revenue', data: cleanData(data.annual.sales, aLabels), borderColor: '#3b82f6', backgroundColor: 'rgba(59,130,246,0.1)', fill: true, tension: 0.35, pointRadius: 4, pointHoverRadius: 7 },
          { label: 'Net Profit', data: cleanData(data.annual.netProfit, aLabels), borderColor: '#22c55e', backgroundColor: 'rgba(34,197,94,0.1)', fill: true, tension: 0.35, pointRadius: 4, pointHoverRadius: 7 }
        ]
      },
      options: defaults
    }));

    // 2. Quarterly Net Profit
    const qLabels = cleanLabels(data.quarterly.dates);
    const qPD = cleanData(data.quarterly.netProfit, qLabels);
    charts.push(new Chart($('chartQuarterly'), {
      type: 'bar',
      data: {
        labels: qLabels,
        datasets: [{ label: 'Net Profit', data: qPD, backgroundColor: qPD.map(v => v >= 0 ? 'rgba(34,197,94,0.7)' : 'rgba(239,68,68,0.7)'), borderColor: qPD.map(v => v >= 0 ? '#22c55e' : '#ef4444'), borderWidth: 1, borderRadius: 4 }]
      },
      options: defaults
    }));

    // 3. Balance Sheet
    const bLabels = cleanLabels(data.balance.dates);
    charts.push(new Chart($('chartBalance'), {
      type: 'bar',
      data: {
        labels: bLabels,
        datasets: [
          { label: 'Equity + Reserves', data: bLabels.map((_, i) => (data.balance.equity[i] || 0) + (data.balance.reserves[i] || 0)), backgroundColor: 'rgba(34,197,94,0.65)', borderRadius: 2 },
          { label: 'Borrowings', data: cleanData(data.balance.borrowings, bLabels), backgroundColor: 'rgba(239,68,68,0.55)', borderRadius: 2 },
          { label: 'Other Liabilities', data: cleanData(data.balance.otherLiab, bLabels), backgroundColor: 'rgba(245,158,11,0.45)', borderRadius: 2 }
        ]
      },
      options: { ...defaults, scales: { ...defaults.scales, x: { ...defaults.scales.x, stacked: true }, y: { ...defaults.scales.y, stacked: true } } }
    }));

    // 4. Cash Flow
    const cfLabels = cleanLabels(data.balance.dates);
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
      options: defaults
    }));

    // 5. Quarterly OPM Trend (NEW)
    const qOpmData = cleanData(data.quarterly.opm, qLabels);
    charts.push(new Chart($('chartOPM'), {
      type: 'line',
      data: {
        labels: qLabels,
        datasets: [{
          label: 'OPM %',
          data: qOpmData,
          borderColor: '#a855f7',
          backgroundColor: 'rgba(168,85,247,0.1)',
          fill: true,
          tension: 0.35,
          pointRadius: 5,
          pointHoverRadius: 8,
          pointBackgroundColor: qOpmData.map(v => v >= 0 ? '#a855f7' : '#ef4444')
        }]
      },
      options: { ...defaults, plugins: { ...defaults.plugins, annotation: {} } }
    }));
  }

  // ── Metrics Table ──
  function renderMetricsTable(data) {
    const headerRow = $('metricsHeader');
    const tbody = $('metricsBody');
    headerRow.innerHTML = ''; tbody.innerHTML = '';

    const dates = data.annual.dates.filter(d => d);
    const n = data.annual.sales.length;
    if (n < 2) return;

    const prevI = n - 2, currI = n - 1;
    headerRow.innerHTML = `<th>Metric</th><th>${dates[prevI] || 'Prev'}</th><th>${dates[currI] || 'Latest'}</th><th>Change</th>`;

    const metrics = [
      { name: 'Sales (Revenue)', prev: data.annual.sales[prevI], curr: data.annual.sales[currI] },
      { name: 'Net Profit', prev: data.annual.netProfit[prevI], curr: data.annual.netProfit[currI] },
      { name: 'Interest', prev: data.annual.interest[prevI], curr: data.annual.interest[currI] },
      { name: 'Borrowings', prev: data.balance.borrowings[prevI], curr: data.balance.borrowings[currI] },
      { name: 'Reserves', prev: data.balance.reserves[prevI], curr: data.balance.reserves[currI] },
      { name: 'Other Expenses', prev: data.annual.expenses[prevI], curr: data.annual.expenses[currI] },
    ];
    const positiveMetrics = ['Net Profit', 'Reserves', 'Sales (Revenue)'];

    metrics.forEach(m => {
      if (m.prev == null && m.curr == null) return;
      const change = (m.prev && m.curr) ? ((m.curr - m.prev) / Math.abs(m.prev) * 100) : null;
      const isPositiveGood = positiveMetrics.includes(m.name);
      const cls = change === null ? '' : change > 5 ? (isPositiveGood ? 'positive' : 'negative') : change < -5 ? (isPositiveGood ? 'negative' : 'positive') : 'neutral';
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${m.name}</td><td>${m.prev != null ? '₹' + fmt(m.prev) + ' Cr' : '—'}</td><td>${m.curr != null ? '₹' + fmt(m.curr) + ' Cr' : '—'}</td><td class="${cls}">${change !== null ? (change >= 0 ? '+' : '') + change.toFixed(1) + '%' : '—'}</td>`;
      tbody.appendChild(tr);
    });
  }

  // ── Compare Table ──
  function renderCompare() {
    const header = $('compareHeader');
    const body = $('compareBody');
    const chips = $('compareChips');
    header.innerHTML = ''; body.innerHTML = ''; chips.innerHTML = '';

    if (compareStocks.length < 2) {
      $('compareBar').classList.add('hidden');
      $('compareSection').classList.add('hidden');
      return;
    }

    // Chips
    compareStocks.forEach((s, i) => {
      const chip = document.createElement('span');
      chip.className = 'compare-chip';
      chip.innerHTML = `${esc(s.data.companyName)} <button class="chip-close" data-idx="${i}">✕</button>`;
      chips.appendChild(chip);
    });
    chips.querySelectorAll('.chip-close').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.idx);
        if (idx === 0) return; // Can't remove primary
        compareStocks.splice(idx, 1);
        renderCompare();
      });
    });

    // Table
    header.innerHTML = '<th>Metric</th>' + compareStocks.map(s => `<th>${esc(s.data.companyName)}</th>`).join('');

    const rows = [
      { label: 'Verdict', values: compareStocks.map(s => `<span class="hi-verdict ${s.analysis.verdictClass}" style="display:inline-block">${s.analysis.verdict}</span>`) },
      { label: 'Score', values: compareStocks.map(s => `${s.analysis.gaugeScore}/10`) },
      { label: 'Price', values: compareStocks.map(s => `₹${s.data.currentPrice}`) },
      { label: 'Market Cap', values: compareStocks.map(s => `₹${fmt(s.data.marketCap)} Cr`) },
      { label: 'Revenue (Latest)', values: compareStocks.map(s => { const v = s.data.annual.sales.filter(v => v !== null).slice(-1)[0]; return v ? `₹${fmt(v)} Cr` : '—'; }) },
      { label: 'Net Profit', values: compareStocks.map(s => { const v = s.data.annual.netProfit.filter(v => v !== null).slice(-1)[0]; return v ? `₹${fmt(v)} Cr` : '—'; }) },
      { label: 'OPM', values: compareStocks.map(s => { const v = s.data.annual.opm.filter(v => v !== null).slice(-1)[0]; return v ? `${v.toFixed(1)}%` : '—'; }) },
    ];

    // Add ratios from first stock's ratio labels
    if (compareStocks[0].analysis.ratios.length) {
      compareStocks[0].analysis.ratios.forEach((r, ri) => {
        rows.push({
          label: r.label,
          values: compareStocks.map(s => s.analysis.ratios[ri] ? s.analysis.ratios[ri].value : '—')
        });
      });
    }

    rows.forEach(r => {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${r.label}</td>` + r.values.map(v => `<td>${v}</td>`).join('');
      body.appendChild(tr);
    });

    $('compareSection').classList.remove('hidden');
  }

  // ── Animated Counters ──
  function animateValue(el, start, end, duration) {
    const startTime = performance.now();
    const update = (now) => {
      const elapsed = now - startTime;
      const progress = Math.min(elapsed / duration, 1);
      const eased = 1 - Math.pow(1 - progress, 3); // easeOutCubic
      const current = Math.round(start + (end - start) * eased);
      el.textContent = current;
      if (progress < 1) requestAnimationFrame(update);
    };
    requestAnimationFrame(update);
  }

  function animateCounters() {
    // Animate ratio cards
    document.querySelectorAll('.ratio-value').forEach(el => {
      const text = el.textContent;
      const match = text.match(/^(₹?)([\d,.]+)(.*)$/);
      if (match) {
        const prefix = match[1];
        const numStr = match[2].replace(/,/g, '');
        const suffix = match[3];
        const num = parseFloat(numStr);
        if (!isNaN(num) && num < 10000) {
          const decimals = numStr.includes('.') ? numStr.split('.')[1].length : 0;
          const startTime = performance.now();
          const duration = 1000;
          const update = (now) => {
            const progress = Math.min((now - startTime) / duration, 1);
            const eased = 1 - Math.pow(1 - progress, 3);
            const val = (num * eased).toFixed(decimals);
            el.textContent = prefix + val + suffix;
            if (progress < 1) requestAnimationFrame(update);
          };
          requestAnimationFrame(update);
        }
      }
    });
  }

  // ════════════════════════════════════════════════
  //  UTILITIES
  // ════════════════════════════════════════════════
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
      return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}`;
    }
    const s = String(v);
    const m = s.match(/(\d{4})[-\/]?(\d{2})?/);
    if (m) return m[2] ? `${m[1]}-${m[2]}` : m[1];
    return s.substring(0, 10);
  }

  function cleanLabels(arr) { return arr.filter(v => v != null && v !== ''); }
  function cleanData(arr, labels) {
    const clean = arr.filter(v => v !== null);
    if (clean.length >= labels.length) return clean.slice(clean.length - labels.length);
    return new Array(labels.length - clean.length).fill(null).concat(clean);
  }

  function fmt(n) {
    if (n === null || n === undefined) return '—';
    const abs = Math.abs(n);
    if (abs >= 100) return n.toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return n.toFixed(2);
  }

  function esc(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }

})();
