/*
 * Saint Daily Reporting — Excel workbook builder (ExcelJS).
 * The "Daily Reporting" sheet is linked by live formulas to the supporting sheets, so every
 * figure can be traced back to its source in Excel. Cached values are written too, so the file
 * shows correct numbers even before Excel recalculates.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.SaintReportExport = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const round2 = v => { v = Number(v) || 0; const s = v < 0 ? -1 : 1; return s * Math.round((Math.abs(v) + Number.EPSILON) * 100) / 100; };
  const utcDate = iso => { const [y, m, d] = iso.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
  const colL = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const q = name => `'${name.replace(/'/g, "''")}'`;
  const fmtDMY = iso => { const [y, m, d] = iso.split('-'); return `${d}/${m}/${y}`; };

  const C = {
    head: 'FFD6E3BC', label: 'FFEAF1DD', total: 'FFC9DAF8', revenue: 'FFD0E0E3', profit: 'FFD9D9D9',
    gstTitle: 'FFFFFF00', gstSub: 'FFD9EAD3', warn: 'FFFFF2CC', white: 'FFFFFFFF'
  };
  const NUM = '#,##0.00;[Red]-#,##0.00';
  const PCT = '0.00%';
  const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const thin = { style: 'thin', color: { argb: 'FFBFBFBF' } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };

  function setF(cell, formula, result, numFmt) {
    cell.value = { formula, result: typeof result === 'number' ? (Number.isFinite(result) ? result : 0) : result };
    if (numFmt) cell.numFmt = numFmt;
  }
  function headerRow(row, argb = C.head) {
    row.eachCell(c => { c.font = { bold: true }; c.fill = fill(argb); c.border = border; c.alignment = { vertical: 'middle', wrapText: true }; });
  }

  function build(result, opts) {
    const ExcelJS = opts.ExcelJS;
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Saint Daily Reporting';
    wb.created = new Date();
    wb.calcProperties = { fullCalcOnLoad: true };

    const dates = result.dates, N = dates.length, rows = result.rows, T = result.totals;
    const has = result.has || { sales: true, net: true, revenue: true, cogs: true, shopCogs: true, merchant: true };
    const notUploaded = (ws, what) => { ws.getCell('A1').value = `${what} was not uploaded for this period, so this sheet is empty.`; ws.getCell('A1').font = { italic: true, color: { argb: 'FF9A6200' } }; ws.getColumn(1).width = 80; };
    const firstMonth = utcDate(dates[0]).getUTCMonth(), lastMonth = utcDate(dates.at(-1)).getUTCMonth();
    const monthName = utcDate(dates[0]).toLocaleString('en', { month: 'long', timeZone: 'UTC' });
    const periodText = `${fmtDMY(dates[0])} to ${fmtDMY(dates.at(-1))}`;

    const S = {
      DR: 'Daily Reporting',
      COGS: 'COGS Calculation as per Supplie',
      CW: 'Countrywise Sales',
      PW: 'Productwise Sales',
      BR: (firstMonth === lastMonth ? `${monthName} Breakdown Sheet` : 'Period Breakdown Sheet').slice(0, 31),
      DIFF: 'Difference COGS',
      FEES: 'Fees Calculation',
      JE: 'Journal Entries',
      RATIO: 'Ratio Calculation Sheet',
      NOTES: 'Checks & Notes'
    };
    // Create in display order.
    const ws = {};
    Object.entries(S).forEach(([k, name]) => ws[k] = wb.addWorksheet(name));

    // ---------------- Breakdown sheet (Sales reconciliation per day) ----------------
    const recCols = result.salesResult.reconciliationColumns;
    const recBy = new Map(result.salesResult.reconciliationRows.map(r => [r.Date, r]));
    const br = ws.BR;
    if (!has.revenue) notUploaded(br, 'Net payments / Total sales');
    else {
    br.addRow(recCols.map(c => c === 'rounding' ? 'Rounding' : c));
    headerRow(br.getRow(1));
    br.getRow(1).height = 45;
    dates.forEach(d => {
      const r = recBy.get(d) || {};
      br.addRow(recCols.map(c => c === 'Date' ? utcDate(d) : round2(r[c] || 0)));
    });
    br.getColumn(1).numFmt = 'd-mmm-yyyy';
    br.getColumn(1).width = 13;
    for (let c = 2; c <= recCols.length; c++) { br.getColumn(c).width = 15; br.getColumn(c).numFmt = NUM; }
    br.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];
    }
    const brLast = colL(Math.max(recCols.length, 2));
    const brCol = name => colL(recCols.indexOf(name) + 1);
    const brHdr = `${q(S.BR)}!$B$1:$${brLast}$1`;
    const brRow = r => `${q(S.BR)}!$B${r}:$${brLast}${r}`;

    // ---------------- Fees Calculation ----------------
    const fe = ws.FEES;
    const feeFirst = 2;
    let feeLast = 2;
    if (!has.merchant) notUploaded(fe, 'The merchant fee workbook');
    else {
    fe.addRow(['Date', 'Payment gateway', 'Number of orders', 'Gross sales (A$)', 'Merchant Exclude GST(A$)', 'GST on Fee', 'Merchant fee (A$)', 'Net deposit/payout (A$)', 'Approx Effective fee rate approx (%)']);
    headerRow(fe.getRow(1));
    fe.getRow(1).height = 32;
    let prevPlatform = null;
    result.merchant.rows.forEach(m => {
      if (prevPlatform && prevPlatform !== m.platform) fe.addRow([]);
      prevPlatform = m.platform;
      const r = fe.addRow([utcDate(m.date), m.platform, m.orders, m.gross, m.exGST, m.gst]);
      const n = r.number;
      setF(r.getCell(7), `E${n}+F${n}`, round2(m.exGST + m.gst));
      setF(r.getCell(8), `D${n}-G${n}`, round2(m.gross - m.exGST - m.gst));
      setF(r.getCell(9), `IFERROR(G${n}/D${n},0)`, m.gross ? (m.exGST + m.gst) / m.gross : 0, '0.0%');
    });
    feeLast = Math.max(fe.rowCount, 2);
    fe.addRow([]);
    const ft = fe.addRow(['Total']);
    const sumMer = k => round2(result.merchant.rows.reduce((z, r) => z + r[k], 0));
    [['C', 'orders'], ['D', 'gross'], ['E', 'exGST'], ['F', 'gst']].forEach(([L, k]) => setF(ft.getCell(L), `SUM(${L}${feeFirst}:${L}${feeLast})`, sumMer(k)));
    setF(ft.getCell('G'), `SUM(G${feeFirst}:G${feeLast})`, round2(sumMer('exGST') + sumMer('gst')));
    setF(ft.getCell('H'), `SUM(H${feeFirst}:H${feeLast})`, round2(sumMer('gross') - sumMer('exGST') - sumMer('gst')));
    setF(ft.getCell('I'), `IFERROR(G${ft.number}/D${ft.number},0)`, sumMer('gross') ? (sumMer('exGST') + sumMer('gst')) / sumMer('gross') : 0, '0.0%');
    ft.eachCell(c => { c.font = { bold: true }; c.fill = fill(C.total); });
    fe.getColumn(1).numFmt = 'd-mmm-yyyy';
    [13, 16, 15, 16, 20, 12, 16, 20, 20].forEach((w, i) => fe.getColumn(i + 1).width = w);
    ['D', 'E', 'F', 'G', 'H'].forEach(L => fe.getColumn(L).numFmt = NUM);
    fe.getColumn('I').numFmt = '0.0%';
    fe.views = [{ state: 'frozen', ySplit: 1 }];
    }
    const feeRange = L => `${q(S.FEES)}!$${L}$${feeFirst}:$${L}$${Math.max(feeLast, feeFirst)}`;

    // ---------------- COGS Calculation ----------------
    const cg = ws.COGS;
    const cogsRow = {};
    if (!has.cogs) notUploaded(cg, 'The supplier COGS workbook');
    else {
    cg.addRow(['Supplier COGS (Saint COGS output) — ' + periodText]);
    cg.getRow(1).font = { bold: true };
    cg.addRow(['Date', 'Product Price($)', 'Shipping Cost ($)', 'Total(Product & Shipping cost)', 'FX rate (USD→AUD)', 'Product Cost AUD', 'Shipping Cost AUD', 'Convert It in AUD', 'GST', 'COGS ex GST', 'Status', 'Note']);
    headerRow(cg.getRow(2));
    cg.getRow(2).height = 32;
    dates.forEach(d => {
      const c = result.cogs[d];
      const r = cg.addRow([utcDate(d), c.productUSD, c.shippingUSD]);
      const n = r.number; cogsRow[d] = n;
      setF(r.getCell(4), `B${n}+C${n}`, round2(c.productUSD + c.shippingUSD));
      if (c.rate) {
        r.getCell(5).value = c.rate;
        setF(r.getCell(6), `ROUND(B${n}*E${n},2)`, c.productAUD);
        setF(r.getCell(7), `ROUND(C${n}*E${n},2)`, c.shippingAUD);
      } else { r.getCell(6).value = c.productAUD; r.getCell(7).value = c.shippingAUD; }
      const tot = round2(c.productAUD + c.shippingAUD);
      setF(r.getCell(8), `F${n}+G${n}`, tot);
      setF(r.getCell(9), `H${n}/110*10`, tot / 110 * 10);
      setF(r.getCell(10), `H${n}/110*100`, tot / 110 * 100);
      r.getCell(11).value = c.status;
      r.getCell(12).value = c.note || '';
      if (c.status !== 'Ready') r.eachCell(x => x.fill = fill(C.warn));
    });
    const cgFirst = 3, cgLast = cg.rowCount;
    const cgt = cg.addRow(['Total']);
    const sumC = k => round2(dates.reduce((z, d) => z + result.cogs[d][k], 0));
    const cgTotVals = { B: sumC('productUSD'), C: sumC('shippingUSD'), F: sumC('productAUD'), G: sumC('shippingAUD') };
    cgTotVals.D = round2(cgTotVals.B + cgTotVals.C); cgTotVals.H = round2(cgTotVals.F + cgTotVals.G); cgTotVals.I = cgTotVals.H / 11; cgTotVals.J = cgTotVals.H / 1.1;
    Object.entries(cgTotVals).forEach(([L, v]) => setF(cgt.getCell(L), `SUM(${L}${cgFirst}:${L}${cgLast})`, v));
    cgt.eachCell(c => { c.font = { bold: true }; c.fill = fill(C.total); });
    if (result.cogsUnpriced.length) {
      cg.addRow([]);
      cg.addRow(['Lines with no product price in the COGS file (#N/A)']).font = { bold: true };
      const h = cg.addRow(['Date', 'Order', 'SKU', 'Product', 'Quantity', 'USD price used', 'Source row']);
      headerRow(h, C.warn);
      result.cogsUnpriced.forEach(u => cg.addRow([utcDate(u.date), u.order, u.sku, u.product, u.quantity, u.override === 'SKIP' ? 'SKIPPED — $0 as in file' : u.override != null && u.override !== '' ? Number(u.override) : 'NOT PRICED (0)', `${u.sheet}!${u.row}`]));
    }
    cg.getColumn(1).numFmt = 'd-mmm-yyyy';
    [13, 16, 16, 18, 14, 16, 16, 16, 12, 14, 10, 34].forEach((w, i) => cg.getColumn(i + 1).width = w);
    ['B', 'C', 'D', 'F', 'G', 'H', 'I', 'J'].forEach(L => cg.getColumn(L).numFmt = NUM);
    cg.getColumn('E').numFmt = '0.0000';
    cg.views = [{ state: 'frozen', ySplit: 2 }];
    }

    // ---------------- Daily Reporting ----------------
    const dr = ws.DR;
    const TC = N + 2;                  // total column index
    const TL = colL(TC), lastL = colL(N + 1);
    const L = i => colL(i + 2);        // date i → column letter
    const R = {};                      // key → row number
    const layout = [
      ['orderRange', 'Order range'], ['numberOrders', 'Number of orders'], ['gross', 'Gross sales (A$)'], ['discounts', 'Discounts (A$)'], ['returns', 'Returns'],
      ['netSales', 'Net sales (excluding shipping, A$)'], ['shipping', 'Shipping charges collected (A$)'], ['roundingTax', 'Rounding(US Tax)'], ['totalSales', 'Total sales (A$)'], [],
      ['productRevenue', 'Product revenue (ex GST, A$)'], ['shippingAU', 'Shipping revenue (ex GST, A$) - AU'], ['shippingOther', 'Shipping Revenue - Others'], ['refund', 'Refund'], ['refundExGST', 'Refund Ex GST'],
      ['totalRevenue', 'Total revenue (ex GST, A$)'], ['cogsProduct', 'COGS Product Cost Sheet – (A$)'], ['cogsShipping', 'COGS Shipping Cost Sheet – (A$)'], ['cogsShopify', 'COGS Shopify – (A$)'],
      ['grossProfit', 'Gross profit –  COGS (A$)'], ['merchantFees', 'Merchant fees – total (A$)'],
      ['meta', 'Advertisment Spend - Meta'], ['google', 'Advertisment Spend - Google'], ['pinterest', 'Advertisment Spend - Pinterest'], ['applovin', 'Advertisment Spend - Applovin'], ['snap', 'Advertisment Spend - Snap'], ['tiktok', 'Advertisment Spend - Tiktok'],
      ['adSpend', 'Advertisement spend (A$)'], ['netProfit', 'Net profit as per Suppliers COGS (A$)'], ['profitMargin', 'Profit Margin (as per Supplier COGS (A$))'], ['roas', 'ROAS (Revenue/Ad spend)'], [],
      ['_gst', 'GST Calculation'], ['_out', 'Output'], ['gstSales', 'GST on Sales (incl. GST on refunded sales)'], ['gstShipping', 'GST on Shipping '], ['_in', 'Input'], ['gstFee', 'GST on Fee '], ['gstRefundAU', 'GST on Refund- AU'], [],
      ['netGST', 'Net GST (Output − Input)'], [], ['netProfitAfterGST', 'Net Profit After GST remittance – cash view (Net profit − Net GST)'], [],
      ['_sum1', 'PERIOD PROFIT SUM (as per Supplier COGS) – accounting profit'], ['_sum2', 'PERIOD CASH AFTER GST REMITTANCE (cash view)']
    ];
    dr.addRow(['Metric', ...dates.map(() => 'Value'), 'Period total']);
    dr.addRow(['Date', ...dates.map(utcDate), `${dates.length} day(s)`]);
    layout.forEach(([k, label]) => { const r = dr.addRow(label ? [label] : []); if (k) R[k] = r.number; });

    const F = {   // formula templates per row; c = column letter, i = date index
      netSales: c => `SUM(${c}${R.gross}:${c}${R.returns})`,
      totalSales: c => `${c}${R.netSales}+${c}${R.shipping}+${c}${R.roundingTax}`,
      productRevenue: (c, i) => `SUMIF(${brHdr},"Sales Revenue (*",${brRow(i + 2)})`,
      shippingAU: (c, i) => `${q(S.BR)}!${brCol('Shipping (Australia, excl. GST)')}${i + 2}`,
      shippingOther: (c, i) => `SUMIF(${brHdr},"Shipping (*",${brRow(i + 2)})-${c}${R.shippingAU}`,
      refundExGST: (c, i) => `-SUMIF(${brHdr},"Refunds (*",${brRow(i + 2)})`,
      refund: c => `${c}${R.refundExGST}+${c}${R.gstRefundAU}`,
      totalRevenue: c => `${c}${R.productRevenue}+${c}${R.shippingAU}+${c}${R.shippingOther}-${c}${R.refundExGST}`,
      cogsProduct: (c, i) => `${q(S.COGS)}!F${cogsRow[dates[i]]}`,
      cogsShipping: (c, i) => `${q(S.COGS)}!G${cogsRow[dates[i]]}`,
      grossProfit: c => `${c}${R.totalRevenue}-${c}${R.cogsProduct}-${c}${R.cogsShipping}`,
      merchantFees: c => `SUMIFS(${feeRange('E')},${feeRange('A')},${c}$2)`,
      adSpend: c => `SUM(${c}${R.meta}:${c}${R.tiktok})`,
      netProfit: c => `${c}${R.grossProfit}-${c}${R.merchantFees}-${c}${R.adSpend}`,
      profitMargin: c => `IFERROR(${c}${R.netProfit}/${c}${R.totalRevenue},0)`,
      roas: c => `IFERROR(${c}${R.totalRevenue}/${c}${R.adSpend},0)`,
      gstSales: (c, i) => `${q(S.BR)}!${brCol('GST on Sales')}${i + 2}`,
      gstShipping: (c, i) => `${q(S.BR)}!${brCol('GST on Shipping')}${i + 2}`,
      gstFee: c => `SUMIFS(${feeRange('F')},${feeRange('A')},${c}$2)`,
      gstRefundAU: (c, i) => `-${q(S.BR)}!${brCol('GST on Refund')}${i + 2}`,
      netGST: c => `${c}${R.gstSales}+${c}${R.gstShipping}-${c}${R.gstFee}-${c}${R.gstRefundAU}`,
      netProfitAfterGST: c => `${c}${R.netProfit}-${c}${R.netGST}`
    };
    const ratioKeys = new Set(['profitMargin', 'roas']);
    const AD_KEYS = new Set(['meta', 'google', 'pinterest', 'applovin', 'snap', 'tiktok']);
    const fmtFor = k => k === 'profitMargin' ? PCT : k === 'roas' ? '#,##0.00' : k === 'numberOrders' ? '#,##0' : NUM;

    // Rows calculated from other rows. They keep their formula even when inputs are blank, wrapped in a
    // COUNT guard so they show nothing until every input is there — and fill themselves in if the blanks
    // are typed into Excel later. Rows that come straight from a file are simply left blank.
    const DERIVED = {
      netSales: c => [`${c}${R.gross}`, `${c}${R.discounts}`, `${c}${R.returns}`],
      totalSales: c => [`${c}${R.netSales}`, `${c}${R.shipping}`, `${c}${R.roundingTax}`],
      refund: c => [`${c}${R.refundExGST}`, `${c}${R.gstRefundAU}`],
      totalRevenue: c => [`${c}${R.productRevenue}`, `${c}${R.shippingAU}`, `${c}${R.shippingOther}`, `${c}${R.refundExGST}`],
      grossProfit: c => [`${c}${R.totalRevenue}`, `${c}${R.cogsProduct}`, `${c}${R.cogsShipping}`],
      netProfit: c => [`${c}${R.grossProfit}`, `${c}${R.merchantFees}`, `${c}${R.adSpend}`],
      profitMargin: c => [`${c}${R.netProfit}`, `${c}${R.totalRevenue}`],
      roas: c => [`${c}${R.totalRevenue}`, `${c}${R.adSpend}`],
      netGST: c => [`${c}${R.gstSales}`, `${c}${R.gstShipping}`, `${c}${R.gstFee}`, `${c}${R.gstRefundAU}`],
      netProfitAfterGST: c => [`${c}${R.netProfit}`, `${c}${R.netGST}`]
    };
    const guard = (cells, expr) => `IF(COUNT(${cells.join(',')})=${cells.length},${expr},"")`;
    const blank = (cell, k) => { cell.value = null; cell.numFmt = fmtFor(k); cell.fill = fill(C.warn); };

    layout.forEach(([k]) => {
      if (!k || k.startsWith('_')) return;
      const row = dr.getRow(R[k]);
      rows.forEach((x, i) => {
        const c = L(i), cell = row.getCell(i + 2);
        if (k === 'orderRange') { cell.value = x.orderRange || null; if (!x.orderRange) cell.fill = fill(C.warn); return; }
        if (k === 'adSpend') {   // blank until at least one ad figure is entered
          setF(cell, `IF(COUNT(${c}${R.meta}:${c}${R.tiktok})=0,"",SUM(${c}${R.meta}:${c}${R.tiktok}))`, x.adSpend == null ? '' : x.adSpend, fmtFor(k));
          return;
        }
        if (DERIVED[k]) { setF(cell, guard(DERIVED[k](c), F[k](c, i)), x[k] == null ? '' : x[k], fmtFor(k)); return; }
        if (x[k] == null) { blank(cell, k); return; }          // source row with no file → leave empty
        if (F[k]) setF(cell, F[k](c, i), x[k], fmtFor(k));
        else { cell.value = x[k]; cell.numFmt = fmtFor(k); }
      });
      // Period total: only once every day in the period has a figure.
      const tcell = row.getCell(TC);
      if (k === 'orderRange') { tcell.value = T.orderRange || null; return; }
      const span = `B${R[k]}:${lastL}${R[k]}`;
      if (ratioKeys.has(k)) setF(tcell, guard(DERIVED[k](TL), F[k](TL)), T[k] == null ? '' : T[k], fmtFor(k));
      else setF(tcell, `IF(COUNT(${span})=${N},SUM(${span}),"")`, T[k] == null ? '' : T[k], fmtFor(k));
    });
    const s1 = dr.getRow(R._sum1).getCell(2), s2 = dr.getRow(R._sum2).getCell(2);
    const spanSum = r => `IF(COUNT(B${r}:${lastL}${r})=${N},SUM(B${r}:${lastL}${r}),"")`;
    setF(s1, spanSum(R.netProfit), T.netProfit == null ? '' : T.netProfit, NUM);
    setF(s2, spanSum(R.netProfitAfterGST), T.netProfitAfterGST == null ? '' : T.netProfitAfterGST, NUM);

    // styling
    dr.getColumn(1).width = 44;
    for (let c = 2; c <= TC; c++) dr.getColumn(c).width = 15;
    dr.getColumn(TC).width = 17;
    dr.getRow(2).eachCell((c, n) => { if (n > 1 && n < TC) c.numFmt = 'd-mmm-yyyy'; });
    headerRow(dr.getRow(1)); headerRow(dr.getRow(2), C.label);
    for (let r = 3; r <= dr.rowCount; r++) {
      const a = dr.getRow(r).getCell(1);
      if (a.value) { a.fill = fill(C.label); a.border = border; }
      dr.getRow(r).getCell(TC).font = { bold: true };
    }
    const band = (key, argb, bold = true) => { const row = dr.getRow(R[key]); for (let c = 1; c <= TC; c++) { const x = row.getCell(c); x.fill = fill(argb); if (bold) x.font = { bold: true }; } };
    band('totalSales', C.total); band('totalRevenue', C.revenue); band('netProfit', C.profit); band('profitMargin', C.profit);
    band('netGST', C.gstSub); band('netProfitAfterGST', C.profit);
    dr.getRow(R._gst).getCell(1).fill = fill(C.gstTitle); dr.getRow(R._gst).getCell(1).font = { bold: true };
    [R._out, R._in].forEach(r => { dr.getRow(r).getCell(1).fill = fill(C.gstSub); dr.getRow(r).getCell(1).font = { italic: true }; });
    [R._sum1, R._sum2].forEach(r => { const row = dr.getRow(r); row.getCell(1).fill = fill(C.head); row.getCell(1).font = { bold: true, size: 13 }; row.getCell(2).font = { bold: true, size: 13 }; });
    dr.getColumn(TC).eachCell(c => { if (!c.fill || c.fill.fgColor?.argb === undefined) c.fill = fill('FFF3F6FA'); });
    dr.views = [{ state: 'frozen', xSplit: 1, ySplit: 2 }];

    // ---------------- Countrywise Sales ----------------
    const cw = ws.CW;
    const cs = opts.countrySummary;
    if (!has.revenue) notUploaded(cw, 'Net payments / Total sales');
    else {
    cw.getCell('B1').value = periodText; cw.getCell('B1').font = { bold: true };
    cw.getRow(3).values = [, 'Country', 'Revenue (ex GST)', 'Percentage (ex GST)'];
    headerRow(cw.getRow(3));
    const cwFirst = 4, cwLast = 3 + cs.rows.length, cwTot = cwLast + 2;
    cs.rows.forEach((r, i) => {
      const n = cwFirst + i, col = brCol(r.column);
      cw.getCell(`B${n}`).value = `Product Revenue ${r.country}`;
      setF(cw.getCell(`C${n}`), `SUM(${q(S.BR)}!${col}2:${col}${N + 1})`, r.revenue, NUM);
      setF(cw.getCell(`D${n}`), `IFERROR(C${n}/$C$${cwTot}*100,0)`, r.percentage * 100, '0.00');
    });
    cw.getCell(`B${cwTot}`).value = 'Total';
    setF(cw.getCell(`C${cwTot}`), `SUM(C${cwFirst}:C${cwLast})`, cs.total, NUM);
    setF(cw.getCell(`D${cwTot}`), `SUM(D${cwFirst}:D${cwLast})`, cs.total ? 100 : 0, '0.00');
    ['B', 'C', 'D'].forEach(x => { cw.getCell(`${x}${cwTot}`).font = { bold: true }; cw.getCell(`${x}${cwTot}`).fill = fill(C.total); });
    cw.getCell(`B${cwTot + 2}`).value = 'Check: difference to Daily Reporting product revenue (should be 0.00)';
    setF(cw.getCell(`C${cwTot + 2}`), `ROUND(C${cwTot}-${q(S.DR)}!${TL}${R.productRevenue},2)`, round2(cs.total - T.productRevenue), NUM);
    cw.getCell(`B${cwTot + 2}`).font = { italic: true, color: { argb: 'FF666666' } };
    cw.getColumn(1).width = 3; cw.getColumn(2).width = 40; cw.getColumn(3).width = 18; cw.getColumn(4).width = 20;
    }

    // ---------------- Productwise Sales ----------------
    const pw = ws.PW;
    const ps = opts.productSummary;
    if (!has.sales || !ps.rows.length) notUploaded(pw, 'The Total sales export');
    else {
    pw.getCell('B1').value = periodText; pw.getCell('B1').font = { bold: true };
    pw.getRow(2).values = ['', 'Product title', 'Gross sales (A$)', 'Discounts', 'Refunds', 'Net sales (A$)', 'Share of gross sales', 'Share of net sales'];
    headerRow(pw.getRow(2));
    const pwFirst = 3, pwLast = 2 + ps.rows.length, pwTot = pwLast + 1;
    ps.rows.forEach((r, i) => {
      const n = pwFirst + i;
      pw.getRow(n).values = ['', r.product, r.gross, r.discounts, r.refunds, r.net];
      setF(pw.getCell(`G${n}`), `IFERROR(C${n}/$C$${pwTot}*100,0)`, r.shareGross * 100, '0.00');
      setF(pw.getCell(`H${n}`), `IFERROR(F${n}/F$${pwTot}*100,0)`, r.shareNet * 100, '0.00');
    });
    pw.getCell(`B${pwTot}`).value = 'Total Sales';
    [['C', 'gross'], ['D', 'discounts'], ['E', 'refunds'], ['F', 'net']].forEach(([x, k]) => setF(pw.getCell(`${x}${pwTot}`), `SUM(${x}${pwFirst}:${x}${pwLast})`, ps.total[k], NUM));
    setF(pw.getCell(`G${pwTot}`), `SUM(G${pwFirst}:G${pwLast})`, ps.total.gross ? 100 : 0, '0.00');
    setF(pw.getCell(`H${pwTot}`), `SUM(H${pwFirst}:H${pwLast})`, ps.total.net ? 100 : 0, '0.00');
    for (let c = 2; c <= 8; c++) { const x = pw.getRow(pwTot).getCell(c); x.font = { bold: true }; x.fill = fill(C.total); }
    pw.getCell(`B${pwTot + 2}`).value = 'Check: difference to Daily Reporting gross sales (should be 0.00)';
    setF(pw.getCell(`C${pwTot + 2}`), `ROUND(C${pwTot}-${q(S.DR)}!${TL}${R.gross},2)`, round2(ps.total.gross - T.gross), NUM);
    pw.getCell(`B${pwTot + 2}`).font = { italic: true, color: { argb: 'FF666666' } };
    ['C', 'D', 'E', 'F'].forEach(x => pw.getColumn(x).numFmt = NUM);
    [3, 38, 16, 14, 14, 16, 18, 18].forEach((w, i) => pw.getColumn(i + 1).width = w);
    pw.views = [{ state: 'frozen', ySplit: 2 }];
    }

    // ---------------- Difference COGS ----------------
    const df = ws.DIFF;
    df.addRow(['COGS Difference', ...dates.map(utcDate), 'Period total']);
    const drCell = (i, row) => `${q(S.DR)}!${L(i)}${row}`;
    const dRows = [
      ['Supplier COGS (product + shipping)', i => `${drCell(i, R.cogsProduct)}+${drCell(i, R.cogsShipping)}`,
        x => x.cogsProduct == null ? null : round2(x.cogsProduct + x.cogsShipping), i => [drCell(i, R.cogsProduct), drCell(i, R.cogsShipping)]],
      ['COGS Shopify', i => drCell(i, R.cogsShopify), x => x.cogsShopify, i => [drCell(i, R.cogsShopify)]],
      ['Difference (Supplier − Shopify)', i => `${L(i)}2-${L(i)}3`,
        x => (x.cogsProduct == null || x.cogsShopify == null) ? null : round2(x.cogsProduct + x.cogsShipping - x.cogsShopify), i => [`${L(i)}2`, `${L(i)}3`]]];
    dRows.forEach(([label, f, v, inputs]) => {
      const r = df.addRow([label]);
      rows.forEach((x, i) => {
        const val = v(x);
        setF(r.getCell(i + 2), `IF(COUNT(${inputs(i).join(',')})=${inputs(i).length},${f(i)},"")`, val == null ? '' : val, NUM);
      });
      const span = `B${r.number}:${lastL}${r.number}`;
      const tot = rows.some(x => v(x) == null) ? '' : round2(rows.reduce((z, x) => z + v(x), 0));
      setF(r.getCell(TC), `IF(COUNT(${span})=${N},SUM(${span}),"")`, tot, NUM);
    });
    headerRow(df.getRow(1));
    df.getRow(1).eachCell((c, n) => { if (n > 1 && n < TC) c.numFmt = 'd-mmm-yyyy'; });
    df.getRow(4).font = { bold: true };
    df.getColumn(1).width = 34; for (let c = 2; c <= TC; c++) df.getColumn(c).width = 14;

    // ---------------- Journal Entries ----------------
    const je = ws.JE;
    if (!has.revenue && !has.merchant) notUploaded(je, 'Net payments / Total sales and the merchant workbook');
    else {
    je.addRow(['Date', 'Account code', 'Account', 'Debit (A$)', 'Credit (A$)', 'Description']);
    headerRow(je.getRow(1));
    const gw = result.salesResult.config.gateways;
    const clearing = { Shopify: gw['Shopify Payments'], Paypal: gw['PayPal'], Afterpay: gw['Afterpay'] };
    const jeTotals = [];
    dates.forEach(d => {
      const start = je.rowCount + 1;
      result.salesResult.journalLineRows.filter(x => x.Date === d).forEach(x =>
        je.addRow([utcDate(d), x['Account Code'], x.Description, x.Debit === '' ? null : x.Debit, x.Credit === '' ? null : x.Credit, 'Sales reconciliation']));
      Object.values(result.merchant.daily[d].platforms).forEach(p => {
        const cl = clearing[p.platform] || {};
        const name = p.platform === 'Paypal' ? 'PayPal' : p.platform;
        if (p.exGST) je.addRow([utcDate(d), '', `Merchant Fees – ${name}`, p.exGST, null, `${name} merchant fees net of GST`]);
        if (p.gst) je.addRow([utcDate(d), 820, 'GST Input Tax Credit', p.gst, null, `GST credit on ${name} fees`]);
        if (p.exGST || p.gst) je.addRow([utcDate(d), cl.accountCode || '', cl.journalDescription || `${name} Clearing`, null, round2(p.exGST + p.gst), `${name} fees deducted from payout`]);
      });
      const end = je.rowCount;
      const t = je.addRow(['', '', `Total ${fmtDMY(d)}`]);
      let deb = 0, cre = 0;
      for (let r = start; r <= end; r++) { deb += Number(je.getRow(r).getCell(4).value) || 0; cre += Number(je.getRow(r).getCell(5).value) || 0; }
      setF(t.getCell(4), `SUM(D${start}:D${end})`, round2(deb), NUM);
      setF(t.getCell(5), `SUM(E${start}:E${end})`, round2(cre), NUM);
      setF(t.getCell(6), `IF(ROUND(D${t.number}-E${t.number},2)=0,"Balanced","NOT BALANCED")`, Math.abs(round2(deb - cre)) < 0.005 ? 'Balanced' : 'NOT BALANCED');
      t.eachCell(c => { c.font = { bold: true }; c.fill = fill(C.total); });
      jeTotals.push(t.number);
      je.addRow([]);
    });
    je.getColumn(1).numFmt = 'd-mmm-yyyy';
    ['D', 'E'].forEach(x => je.getColumn(x).numFmt = NUM);
    [13, 12, 34, 15, 15, 40].forEach((w, i) => je.getColumn(i + 1).width = w);
    je.views = [{ state: 'frozen', ySplit: 1 }];
    }

    // ---------------- Ratio Calculation Sheet ----------------
    const ra = ws.RATIO;
    const DRT = k => `${q(S.DR)}!${TL}${R[k]}`;
    const cogsT = T.cogsProduct == null ? null : round2(T.cogsProduct + T.cogsShipping);
    const n2 = (...xs) => xs.some(v => v == null) ? null : xs;
    const ratioRows = [
      ['h', '1. Gross Profit (GP) & Gross Profit Margin'],
      ['Revenue (ex GST)', DRT('totalRevenue'), [DRT('totalRevenue')], T.totalRevenue, NUM],
      ['Cost of Goods Sold (Supplier COGS)', `${DRT('cogsProduct')}+${DRT('cogsShipping')}`, [DRT('cogsProduct'), DRT('cogsShipping')], cogsT, NUM],
      ['Gross Profit = Revenue − COGS', 'B3-B4', ['B3', 'B4'], (T.totalRevenue == null || cogsT == null) ? null : round2(T.totalRevenue - cogsT), NUM],
      ['Gross Profit Margin = Gross Profit ÷ Revenue', 'B5/B3', ['B5', 'B3'], (T.totalRevenue && cogsT != null) ? (T.totalRevenue - cogsT) / T.totalRevenue : null, PCT],
      [],
      ['h', '2. Net Profit & Net Profit Margin'],
      ['Net Profit (as per Supplier COGS)', DRT('netProfit'), [DRT('netProfit')], T.netProfit, NUM],
      ['Net Profit Margin = Net Profit ÷ Revenue', 'B9/B3', ['B9', 'B3'], T.profitMargin, PCT],
      ['Net Profit After GST remittance (cash view)', DRT('netProfitAfterGST'), [DRT('netProfitAfterGST')], T.netProfitAfterGST, NUM],
      ['Cash-view margin after GST remittance', 'B11/B3', ['B11', 'B3'], (T.totalRevenue && T.netProfitAfterGST != null) ? T.netProfitAfterGST / T.totalRevenue : null, PCT],
      [],
      ['h', '3. Cost of Goods Sold (COGS) Ratio'],
      ['COGS Ratio = COGS ÷ Revenue', 'B4/B3', ['B4', 'B3'], (T.totalRevenue && cogsT != null) ? cogsT / T.totalRevenue : null, PCT],
      [],
      ['h', '4. Operating Expense Ratio'],
      ['Operating Expenses = Merchant fees + Advertising', `${DRT('merchantFees')}+${DRT('adSpend')}`, [DRT('merchantFees'), DRT('adSpend')],
        (T.merchantFees == null || T.adSpend == null) ? null : round2(T.merchantFees + T.adSpend), NUM],
      ['Operating Expense Ratio = Operating Expenses ÷ Revenue', 'B18/B3', ['B18', 'B3'],
        (T.totalRevenue && T.merchantFees != null && T.adSpend != null) ? (T.merchantFees + T.adSpend) / T.totalRevenue : null, PCT],
      [],
      ['h', '5. Advertising'],
      ['ROAS = Revenue ÷ Ad spend', DRT('roas'), [DRT('roas')], T.roas, '#,##0.00']
    ];
    ratioRows.forEach((r, i) => {
      const n = i + 2;
      if (!r.length) return;
      if (r[0] === 'h') { const c = ra.getCell(`A${n}`); c.value = r[1]; c.font = { bold: true }; c.fill = fill(C.head); ra.getCell(`B${n}`).fill = fill(C.head); return; }
      const [label, expr, deps, value, fmt] = r;
      ra.getCell(`A${n}`).value = label;
      setF(ra.getCell(`B${n}`), `IFERROR(IF(COUNT(${deps.join(',')})=${deps.length},${expr},""),"")`, value == null ? '' : value, fmt);
    });
    ra.getCell('A1').value = `Ratios for ${periodText}`; ra.getCell('A1').font = { bold: true, size: 13 };
    ra.getColumn(1).width = 54; ra.getColumn(2).width = 18;

    // ---------------- Checks & Notes ----------------
    const nt = ws.NOTES;
    nt.getColumn(1).width = 26; nt.getColumn(2).width = 22; nt.getColumn(3).width = 90;
    nt.addRow(['Saint Daily Reporting']).font = { bold: true, size: 14 };
    nt.addRow(['Period', periodText]);
    nt.addRow(['Generated', new Date().toLocaleString('en-AU')]);
    nt.addRow([]);
    nt.addRow(['Source files']).font = { bold: true };
    Object.entries(opts.fileNames || {}).forEach(([k, v]) => nt.addRow([k, v || '(not uploaded)']));
    nt.addRow([]);
    nt.addRow(['Checks at export']).font = { bold: true };
    const ih = nt.addRow(['Severity', 'Source / date', 'Message']); headerRow(ih);
    const issues = opts.issues || [];
    if (!issues.length) nt.addRow(['OK', '', 'No gaps or review items.']);
    issues.forEach(x => { const r = nt.addRow([x.severity, `${x.source}${x.date ? ' · ' + x.date : ''}`, x.message]); if (x.severity !== 'INFO') r.getCell(1).fill = fill(x.severity === 'ERROR' ? 'FFF4CCCC' : C.warn); });
    nt.addRow([]);
    nt.addRow(['How key rows are calculated']).font = { bold: true };
    [
      ['Product revenue (ex GST)', 'Sum of every "Sales Revenue (…)" column on the Breakdown sheet. Australia is ex GST (÷1.1); other countries exclude their Other Tax.'],
      ['Refund / Refund Ex GST', 'Refund Ex GST = all "Refunds (…)" columns. Refund = Refund Ex GST + GST on AU refunds.'],
      ['Total revenue (ex GST)', 'Product revenue + Shipping AU (ex GST) + Shipping others − Refund Ex GST.'],
      ['COGS (supplier)', 'COGS sheet: USD × FX rate from the Saint COGS file, rounded to cents. Lines priced #N/A are listed on that sheet.'],
      ['Merchant fees', 'SUMIFS of "Merchant Exclude GST" on Fees Calculation for the date (Afterpay + PayPal + Shopify). PayPal has no GST.'],
      ['Net profit (supplier COGS)', 'Gross profit − Merchant fees (ex GST) − Advertising.'],
      ['GST output', 'GST on product sales (before refunds) + GST on AU shipping — taken from the Breakdown sheet.'],
      ['GST input', 'GST on merchant fees + GST on AU refunds.'],
      ['Net GST', 'Output − Input.'],
      ['Net profit after GST remittance', 'Cash view only: Net profit (supplier COGS) − Net GST payable. Not an accounting profit measure.'],
      ['Accounting basis', 'Revenue, COGS and fees are all ex GST. GST collected is a liability, not income, and Net GST is paid out of it, so it is not an expense. The accounting profit is "Net profit as per Suppliers COGS". GST on fees uses the actual GST charged (GSTOnFee in the merchant file). COGS uses the USD→AUD rate in the processed COGS file; if a date has no rate, the rate from the most recent earlier date is used and noted on the COGS sheet. #N/A COGS lines stay unpriced unless a price is entered.']
    ].forEach(([a, b]) => { const r = nt.addRow([a, '', b]); r.getCell(3).alignment = { wrapText: true }; });

    return wb;
  }

  return { build };
});
