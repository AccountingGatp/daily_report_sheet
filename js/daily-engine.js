/*
 * Saint Daily Reporting — calculation engine (v3)
 * Works in the browser (window.SaintDailyEngine) and in Node (module.exports).
 * Needs SaintSalesEngine (sales-engine.js) and SheetJS (XLSX) for workbook inputs.
 *
 * Nothing in here is period-specific: every figure is derived from the uploaded files
 * and the manual advertising-spend inputs.
 */
(function (root, factory) {
  const api = factory(root.SaintSalesEngine, root.XLSX);
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.SaintDailyEngine = api;
})(typeof self !== 'undefined' ? self : this, function (SalesEngine, XLSXRef) {
  'use strict';

  const GST_RATE = 0.10;
  const X = () => XLSXRef || (typeof XLSX !== 'undefined' ? XLSX : null);

  // ---------- small helpers ----------
  const round2 = v => { v = Number(v) || 0; const s = v < 0 ? -1 : 1; return s * Math.round((Math.abs(v) + Number.EPSILON) * 100) / 100; };
  const isErr = v => typeof v === 'string' && /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NULL!|NUM!|ERR)/i.test(v.trim());
  function toNum(v) {
    if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
    let s = String(v ?? '').trim();
    if (!s || /^[-–—]$/.test(s) || s.startsWith('#')) return 0;
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    s = s.replace(/[$£€%\s,]|A\$|AUD|USD/gi, '');
    const n = Number(s);
    return Number.isFinite(n) ? (neg ? -n : n) : 0;
  }
  const parseCSV = text => (SalesEngine && SalesEngine.parseCSV ? SalesEngine.parseCSV(text) : []);
  const norm = h => String(h ?? '').trim().toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  const pad = n => String(n).padStart(2, '0');

  /** Returns YYYY-MM-DD, or "--MM-DD" when the source only has month/day (e.g. "09/04"). */
  function isoDate(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
      return d.toISOString().slice(0, 10);
    }
    const s = String(v).trim();
    let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (m) return `${m[1]}-${pad(+m[2])}-${pad(+m[3])}`;
    m = s.match(/^(\d{1,2})[\/-](\d{1,2})(?:[\/-](\d{2,4}))?$/);
    if (m) {
      let y = m[3] ? +m[3] : null;
      if (y != null && y < 100) y += 2000;
      return y == null ? `--${pad(+m[1])}-${pad(+m[2])}` : `${y}-${pad(+m[1])}-${pad(+m[2])}`;
    }
    const d = new Date(s);
    return isNaN(d) ? '' : `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }
  /** Resolves a possibly year-less date against the report period. */
  function inPeriod(v, dates, wanted) {
    let d = isoDate(v);
    if (d.startsWith('--')) d = dates.find(x => x.slice(4) === d.slice(1)) || '';
    return wanted.has(d) ? d : '';
  }
  /** Like isoDate, but a year-less "09/03" gets the report period's year (so earlier dates resolve too). */
  function resolveYear(v, dates) {
    const d = isoDate(v);
    return d.startsWith('--') ? `${dates[0].slice(0, 4)}${d.slice(1)}` : d;
  }
  function dateRange(from, to) {
    const a = new Date(from + 'T00:00:00Z'), b = new Date(to + 'T00:00:00Z');
    if (isNaN(a) || isNaN(b)) throw new Error('Choose a valid From and To date.');
    if (a > b) throw new Error('"From" date is after the "To" date.');
    const out = [];
    for (let d = new Date(a); d <= b; d.setUTCDate(d.getUTCDate() + 1)) out.push(d.toISOString().slice(0, 10));
    if (out.length > 62) throw new Error('Report period is longer than 62 days. Split it into months.');
    return out;
  }
  function findCol(headers, cands) {
    const idx = new Map(headers.map((h, i) => [norm(h), i]));
    for (const c of cands) { const i = idx.get(norm(c)); if (i != null) return i; }
    return -1;
  }
  const pick = (row, names) => { for (const n of names) if (row[n] !== undefined && row[n] !== '') return row[n]; const keys = Object.keys(row); for (const n of names) { const k = keys.find(k => norm(k) === norm(n)); if (k) return row[k]; } return undefined; };
  const orderNumber = s => { const m = String(s || '').trim().match(/^#?(\d+)$/); return m ? +m[1] : null; };

  /** Sheet → 2-D array. Excel error cells come back as their text ("#N/A") instead of blank. */
  function sheetMatrix(ws) {
    const XL = X();
    if (!ws || !ws['!ref']) return [];
    const r = XL.utils.decode_range(ws['!ref']);
    const out = [];
    for (let R = r.s.r; R <= r.e.r; R++) {
      const row = [];
      for (let C = r.s.c; C <= r.e.c; C++) {
        const c = ws[XL.utils.encode_cell({ r: R, c: C })];
        row.push(!c ? null : c.t === 'e' ? (c.w || '#ERR') : c.v);
      }
      out.push(row);
    }
    return out;
  }

  // ---------- Total sales (Shopify "Total sales by order") ----------
  function aggregateSales(salesRows, dates) {
    const set = new Set(dates), out = {};
    dates.forEach(d => out[d] = { rows: 0, gross: 0, discounts: 0, returns: 0, netSales: 0, shipping: 0, taxes: 0, totalSales: 0, minOrder: null, maxOrder: null });
    for (const r of salesRows) {
      const d = isoDate(pick(r, ['Day', 'Date']));
      if (!set.has(d)) continue;
      const x = out[d];
      const gross = toNum(pick(r, ['Gross sales']));
      x.rows++;
      x.gross += gross;
      x.discounts += toNum(pick(r, ['Discounts']));
      x.returns += toNum(pick(r, ['Sales reversals', 'Returns']));
      x.netSales += toNum(pick(r, ['Net sales']));
      x.shipping += toNum(pick(r, ['Shipping charges']));
      x.taxes += toNum(pick(r, ['Taxes', 'Tax']));
      x.totalSales += toNum(pick(r, ['Total sales']));
      if (gross > 0) {
        const n = orderNumber(pick(r, ['Order name', 'Order']));
        if (n != null) { x.minOrder = x.minOrder == null ? n : Math.min(x.minOrder, n); x.maxOrder = x.maxOrder == null ? n : Math.max(x.maxOrder, n); }
      }
    }
    Object.values(out).forEach(x => {
      ['gross', 'discounts', 'returns', 'netSales', 'shipping', 'taxes', 'totalSales'].forEach(k => x[k] = round2(x[k]));
      const has = x.minOrder != null;
      x.orderRange = has ? `#${x.minOrder}-#${x.maxOrder}` : '';
      x.orderCount = has ? x.maxOrder - x.minOrder + 1 : 0;
    });
    return out;
  }

  // ---------- COGS Shopify ----------
  function aggregateShopifyCogs(rows, dates) {
    const set = new Set(dates), out = Object.fromEntries(dates.map(d => [d, 0])), seen = Object.fromEntries(dates.map(d => [d, 0]));
    for (const r of rows) {
      const d = isoDate(pick(r, ['Day', 'Date']));
      if (!set.has(d)) continue;
      out[d] += toNum(pick(r, ['Cost of goods sold', 'COGS']));
      seen[d]++;
    }
    dates.forEach(d => out[d] = round2(out[d]));
    return { byDate: out, rowsByDate: seen };
  }

  // ---------- Merchant fees (Saint Merchant output) ----------
  function parseMerchantWorkbook(wb, dates) {
    const wanted = new Set(dates);
    const daily = Object.fromEntries(dates.map(d => [d, { exGST: 0, gst: 0, inclGST: 0, platforms: {} }]));
    const rows = [];
    for (const s of wb.SheetNames) {
      const m = sheetMatrix(wb.Sheets[s]);
      for (let i = 0; i < m.length; i++) {
        const hdr = m[i].map(v => String(v ?? ''));
        const dateIdx = findCol(hdr, ['Row Labels', 'Date']);
        if (dateIdx < 0) continue;
        const grossIdx = findCol(hdr, ['Sum of GrossAmountAUD', 'GrossAmountAUD', 'Gross sales (A$)']);
        const countIdx = findCol(hdr, ['Count of OrderID', 'Number of orders']);
        const cols = {
          Paypal: { ex: findCol(hdr, ['Sum of FeeAmountAUD', 'FeeAmountAUD']), gst: -1, incl: -1 },
          Afterpay: { ex: findCol(hdr, ['Sum of MerchantFeeExclGST', 'MerchantFeeExclGST']), gst: findCol(hdr, ['Sum of MerchantFeeGST', 'MerchantFeeGST']), incl: findCol(hdr, ['Sum of MerchantFeeInclGST', 'MerchantFeeInclGST']) },
          Shopify: { ex: findCol(hdr, ['Sum of FeeExGST', 'FeeExGST']), gst: findCol(hdr, ['Sum of GSTOnFee', 'GSTOnFee']), incl: findCol(hdr, ['Sum of FeeInclGST', 'FeeInclGST']) }
        };
        const platform = Object.keys(cols).find(p => cols[p].ex >= 0 || cols[p].incl >= 0);
        if (!platform) continue;
        const c = cols[platform];
        for (let j = i + 1; j < m.length; j++) {
          const row = m[j];
          if (!row || row.every(v => v == null || v === '')) break;
          if (/grand total/i.test(String(row[dateIdx] ?? ''))) break;
          const d = inPeriod(row[dateIdx], dates, wanted);
          if (!d) continue;
          let ex = c.ex >= 0 ? toNum(row[c.ex]) : 0, gst = c.gst >= 0 ? toNum(row[c.gst]) : 0, incl = c.incl >= 0 ? toNum(row[c.incl]) : ex + gst;
          if (!ex && incl) ex = incl - gst;
          if (!incl) incl = ex + gst;
          const rec = { date: d, platform, orders: countIdx >= 0 ? toNum(row[countIdx]) : 0, gross: grossIdx >= 0 ? round2(toNum(row[grossIdx])) : 0, exGST: round2(ex), gst: round2(gst), inclGST: round2(incl) };
          const prev = daily[d].platforms[platform];
          if (prev) { // same platform/date twice → add, don't overwrite
            ['orders', 'gross', 'exGST', 'gst', 'inclGST'].forEach(k => prev[k] = round2(prev[k] + rec[k]));
          } else { daily[d].platforms[platform] = rec; rows.push(rec); }
        }
      }
    }
    dates.forEach(d => {
      const x = daily[d], p = Object.values(x.platforms);
      x.exGST = round2(p.reduce((s, r) => s + r.exGST, 0));
      x.gst = round2(p.reduce((s, r) => s + r.gst, 0));
      x.inclGST = round2(p.reduce((s, r) => s + r.inclGST, 0));
    });
    const order = { Afterpay: 1, Paypal: 2, Shopify: 3 };
    rows.sort((a, b) => order[a.platform] - order[b.platform] || a.date.localeCompare(b.date));
    return { daily, rows };
  }

  // ---------- Supplier COGS (Saint COGS output) ----------
  /**
   * overrides: { [lineKey]: usdPrice } for detail lines whose Price is #N/A (SKU not priced in Saint COGS).
   * Returns { byDate, unpriced }.
   */
  function parseCogsWorkbook(wb, dates, overrides = {}) {
    const wanted = new Set(dates);
    const byDate = Object.fromEntries(dates.map(d => [d, { productUSD: 0, shippingUSD: 0, rate: 0, productAUD: 0, shippingAUD: 0, source: '', status: 'Missing', note: '' }]));
    const matrices = wb.SheetNames.map(s => ({ name: s, m: sheetMatrix(wb.Sheets[s]) }));

    // 1) Detail rows: Date / Price / Total cost (one detail sheet only, prefer the one with SKUs).
    const detail = Object.fromEntries(dates.map(d => [d, { productUSD: 0, shippingUSD: 0, lines: 0 }]));
    const unpriced = [];
    let best = null;
    for (const { name, m } of matrices) {
      for (let i = 0; i < Math.min(m.length, 30); i++) {
        const h = m[i].map(v => String(v ?? ''));
        const di = findCol(h, ['Date']), pi = findCol(h, ['Price']), si = findCol(h, ['Total cost']);
        if (di < 0 || pi < 0 || si < 0) continue;
        const cand = { name, m, hr: i, di, pi, si, sku: findCol(h, ['SKU']), prod: findCol(h, ['Product Name']), ord: findCol(h, ['Clients Order Number', 'Order']), qty: findCol(h, ['Quantity']) };
        if (!best || (best.sku < 0 && cand.sku >= 0)) best = cand;
        break;
      }
    }
    if (best) {
      let lastOrder = '';
      for (let j = best.hr + 1; j < best.m.length; j++) {
        const row = best.m[j];
        if (!row) continue;
        if (best.ord >= 0 && row[best.ord]) lastOrder = String(row[best.ord]);
        const d = inPeriod(row[best.di], dates, wanted);
        if (!d) continue;
        const price = row[best.pi], ship = row[best.si];
        detail[d].lines++;
        detail[d].shippingUSD += toNum(ship);
        if (isErr(price)) {
          const key = `${d}|${j + 1}`;
          const line = { key, date: d, row: j + 1, sheet: best.name, order: lastOrder, sku: best.sku >= 0 ? String(row[best.sku] ?? '') : '', product: best.prod >= 0 ? String(row[best.prod] ?? '') : '', quantity: best.qty >= 0 ? toNum(row[best.qty]) : '', override: overrides[key] };
          unpriced.push(line);
          if (overrides[key] != null && overrides[key] !== '' && overrides[key] !== 'SKIP') detail[d].productUSD += toNum(overrides[key]);
        } else detail[d].productUSD += toNum(price);
      }
    }
    Object.values(detail).forEach(x => { x.productUSD = round2(x.productUSD); x.shippingUSD = round2(x.shippingUSD); });

    // 2) Summary tables (pivot "Row Labels / Sum of Price / Sum of Total cost" + FX rate, or labelled AUD tables).
    const summary = {}, allRates = {};
    for (const { m } of matrices) {
      for (let i = 0; i < m.length; i++) {
        const h = m[i].map(v => String(v ?? ''));
        const di = findCol(h, ['Row Labels', 'Date']);
        const pu = findCol(h, ['Sum of Price', 'Sum of Product Price USD', 'Product Price($)', 'Product Cost USD']);
        const su = findCol(h, ['Sum of Total cost', 'Sum of Shipping Cost USD', 'Shipping Cost ($)', 'Shipping Cost USD']);
        const pa = findCol(h, ['Sum of Product Price AUD', 'Product Cost AUD', 'Product Cost (AUD)']);
        const sa = findCol(h, ['Sum of Shipping Cost AUD', 'Shipping Cost AUD', 'Shipping Cost (AUD)']);
        const ri = findCol(h, ['Rate(USD to AUD', 'Rate(USD to AUD)', 'Rate (USD to AUD)', 'FX Rate']);
        if (di < 0 || ((pu < 0 || su < 0) && (pa < 0 || sa < 0))) continue;
        for (let j = i + 1; j < m.length; j++) {
          const row = m[j];
          if (!row) continue;
          const label = String(row[di] ?? '');
          if (/grand total|\(blank\)/i.test(label)) break;
          if (!label) { if (j > i + 1) break; else continue; }
          // Keep every dated FX rate in the file (also dates before the period) for carry-forward.
          const any = resolveYear(row[di], dates);
          if (any && ri >= 0 && !isErr(row[ri]) && toNum(row[ri]) > 0 && allRates[any] == null) allRates[any] = toNum(row[ri]);
          const d = inPeriod(row[di], dates, wanted);
          if (!d) continue;
          const s = summary[d] || (summary[d] = {});
          const set = (k, v) => { if (v == null || isErr(v) || v === '') return; if (s[k] == null) s[k] = toNum(v); };
          if (pu >= 0) set('productUSD', row[pu]);
          if (su >= 0) set('shippingUSD', row[su]);
          if (pa >= 0 && typeof row[pa] === 'number') set('productAUD', row[pa]);
          if (sa >= 0 && typeof row[sa] === 'number') set('shippingAUD', row[sa]);
          if (ri >= 0) set('rate', row[ri]);
          if (pu >= 0 && isErr(row[pu])) s.productErr = true;
        }
      }
    }

    // 3) Combine. Detail rows win for USD amounts (they include priced overrides); summary gives the FX rate.
    for (const d of dates) {
      const s = summary[d] || {}, det = detail[d], out = byDate[d];
      const haveDetail = det.lines > 0;
      out.productUSD = haveDetail ? det.productUSD : round2(s.productUSD || 0);
      out.shippingUSD = haveDetail ? det.shippingUSD : round2(s.shippingUSD || 0);
      out.rate = s.rate || 0;
      out.rateFrom = '';
      if (!out.rate && s.productAUD && out.productUSD) out.rate = Math.round(s.productAUD / out.productUSD * 1e6) / 1e6;  // rate implied by the file's own AUD column
      if (!out.rate) {
        // No rate for this date → use the most recent earlier date that has one.
        const prev = Object.keys(allRates).filter(x => x < d).sort().pop();
        if (prev) { out.rate = allRates[prev]; out.rateFrom = prev; }
      }
      if (out.rate) {
        out.productAUD = round2(out.productUSD * out.rate);
        out.shippingAUD = round2(out.shippingUSD * out.rate);
        out.source = 'USD × FX rate';
      } else if (s.productAUD != null || s.shippingAUD != null) {
        out.productAUD = round2(s.productAUD || 0);
        out.shippingAUD = round2(s.shippingAUD || 0);
        out.source = 'AUD from file';
      }
      const missingLines = unpriced.filter(u => u.date === d && (u.override == null || u.override === ''));
      const skippedLines = unpriced.filter(u => u.date === d && u.override === 'SKIP');
      if (!haveDetail && !summary[d]) { out.status = 'Missing'; out.note = 'Date not found in COGS file'; }
      else if (!out.rate && !out.source) { out.status = 'Review'; out.note = 'No USD→AUD rate for this date or any earlier date'; }
      else if (missingLines.length) { out.status = 'Review'; out.note = `${missingLines.length} line(s) have no product price (#N/A)`; }
      else if (!(out.productAUD > 0) || !(out.shippingAUD > 0)) { out.status = 'Review'; out.note = 'Product or shipping cost is zero'; }
      else {
        out.status = 'Ready';
        const priced = unpriced.filter(u => u.date === d && u.override !== 'SKIP').length;
        out.note = [skippedLines.length ? `${skippedLines.length} unpriced line(s) skipped — counted as $0 (as in file)` : '', priced ? `${priced} line(s) priced manually` : ''].filter(Boolean).join('; ');
      }
      if (out.rateFrom) out.note = [`FX rate ${out.rate} carried from ${out.rateFrom} (no rate for this date)`, out.note].filter(Boolean).join('; ');
    }
    return { byDate, unpriced };
  }

  // ---------- Uploaded Sales reconciliation vs recomputed ----------
  function compareRecon(uploadedRows, salesResult, dates) {
    const out = { checked: 0, diffs: [], datesFound: 0 };
    if (!uploadedRows || !uploadedRows.length) return out;
    const wanted = new Set(dates);
    const eng = new Map(salesResult.reconciliationRows.map(r => [r.Date, r]));
    for (const u of uploadedRows) {
      const d = isoDate(pick(u, ['Date', 'Day']));
      if (!wanted.has(d)) continue;
      out.datesFound++;
      const e = eng.get(d) || {};
      for (const k of Object.keys(u)) {
        if (/^(date|day)$/i.test(k.trim()) || /^rounding$/i.test(k.trim())) continue;
        if (!(k in e)) continue;
        out.checked++;
        const a = round2(toNum(u[k])), b = round2(e[k]);
        if (Math.abs(a - b) > 0.011) out.diffs.push({ date: d, column: k, uploaded: a, calculated: b });
      }
    }
    return out;
  }

  // ---------- Build ----------
  function buildDaily({ from, to, netRows, salesRows, shopifyCogsRows, merchantWb, cogsWb, adSpend = {}, cogsOverrides = {}, reconRows = null }) {
    if (!SalesEngine) throw new Error('Sales engine did not load.');
    const dates = dateRange(from, to);
    const salesAgg = aggregateSales(salesRows, dates);
    const shop = aggregateShopifyCogs(shopifyCogsRows, dates);
    const merchant = parseMerchantWorkbook(merchantWb, dates);
    const cogsParsed = parseCogsWorkbook(cogsWb, dates, cogsOverrides);
    const salesResult = SalesEngine.process(netRows, salesRows, null);
    const recByDate = new Map(salesResult.reconciliationRows.map(r => [r.Date, r]));
    const countries = salesResult.countries || [];
    const netRowsByDate = Object.fromEntries(dates.map(d => [d, 0]));
    salesResult.netRows.forEach(r => { if (r.day in netRowsByDate) netRowsByDate[r.day]++; });

    const rows = dates.map(d => {
      const s = salesAgg[d], r = recByDate.get(d) || {}, c = cogsParsed.byDate[d], m = merchant.daily[d], a = adSpend[d] || {};
      const sumC = prefix => countries.reduce((z, k) => z + (r[`${prefix} (${k})`] || 0), 0);
      const productRevenue = round2((r['Sales Revenue (Australia)'] || 0) + sumC('Sales Revenue'));
      const shippingAU = round2(r['Shipping (Australia, excl. GST)'] || 0);
      const shippingOther = round2(sumC('Shipping'));
      const gstRefundAU = round2(-(r['GST on Refund'] || 0));             // positive
      const refundExGST = round2(-((r['Refunds (Australia, excl. GST)'] || 0) + sumC('Refunds')));
      const refund = round2(refundExGST + gstRefundAU);                    // incl. AU GST
      const totalRevenue = round2(productRevenue + shippingAU + shippingOther - refundExGST);
      const grossProfit = round2(totalRevenue - c.productAUD - c.shippingAUD);
      const ads = { meta: toNum(a.meta), google: toNum(a.google), pinterest: toNum(a.pinterest), snap: toNum(a.snap), tiktok: toNum(a.tiktok) };
      const adTotal = round2(ads.meta + ads.google + ads.pinterest + ads.snap + ads.tiktok);
      const netProfit = round2(grossProfit - m.exGST - adTotal);
      // GST — output: GST on product sales (before refunds) + GST on AU shipping; input: GST on fees + GST on AU refunds.
      const gstSales = round2(r['GST on Sales'] || 0);
      const gstShipping = round2(r['GST on Shipping'] || 0);
      const netGST = round2(gstSales + gstShipping - m.gst - gstRefundAU);
      return {
        date: d, orderRange: s.orderRange, numberOrders: s.orderCount,
        gross: s.gross, discounts: s.discounts, returns: s.returns, netSales: round2(s.gross + s.discounts + s.returns),
        shipping: s.shipping, roundingTax: s.taxes, totalSales: round2(s.gross + s.discounts + s.returns + s.shipping + s.taxes),
        productRevenue, shippingAU, shippingOther, refund, refundExGST, totalRevenue,
        cogsProduct: c.productAUD, cogsShipping: c.shippingAUD, cogsShopify: shop.byDate[d], grossProfit,
        merchantFees: m.exGST, ...ads, adSpend: adTotal, netProfit,
        profitMargin: totalRevenue ? netProfit / totalRevenue : 0, roas: adTotal ? totalRevenue / adTotal : 0,
        gstSales, gstShipping, gstFee: m.gst, gstRefundAU, netGST, netProfitAfterGST: round2(netProfit - netGST)
      };
    });
    const sum = k => round2(rows.reduce((z, r) => z + (Number(r[k]) || 0), 0));
    const totals = {};
    ['numberOrders', 'gross', 'discounts', 'returns', 'netSales', 'shipping', 'roundingTax', 'totalSales', 'productRevenue', 'shippingAU', 'shippingOther', 'refund', 'refundExGST', 'totalRevenue', 'cogsProduct', 'cogsShipping', 'cogsShopify', 'grossProfit', 'merchantFees', 'meta', 'google', 'pinterest', 'snap', 'tiktok', 'adSpend', 'netProfit', 'gstSales', 'gstShipping', 'gstFee', 'gstRefundAU', 'netGST', 'netProfitAfterGST'].forEach(k => totals[k] = sum(k));
    totals.profitMargin = totals.totalRevenue ? totals.netProfit / totals.totalRevenue : 0;
    totals.roas = totals.adSpend ? totals.totalRevenue / totals.adSpend : 0;
    const withRange = rows.filter(r => r.orderRange);
    totals.orderRange = withRange.length ? `${withRange[0].orderRange.split('-')[0]}-${withRange.at(-1).orderRange.split('-')[1]}` : '';

    const reconCheck = compareRecon(reconRows, salesResult, dates);
    return { dates, rows, totals, salesResult, salesAgg, merchant, cogs: cogsParsed.byDate, cogsUnpriced: cogsParsed.unpriced, shopCogs: shop.byDate, shopCogsRows: shop.rowsByDate, netRowsByDate, reconCheck, reconUploaded: !!(reconRows && reconRows.length), adSpendEntered: rows.some(r => r.adSpend > 0) };
  }

  // ---------- Summaries for the workbook ----------
  function countrySummary(result) {
    const dates = new Set(result.dates);
    const list = ['Australia', ...(result.salesResult.countries || [])];
    const rows = list.map(c => {
      const col = c === 'Australia' ? 'Sales Revenue (Australia)' : `Sales Revenue (${c})`;
      const v = result.salesResult.reconciliationRows.filter(r => dates.has(r.Date)).reduce((z, r) => z + (r[col] || 0), 0);
      return { country: c, column: col, revenue: round2(v) };
    });
    const total = round2(rows.reduce((z, r) => z + r.revenue, 0));
    rows.forEach(r => r.percentage = total ? r.revenue / total : 0);
    return { rows, total };
  }
  function productSummary(salesRows, dates) {
    const set = new Set(dates), m = new Map();
    for (const r of salesRows) {
      if (!set.has(isoDate(pick(r, ['Day', 'Date'])))) continue;
      const p = String(pick(r, ['Product title at time of sale', 'Product title', 'Product']) ?? '').trim();
      if (!p) continue;
      const x = m.get(p) || { product: p, gross: 0, discounts: 0, refunds: 0, net: 0 };
      x.gross += toNum(pick(r, ['Gross sales'])); x.discounts += toNum(pick(r, ['Discounts']));
      x.refunds += toNum(pick(r, ['Sales reversals', 'Returns'])); x.net += toNum(pick(r, ['Net sales']));
      m.set(p, x);
    }
    const rows = [...m.values()].sort((a, b) => a.product.localeCompare(b.product));
    rows.forEach(r => ['gross', 'discounts', 'refunds', 'net'].forEach(k => r[k] = round2(r[k])));
    const tot = k => round2(rows.reduce((z, r) => z + r[k], 0));
    const total = { gross: tot('gross'), discounts: tot('discounts'), refunds: tot('refunds'), net: tot('net') };
    rows.forEach(r => { r.shareGross = total.gross ? r.gross / total.gross : 0; r.shareNet = total.net ? r.net / total.net : 0; });
    return { rows, total };
  }

  // ---------- Checks ----------
  const SOURCES = ['Total sales', 'Net payments', 'Sales reconciliation', 'COGS', 'COGS Shopify', 'Merchant fees', 'Ad spend'];
  const GATEWAY_TO_PLATFORM = { 'Shopify Payments': 'Shopify', 'PayPal': 'Paypal', 'Afterpay': 'Afterpay' };

  /** Per-source, per-date status: 'ok' | 'review' | 'gap', plus a short reason. */
  function coverage(result) {
    const cov = {};
    SOURCES.forEach(s => cov[s] = {});
    const recBy = new Map(result.salesResult.reconciliationRows.map(r => [r.Date, r]));
    const reconDiffDates = new Set(result.reconCheck.diffs.map(x => x.date));
    for (const r of result.rows) {
      const d = r.date, rec = recBy.get(d) || {};
      cov['Total sales'][d] = result.salesAgg[d].rows ? { s: 'ok' } : { s: 'gap', why: 'No rows for this date' };
      cov['Net payments'][d] = result.netRowsByDate[d] ? { s: 'ok' } : { s: 'gap', why: 'No rows for this date' };
      cov['Sales reconciliation'][d] = !result.reconUploaded ? { s: 'review', why: 'Not uploaded (report still calculated from Shopify files)' } : reconDiffDates.has(d) ? { s: 'review', why: 'Differs from recalculation' } : { s: 'ok' };
      const c = result.cogs[d];
      cov['COGS'][d] = c.status === 'Ready' ? { s: 'ok', why: c.note } : { s: c.status === 'Missing' ? 'gap' : 'review', why: c.note };
      cov['COGS Shopify'][d] = result.shopCogsRows[d] ? { s: 'ok' } : { s: 'gap', why: 'No rows for this date' };
      const pl = result.merchant.daily[d].platforms;
      const missing = Object.entries(GATEWAY_TO_PLATFORM).filter(([g, p]) => Math.abs(rec[g] || 0) > 0.005 && !pl[p]).map(([, p]) => p);
      cov['Merchant fees'][d] = !Object.keys(pl).length ? { s: 'gap', why: 'No fee rows for this date' } : missing.length ? { s: 'review', why: `Missing ${missing.join(', ')} fees` } : { s: 'ok' };
      cov['Ad spend'][d] = r.adSpend > 0 ? { s: 'ok' } : { s: 'info', why: 'Left blank — fill in manually in the Excel file' };
    }
    return cov;
  }

  function validate(result) {
    const issues = [];
    const add = (severity, source, date, message) => issues.push({ severity, source, date: date || '', message });
    const cov = coverage(result);
    for (const src of SOURCES) {
      const bad = Object.entries(cov[src]).filter(([, v]) => v.s !== 'ok');
      if (!bad.length) continue;
      const severity = bad.every(([, v]) => v.s === 'info') ? 'INFO' : bad.some(([, v]) => v.s === 'gap') && ['Total sales', 'Net payments'].includes(src) ? 'ERROR' : 'REVIEW';
      const groups = new Map();
      bad.forEach(([d, v]) => { const k = v.why || 'Check'; if (src === 'COGS' && /no product price/.test(k)) return; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(d); });
      groups.forEach((ds, why) => add(severity, src, ds.length === result.dates.length ? 'All dates' : ds.join(', '), why));
    }
    result.dates.filter(d => result.cogs[d].rateFrom).forEach(d => add('INFO', 'COGS', d, `No FX rate for this date in the COGS file — used ${result.cogs[d].rate} from ${result.cogs[d].rateFrom}.`));
    result.cogsUnpriced.filter(u => u.override === 'SKIP').forEach(u => add('INFO', 'COGS', u.date, `Skipped: ${u.sku || 'SKU'}${u.order ? `, order ${u.order}` : ''} has no price in the COGS file and is counted as $0.`));
    result.cogsUnpriced.filter(u => u.override == null || u.override === '').forEach(u =>
      add('REVIEW', 'COGS', u.date, `No price for ${u.sku || 'a SKU'}${u.product ? ` (${u.product})` : ''}${u.order ? `, order ${u.order}` : ''}. Enter a USD price or choose Skip in the Overview.`));
    result.reconCheck.diffs.slice(0, 8).forEach(x => add('REVIEW', 'Sales reconciliation', x.date, `${x.column}: file ${x.uploaded.toFixed(2)} vs recalculated ${x.calculated.toFixed(2)}`));
    if (result.reconCheck.diffs.length > 8) add('REVIEW', 'Sales reconciliation', '', `${result.reconCheck.diffs.length - 8} more differences not shown.`);
    (result.salesResult.audits || []).filter(a => a.severity === 'ERROR').forEach(a => add('ERROR', 'Net payments', a.date, a.message));
    return issues;
  }

  return { GST_RATE, round2, toNum, parseCSV, isoDate, dateRange, sheetMatrix, aggregateSales, aggregateShopifyCogs, parseMerchantWorkbook, parseCogsWorkbook, compareRecon, buildDaily, countrySummary, productSummary, coverage, validate, SOURCES };
});
