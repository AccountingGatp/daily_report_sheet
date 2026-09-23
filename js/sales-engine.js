(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.SaintSalesEngine = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const DEFAULT_CONFIG = {
    settings: { maxRounding: 0.05, australiaGstRate: 0.10 },
    gateways: {
      'Shopify Payments': { accountCode: 251, journalDescription: 'Shopify Clearing Account', xeroDescription: 'Shopify Clearing', taxRate: 'BAS Excluded', reportOrder: 10, journalOrder: 10, aliases: ['shopify'] },
      'PayPal': { accountCode: 252, journalDescription: 'PayPal Clearing Account', xeroDescription: 'PayPal Clearing', taxRate: 'BAS Excluded', reportOrder: 20, journalOrder: 30, aliases: ['paypal'] },
      'Afterpay': { accountCode: 253, journalDescription: 'Afterpay Clearing Account', xeroDescription: 'Afterpay Clearing', taxRate: 'BAS Excluded', reportOrder: 30, journalOrder: 20, aliases: ['afterpay'] },
      'Shop Cash': { accountCode: 272, journalDescription: 'ShopCash Clearing Account', xeroDescription: 'Shop Cash Clearing', taxRate: 'BAS Excluded', reportOrder: 40, journalOrder: 40, aliases: ['shop cash','shop_cash'] }
    },
    australia: {
      refund: { accountCode: 204, description: 'Refunds - Australia', taxRate: 'GST on Income', order: 10 },
      shipping: { accountCode: 254, description: 'Shipping Income - AU', taxRate: 'GST on Income' },
      revenue: { accountCode: 257, description: 'Product Revenue - AU', taxRate: 'GST on Income' }
    },
    gst: {
      accountCode: 820,
      description: 'GST',
      refundDescription: 'GST on Refund',
      taxRate: 'BAS Excluded'
    },
    exportShipping: { accountCode: 250, description: 'Shipping Income - Other', taxRate: 'BAS Excluded' },
    rounding: { accountCode: 860, description: 'Rounding', taxRate: 'BAS Excluded' },
    fallbackRevenue: { accountCode: 269, descriptionPrefix: 'Product Revenue - Other', taxRate: 'BAS Excluded', order: 900 },
    fallbackRefund: { accountCode: '', descriptionPrefix: 'Refunds', taxRate: 'BAS Excluded', order: 900, needsReview: true },
    fallbackOtherTax: { accountCode: '', descriptionPrefix: 'Other Tax', taxRate: 'BAS Excluded', order: 900, needsReview: true },
    countries: {
      'United States': { revenueAccount: 258, revenueDescription: 'Product Revenue - US', revenueTaxRate: 'BAS Excluded', revenueOrder: 10, refundAccount: 208, refundDescription: 'Refunds - United States', refundTaxRate: 'BAS Excluded', refundOrder: 20, taxAccount: 221, taxDescription: 'Other Tax - US', taxRate: 'BAS Excluded', taxOrder: 10 },
      'New Zealand': { revenueAccount: 259, revenueDescription: 'Product Revenue -NZ', revenueTaxRate: 'BAS Excluded', revenueOrder: 20, refundAccount: 207, refundDescription: 'Refunds - New Zealand', refundTaxRate: 'BAS Excluded', refundOrder: 30 },
      'United Kingdom': { revenueAccount: 261, revenueDescription: 'Product Revenue - UK', revenueTaxRate: 'BAS Excluded', revenueOrder: 30, refundAccount: 211, refundDescription: 'Refunds - United Kingdom', refundTaxRate: 'BAS Excluded', refundOrder: 40 },
      'Canada': { revenueAccount: 262, revenueDescription: 'Product Revenue - CA', revenueTaxRate: 'BAS Excluded', revenueOrder: 40, refundAccount: 219, refundDescription: 'Refunds - Canada', refundTaxRate: 'BAS Excluded', refundOrder: 50 },
      'Ireland': { revenueAccount: 279, revenueDescription: 'Product Revenue – Ireland', revenueTaxRate: 'BAS Excluded', revenueOrder: 50 },
      'Germany': { revenueAccount: 280, revenueDescription: 'Product Revenue - Germany', revenueTaxRate: 'BAS Excluded', revenueOrder: 60, refundAccount: 223, refundDescription: 'Refunds - Germany', refundTaxRate: 'BAS Excluded', refundOrder: 60 },
      'Belgium': { revenueAccount: 283, revenueDescription: 'Product Revenuw -Belgium', revenueTaxRate: 'BAS Excluded', revenueOrder: 70 },
      'Italy': { revenueAccount: 284, revenueDescription: 'Product Revenue – Italy', revenueTaxRate: 'BAS Excluded', revenueOrder: 80, refundAccount: 296, refundDescription: 'Refunds - Italy', refundTaxRate: 'BAS Excluded', refundOrder: 70 },
      'Spain': { revenueAccount: 286, revenueDescription: 'Product Revenue – Spain', revenueTaxRate: 'BAS Excluded', revenueOrder: 90, refundAccount: 337, refundDescription: 'Refunds - Spain', refundTaxRate: 'BAS Excluded', refundOrder: 80 },
      'France': { revenueAccount: 303, revenueDescription: 'Product Revenue - France', revenueTaxRate: 'BAS Excluded', revenueOrder: 100, refundAccount: 340, refundDescription: 'Refunds - France', refundTaxRate: 'BAS Excluded', refundOrder: 100 },
      'Netherlands': { revenueAccount: 306, revenueDescription: 'Product Revenue – Netherlands', revenueTaxRate: 'BAS Excluded', revenueOrder: 110, refundAccount: 298, refundDescription: 'Refunds - Netherlands', refundTaxRate: 'BAS Excluded', refundOrder: 90 },
      'Portugal': { revenueAccount: 320, revenueDescription: 'Product Revenue - Portugal', revenueTaxRate: 'BAS Excluded', revenueOrder: 120 },
      'Poland': { revenueAccount: 326, revenueDescription: 'Product Revenue - Poland', revenueTaxRate: 'BAS Excluded', revenueOrder: 130 },
      'Isle Of Man': { revenueAccount: 327, revenueDescription: 'Product Revenue - Isle Of Man', revenueTaxRate: 'BAS Excluded', revenueOrder: 140 },
      'Switzerland': { revenueAccount: 336, revenueDescription: 'Product Revenue - Switzerland', revenueTaxRate: 'BAS Excluded', revenueOrder: 150 },
      'Denmark': { revenueAccount: 338, revenueDescription: 'Product Revenue - Denmark', revenueTaxRate: 'BAS Excluded', revenueOrder: 160 },
      'Vietnam': { revenueAccount: 277, revenueDescription: 'Product Revenue - Other (Vietnam)', revenueTaxRate: 'BAS Excluded', revenueOrder: 990 }
    }
  };

  const REQUIRED_NET = ['Day', 'Order name', 'Payment gateway', 'Billing country', 'Gross payments', 'Refunded payments', 'Net payments'];
  const REQUIRED_SALES = ['Day', 'Order name', 'Net sales', 'Shipping charges', 'Taxes'];
  const REQUIRED_GATEWAY_SUMMARY = ['Payment gateway', 'Transactions', 'Gross payments', 'Refunded payments', 'Net payments'];
  const FIELD_ALIASES = {
    'Day': ['Date'],
    'Order name': ['Order Name', 'Order number', 'Order Number'],
    'Payment gateway': ['Payment Gateway', 'Gateway'],
    'Billing country': ['Billing Country', 'Billing country/region'],
    'Gross payments': ['Gross payment', 'Gross Payments'],
    'Refunded payments': ['Refunded payment', 'Refund payments', 'Refunded Payments'],
    'Net payments': ['Net payment', 'Net Payments'],
    'Net sales': ['Net Sales'],
    'Shipping charges': ['Shipping charge', 'Shipping Charges'],
    'Taxes': ['Tax', 'Total taxes', 'Total Taxes'],
    'Transactions': ['Transaction count', 'Transactions count']
  };

  function clone(obj) { return JSON.parse(JSON.stringify(obj)); }

  function parseCSV(text) {
    text = String(text == null ? '' : text).replace(/^\uFEFF/, '');
    const rows = [];
    let row = [], field = '', inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else inQuotes = false;
        } else field += ch;
      } else {
        if (ch === '"') inQuotes = true;
        else if (ch === ',') { row.push(field); field = ''; }
        else if (ch === '\n') { row.push(field.replace(/\r$/, '')); rows.push(row); row = []; field = ''; }
        else field += ch;
      }
    }
    if (field.length || row.length) { row.push(field.replace(/\r$/, '')); rows.push(row); }
    while (rows.length && rows[rows.length - 1].every(v => String(v).trim() === '')) rows.pop();
    if (!rows.length) return [];
    const headers = rows[0].map(h => String(h).trim());
    return rows.slice(1).filter(r => r.some(v => String(v).trim() !== '')).map(r => {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = idx < r.length ? r[idx] : '');
      return obj;
    });
  }

  function toNumber(v) {
    if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
    let s = String(v == null ? '' : v).trim();
    if (!s || /^[-–—]$/.test(s)) return 0;
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    s = s.replace(/[$£€,%\s]/g, '').replace(/,/g, '');
    const n = Number(s);
    if (!Number.isFinite(n)) return 0;
    return neg ? -n : n;
  }

  function round2(v) {
    v = Number(v) || 0;
    const sign = v < 0 ? -1 : 1;
    return sign * Math.round((Math.abs(v) + Number.EPSILON) * 100) / 100;
  }

  function normHeader(h) { return String(h || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, ' '); }

  function resolveColumns(rows, required) {
    if (!rows.length) return { map: {}, missing: required.slice() };
    const actual = Object.keys(rows[0]);
    const idx = new Map(actual.map(h => [normHeader(h), h]));
    const map = {}, missing = [];
    required.forEach(req => {
      const candidates = [req, ...(FIELD_ALIASES[req] || [])];
      let found = '';
      for (const candidate of candidates) {
        found = idx.get(normHeader(candidate)) || '';
        if (found) break;
      }
      if (found) map[req] = found;
      else missing.push(req);
    });
    return { map, missing };
  }

  function normalizeDate(v) {
    const s = String(v == null ? '' : v).trim();
    if (!s) return '';
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
    if (m) return `${m[1]}-${String(+m[2]).padStart(2,'0')}-${String(+m[3]).padStart(2,'0')}`;
    m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
    if (m) {
      let y = +m[3]; if (y < 100) y += y < 70 ? 2000 : 1900;
      return `${y}-${String(+m[1]).padStart(2,'0')}-${String(+m[2]).padStart(2,'0')}`;
    }
    const d = new Date(s);
    if (!Number.isNaN(d.getTime())) return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    return s;
  }

  function normalizeGateway(v, cfg) {
    const s = String(v == null ? '' : v).trim();
    const l = s.toLowerCase();
    const gateways = (cfg && cfg.gateways) || DEFAULT_CONFIG.gateways;
    for (const [name, mapping] of Object.entries(gateways)) {
      if (name.toLowerCase() === l) return name;
      const aliases = Array.isArray(mapping && mapping.aliases) ? mapping.aliases : [];
      if (aliases.some(a => l.includes(String(a).toLowerCase()))) return name;
    }
    if (l.includes('afterpay')) return 'Afterpay';
    if (l.includes('paypal')) return 'PayPal';
    if (l.includes('shopify')) return 'Shopify Payments';
    if (l.includes('shop_cash') || l.includes('shop cash')) return 'Shop Cash';
    return s || 'Unknown';
  }

  function normalizeCountry(v, cfg) {
    const s = String(v == null ? '' : v).trim();
    if (!s) return '';
    const l = s.toLowerCase();
    if (l === 'australia') return 'Australia';
    const countries = (cfg && cfg.countries) || DEFAULT_CONFIG.countries;
    const exact = Object.keys(countries).find(c => c.toLowerCase() === l);
    return exact || s;
  }

  function isValidAccount(v) {
    if (v == null || String(v).trim() === '') return false;
    return Number.isFinite(Number(v));
  }

  function key2(a,b) { return `${a}\u241F${b}`; }
  function key3(a,b,c) { return `${a}\u241F${b}\u241F${c}`; }
  function addMap(map, key, value) { map.set(key, (map.get(key) || 0) + value); }

  function formatDateMDY(iso) {
    const [y,m,d] = String(iso).split('-').map(Number);
    return `${m}/${d}/${y}`;
  }
  function formatNarrationDate(iso) {
    const [y,m,d] = String(iso).split('-');
    return `${String(m).padStart(2,'0')}.${String(d).padStart(2,'0')}.${y}`;
  }

  function getCountryConfig(country, cfg, audits) {
    const c = cfg.countries[country] || {};
    const missingCountry = !cfg.countries[country];
    const out = {
      revenueAccount: c.revenueAccount || cfg.fallbackRevenue.accountCode,
      revenueDescription: c.revenueDescription || `${cfg.fallbackRevenue.descriptionPrefix} (${country})`,
      revenueTaxRate: c.revenueTaxRate || cfg.fallbackRevenue.taxRate,
      revenueOrder: c.revenueOrder ?? cfg.fallbackRevenue.order,
      refundAccount: c.refundAccount ?? cfg.fallbackRefund.accountCode,
      refundDescription: c.refundDescription || `${cfg.fallbackRefund.descriptionPrefix} - ${country}`,
      refundTaxRate: c.refundTaxRate || cfg.fallbackRefund.taxRate,
      refundOrder: c.refundOrder ?? cfg.fallbackRefund.order,
      taxAccount: c.taxAccount ?? cfg.fallbackOtherTax.accountCode,
      taxDescription: c.taxDescription || `${cfg.fallbackOtherTax.descriptionPrefix} - ${country}`,
      taxRate: c.taxRate || cfg.fallbackOtherTax.taxRate,
      taxOrder: c.taxOrder ?? cfg.fallbackOtherTax.order,
      fallback: missingCountry,
      revenueFallback: !c.revenueAccount,
      refundFallback: !isValidAccount(c.refundAccount),
      taxFallback: !isValidAccount(c.taxAccount)
    };
    if (audits && missingCountry && !audits.some(a => a.category === 'GL Mapping' && a.country === country && a.message.includes('No country-specific GL mapping'))) {
      audits.push({ severity: 'REVIEW', category: 'GL Mapping', date: '', order: '', country, message: `No country-specific GL mapping exists for ${country}. Revenue will use the configured Product Revenue - Other fallback; refund/other-tax fallbacks are marked for review.` });
    }
    return out;
  }

  function process(netRowsRaw, salesRowsRaw, gatewaySummaryRaw, userConfig) {
    const cfg = Object.assign(clone(DEFAULT_CONFIG), userConfig || {});
    if (userConfig && userConfig.settings) cfg.settings = Object.assign(clone(DEFAULT_CONFIG.settings), userConfig.settings);
    if (userConfig && userConfig.countries) cfg.countries = Object.assign(clone(DEFAULT_CONFIG.countries), userConfig.countries);
    if (userConfig && userConfig.gateways) cfg.gateways = Object.assign(clone(DEFAULT_CONFIG.gateways), userConfig.gateways);
    const audits = [];

    const netCols = resolveColumns(netRowsRaw, REQUIRED_NET);
    const salesCols = resolveColumns(salesRowsRaw, REQUIRED_SALES);
    if (netCols.missing.length) throw new Error(`Net payments file is missing required column(s): ${netCols.missing.join(', ')}`);
    if (salesCols.missing.length) throw new Error(`Total sales file is missing required column(s): ${salesCols.missing.join(', ')}`);

    const net = netRowsRaw.map((r, idx) => ({
      _row: idx + 2,
      day: normalizeDate(r[netCols.map['Day']]),
      order: String(r[netCols.map['Order name']] || '').trim(),
      gatewayRaw: String(r[netCols.map['Payment gateway']] || '').trim(),
      gateway: normalizeGateway(r[netCols.map['Payment gateway']], cfg),
      country: normalizeCountry(r[netCols.map['Billing country']], cfg),
      gross: toNumber(r[netCols.map['Gross payments']]),
      refund: toNumber(r[netCols.map['Refunded payments']]),
      net: toNumber(r[netCols.map['Net payments']])
    })).filter(r => r.day || r.order || r.gatewayRaw || r.country || r.gross || r.refund || r.net);

    const sales = salesRowsRaw.map((r, idx) => ({
      _row: idx + 2,
      day: normalizeDate(r[salesCols.map['Day']]),
      order: String(r[salesCols.map['Order name']] || '').trim(),
      netSales: toNumber(r[salesCols.map['Net sales']]),
      shipping: toNumber(r[salesCols.map['Shipping charges']]),
      taxes: toNumber(r[salesCols.map['Taxes']])
    })).filter(r => r.day || r.order || r.netSales || r.shipping || r.taxes);

    // Aggregate Total Sales by Order name + Day exactly as the SOP requires.
    const salesAgg = new Map();
    sales.forEach(r => {
      const k = key2(r.day, r.order);
      const x = salesAgg.get(k) || { day: r.day, order: r.order, netSales: 0, shipping: 0, taxes: 0 };
      x.netSales += r.netSales; x.shipping += r.shipping; x.taxes += r.taxes;
      salesAgg.set(k, x);
    });

    // Net Payments is the source of truth for country. Track any order/day that crosses countries.
    const orderCountrySets = new Map();
    net.forEach(r => {
      const k = key2(r.day, r.order);
      if (!orderCountrySets.has(k)) orderCountrySets.set(k, []);
      if (r.country && !orderCountrySets.get(k).includes(r.country)) orderCountrySets.get(k).push(r.country);
    });
    const orderCountry = new Map();
    orderCountrySets.forEach((arr,k) => {
      if (arr.length > 1) audits.push({ severity: 'ERROR', category: 'Join', date: k.split('\u241F')[0], order: k.split('\u241F')[1], country: arr.join(' / '), message: 'The same Order name + Day has more than one Billing country in Net payments. Automatic country allocation is unsafe.' });
      orderCountry.set(k, arr[0] || '');
    });

    let missingSalesJoin = 0;
    net.forEach(r => { if (!salesAgg.has(key2(r.day,r.order))) { missingSalesJoin++; audits.push({ severity: 'REVIEW', category: 'Join', date: r.day, order: r.order, country: r.country, message: 'Order exists in Net payments but not in Total sales. SOP treatment for missing sales-side values is zero; final gross/refund reconciliation is still retained.' }); } });
    let salesWithoutNet = 0;
    salesAgg.forEach((x,k) => { if (!orderCountry.has(k)) salesWithoutNet++; });
    if (salesWithoutNet) audits.push({ severity: 'INFO', category: 'Base Dataset', date: '', order: '', country: '', message: `${salesWithoutNet} aggregated Total sales Order+Day records do not exist in Net payments and are excluded because Net payments is the SOP base dataset.` });

    const netDates = [...new Set(net.map(r => r.day).filter(Boolean))].sort();
    const salesDates = [...new Set(sales.map(r => r.day).filter(Boolean))].sort();
    if (netDates.join('|') !== salesDates.join('|')) audits.push({ severity: 'REVIEW', category: 'Date Range', date: '', order: '', country: '', message: `Input date coverage differs. Net payments: ${netDates[0] || 'n/a'} to ${netDates[netDates.length-1] || 'n/a'}; Total sales: ${salesDates[0] || 'n/a'} to ${salesDates[salesDates.length-1] || 'n/a'}.` });

    const countries = [...new Set(net.map(r => r.country).filter(c => c && c !== 'Australia'))].sort((a,b) => a.localeCompare(b));
    const dates = [...new Set(netDates.concat(salesDates))].sort();

    const gatewayDaily = new Map(), countryGross = new Map(), countryRefund = new Map(), countryNet = new Map();
    const unknownGateways = new Set();
    net.forEach(r => {
      addMap(gatewayDaily, key2(r.day,r.gateway), r.net);
      addMap(countryGross, key3(r.day,r.country,'gross'), r.gross);
      addMap(countryRefund, key3(r.day,r.country,'refund'), r.refund);
      addMap(countryNet, key3(r.day,r.country,'net'), r.net);
      if (!cfg.gateways[r.gateway] || !isValidAccount(cfg.gateways[r.gateway].accountCode)) unknownGateways.add(r.gateway);
    });
    unknownGateways.forEach(g => audits.push({ severity: 'REVIEW', category: 'Gateway Mapping', date: '', order: '', country: '', message: `Unmapped payment gateway '${g}' detected. It is included in reconciliation but needs a clearing-account mapping before Xero posting.` }));

    // Shipping is kept from the aggregated Total Sales data and mapped to the authoritative Net Payments country.
    const shipDailyCountry = new Map();
    const prelimNetSalesDailyCountry = new Map();
    salesAgg.forEach((x,k) => {
      const c = orderCountry.get(k) || '';
      if (!c) return;
      addMap(shipDailyCountry, key2(x.day,c), x.shipping);
      addMap(prelimNetSalesDailyCountry, key2(x.day,c), x.netSales);
    });

    // Reference-output compatibility: Taxes are evaluated on the Net Payments base after the Order+Day join.
    // This reproduces the supplied reconciliation_output exactly, including split-payment orders.
    const taxDailyCountry = new Map();
    net.forEach(r => {
      const x = salesAgg.get(key2(r.day,r.order));
      addMap(taxDailyCountry, key2(r.day,r.country), x ? x.taxes : 0);
    });

    const gatewayColumns = Object.keys(cfg.gateways).sort((a,b) => {
      const ma = cfg.gateways[a] || {}, mb = cfg.gateways[b] || {};
      return (ma.reportOrder ?? ma.order ?? 900) - (mb.reportOrder ?? mb.order ?? 900) || a.localeCompare(b);
    });
    const unknownGatewayColumns = [...unknownGateways].filter(g => !gatewayColumns.includes(g)).sort();
    const detectedGateways = [...new Set(net.map(r => r.gateway).filter(Boolean))].sort();
    const recRows = [];
    dates.forEach(day => {
      const row = { Date: day };
      gatewayColumns.forEach(g => row[g] = round2(gatewayDaily.get(key2(day,g)) || 0));
      unknownGatewayColumns.forEach(g => row[`Gateway (${g})`] = round2(gatewayDaily.get(key2(day,g)) || 0));

      const auGross = round2(countryGross.get(key3(day,'Australia','gross')) || 0);
      const auRefund = round2(countryRefund.get(key3(day,'Australia','refund')) || 0);
      const auShippingInclusive = round2(shipDailyCountry.get(key2(day,'Australia')) || 0);
      const auSalesInclusive = round2(auGross - auShippingInclusive);
      const gstRate = Number(cfg.settings && cfg.settings.australiaGstRate);
      const safeGstRate = Number.isFinite(gstRate) && gstRate >= 0 ? gstRate : 0.10;
      const gstDivisor = 1 + safeGstRate;
      const gstPart = amount => amount * safeGstRate / gstDivisor;
      row['Sales Revenue (Australia)'] = round2(auSalesInclusive / gstDivisor);
      row['Shipping (Australia, excl. GST)'] = round2(auShippingInclusive / gstDivisor);
      row['GST on Shipping'] = round2(gstPart(auShippingInclusive));
      row['GST on Sales'] = round2(gstPart(auSalesInclusive));
      row['Refunds (Australia, excl. GST)'] = round2(auRefund / gstDivisor);
      row['GST on Refund'] = round2(gstPart(auRefund));
      // Combined pre-refund GST is rounded once, matching the supplied Saint reference output.
      row['Total GST Payable'] = round2(gstPart(auGross));

      countries.forEach(country => {
        const gross = round2(countryGross.get(key3(day,country,'gross')) || 0);
        const refund = round2(countryRefund.get(key3(day,country,'refund')) || 0);
        const shipping = round2(shipDailyCountry.get(key2(day,country)) || 0);
        const otherTax = round2(taxDailyCountry.get(key2(day,country)) || 0);
        row[`Sales Revenue (${country})`] = round2(gross - shipping - otherTax);
        row[`Shipping (${country})`] = shipping;
        row[`Other Tax (${country})`] = otherTax;
        row[`Refunds (${country})`] = refund;
      });
      row.rounding = 0;
      recRows.push(row);
    });

    // Audit country-level source tie-outs and adjustment trail.
    recRows.forEach(row => {
      const day = row.Date;
      const auGrossCalc = round2(row['Sales Revenue (Australia)'] + row['Shipping (Australia, excl. GST)'] + row['GST on Shipping'] + row['GST on Sales']);
      const auGrossSrc = round2(countryGross.get(key3(day,'Australia','gross')) || 0);
      const auRefundCalc = round2(row['Refunds (Australia, excl. GST)'] + row['GST on Refund']);
      const auRefundSrc = round2(countryRefund.get(key3(day,'Australia','refund')) || 0);
      if (Math.abs(round2(auGrossSrc - auGrossCalc)) > 0.01 || Math.abs(round2(auRefundSrc - auRefundCalc)) > 0.01) audits.push({ severity:'ERROR', category:'Country Reconciliation', date:day, order:'', country:'Australia', message:`Australia does not tie to Net payments. Gross diff ${round2(auGrossSrc-auGrossCalc)}, refund diff ${round2(auRefundSrc-auRefundCalc)}.` });

      countries.forEach(country => {
        const grossCalc = round2(row[`Sales Revenue (${country})`] + row[`Shipping (${country})`] + row[`Other Tax (${country})`]);
        const grossSrc = round2(countryGross.get(key3(day,country,'gross')) || 0);
        const refundCalc = round2(row[`Refunds (${country})`]);
        const refundSrc = round2(countryRefund.get(key3(day,country,'refund')) || 0);
        if (Math.abs(round2(grossSrc-grossCalc)) > 0.01 || Math.abs(round2(refundSrc-refundCalc)) > 0.01) audits.push({ severity:'ERROR', category:'Country Reconciliation', date:day, order:'', country, message:`Country does not tie to Net payments. Gross diff ${round2(grossSrc-grossCalc)}, refund diff ${round2(refundSrc-refundCalc)}.` });

        const preliminary = round2(prelimNetSalesDailyCountry.get(key2(day,country)) || 0);
        const finalSalesNetOfRefund = round2(row[`Sales Revenue (${country})`] + row[`Refunds (${country})`]);
        const adjustment = round2(finalSalesNetOfRefund - preliminary);
        if (Math.abs(adjustment) >= 0.01) audits.push({ severity:'INFO', category:'Sales Adjustment', date:day, order:'', country, message:`Reference reconciliation adjusts Sales while retaining Shipping so Gross payments tie to Net payments. Net-sales comparison adjustment: ${adjustment.toFixed(2)}.` });
      });
    });

    // Optional Payments by gateway summary validation.
    const gatewayCheck = [];
    if (gatewaySummaryRaw && gatewaySummaryRaw.length) {
      const gcols = resolveColumns(gatewaySummaryRaw, REQUIRED_GATEWAY_SUMMARY);
      if (gcols.missing.length) audits.push({ severity:'REVIEW', category:'Gateway Summary', date:'', order:'', country:'', message:`Optional Payments by gateway summary is missing: ${gcols.missing.join(', ')}.` });
      else {
        const source = new Map();
        gatewaySummaryRaw.forEach(r => {
          const g = normalizeGateway(r[gcols.map['Payment gateway']], cfg);
          source.set(g, {
            transactions: toNumber(r[gcols.map['Transactions']]),
            gross: toNumber(r[gcols.map['Gross payments']]),
            refund: toNumber(r[gcols.map['Refunded payments']]),
            net: toNumber(r[gcols.map['Net payments']])
          });
        });
        const netGrouped = new Map();
        net.forEach(r => {
          const x = netGrouped.get(r.gateway) || { transactions:0,gross:0,refund:0,net:0 };
          x.transactions++; x.gross += r.gross; x.refund += r.refund; x.net += r.net; netGrouped.set(r.gateway,x);
        });
        [...new Set([...source.keys(), ...netGrouped.keys()])].sort().forEach(g => {
          const a = source.get(g) || {transactions:0,gross:0,refund:0,net:0};
          const b = netGrouped.get(g) || {transactions:0,gross:0,refund:0,net:0};
          const ok = a.transactions === b.transactions && Math.abs(round2(a.gross-b.gross))<0.01 && Math.abs(round2(a.refund-b.refund))<0.01 && Math.abs(round2(a.net-b.net))<0.01;
          gatewayCheck.push({ Gateway:g, Status:ok?'PASS':'REVIEW', 'Summary Transactions':a.transactions, 'Net File Transactions':b.transactions, 'Summary Gross':round2(a.gross), 'Net File Gross':round2(b.gross), 'Summary Refund':round2(a.refund), 'Net File Refund':round2(b.refund), 'Summary Net':round2(a.net), 'Net File Net':round2(b.net) });
          if (!ok) audits.push({ severity:'REVIEW', category:'Gateway Summary', date:'', order:'', country:'', message:`Gateway summary does not match Net payments for ${g}.` });
        });
      }
    }

    // Apply mappings and construct balanced journal + Xero import.
    const built = buildJournal(recRows, countries, gatewayColumns, unknownGatewayColumns, cfg, audits);

    const summaryRows = countries.map(country => ({
      Country: country,
      'Total Sales Revenue': round2(recRows.reduce((s,r)=>s + (r[`Sales Revenue (${country})`]||0),0)),
      'Total Shipping': round2(recRows.reduce((s,r)=>s + (r[`Shipping (${country})`]||0),0)),
      'Total Other Tax': round2(recRows.reduce((s,r)=>s + (r[`Other Tax (${country})`]||0),0)),
      'Total Refunds': round2(recRows.reduce((s,r)=>s + (r[`Refunds (${country})`]||0),0))
    }));

    const recColumns = ['Date', ...gatewayColumns, ...unknownGatewayColumns.map(g=>`Gateway (${g})`), 'Sales Revenue (Australia)', 'Shipping (Australia, excl. GST)', 'GST on Shipping', 'GST on Sales', 'Refunds (Australia, excl. GST)', 'GST on Refund', 'Total GST Payable'];
    countries.forEach(c => recColumns.push(`Sales Revenue (${c})`,`Shipping (${c})`,`Other Tax (${c})`,`Refunds (${c})`));
    recColumns.push('rounding');

    return {
      config: cfg,
      netRows: net,
      salesRows: sales,
      dates,
      countries,
      detectedGateways,
      unmappedGateways: [...unknownGateways].sort(),
      reconciliationRows: recRows,
      reconciliationColumns: recColumns,
      summaryRows,
      journalRows: built.journalRows,
      journalLineRows: built.journalLineRows,
      xeroRows: built.xeroRows,
      audits,
      gatewayCheck,
      metrics: {
        netRows: net.length,
        salesRows: sales.length,
        aggregatedSalesOrders: salesAgg.size,
        countries: countries.length + (net.some(r=>r.country==='Australia') ? 1 : 0),
        startDate: dates[0] || '',
        endDate: dates[dates.length-1] || '',
        missingSalesJoin,
        salesWithoutNet,
        grossPayments: round2(net.reduce((s,r)=>s+r.gross,0)),
        refundedPayments: round2(net.reduce((s,r)=>s+r.refund,0)),
        netPayments: round2(net.reduce((s,r)=>s+r.net,0)),
        journalBalanced: built.journalBalanced,
        xeroReady: built.xeroReady,
        unresolvedMappings: built.unresolvedCount,
        maxAbsRounding: Math.max(0,...recRows.map(r=>Math.abs(r.rounding||0)))
      }
    };
  }

  function buildJournal(recRows, countries, gatewayColumns, unknownGateways, cfg, audits) {
    const journalRows = [], journalLineRows = [], xeroRows = [];
    let grandDebit = 0, grandCredit = 0, allBalanced = true, unresolvedCount = 0, xeroReady = true;

    function actualLine(date, code, journalDescription, xeroDescription, taxRate, conceptualSide, amount, debitLines, creditLines, meta) {
      amount = round2(amount);
      if (Math.abs(amount) < 0.005) return;
      const isDebit = (conceptualSide === 'debit' && amount > 0) || (conceptualSide === 'credit' && amount < 0);
      const v = Math.abs(amount);
      const line = { Date: date, 'Account Code': code, Description: journalDescription, Debit: isDebit ? v : '', Credit: isDebit ? '' : v, _xeroDescription: xeroDescription || journalDescription, _taxRate: taxRate || 'BAS Excluded', _meta: meta || '', _unresolved: false };
      (isDebit ? debitLines : creditLines).push(line);
    }

    function unresolvedLine(date, description, conceptualSide, amount, debitLines, creditLines, meta) {
      amount = round2(amount);
      if (Math.abs(amount) < 0.005) return;
      const isDebit = (conceptualSide === 'debit' && amount > 0) || (conceptualSide === 'credit' && amount < 0);
      const v = Math.abs(amount);
      const line = { Date: date, 'Account Code': '', Description: `UNMAPPED — ${description}`, Debit: isDebit ? v : '', Credit: isDebit ? '' : v, _xeroDescription: description, _taxRate: '', _meta: meta || '', _unresolved: true };
      (isDebit ? debitLines : creditLines).push(line);
      unresolvedCount++; xeroReady = false;
    }

    recRows.forEach(row => {
      const day = row.Date;
      const debits = [], credits = [];

      // Gateway clearing debits. Mapping order comes from configuration, not the source period.
      const journalGateways = gatewayColumns.slice().sort((a,b) => {
        const ma = cfg.gateways[a] || {}, mb = cfg.gateways[b] || {};
        return (ma.journalOrder ?? ma.order ?? 900) - (mb.journalOrder ?? mb.order ?? 900) || a.localeCompare(b);
      });
      journalGateways.forEach(g => {
        const amount = row[g] || 0;
        if (Math.abs(amount) < 0.005) return;
        const m = cfg.gateways[g] || {};
        if (isValidAccount(m.accountCode)) actualLine(day,m.accountCode,m.journalDescription || `${g} Clearing Account`,m.xeroDescription || `${g} Clearing`,m.taxRate || 'BAS Excluded','debit',amount,debits,credits,'gateway');
        else {
          unresolvedLine(day,`${g} Clearing`,'debit',amount,debits,credits,'gateway');
          audits.push({ severity:'REVIEW', category:'Gateway Mapping', date:day, order:'', country:'', message:`${g} has ${amount.toFixed(2)} net payments on ${day} but no clearing GL. Map this gateway before Xero export.` });
        }
      });
      unknownGateways.forEach(g => {
        const amount = row[`Gateway (${g})`] || 0;
        if (Math.abs(amount) < 0.005) return;
        unresolvedLine(day,`${g} Clearing`,'debit',amount,debits,credits,'gateway');
        audits.push({ severity:'REVIEW', category:'Gateway Mapping', date:day, order:'', country:'', message:`${g} has ${amount.toFixed(2)} net payments on ${day} but no clearing GL. Map this gateway before Xero export.` });
      });

      // Refund debits. Australia is ex GST and refund GST is posted separately.
      actualLine(day,cfg.australia.refund.accountCode,cfg.australia.refund.description,cfg.australia.refund.description,cfg.australia.refund.taxRate,'debit',-(row['Refunds (Australia, excl. GST)']||0),debits,credits,'refund');
      const refundCountries = countries.slice().sort((a,b) => {
        const ca = getCountryConfig(a,cfg), cb = getCountryConfig(b,cfg);
        return (ca.refundOrder ?? 900) - (cb.refundOrder ?? 900) || a.localeCompare(b);
      });
      refundCountries.forEach(country => {
        const amount = -(row[`Refunds (${country})`] || 0);
        if (Math.abs(amount) < 0.005) return;
        const c = getCountryConfig(country,cfg,audits);
        if (isValidAccount(c.refundAccount)) actualLine(day,c.refundAccount,c.refundDescription,c.refundDescription,c.refundTaxRate || 'BAS Excluded','debit',amount,debits,credits,'refund');
        else {
          unresolvedLine(day,c.refundDescription || `Refunds - ${country}`,'debit',amount,debits,credits,'refund');
          audits.push({ severity:'REVIEW', category:'GL Mapping', date:day, order:'', country, message:`Refund for ${country} has no configured GL account. Map it before Xero export.` });
        }
      });
      actualLine(day,cfg.gst.accountCode,cfg.gst.refundDescription,cfg.gst.refundDescription,cfg.gst.taxRate,'debit',-(row['GST on Refund']||0),debits,credits,'gst_refund');

      // Credits and income lines.
      actualLine(day,cfg.australia.shipping.accountCode,cfg.australia.shipping.description,cfg.australia.shipping.description,cfg.australia.shipping.taxRate,'credit',row['Shipping (Australia, excl. GST)']||0,debits,credits,'shipping_au');
      actualLine(day,cfg.australia.revenue.accountCode,cfg.australia.revenue.description,cfg.australia.revenue.description,cfg.australia.revenue.taxRate,'credit',row['Sales Revenue (Australia)']||0,debits,credits,'revenue_au');
      actualLine(day,cfg.gst.accountCode,cfg.gst.description,cfg.gst.description,cfg.gst.taxRate,'credit',row['Total GST Payable']||0,debits,credits,'gst');

      // Other tax is separate where a country-specific tax GL exists (US=221 in the supplied reference).
      const taxCountries = countries.slice().sort((a,b) => {
        const ca = getCountryConfig(a,cfg), cb = getCountryConfig(b,cfg);
        return (ca.taxOrder ?? 900) - (cb.taxOrder ?? 900) || a.localeCompare(b);
      });
      taxCountries.forEach(country => {
        const amount = row[`Other Tax (${country})`] || 0;
        if (Math.abs(amount) < 0.005) return;
        const c = getCountryConfig(country,cfg,audits);
        const desc = c.taxDescription || `${cfg.fallbackOtherTax.descriptionPrefix} - ${country}`;
        if (isValidAccount(c.taxAccount)) actualLine(day,c.taxAccount,desc,desc,c.taxRate || cfg.fallbackOtherTax.taxRate,'credit',amount,debits,credits,'other_tax');
        else {
          unresolvedLine(day,desc,'credit',amount,debits,credits,'other_tax');
          audits.push({ severity:'REVIEW', category:'GL Mapping', date:day, order:'', country, message:`Other Tax for ${country} has no configured GL account. Map it before Xero export.` });
        }
      });

      const revenueCountries = countries.slice().sort((a,b) => {
        const ca = getCountryConfig(a,cfg), cb = getCountryConfig(b,cfg);
        return (ca.revenueOrder ?? 900) - (cb.revenueOrder ?? 900) || a.localeCompare(b);
      });
      revenueCountries.forEach(country => {
        const amount = row[`Sales Revenue (${country})`] || 0;
        if (Math.abs(amount) < 0.005) return;
        const c = getCountryConfig(country,cfg,audits);
        if (isValidAccount(c.revenueAccount)) actualLine(day,c.revenueAccount,c.revenueDescription,c.revenueDescription,c.revenueTaxRate || 'BAS Excluded','credit',amount,debits,credits,'revenue_export');
        else {
          unresolvedLine(day,c.revenueDescription || `Product Revenue - Other (${country})`,'credit',amount,debits,credits,'revenue_export');
          audits.push({ severity:'REVIEW', category:'GL Mapping', date:day, order:'', country, message:`Revenue for ${country} has no configured GL account. Map it before Xero export.` });
        }
      });

      const exportShipping = round2(countries.reduce((s,c)=>s+(row[`Shipping (${c})`]||0),0));
      actualLine(day,cfg.exportShipping.accountCode,cfg.exportShipping.description,cfg.exportShipping.description,cfg.exportShipping.taxRate,'credit',exportShipping,debits,credits,'shipping_export');

      const debitBefore = round2(debits.reduce((s,l)=>s+(Number(l.Debit)||0),0));
      const creditBefore = round2(credits.reduce((s,l)=>s+(Number(l.Credit)||0),0));
      const balance = round2(debitBefore - creditBefore);
      row.rounding = balance;
      const maxRounding = Math.abs(Number(cfg.settings && cfg.settings.maxRounding) || 0.05);
      if (Math.abs(balance) <= maxRounding) {
        if (balance < 0) actualLine(day,cfg.rounding.accountCode,cfg.rounding.description,cfg.rounding.description,cfg.rounding.taxRate,'debit',-balance,debits,credits,'rounding');
        if (balance > 0) actualLine(day,cfg.rounding.accountCode,cfg.rounding.description,cfg.rounding.description,cfg.rounding.taxRate,'credit',balance,debits,credits,'rounding');
      } else {
        allBalanced = false; xeroReady = false;
        audits.push({ severity:'ERROR', category:'Rounding', date:day, order:'', country:'', message:`Unbalanced difference is ${balance.toFixed(2)}, above the configured rounding limit ${maxRounding.toFixed(2)}. It was NOT posted to account ${cfg.rounding.accountCode}.` });
      }

      const lines = debits.concat(credits);
      const totalDebit = round2(lines.reduce((s,l)=>s+(Number(l.Debit)||0),0));
      const totalCredit = round2(lines.reduce((s,l)=>s+(Number(l.Credit)||0),0));
      if (Math.abs(round2(totalDebit-totalCredit)) >= 0.01) { allBalanced = false; xeroReady = false; audits.push({ severity:'ERROR', category:'Journal Balance', date:day, order:'', country:'', message:`Journal is not balanced. Debit ${totalDebit.toFixed(2)}, Credit ${totalCredit.toFixed(2)}.` }); }

      lines.forEach(line => {
        journalLineRows.push(line);
        journalRows.push({ Date:line.Date, 'Account Code':line['Account Code'], Description:line.Description, Debit:line.Debit, Credit:line.Credit });
      });
      // The supplied Xero-import reference always places Rounding last for the day,
      // even when the review journal shows a debit rounding line with the debit block.
      const xeroOrder = lines.filter(l => l._meta !== 'rounding').concat(lines.filter(l => l._meta === 'rounding'));
      xeroOrder.filter(line => !line._unresolved).forEach(line => {
        xeroRows.push({
          '*Narration': `Sales for the period ${formatNarrationDate(day)}`,
          '*Date': formatDateMDY(day),
          'Description': line._xeroDescription,
          '*AccountCode': line['Account Code'],
          '*TaxRate': line._taxRate,
          '*Amount': round2((Number(line.Debit)||0) - (Number(line.Credit)||0)),
          'TrackingName1': '', 'TrackingOption1': '', 'TrackingName2': '', 'TrackingOption2': ''
        });
      });
      journalRows.push({ Date:day, 'Account Code':'', Description:`Total ${day}`, Debit:totalDebit, Credit:totalCredit });
      journalRows.push({ Date:'', 'Account Code':'', Description:'', Debit:'', Credit:'' });
      grandDebit += totalDebit; grandCredit += totalCredit;
    });
    journalRows.push({ Date:'', 'Account Code':'', Description:'GRAND TOTAL', Debit:round2(grandDebit), Credit:round2(grandCredit) });

    if (unresolvedCount) xeroReady = false;
    return { journalRows, journalLineRows, xeroRows, journalBalanced: allBalanced, xeroReady: xeroReady && allBalanced, unresolvedCount };
  }

  function toCSV(rows, columns, options) {
    options = options || {};
    if (!columns) columns = rows.length ? Object.keys(rows[0]) : [];
    const q = v => {
      if (v == null) return '';
      let s = String(v);
      if (options.twoDecimals && typeof v === 'number' && Number.isFinite(v)) s = v.toFixed(2);
      if (/[",\r\n]/.test(s)) s = `"${s.replace(/"/g,'""')}"`;
      return s;
    };
    return [columns.map(q).join(','), ...rows.map(r=>columns.map(c=>q(r[c])).join(','))].join('\r\n');
  }

  function reconciliationCSV(result) { return toCSV(result.reconciliationRows, result.reconciliationColumns); }
  function journalCSV(result) { return toCSV(result.journalRows, ['Date','Account Code','Description','Debit','Credit']); }
  function xeroCSV(result) { return toCSV(result.xeroRows, ['*Narration','*Date','Description','*AccountCode','*TaxRate','*Amount','TrackingName1','TrackingOption1','TrackingName2','TrackingOption2']); }

  return {
    DEFAULT_CONFIG,
    parseCSV,
    toCSV,
    reconciliationCSV,
    journalCSV,
    xeroCSV,
    process,
    normalizeGateway,
    normalizeCountry,
    normalizeDate,
    round2
  };
});
