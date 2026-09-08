// ==========================================
// OWNER DASHBOARD - CONSOLIDATED API (V4 - OPTIMIZED)
// Perubahan utama dari V3:
// 1. getSheetData() -> setiap sheet hanya dibaca SEKALI per request,
//    dipakai ulang oleh semua fungsi kalkulasi yang butuh sheet itu.
// 2. Cache CacheService dipasang di level handleApiRequest untuk SEMUA
//    action (bukan cuma 'overview'), dengan opsi bypass ?refresh=1.
// 3. calcManualPL + getTrend digabung jadi calcPLAndTrend() -> sheet
//    yang sama (Penjualan_Header, Sales_*, Broker Orders, Jurnal_*)
//    cuma di-scan SEKALI untuk menghasilkan PL dan Trend bersamaan.
//    Action baru 'report' mengembalikan { pl, trend } dalam 1 call.
// 4. Lookup Chart of Account di loop jurnal diubah dari Array.find()
//    (O(n*m)) jadi Map lookup (O(n+m)).
// ==========================================
const OWNER_CACHE_TTL = 300; // 5 menit
const OWNER_SS = SpreadsheetApp.getActiveSpreadsheet();

// ------------------------------------------
// REQUEST-SCOPED SHEET CACHE
// PENTING: Apps Script bisa reuse global state antar eksekusi di
// container yang sama, jadi cache ini WAJIB direset di awal setiap
// request (lihat resetSheetCache() di handleApiRequest), supaya tidak
// pernah menyajikan data basi dari request sebelumnya.
// ------------------------------------------
let _sheetCache = {};

function resetSheetCache() {
  _sheetCache = {};
}

function getSheetData(name, numCols) {
  const key = name + '::' + numCols;
  if (_sheetCache.hasOwnProperty(key)) return _sheetCache[key];
  const sheet = OWNER_SS.getSheetByName(name);
  if (!sheet || sheet.getLastRow() <= 1) {
    _sheetCache[key] = [];
    return _sheetCache[key];
  }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
  _sheetCache[key] = data;
  return data;
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    const result = handleApiRequest(e);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createTemplateFromFile('index_owner')
    .evaluate()
    .setTitle('Dashboard Owner')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleApiRequest(e) {
  const action = e.parameter.action;
  const start = e.parameter.start || null;
  const end = e.parameter.end || null;
  const sub = e.parameter.sub || null;
  const forceRefresh = e.parameter.refresh === '1';

  const cache = CacheService.getScriptCache();
  const cacheKey = `own_${action}_v4_${start || 'x'}_${end || 'x'}_${sub || 'x'}`;

  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  }

  // Reset sheet cache di awal SETIAP request -> memastikan data segar,
  // sambil tetap dedup baca sheet DI DALAM request yang sama.
  resetSheetCache();

  let result;
  try {
    switch (action) {
      case 'overview': result = getOwnerOverview(start, end); break;
      case 'piutang': result = getPiutangDetail(start, end, sub); break;
      case 'hutang': result = getHutangDetail(start, end, sub); break;
      case 'deposit': result = calcDepositDetail(); break;
      case 'kontrak': result = calcKontrakJaminan(); break;
      case 'stok': result = calcStokValue(); break;
      case 'kas': result = calcKasSaldo(); break;
      case 'pl': result = calcManualPL(start, end); break;
      case 'trend': result = getTrend(start, end); break;
      case 'report': result = getReport(start, end); break; // pl + trend sekaligus
      default: result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  // Best-effort cache: kalau hasil terlalu besar (>100KB), put() akan
  // gagal diam-diam lewat try/catch ini, request tetap jalan normal.
  try { cache.put(cacheKey, JSON.stringify(result), OWNER_CACHE_TTL); } catch (err) {}

  return result;
}

// ==========================================
// HELPER: PARSE TANGGAL AMAN
// ==========================================
function parseOwnerDate(str) {
  if (!str) return null;
  const p = String(str).split('-');
  if (p.length !== 3) return null;
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function ownerDateRange(start, end) {
  const s = parseOwnerDate(start);
  const e = parseOwnerDate(end);
  if (e) e.setHours(23, 59, 59, 999);
  return { s, e };
}

function fmtDate(d) {
  if (!d) return '-';
  if (typeof d === 'string') return d;
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// ==========================================
// 1. OVERVIEW UTAMA
// ==========================================
function getOwnerOverview(startDate, endDate) {
  const brokerPiutang = calcBrokerBuyerPiutang();
  const salesPiutang = calcSalesCustomerPiutang();
  const posPiutang = calcPosCustomerPiutang();
  const totalPiutang = brokerPiutang.total + salesPiutang.total + posPiutang.total;

  const brokerHutang = calcBrokerFactoryHutang();
  const poHutang = calcPOSupplierHutang();
  const totalHutang = brokerHutang.total + poHutang.total;

  const deposit = calcDepositDetail();
  const kontrak = calcKontrakJaminan();
  const stok = calcStokValue();
  const kas = calcKasSaldo();
  const period = calcManualPL(startDate, endDate);

  return {
    piutang: {
      total: totalPiutang,
      breakdown: [
        { label: 'Broker (Buyer)', amount: brokerPiutang.total, count: brokerPiutang.count, sub: 'broker' },
        { label: 'Sales', amount: salesPiutang.total, count: salesPiutang.count, sub: 'sales' },
        { label: 'POS / Ecer', amount: posPiutang.total, count: posPiutang.count, sub: 'pos' }
      ]
    },
    hutang: {
      total: totalHutang,
      breakdown: [
        { label: 'Broker (Pabrik)', amount: brokerHutang.total, count: brokerHutang.count, sub: 'broker' },
        { label: 'Purchase Order', amount: poHutang.total, count: poHutang.count, sub: 'po' }
      ]
    },
    deposit: { total: deposit.total, count: deposit.count },
    kontrak: { total: kontrak.total, count: kontrak.count, nearMaturity: kontrak.nearMaturity },
    stok: { totalValue: stok.totalValue, totalQty: stok.totalQty, itemCount: stok.itemCount, lowStock: stok.lowStock },
    kas: { total: kas.total, accountCount: kas.accounts ? kas.accounts.length : 0 },
    period: period,
    asOf: fmtDate(new Date()),
    filter: { start: startDate, end: endDate }
  };
}

// ==========================================
// CALCULATIONS (SYNCED WITH POS SYSTEM)
// ==========================================

// --- PIUTANG ---
function calcBrokerBuyerPiutang() {
  try {
    const data = getSheetData('Broker Orders', 26);
    if (data.length === 0) return { total: 0, count: 0, list: [] };

    const custMap = {};
    getSheetData('Broker Customers', 2).forEach(r => custMap[r[0]] = r[1]);

    const map = {};
    data.forEach(r => {
      if (String(r[24] || '').trim() === 'Dibatalkan') return;
      const cCode = String(r[3] || '').trim();
      if (!cCode) return;
      const totalBill = Number(r[20]) || 0;
      const paid = Number(r[22]) || 0;
      const remaining = Math.max(0, totalBill - paid);
      if (remaining <= 0) return;

      if (!map[cCode]) map[cCode] = { code: cCode, name: custMap[cCode] || cCode, total: 0, paid: 0, remaining: 0, orders: 0 };
      map[cCode].total += totalBill;
      map[cCode].paid += paid;
      map[cCode].remaining += remaining;
      map[cCode].orders++;
    });

    const list = Object.values(map).sort((a, b) => b.remaining - a.remaining);
    return { total: list.reduce((s, d) => s + d.remaining, 0), count: list.length, list: list };
  } catch (e) { return { total: 0, count: 0, list: [], error: e.toString() }; }
}

function calcSalesCustomerPiutang() {
  try {
    const data = getSheetData('Piutang_Sales', 9);
    if (data.length === 0) return { total: 0, count: 0, list: [] };
    const list = data.filter(r => Number(r[4]) > 0).map(r => ({
      name: r[1], phone: r[2], address: r[3], piutang: Number(r[4]) || 0, limit: Number(r[5]) || 0, salesName: r[7] || 'Tanpa Sales'
    })).sort((a, b) => b.piutang - a.piutang);
    return { total: list.reduce((s, d) => s + d.piutang, 0), count: list.length, list: list };
  } catch (e) { return { total: 0, count: 0, list: [], error: e.toString() }; }
}

function calcPosCustomerPiutang() {
  try {
    const data = getSheetData('Penjualan_Header', 17);
    if (data.length === 0) return { total: 0, count: 0, list: [] };
    const map = {};

    data.forEach(r => {
      const status = String(r[11] || '').trim();
      if (status === 'Dibatalkan') return;

      if (status === 'Belum Lunas' || status === 'Sebagian') {
        const customer = String(r[2] || '').trim() || 'Umum';
        const hutang = Number(r[15]) || Number(r[8]) || 0;
        const bayar = Number(r[16]) || 0;
        const sisa = Math.max(0, hutang - bayar);

        if (sisa > 0) {
          if (!map[customer]) map[customer] = { name: customer, piutang: 0, transactions: 0 };
          map[customer].piutang += sisa;
          map[customer].transactions++;
        }
      }
    });

    const list = Object.values(map).sort((a, b) => b.piutang - a.piutang);
    return { total: list.reduce((s, d) => s + d.piutang, 0), count: list.length, list: list };
  } catch (e) { return { total: 0, count: 0, list: [], error: e.toString() }; }
}

// --- HUTANG ---
function calcBrokerFactoryHutang() {
  try {
    if (typeof getFactoryDebtReport === 'function') {
      const r = getFactoryDebtReport();
      return { total: r.reduce((s, d) => s + Number(d.remaining || 0), 0), count: r.length, list: r };
    }
    return { total: 0, count: 0, list: [] };
  } catch (e) { return { total: 0, count: 0, list: [], error: e.toString() }; }
}

function calcPOSupplierHutang() {
  try {
    const data = getSheetData('Purchase Order', 23);
    if (data.length === 0) return { total: 0, count: 0, list: [] };

    const depositCodes = new Set();
    const suppMap = {};
    getSheetData('Suppliers', 4).forEach(r => {
      if (String(r[1] || '').trim().toLowerCase() === 'deposit') depositCodes.add(String(r[0]).trim());
      suppMap[String(r[0]).trim()] = r[2];
    });

    const map = {};
    const seenRef = new Set();
    data.forEach(r => {
      const ref = String(r[1] || '').trim();
      const suppCode = String(r[2] || '').trim();
      if (String(r[18] || '').toLowerCase() === 'cancel') return;
      if (depositCodes.has(suppCode)) return;
      if (seenRef.has(ref)) return;
      seenRef.add(ref);

      const total = Number(r[12]) || 0;
      const paid = Number(r[15]) || 0;
      const sisa = Number(r[21]) || 0;

      if (sisa <= 0) return;
      if (!map[suppCode]) map[suppCode] = { code: suppCode, name: suppMap[suppCode] || suppCode, total: 0, paid: 0, sisa: 0, orders: 0 };
      map[suppCode].total += total;
      map[suppCode].paid += paid;
      map[suppCode].sisa += sisa;
      map[suppCode].orders++;
    });

    const list = Object.values(map).sort((a, b) => b.sisa - a.sisa);
    return { total: list.reduce((s, d) => s + d.sisa, 0), count: list.length, list: list };
  } catch (e) { return { total: 0, count: 0, list: [], error: e.toString() }; }
}

// --- DEPOSIT ---
function calcDepositDetail() {
  try {
    if (typeof getDepositBalances === 'function') {
      const balances = getDepositBalances();
      const total = balances.reduce((sum, item) => sum + (item.balance || 0), 0);
      return { total: total, count: balances.length, list: balances };
    }

    const depositCodes = new Set();
    const supplierMap = {};
    getSheetData('Suppliers', 4).forEach(r => {
      const code = String(r[0] || '').trim();
      if (String(r[1] || '').trim().toLowerCase() === 'deposit') depositCodes.add(code);
      supplierMap[code] = r[2] || code;
    });

    let transactions = [];

    getSheetData('Deposit Supplier', 8).forEach(r => {
      const code = String(r[2] || '').trim();
      if (depositCodes.has(code)) {
        transactions.push({ supplierCode: code, supplierName: r[3] || supplierMap[code] || code, nominal: Number(r[7]) || 0 });
      }
    });

    const poData = getSheetData('Purchase Order', 23);
    const processedRefNos = new Set();
    poData.forEach(r => {
      const suppCode = String(r[2] || '').trim();
      const refNo = String(r[1] || '').trim();
      const status = String(r[18] || '').toLowerCase();
      if (depositCodes.has(suppCode) && status !== 'cancel' && !processedRefNos.has(refNo)) {
        const amountPaid = Number(r[12]) || 0;
        if (amountPaid > 0) {
          processedRefNos.add(refNo);
          transactions.push({ supplierCode: suppCode, supplierName: supplierMap[suppCode] || suppCode, nominal: -amountPaid });
        }
      }
    });

    getSheetData('Broker Orders', 26).forEach(r => {
      const factoryCode = String(r[4] || '').trim();
      const statusbroker = String(r[24] || '').toLowerCase();
      if (depositCodes.has(factoryCode) && statusbroker !== 'dibatalkan') {
        const sellPrice = Number(r[19]) || 0;
        if (sellPrice > 0) {
          transactions.push({ supplierCode: factoryCode, supplierName: supplierMap[factoryCode] || factoryCode, nominal: -sellPrice });
        }
      }
    });

    const map = {};
    depositCodes.forEach(code => { map[code] = { code: code, name: supplierMap[code] || code, balance: 0 }; });
    transactions.forEach(t => { if (map[t.supplierCode]) map[t.supplierCode].balance += t.nominal; });

    const list = Object.values(map).filter(d => d.balance !== 0).sort((a, b) => b.balance - a.balance);
    return { total: list.reduce((s, d) => s + d.balance, 0), count: list.length, list: list };
  } catch (e) {
    return { total: 0, count: 0, list: [], error: e.toString() };
  }
}

// --- KONTRAK JAMINAN ---
function calcKontrakJaminan() {
  try {
    const data = getSheetData('Kontrak Jaminan Supplier', 13);
    if (data.length === 0) return { total: 0, count: 0, nearMaturity: 0, list: [] };

    const now = new Date();
    const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const next7 = new Date(); next7.setDate(now.getDate() + 7);
    const next7Str = Utilities.formatDate(next7, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    let total = 0, nearMaturity = 0;
    const list = [];
    data.forEach(r => {
      if (String(r[9] || '').trim() !== 'Aktif') return;
      const nominal = Number(r[5]) || 0;
      total += nominal;
      const category = String(r[1] || '').trim();
      const dueDateObj = r[7];
      let dueStr = '', isNear = false;
      if (dueDateObj && category !== 'Unlimited') {
        dueStr = Utilities.formatDate(new Date(dueDateObj), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        if (dueStr >= todayStr && dueStr <= next7Str) { nearMaturity++; isNear = true; }
      }
      list.push({ code: r[0], category, supplierName: r[3], nominal, dueDate: dueStr, nearMaturity: isNear });
    });
    return { total, count: list.length, nearMaturity, list };
  } catch (e) { return { total: 0, count: 0, nearMaturity: 0, list: [], error: e.toString() }; }
}

// --- STOK ---
function calcStokValue() {
  try {
    const data = getSheetData('Stok Barang', 8);
    if (data.length === 0) return { totalValue: 0, totalQty: 0, itemCount: 0, lowStock: 0, list: [] };

    let totalValue = 0, totalQty = 0, lowStock = 0;
    const list = data.map(r => {
      const qty = Number(r[5]) || 0;
      const hpp = Number(r[6]) || 0;
      const value = qty * hpp;
      totalValue += value; totalQty += qty;
      if (qty <= 0) lowStock++;
      return { code: String(r[0]).trim(), name: r[1], stock: qty, hpp: hpp, value: value };
    }).sort((a, b) => b.value - a.value);
    return { totalValue, totalQty, itemCount: list.length, lowStock, list };
  } catch (e) { return { totalValue: 0, totalQty: 0, itemCount: 0, lowStock: 0, list: [], error: e.toString() }; }
}

// --- KAS & BANK ---
function calcKasSaldo() {
  try {
    const dashboardData = AkuntansiCore.getCashDashboardData();
    return {
      total: dashboardData.totalCash || 0,
      accounts: (dashboardData.accountDetails || []).filter(a => a.balance !== 0).sort((a, b) => b.balance - a.balance)
    };
  } catch (e) {
    return { total: 0, accounts: [], error: "Gagal memuat data dari Library Akuntansi: " + e.toString() };
  }
}

// ==========================================
// PROFIT/LOSS + TREND (DIGABUNG - satu kali scan sheet untuk dua-duanya)
// ==========================================
function calcPLAndTrend(startDate, endDate) {
  try {
    const { s, e } = ownerDateRange(startDate, endDate);
    const start = s ? s.getTime() : 0;
    const end = e ? e.getTime() : 9999999999999;
    const rangeDays = (end - start) / (1000 * 60 * 60 * 24);
    const useMonthly = rangeDays > 90;

    let revenue = 0, expense = 0, cogs = 0;
    let segDRev = 0, segDCogs = 0, segBRev = 0, segBCosts = 0;
    const revenueByAcc = {}, expenseByAcc = {};
    const trendMap = {};

    const addRev = (accName, val) => { val = Number(val) || 0; if (val > 0) { revenue += val; revenueByAcc[accName] = (revenueByAcc[accName] || 0) + val; } };
    const addExp = (accName, val) => { val = Number(val) || 0; if (val > 0) { expense += val; expenseByAcc[accName] = (expenseByAcc[accName] || 0) + val; } };
    const addTrend = (date, rev, exp) => {
      const d = new Date(date);
      const key = useMonthly
        ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`
        : `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
      if (!trendMap[key]) trendMap[key] = { rev: 0, exp: 0 };
      trendMap[key].rev += rev; trendMap[key].exp += exp;
    };

    // 1. Penjualan POS
    getSheetData('Penjualan_Header', 17).forEach(r => {
      if (String(r[11] || '').toLowerCase() === 'dibatalkan') return;
      const dt = new Date(r[1]).getTime();
      if (dt < start || dt > end) return;

      const grandTotal = Number(r[8]) || 0;
      const totalHpp = Number(r[9]) || 0;

      if (grandTotal > 0) { addRev('Pendapatan Penjualan POS', grandTotal); segDRev += grandTotal; }
      if (totalHpp > 0) { addExp('HPP Penjualan POS', totalHpp); cogs += totalHpp; segDCogs += totalHpp; }
      addTrend(r[1], grandTotal, totalHpp);
    });

    // 2. Sales
    const dMap = {};
    getSheetData('Sales_Detail', 11).forEach(r => {
      dMap[r[0]] = (dMap[r[0]] || 0) + (Number(r[3]) || 0) * (Number(r[4]) || 0);
    });
    getSheetData('Sales_Header', 20).forEach(r => {
      if (String(r[11] || '').toLowerCase() === 'dibatalkan') return;
      const dt = new Date(r[1]).getTime();
      if (dt < start || dt > end) return;
      const rev = (Number(r[4]) || 0) + (Number(r[6]) || 0) + (Number(r[7]) || 0);
      const hpp = dMap[r[0]] || (Number(r[9]) || 0);
      if (rev > 0) { addRev('Pendapatan Penjualan Sales', rev); segDRev += rev; }
      if (hpp > 0) { addExp('HPP Penjualan Sales', hpp); cogs += hpp; segDCogs += hpp; }
      addTrend(r[1], rev, hpp);
    });

    // 3. Broker
    getSheetData('Broker Orders', 26).forEach(r => {
      if (String(r[24] || '').toLowerCase() === 'dibatalkan') return;
      const dt = new Date(r[2]).getTime();
      if (dt < start || dt > end) return;
      const rev = (Number(r[6]) || 0) + (Number(r[10]) || 0) + (Number(r[14]) || 0) + (Number(r[17]) || 0);
      const exp = (Number(r[7]) || 0) + (Number(r[9]) || 0) + (Number(r[13]) || 0) + (Number(r[16]) || 0);
      if (rev > 0) { addRev('Pendapatan Broker', rev); segBRev += rev; }
      if (exp > 0) { addExp('HPP & Beban Broker', exp); cogs += (Number(r[7]) || 0); segBCosts += exp; }
      addTrend(r[2], rev, exp);
    });

    // 4. Jurnal Lainnya (hanya masuk PL, tidak masuk trend harian - sama seperti versi asli)
    const coaMap = {};
    getSheetData('Chart of Account', 9).forEach(r => {
      const code = String(r[0]);
      coaMap[code] = { name: String(r[4] || ''), balancePos: String(r[6] || '').toLowerCase(), reportType: String(r[7] || '') };
    });

    ['Jurnal Umum', 'Jurnal Manual', 'Jurnal Otomatis'].forEach(sheetName => {
      getSheetData(sheetName, 10).forEach(r => {
        const dt = new Date(r[0]).getTime();
        if (dt < start || dt > end) return;
        if (sheetName === 'Jurnal Umum') {
          const debitCode = String(r[5] || ''), creditCode = String(r[4] || ''), amount = Number(r[7]) || 0;
          const accD = coaMap[debitCode];
          if (accD && accD.reportType === 'Laba Rugi' && amount > 0) {
            if (accD.balancePos === 'kredit') addRev(accD.name, amount); else addExp(accD.name, amount);
          }
          const accC = coaMap[creditCode];
          if (accC && accC.reportType === 'Laba Rugi' && amount > 0) {
            if (accC.balancePos === 'kredit') addRev(accC.name, amount); else addExp(accC.name, amount);
          }
        } else {
          const code = String(r[3] || ''), debit = Number(r[5]) || 0, credit = Number(r[6]) || 0;
          const acc = coaMap[code];
          if (acc && acc.reportType === 'Laba Rugi') {
            if (acc.balancePos === 'kredit' && credit > 0) addRev(acc.name, credit);
            else if (acc.balancePos !== 'kredit' && debit > 0) addExp(acc.name, debit);
          }
        }
      });
    });

    const profit = revenue - expense;
    const margin = revenue > 0 ? (profit / revenue) * 100 : 0;
    const trendKeys = Object.keys(trendMap).sort();
    const overhead = Math.max(0, expense - segDCogs - segBCosts);
    const otherRevenue = Math.max(0, revenue - segDRev - segBRev);

    return {
      pl: {
        revenue, expense, cogs, profit, margin,
        segments: {
          direct: { revenue: segDRev, costs: segDCogs },
          broker: { revenue: segBRev, costs: segBCosts },
          otherRevenue,
          overhead
        },
        topRevenue: Object.entries(revenueByAcc).map(([k, v]) => ({ name: k, amount: v })).sort((a, b) => b.amount - a.amount).slice(0, 5),
        topExpense: Object.entries(expenseByAcc).map(([k, v]) => ({ name: k, amount: v })).sort((a, b) => b.amount - a.amount).slice(0, 5),
        filter: { start: startDate, end: endDate }
      },
      trend: {
        labels: trendKeys,
        revenue: trendKeys.map(k => trendMap[k].rev),
        expense: trendKeys.map(k => trendMap[k].exp),
        profit: trendKeys.map(k => trendMap[k].rev - trendMap[k].exp),
        isMonthly: useMonthly
      }
    };
  } catch (e) {
    return {
      pl: { error: e.toString() },
      trend: { labels: [], revenue: [], expense: [], profit: [], error: e.toString() }
    };
  }
}

function calcManualPL(startDate, endDate) { return calcPLAndTrend(startDate, endDate).pl; }
function getTrend(startDate, endDate) { return calcPLAndTrend(startDate, endDate).trend; }
function getReport(startDate, endDate) { return calcPLAndTrend(startDate, endDate); } // { pl, trend } dalam 1x scan

// --- ROUTERS ---
function getPiutangDetail(start, end, sub) {
  if (sub === 'broker') return calcBrokerBuyerPiutang();
  if (sub === 'sales') return calcSalesCustomerPiutang();
  if (sub === 'pos') return calcPosCustomerPiutang();
  return { broker: calcBrokerBuyerPiutang(), sales: calcSalesCustomerPiutang(), pos: calcPosCustomerPiutang() };
}

function getHutangDetail(start, end, sub) {
  if (sub === 'broker') return calcBrokerFactoryHutang();
  if (sub === 'po') return calcPOSupplierHutang();
  return { broker: calcBrokerFactoryHutang(), po: calcPOSupplierHutang() };
}