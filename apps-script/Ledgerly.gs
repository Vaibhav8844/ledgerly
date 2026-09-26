/**
 * LEDGERLY PORTFOLIO — ISOLATED GOOGLE APPS SCRIPT BACKEND
 *
 * This project owns ONLY investment data. Finance Assistant remains a
 * separate Apps Script project + Google Sheet and is accessed through a
 * private server-to-server API.
 */

const LEDGERLY_CONFIG = {
  spreadsheetId: "1nw8QpT85epAlmAmtNf3yLZ8JrM5uOb3LqHSTgukcf2Q",
  financeApiUrlProperty: "FINANCE_ASSISTANT_API_URL",
  financeSecretProperty: "FINANCE_ASSISTANT_API_SECRET"
};

const LEDGERLY_PERF = {
  portfolioCacheSeconds: 15, dataCacheSeconds: 300, quoteCacheSeconds: 30,
  txKey: "ledgerly_tx_v1", quotesKey: "ledgerly_quotes_v1",
  snapshotsKey: "ledgerly_snapshots_v1", mfKey: "ledgerly_mf_v1", stocksKey: "ledgerly_stocks_v1",
  refreshKey: "LEDGERLY_LAST_QUOTE_REFRESH"
};

const SHEETS = {
  transactions: 'Transactions',
  quotes: 'Quotes',
  snapshots: 'Snapshots',
  lots: 'Lots',
  settings: 'Settings',
  holidays: 'MarketHolidays',
  mutualFunds: 'MutualFunds',
  stocks: 'ImportedStocks'
};

const MUTUAL_FUND_HEADERS = ['Scheme Name','AMC','Category','Sub-category','Folio No.','Source','Units','Invested Value','Current Value','Returns','XIRR'];
const STOCK_HEADERS = ['Stock Name','ISIN','Quantity','Average buy price','Buy value','Closing price','Closing value','Unrealised P&L'];
const TRANSACTION_HEADERS = ['ID','Date','Symbol','Asset Type','Transaction Type','Quantity','Price','Charges','Funding Account','Notes','Created At','Finance Transaction ID'];

function setupLedgerlySheet() {
  const ss = getLedgerlySpreadsheet_();
  ensureSheet_(ss, SHEETS.transactions, TRANSACTION_HEADERS);
  ensureSheet_(ss, SHEETS.quotes, ['Symbol','LTP','Previous Close','Updated At']);
  ensureSheet_(ss, SHEETS.snapshots, ['Date','Portfolio Value','Invested','P&L']);
  ensureSheet_(ss, SHEETS.lots, ['Lot ID','Buy Transaction ID','Symbol','Asset Type','Buy Date','Original Qty','Remaining Qty','Buy Price','Cost','Status']);
  ensureSheet_(ss, SHEETS.settings, ['Key','Value']);
  ensureSheet_(ss, SHEETS.holidays, ['Date','Description']);
  ensureSheet_(ss, SHEETS.mutualFunds, MUTUAL_FUND_HEADERS);
  ensureSheet_(ss, SHEETS.stocks, STOCK_HEADERS);
  const settings = ss.getSheetByName(SHEETS.settings);
  if (settings.getLastRow() === 1) settings.getRange(2,1,4,2).setValues([
    ['Accounting Method','FIFO'], ['Currency','INR'], ['Price Source','GoogleFinance / NSE'], ['Created','Ledgerly Portfolio']
  ]);
  seed2026Holidays_();
  rebuildLots_();
  installLedgerlyMaintenanceTrigger_();
  return {success:true, message:'Isolated Ledgerly sheets are ready.'};
}

function installLedgerlyMaintenanceTrigger_(){
  ScriptApp.getProjectTriggers().forEach(t=>{
    if(t.getHandlerFunction()==='ledgerlyMaintenance_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('ledgerlyMaintenance_').timeBased().everyMinutes(5).create();
}

function ledgerlyMaintenance_(){
  try{rebuildLots_();}catch(err){console.error(err);}
}

function setFinanceAssistantConnection(apiUrl, secret) {
  apiUrl = String(apiUrl || '').trim().replace(/\/$/, '');
  secret = String(secret || '').trim();
  if (!apiUrl || !/^https:\/\//i.test(apiUrl)) throw new Error('Finance Assistant API URL must be an HTTPS /exec URL.');
  if (!secret || secret.length < 24) throw new Error('Finance Assistant API secret must be at least 24 characters.');
  PropertiesService.getScriptProperties().setProperties({
    [LEDGERLY_CONFIG.financeApiUrlProperty]: apiUrl,
    [LEDGERLY_CONFIG.financeSecretProperty]: secret
  }, true);
  return {success:true,message:'Finance Assistant connection saved in Ledgerly Script Properties.'};
}

function getFinanceConnection_() {
  const p = PropertiesService.getScriptProperties();
  const url = String(p.getProperty(LEDGERLY_CONFIG.financeApiUrlProperty) || '').trim().replace(/\/$/, '');
  const secret = String(p.getProperty(LEDGERLY_CONFIG.financeSecretProperty) || '').trim();
  if (!url || !secret) throw new Error('Finance Assistant connection is not configured. Run setFinanceAssistantConnection(url, secret).');
  return {url, secret};
}

function callFinanceAssistant_(payload) {
  const c = getFinanceConnection_();
  const body = Object.assign({}, payload, {apiKey:c.secret});
  const response = UrlFetchApp.fetch(c.url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
    followRedirects: true
  });
  const code = response.getResponseCode();
  const raw = response.getContentText();
  let data;
  try { data = JSON.parse(raw); } catch (_) { throw new Error('Finance Assistant returned a non-JSON response (HTTP '+code+').'); }
  if (code < 200 || code >= 300 || data.success === false) {
    throw new Error(data.message || data.error || ('Finance Assistant HTTP '+code));
  }
  return data;
}

function doGet(e) { return handle_(e, 'GET'); }
function doPost(e) { return handle_(e, 'POST'); }

function handle_(e, method) {
  try {
    const params = (e && e.parameter) || {};
    let body = {};
    if (method === 'POST' && e && e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch (_) { body = params; }
    }
    const action = body.action || params.action || 'portfolio';
    let result;
    if (action === 'portfolio') result = getPortfolio_(body.period || params.period || 'this-month');
    else if (action === 'transactions') result = {success:true,transactions: readTransactions_()};
    else if (action === 'quotes') result = {success:true,quotes: readQuotes_()};
    else if (action === 'lots') result = {success:true,lots: calculateAllLots_().map(l=>({lotId:l.lotId,transactionId:l.transactionId,symbol:l.symbol,assetType:l.assetType,buyDate:l.buyDate,originalQty:l.originalQty,remainingQty:l.remainingQty,buyPrice:l.buyPrice,cost:l.originalQty*l.buyPrice,status:l.remainingQty>0?'OPEN':'CLOSED'}))};
    else if (action === 'marketStatus') result = getMarketStatus_();
    else if (action === 'refreshQuotes') result = refreshLiveQuotes_();
    else if (action === 'transaction') result = addTransaction_(body);
    else if (action === 'deleteTransaction') result = deleteTransaction_(body.id || params.id);
    else if (action === 'quote') result = upsertQuote_(body);
    else if (action === 'snapshot') result = saveSnapshot_(body);
    else if (action === 'importMutualFunds') result = importMutualFunds_(body.rows || []);
    else if (action === 'importData') result = importData_(body.kind, body.rows || []);
    else if (action === 'updateHolding') result = updateHolding_(body);
    else if (action === 'updateBankAccount') result = updateBankAccount_(body);
    else if (action === 'setup') result = setupLedgerlySheet();
    else if (action === 'connectionTest') result = connectionTest_();
    else throw new Error('Unknown action: ' + action);
    return json_(result, params.callback || body.callback);
  } catch (err) {
    return json_({success:false,error:String(err.message || err)}, (e && e.parameter && e.parameter.callback));
  }
}

function connectionTest_() {
  const d = callFinanceAssistant_({action:'ledgerlyDashboardData', period:'this-month'});
  return {success:true,financeAssistant:true,bankAccounts:d.bankAccounts||[],expenseSummary:d.expenseSummary||{}};
}

function updateBankAccount_(input) {
  const accountId = String(input.accountId || input.id || '').trim();
  const name = String(input.name || input.accountName || '').trim();
  if (!accountId || !name) throw new Error('Bank account ID and name are required.');
  const result = callFinanceAssistant_({
    action:'ledgerlyUpdateBankAccount',
    accountId,
    id:accountId,
    name,
    accountName:name,
    type:String(input.type || 'Bank account').trim(),
    balance:Number(input.balance || 0)
  });
  return {success:true,account:result.account||result.bankAccount||null};
}

function testLedgerlyFinanceConnection() {
  const result = connectionTest_();
  Logger.log(JSON.stringify(result, null, 2));
}

function json_(obj, callback) {
  const text = JSON.stringify(obj);
  if (callback) return ContentService.createTextOutput(callback + '(' + text + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function getLedgerlySpreadsheet_() { return SpreadsheetApp.openById(LEDGERLY_CONFIG.spreadsheetId); }

function cacheGetJson_(key){try{const raw=CacheService.getScriptCache().get(key);return raw?JSON.parse(raw):null;}catch(_){return null;}}
function cachePutJson_(key,value,ttl){try{const raw=JSON.stringify(value);if(raw.length<95000)CacheService.getScriptCache().put(key,raw,ttl);}catch(_){}}
function cacheRemove_(key){try{CacheService.getScriptCache().remove(key);}catch(_) {}}
function invalidateLedgerlyCaches_(){[LEDGERLY_PERF.txKey,LEDGERLY_PERF.quotesKey,LEDGERLY_PERF.snapshotsKey,LEDGERLY_PERF.mfKey,LEDGERLY_PERF.stocksKey,'ledgerly_portfolio_this-month','ledgerly_portfolio_last-month','ledgerly_portfolio_last-3-months','ledgerly_portfolio_last-6-months','ledgerly_portfolio_this-year','ledgerly_portfolio_all','ledgerly_portfolio_all-time'].forEach(cacheRemove_);}

function getPortfolio_(period) {
  period = String(period || "this-month");
  const cacheKey='ledgerly_portfolio_'+period;
  const cached=cacheGetJson_(cacheKey);
  if(cached)return cached;

  const transactions = readTransactions_();
  const quotes = readQuotes_();
  const snapshots = readSnapshots_();

  const calc = calculate_(transactions, quotes);

  // The transaction/lot ledger is primary. Imported holdings are also surfaced
  // when the same security is not already represented by a Ledgerly BUY/SELL
  // ledger, so an existing portfolio export can populate the dashboard.
  // If a symbol exists in the FIFO ledger, the transaction-derived position wins
  // to prevent double counting.
  const mutualFunds = readMutualFunds_();
  const importedStocks = readImportedStocks_();
  const transactionSymbols = new Set(calc.holdings.map(h => String(h.symbol||'').toUpperCase()));
  const importedMfHoldings = mutualFunds
    .filter(f => !transactionSymbols.has(String(f.schemeName||f.folioNo||'').toUpperCase()))
    .map(mutualFundHolding_);
  const importedStockHoldings = importedStocks
    .filter(s => !transactionSymbols.has(String(s.symbol||'').toUpperCase()))
    .map(s => importedStockHolding_(s, quotes));

  const holdings = calc.holdings
    .concat(importedMfHoldings, importedStockHoldings)
    .sort((a, b) => b.value - a.value);

  const invested = holdings.reduce(
    (s, h) => s + Number(h.invested || 0),
    0
  );

  const value = holdings.reduce(
    (s, h) => s + Number(h.value || 0),
    0
  );

  const todayPnl = holdings.reduce(
    (s, h) => s + Number(h.dayPnl || 0),
    0
  );

  // Finance Assistant remains the source of truth for bank cash flow,
  // income, expenses and investment cash flows.
  const finance = callFinanceAssistant_({
    action: "ledgerlyDashboardData",
    period: period
  });

  const bankAccounts = finance.bankAccounts || [];
  const bankTotal = Number(finance.bankTotal || 0);

  const stocks = holdings
    .filter(h =>
      h.assetType === "STOCK" ||
      h.assetType === "EQUITY" ||
      h.assetType === "ETF"
    )
    .reduce((s, h) => s + Number(h.value || 0), 0);

  const mfs = holdings
    .filter(h =>
      h.assetType === "MF" ||
      h.assetType === "MUTUAL FUND" ||
      h.assetType === "MUTUAL_FUND"
    )
    .reduce((s, h) => s + Number(h.value || 0), 0);

  const other = value - stocks - mfs;
  const liabilities = 0;

  const result = {
    success: true,

    // Investment data
    transactions,
    quotes,
    snapshots,
    mutualFunds,
    importedStocks,
    importedHoldings: {
      mutualFunds: importedMfHoldings,
      stocks: importedStockHoldings
    },
    holdings,
    lots: calc.lots,

    totals: {
      invested,
      value,
      pnl: value - invested,
      returnPct: invested
        ? (value - invested) / invested * 100
        : 0,
      realized: calc.realized,
      todayPnl
    },

    // Finance Assistant data
    bankAccounts,
    bankTotal,

    // Combined net worth
    netWorth: {
      total: bankTotal + stocks + mfs + other - liabilities,
      bank: bankTotal,
      stocks,
      mutualFunds: mfs,
      otherAssets: other,
      liabilities
    },

    // Detailed finance analytics for the React dashboard
    finance,

    market: getMarketStatus_()
  };
  cachePutJson_(cacheKey,result,LEDGERLY_PERF.portfolioCacheSeconds);
  return result;
}

function addTransaction_(input) {
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const type=String(input.type || input.transactionType || '').toUpperCase(), symbol=String(input.symbol || '').trim().toUpperCase(), date=String(input.date || '').trim(), assetType=String(input.assetType || 'STOCK').toUpperCase();
    const qty=Number(input.quantity), price=Number(input.price), charges=Number(input.charges || 0), funding=String(input.fundingAccount || input.account || '').trim(), note=String(input.note || input.notes || '');
    if(!['BUY','SELL'].includes(type))throw new Error('Transaction type must be BUY or SELL.');
    if(!symbol||!date||!(qty>0)||!(price>0)||!(charges>=0))throw new Error('Check symbol, date, quantity, price and charges.');
    if(type==='SELL'){const available=calculate_(readTransactions_(),readQuotes_()).lots.filter(l=>l.symbol===symbol&&l.remainingQty>0).reduce((s,l)=>s+l.remainingQty,0);if(qty>available+1e-9)throw new Error('Cannot sell more than the available quantity for '+symbol+'.');}
    const amount=type==='BUY'?qty*price+charges:qty*price-charges;if(!(amount>0))throw new Error('Sale proceeds after charges must be positive.');
    const finance=callFinanceAssistant_({action:'ledgerlyCashFlow',financeType:type==='BUY'?'Investment':'Income',amount,account:funding,date,symbol,assetType,remarks:'Ledgerly '+(type==='BUY'?'investment purchase':'investment sale proceeds')});
    const financeId=String(finance.financeTransactionId||'');if(!financeId)throw new Error('Finance Assistant did not return a transaction ID.');
    const id='TXN-'+Utilities.getUuid().slice(0,8).toUpperCase();
    try{getLedgerlySpreadsheet_().getSheetByName(SHEETS.transactions).appendRow([id,date,symbol,assetType,type,qty,price,charges,funding,note,new Date(),financeId]);invalidateLedgerlyCaches_();}
    catch(err){try{callFinanceAssistant_({action:'ledgerlyDeleteFinanceTransaction',financeTransactionId:financeId});}catch(_){}throw err;}
    return {success:true,id,transaction:{id,date,symbol,assetType,type,quantity:qty,price,charges,fundingAccount:funding,note,createdAt:Date.now(),financeTransactionId:financeId}};
  } finally {lock.releaseLock();}
}

function deleteTransaction_(id) {
  if(!id) throw new Error('Transaction ID is required.');
  const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.transactions);
  const values=sh.getDataRange().getValues();
  for(let r=1;r<values.length;r++) if(String(values[r][0])===String(id)){
    const financeId=String(values[r][11]||'');
    sh.deleteRow(r+1);
    if(financeId) callFinanceAssistant_({action:'ledgerlyDeleteFinanceTransaction',financeTransactionId:financeId});
    invalidateLedgerlyCaches_();
    return {success:true};
  }
  throw new Error('Transaction not found.');
}

function readTransactions_(){const cached=cacheGetJson_(LEDGERLY_PERF.txKey);if(cached)return cached;const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.transactions);if(!sh||sh.getLastRow()<2)return [];const rows=sh.getDataRange().getValues(),headers=rows[0].map(v=>String(v).trim()),idx={};headers.forEach((h,i)=>idx[h]=i);const out=rows.slice(1).filter(r=>r[idx.ID]).map(r=>({id:String(r[idx.ID]),date:formatDate_(r[idx.Date]),symbol:String(r[idx.Symbol]).toUpperCase(),assetType:String(r[idx['Asset Type']]||'STOCK').toUpperCase(),type:String(r[idx['Transaction Type']]).toUpperCase(),quantity:Number(r[idx.Quantity]),price:Number(r[idx.Price]),charges:Number(r[idx.Charges]||0),fundingAccount:String(r[idx['Funding Account']]||''),note:String(r[idx.Notes]||''),createdAt:new Date(r[idx['Created At']]||r[idx.Date]).getTime(),financeTransactionId:String(r[idx['Finance Transaction ID']]||''),total:Number(r[idx.Quantity])*Number(r[idx.Price])+(String(r[idx['Transaction Type']]).toUpperCase()==='SELL'?-Number(r[idx.Charges]||0):Number(r[idx.Charges]||0))})).sort((a,b)=>b.date.localeCompare(a.date)||b.createdAt-a.createdAt);cachePutJson_(LEDGERLY_PERF.txKey,out,LEDGERLY_PERF.dataCacheSeconds);return out;}

function calculate_(transactions, quotes) {
  const lots = [], realized = {value:0};
  transactions.slice().sort((a,b)=>a.date.localeCompare(b.date)||a.createdAt-b.createdAt).forEach(t=>{
    if (t.type==='BUY') lots.push({lotId:'LOT-'+t.id,symbol:t.symbol,buyDate:t.date,originalQty:t.quantity,remainingQty:t.quantity,buyPrice:t.price,cost:t.quantity*t.price+t.charges,buyCharges:t.charges,assetType:t.assetType,transactionId:t.id});
    else {
      let remaining=t.quantity;
      for (const lot of lots.filter(l=>l.symbol===t.symbol&&l.remainingQty>0)) {
        if(remaining<=0) break;
        const used=Math.min(remaining,lot.remainingQty);
        const buyUnitCost=lot.buyPrice+(lot.buyCharges/lot.originalQty);
        realized.value += used*(t.price-buyUnitCost)-(t.charges*used/t.quantity);
        lot.remainingQty-=used; remaining-=used;
      }
    }
  });
  const bySymbol={};
  lots.filter(l=>l.remainingQty>1e-9).forEach(l=>{
    if (!bySymbol[l.symbol]) bySymbol[l.symbol] = {symbol:l.symbol,qty:0,invested:0,lots:[],assetType:l.assetType};
    bySymbol[l.symbol].qty += l.remainingQty;
    bySymbol[l.symbol].invested += l.remainingQty*l.buyPrice;
    bySymbol[l.symbol].lots.push(l);
  });
  const holdings=Object.values(bySymbol).map(h=>{
    const q=Number(quotes[h.symbol]?.price||0), prev=Number(quotes[h.symbol]?.previousClose||0);
    const value=q>0?h.qty*q:h.invested, pnl=value-h.invested, dayPnl=(q>0&&prev>0)?h.qty*(q-prev):0;
    return {...h,avg:h.qty?h.invested/h.qty:0,value,pnl,returnPct:h.invested?pnl/h.invested*100:0,ltp:q,previousClose:prev,dayPnl};
  }).sort((a,b)=>b.value-a.value);
  return {holdings,lots,realized:realized.value};
}

function rebuildLots_() {
  const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.lots); if(!sh)return;
  const allLots=calculateAllLots_();
  if(sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();
  if(allLots.length) sh.getRange(2,1,allLots.length,10).setValues(allLots.map(l=>[
    l.lotId,l.transactionId,l.symbol,l.assetType,l.buyDate,l.originalQty,l.remainingQty,
    l.buyPrice,l.originalQty*l.buyPrice,l.remainingQty>0?'OPEN':'CLOSED'
  ]));
}
function calculateAllLots_(){
  const transactions=readTransactions_().sort((a,b)=>a.createdAt-b.createdAt),lots=[];
  transactions.forEach(t=>{
    if(t.type==='BUY') lots.push({
      lotId:'LOT-'+t.id,transactionId:t.id,symbol:t.symbol,assetType:t.assetType,
      buyDate:t.date,originalQty:t.quantity,remainingQty:t.quantity,buyPrice:t.price,buyCharges:t.charges
    });
    else {
      let rem=t.quantity;
      for(const l of lots.filter(x=>x.symbol===t.symbol&&x.remainingQty>0)){
        const used=Math.min(rem,l.remainingQty); l.remainingQty-=used; rem-=used;
        if(rem<=0)break;
      }
    }
  });
  return lots;
}

function readQuotes_(){const cached=cacheGetJson_(LEDGERLY_PERF.quotesKey);if(cached)return cached;const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.quotes),out={};if(!sh||sh.getLastRow()<2)return out;sh.getRange(2,1,sh.getLastRow()-1,4).getValues().filter(r=>r[0]).forEach(r=>out[String(r[0]).toUpperCase()]={price:Number(r[1]),previousClose:Number(r[2]||0),updatedAt:formatDateTime_(r[3])});cachePutJson_(LEDGERLY_PERF.quotesKey,out,LEDGERLY_PERF.dataCacheSeconds);return out;}
function readMutualFunds_(){
  const cached=cacheGetJson_(LEDGERLY_PERF.mfKey);
  if(cached)return cached;
  const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.mutualFunds);
  if(!sh||sh.getLastRow()<2)return [];
  const out=sh.getRange(2,1,sh.getLastRow()-1,MUTUAL_FUND_HEADERS.length).getValues()
    .filter(r=>r[0])
    .map(r=>({
      schemeName:String(r[0]||''),
      amc:String(r[1]||''),
      category:String(r[2]||''),
      subCategory:String(r[3]||''),
      folioNo:String(r[4]||''),
      source:String(r[5]||''),
      units:Number(r[6])||0,
      invested:Number(r[7])||0,
      value:Number(r[8])||0,
      returns:Number(r[9])||0,
      xirr:Number(r[10])||0
    }));
  cachePutJson_(LEDGERLY_PERF.mfKey,out,LEDGERLY_PERF.dataCacheSeconds);
  return out;
}
function readImportedStocks_(){
  const cached=cacheGetJson_(LEDGERLY_PERF.stocksKey);
  if(cached)return cached;
  const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.stocks);
  if(!sh||sh.getLastRow()<2)return [];
  const out=sh.getRange(2,1,sh.getLastRow()-1,STOCK_HEADERS.length).getValues()
    .filter(r=>r[0])
    .map(r=>({
      symbol:String(r[0]||''),
      isin:String(r[1]||''),
      qty:Number(r[2])||0,
      avg:Number(r[3])||0,
      invested:Number(r[4])||0,
      ltp:Number(r[5])||0,
      value:Number(r[6])||0,
      pnl:Number(r[7])||0
    }));
  cachePutJson_(LEDGERLY_PERF.stocksKey,out,LEDGERLY_PERF.dataCacheSeconds);
  return out;
}
function mutualFundHolding_(fund){
  const symbol=fund.schemeName||fund.folioNo||'Mutual fund';
  const qty=Number(fund.units||0);
  const avg=qty?Number(fund.invested||0)/qty:0;
  const value=Number(fund.value||0);
  const invested=Number(fund.invested||0);
  const pnl=value-invested;
  return {symbol,displayName:fund.schemeName,qty,invested,value,pnl,returnPct:invested?pnl/invested*100:0,avg,ltp:qty?value/qty:0,previousClose:0,dayPnl:0,assetType:'MF',lots:[],mutualFund:fund};
}
function importedStockHolding_(stock,quotes){
  const key=String(stock.symbol||'').toUpperCase();
  const q=quotes&&quotes[key]?quotes[key]:null;
  const ltp=q&&Number(q.price)>0?Number(q.price):Number(stock.ltp||0);
  const qty=Number(stock.qty||0);
  const invested=Number(stock.invested||0);
  const value=ltp>0?qty*ltp:Number(stock.value||0);
  const pnl=value-invested;
  const previousClose=q?Number(q.previousClose||0):0;
  const dayPnl=(ltp>0&&previousClose>0)?qty*(ltp-previousClose):0;
  return {symbol:stock.symbol,displayName:stock.symbol,qty,invested,value,pnl,returnPct:invested?pnl/invested*100:0,avg:Number(stock.avg||0),ltp,previousClose,dayPnl,assetType:'STOCK',lots:[],importedStock:stock};
}
function readSnapshots_(){const cached=cacheGetJson_(LEDGERLY_PERF.snapshotsKey);if(cached)return cached;const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.snapshots);if(!sh||sh.getLastRow()<2)return [];const out=sh.getRange(2,1,sh.getLastRow()-1,4).getValues().filter(r=>r[0]).map(r=>({date:formatDate_(r[0]),value:Number(r[1]),invested:Number(r[2]),pnl:Number(r[3])}));cachePutJson_(LEDGERLY_PERF.snapshotsKey,out,LEDGERLY_PERF.dataCacheSeconds);return out;}
function importData_(kind, rows){if(kind==='stocks')return importStocks_(rows);if(kind==='mutualFunds')return importMutualFunds_(rows);throw new Error('Unsupported import type.');}
function updateHolding_(input){
  const kind=String(input.kind||'');
  if(kind==='stocks'){
    const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.stocks), values=sh&&sh.getDataRange().getValues();
    if(!sh||!input.symbol)throw new Error('Stock holding not found.');
    for(let row=1;row<values.length;row++)if(String(values[row][0])===String(input.symbol)){sh.getRange(row+1,1,1,STOCK_HEADERS.length).setValues([[String(input.symbol),String(input.isin||''),Number(input.qty)||0,Number(input.avg)||0,Number(input.invested)||0,Number(input.ltp)||0,Number(input.value)||0,Number(input.value||0)-Number(input.invested||0)]]);return {ok:true};}
  }
  if(kind==='mutualFunds'){
    const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.mutualFunds), values=sh&&sh.getDataRange().getValues();
    if(!sh||!input.schemeName)throw new Error('Mutual-fund holding not found.');
    for(let row=1;row<values.length;row++)if(String(values[row][0])===String(input.schemeName)&&String(values[row][4]||'')===String(input.folioNo||'')){sh.getRange(row+1,1,1,MUTUAL_FUND_HEADERS.length).setValues([[String(input.schemeName),String(input.amc||''),String(input.category||''),String(input.subCategory||''),String(input.folioNo||''),String(input.source||''),Number(input.units)||0,Number(input.invested)||0,Number(input.value)||0,Number(input.value||0)-Number(input.invested||0),Number(input.xirr)||0]]);return {ok:true};}
  }
  throw new Error('Holding not found or cannot be edited.');
}
function importStocks_(rows){
  if(!Array.isArray(rows))throw new Error('Stock rows must be an array.');
  const ss=getLedgerlySpreadsheet_(),sh=ensureSheet_(ss,SHEETS.stocks,STOCK_HEADERS);
  const clean=rows.filter(r=>r&&String(r.symbol||'').trim()).map(r=>STOCK_HEADERS.map((_,i)=>[r.symbol,r.isin,r.qty,r.avg,r.invested,r.ltp,r.value,r.pnl][i]));
  if(!clean.length)throw new Error('No stock rows found. Check the header names and data.');
  if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,STOCK_HEADERS.length).clearContent();
  sh.getRange(2,1,clean.length,STOCK_HEADERS.length).setValues(clean);
  cacheRemove_(LEDGERLY_PERF.stocksKey);invalidateLedgerlyCaches_();
  return {ok:true,count:clean.length,stocks:readImportedStocks_()};
}
function importMutualFunds_(rows){
  if(!Array.isArray(rows))throw new Error('Mutual fund rows must be an array.');
  const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.mutualFunds)||getLedgerlySpreadsheet_().insertSheet(SHEETS.mutualFunds);
  ensureSheet_(getLedgerlySpreadsheet_(),SHEETS.mutualFunds,MUTUAL_FUND_HEADERS);
  const required=['schemeName','amc','category','subCategory','folioNo','source','units','invested','value','returns','xirr'];
  const clean=rows.filter(r=>r&&String(r.schemeName||'').trim()).map(r=>required.map(key=>key==='schemeName'||key==='amc'||key==='category'||key==='subCategory'||key==='folioNo'||key==='source'?String(r[key]||'').trim():Number(r[key])||0));
  if(!clean.length)throw new Error('No mutual fund rows found. Check the header names and data.');
  if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,MUTUAL_FUND_HEADERS.length).clearContent();
  sh.getRange(2,1,clean.length,MUTUAL_FUND_HEADERS.length).setValues(clean);
  cacheRemove_(LEDGERLY_PERF.mfKey);invalidateLedgerlyCaches_();
  return {ok:true,count:clean.length,mutualFunds:readMutualFunds_()};
}
function upsertQuote_(input){
  const symbol=String(input.symbol||'').trim().toUpperCase(), price=Number(input.price); if(!symbol||!(price>0))throw new Error('Symbol and positive price are required.');
  const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.quotes),vals=sh.getDataRange().getValues();
  for(let r=1;r<vals.length;r++)if(String(vals[r][0]).toUpperCase()===symbol){sh.getRange(r+1,1,1,4).setValues([[symbol,price,Number(input.previousClose||vals[r][2]||0),new Date()]]);return {ok:true};}
  sh.appendRow([symbol,price,Number(input.previousClose||0),new Date()]); return {ok:true};
}
function saveSnapshot_(input){const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.snapshots);sh.appendRow([input.date||new Date(),Number(input.value||0),Number(input.invested||0),Number(input.pnl||0)]);return {ok:true};}

/** Refresh all currently held STOCK symbols using GoogleFinance formulas.
 * Google documents the price quote as real-time but delayed by up to 20 minutes.
 */
function refreshLiveQuotes_(){
  const last=Number(PropertiesService.getScriptProperties().getProperty(LEDGERLY_PERF.refreshKey)||0),now=Date.now();
  if(last&&now-last<LEDGERLY_PERF.quoteCacheSeconds*1000)return {ok:true,quotes:readQuotes_(),market:getMarketStatus_(),updated:0,cached:true,message:'Recent quote refresh already available.'};
  const status=getMarketStatus_(),txs=readTransactions_(),quotes=readQuotes_(),calc=calculate_(txs,quotes);
  const imported=readImportedStocks_().map(s=>String(s.symbol||'').toUpperCase()).filter(Boolean);
  const active=calc.holdings.filter(h=>['STOCK','EQUITY','ETF'].includes(h.assetType)&&Number(h.qty)>1e-9).map(h=>h.symbol);
  const symbols=[...new Set(active.concat(imported))];
  if(!symbols.length)return {ok:true,quotes,market:status,updated:0,message:'No stock holdings to refresh.'};
  const ss=getLedgerlySpreadsheet_(),tempName='_LedgerlyLiveQuotes';let temp=ss.getSheetByName(tempName);if(!temp)temp=ss.insertSheet(tempName);temp.clear();
  temp.getRange(1,1,1,4).setValues([['Symbol','LTP','Previous Close','Trade Time']]);temp.getRange(2,1,symbols.length,1).setValues(symbols.map(s=>[s]));
  temp.getRange(2,2,symbols.length,3).setFormulas(symbols.map(symbol=>[`=IFERROR(GOOGLEFINANCE("NSE:${symbol}","price"),"")`,`=IFERROR(GOOGLEFINANCE("NSE:${symbol}","closeyest"),"")`,`=IFERROR(GOOGLEFINANCE("NSE:${symbol}","tradetime"),"")`]));
  SpreadsheetApp.flush();Utilities.sleep(1500);
  const rows=temp.getRange(2,1,symbols.length,4).getValues(),quoteSheet=ss.getSheetByName(SHEETS.quotes);const existing=quoteSheet.getDataRange().getValues();const map={};for(let i=1;i<existing.length;i++)map[String(existing[i][0]).toUpperCase()]=i-1;
  const out=existing.slice(1).map(r=>r.slice(0,4));let updated=0;rows.forEach(r=>{const symbol=String(r[0]).toUpperCase(),price=Number(r[1]),prev=Number(r[2]);if(!(price>0))return;const row=[symbol,price,prev||Number(quotes[symbol]?.previousClose||0),new Date()];if(map[symbol]!==undefined)out[map[symbol]]=row;else out.push(row);updated++;});
  if(quoteSheet.getLastRow()>1)quoteSheet.getRange(2,1,quoteSheet.getLastRow()-1,4).clearContent();if(out.length)quoteSheet.getRange(2,1,out.length,4).setValues(out);
  cacheRemove_(LEDGERLY_PERF.quotesKey);PropertiesService.getScriptProperties().setProperty(LEDGERLY_PERF.refreshKey,String(now));invalidateLedgerlyCaches_();
  return {ok:true,quotes:readQuotes_(),market:getMarketStatus_(),updated,source:'GoogleFinance / NSE',delay:'Up to 20 minutes'};
}

function getMarketStatus_(){
  const now=new Date(), tz='Asia/Kolkata', day=Number(Utilities.formatDate(now,tz,'u')), date=Utilities.formatDate(now,tz,'yyyy-MM-dd'), time=Utilities.formatDate(now,tz,'HH:mm:ss');
  const holiday=getHoliday_(date);
  if(day>=6)return {open:false,label:'Market closed',reason:'Weekend',date,time,timeZone:'IST',nextOpen:'Monday 09:15 IST'};
  if(holiday)return {open:false,label:'Market closed',reason:holiday,date,time,timeZone:'IST',nextOpen:'Next trading day 09:15 IST'};
  if(time<'09:15:00')return {open:false,label:'Market closed',reason:'Before regular session',date,time,timeZone:'IST',nextOpen:'09:15 IST'};
  if(time<'15:30:00')return {open:true,label:'Market open',reason:'NSE regular equity session',date,time,timeZone:'IST',closesAt:'15:30 IST'};
  return {open:false,label:'Market closed',reason:'Regular session ended',date,time,timeZone:'IST',nextOpen:'Next trading day 09:15 IST'};
}
function getHoliday_(date){const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.holidays);if(!sh||sh.getLastRow()<2)return '';const vals=sh.getRange(2,1,sh.getLastRow()-1,2).getValues();for(const r of vals){if(formatDate_(r[0])===date)return String(r[1]||'NSE holiday');}return '';}
function seed2026Holidays_(){
  const sh=getLedgerlySpreadsheet_().getSheetByName(SHEETS.holidays); if(sh.getLastRow()>1)return;
  const rows=[
    ['2026-01-15','Municipal Corporation Election in Maharashtra'],['2026-01-26','Republic Day'],['2026-02-19','Chhatrapati Shivaji Maharaj Jayanti'],['2026-03-03','Holi'],['2026-03-19','Gudhi Padwa'],['2026-03-26','Ram Navami'],['2026-03-31','Mahavir Jayanti'],['2026-04-01','Annual Bank Closing'],['2026-04-03','Good Friday'],['2026-04-14','Dr. Babasaheb Ambedkar Jayanti'],['2026-05-01','Maharashtra Din / Buddha Pournima'],['2026-05-28','Bakri Id'],['2026-06-26','Muharram'],['2026-08-26','Id-E-Milad'],['2026-09-14','Ganesh Chaturthi'],['2026-10-02','Mahatma Gandhi Jayanti'],['2026-10-20','Dussehra'],['2026-11-10','Diwali (Bali Pratipada)'],['2026-11-24','Guru Nanak Jayanti'],['2026-12-25','Christmas']
  ]; sh.getRange(2,1,rows.length,2).setValues(rows.map(r=>[r[0],r[1]]));
}
function ensureSheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);return sh;}
function formatDate_(v){if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');return String(v||'');}
function formatDateTime_(v){if(!v)return '';if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm');return String(v);}