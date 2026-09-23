/** Ledgerly Portfolio â€” Google Apps Script backend
 * Live pricing: GoogleFinance-backed NSE quotes (delayed by up to 20 minutes).
 * Market status: NSE equity regular session, IST, with 2026 NSE holidays seeded.
 */

const SHEETS = {
  transactions: 'Transactions', quotes: 'Quotes', snapshots: 'Snapshots', lots: 'Lots',
  settings: 'Settings', holidays: 'MarketHolidays', mutualFunds: 'MutualFunds', stocks: 'ImportedStocks'
};

const MUTUAL_FUND_HEADERS = ['Scheme Name','AMC','Category','Sub-category','Folio No.','Source','Units','Invested Value','Current Value','Returns','XIRR'];
const STOCK_HEADERS = ['Stock Name','ISIN','Quantity','Average buy price','Buy value','Closing price','Closing value','Unrealised P&L'];

function setupLedgerlySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEETS.transactions, ['ID','Date','Symbol','Asset Type','Transaction Type','Quantity','Price','Charges','Notes','Created At']);
  ensureSheet_(ss, SHEETS.quotes, ['Symbol','LTP','Previous Close','Updated At']);
  ensureSheet_(ss, SHEETS.snapshots, ['Date','Portfolio Value','Invested','P&L']);
  ensureSheet_(ss, SHEETS.lots, ['Lot ID','Buy Transaction ID','Symbol','Buy Date','Original Qty','Remaining Qty','Buy Price','Cost','Status']);
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
  return {ok:true, message:'Ledgerly sheets are ready.'};
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
    if (action === 'portfolio') result = getPortfolio_();
    else if (action === 'transactions') result = {transactions: readTransactions_()};
    else if (action === 'quotes') result = {quotes: readQuotes_()};
    else if (action === 'lots') result = {lots: readLots_()};
    else if (action === 'marketStatus') result = getMarketStatus_();
    else if (action === 'refreshQuotes') result = refreshLiveQuotes_();
    else if (action === 'transaction') result = method === 'DELETE' ? deleteTransaction_(body.id || params.id) : addTransaction_(body);
    else if (action === 'deleteTransaction') result = deleteTransaction_(body.id || params.id);
    else if (action === 'quote') result = upsertQuote_(body);
    else if (action === 'snapshot') result = saveSnapshot_(body);
    else if (action === 'importMutualFunds') result = importMutualFunds_(body.rows || []);
    else if (action === 'importData') result = importData_(body.kind, body.rows || []);
    else if (action === 'updateHolding') result = updateHolding_(body);
    else if (action === 'setup') result = setupLedgerlySheet();
    else throw new Error('Unknown action: ' + action);
    return json_(result, params.callback || body.callback);
  } catch (err) {
    return json_({ok:false, error:String(err.message || err)}, (e && e.parameter && e.parameter.callback));
  }
}

function json_(obj, callback) {
  const text = JSON.stringify(obj);
  if (callback) return ContentService.createTextOutput(callback + '(' + text + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function getPortfolio_() {
  const transactions = readTransactions_();
  const quotes = readQuotes_();
  const snapshots = readSnapshots_();
  const calc = calculate_(transactions, quotes);
  const mutualFunds = readMutualFunds_();
  const importedStocks = readImportedStocks_();
  const holdings = calc.holdings.concat(mutualFunds.map(mutualFundHolding_), importedStocks.map(importedStockHolding_)).sort((a,b)=>b.value-a.value);
  const invested = holdings.reduce((s,h) => s + h.invested, 0);
  const value = holdings.reduce((s,h) => s + h.value, 0);
  const todayPnl = holdings.reduce((s,h) => s + (h.dayPnl || 0), 0);
  return {ok:true, transactions, quotes, snapshots, mutualFunds, importedStocks, holdings, lots:calc.lots,
    totals:{invested,value,pnl:value-invested,returnPct:invested ? (value-invested)/invested*100 : 0,realized:calc.realized,todayPnl},
    market:getMarketStatus_()};
}

function addTransaction_(input) {
  const type = String(input.type || '').toUpperCase();
  const symbol = String(input.symbol || '').trim().toUpperCase();
  const date = String(input.date || '').trim();
  const assetType = String(input.assetType || 'STOCK').toUpperCase();
  const qty = Number(input.quantity), price = Number(input.price), charges = Number(input.charges || 0);
  const note = String(input.note || '');
  if (!['BUY','SELL'].includes(type)) throw new Error('Transaction type must be BUY or SELL.');
  if (!symbol || !date || !(qty > 0) || !(price > 0) || !(charges >= 0)) throw new Error('Check symbol, date, quantity, price and charges.');
  const txs = readTransactions_();
  if (type === 'SELL') {
    const available = txs.filter(t => t.symbol === symbol).reduce((sum,t) => sum + (t.type === 'BUY' ? t.quantity : -t.quantity), 0);
    if (qty > available + 1e-9) throw new Error('Cannot sell more than the available quantity for ' + symbol + '.');
  }
  const id = 'TXN-' + Utilities.getUuid().slice(0,8).toUpperCase();
  SpreadsheetApp.getActive().getSheetByName(SHEETS.transactions).appendRow([id,date,symbol,assetType,type,qty,price,charges,note,new Date()]);
  rebuildLots_();
  return {ok:true, transaction: readTransactions_().find(t => t.id === id)};
}

function deleteTransaction_(id) {
  if (!id) throw new Error('Transaction ID is required.');
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.transactions), values = sh.getDataRange().getValues();
  for (let r=1;r<values.length;r++) if (String(values[r][0])===String(id)) { sh.deleteRow(r+1); rebuildLots_(); return {ok:true}; }
  throw new Error('Transaction not found.');
}

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
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.lots); if(!sh)return;
  const allLots=calculateAllLots_();
  if(sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();
  if(allLots.length) sh.getRange(2,1,allLots.length,9).setValues(allLots.map(l=>[l.lotId,l.transactionId,l.symbol,l.buyDate,l.originalQty,l.remainingQty,l.buyPrice,l.originalQty*l.buyPrice,l.remainingQty>0?'OPEN':'CLOSED']));
}
function calculateAllLots_(){
  const transactions=readTransactions_().sort((a,b)=>a.createdAt-b.createdAt),lots=[];
  transactions.forEach(t=>{if(t.type==='BUY')lots.push({lotId:'LOT-'+t.id,transactionId:t.id,symbol:t.symbol,buyDate:t.date,originalQty:t.quantity,remainingQty:t.quantity,buyPrice:t.price,buyCharges:t.charges});else{let rem=t.quantity;for(const l of lots.filter(x=>x.symbol===t.symbol&&x.remainingQty>0)){const used=Math.min(rem,l.remainingQty);l.remainingQty-=used;rem-=used;if(rem<=0)break;}}});
  return lots;
}

function readTransactions_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.transactions); if(!sh||sh.getLastRow()<2)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,10).getValues().filter(r=>r[0]).map(r=>({id:String(r[0]),date:formatDate_(r[1]),symbol:String(r[2]).toUpperCase(),assetType:String(r[3]||'STOCK').toUpperCase(),type:String(r[4]).toUpperCase(),quantity:Number(r[5]),price:Number(r[6]),charges:Number(r[7]||0),note:String(r[8]||''),createdAt:new Date(r[9]||r[1]).getTime(),total:Number(r[5])*Number(r[6])+Number(r[7]||0)})).sort((a,b)=>b.date.localeCompare(a.date)||b.createdAt-a.createdAt);
}
function readQuotes_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.quotes),out={}; if(!sh||sh.getLastRow()<2)return out;
  sh.getRange(2,1,sh.getLastRow()-1,4).getValues().filter(r=>r[0]).forEach(r=>out[String(r[0]).toUpperCase()]={price:Number(r[1]),previousClose:Number(r[2]||0),updatedAt:formatDateTime_(r[3])}); return out;
}
function readSnapshots_(){const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.snapshots);if(!sh||sh.getLastRow()<2)return [];return sh.getRange(2,1,sh.getLastRow()-1,4).getValues().filter(r=>r[0]).map(r=>({date:formatDate_(r[0]),value:Number(r[1]),invested:Number(r[2]),pnl:Number(r[3])}));}
function readLots_(){const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.lots);if(!sh||sh.getLastRow()<2)return [];return sh.getRange(2,1,sh.getLastRow()-1,9).getValues().filter(r=>r[0]).map(r=>({lotId:String(r[0]),transactionId:String(r[1]),symbol:String(r[2]),buyDate:formatDate_(r[3]),originalQty:Number(r[4]),remainingQty:Number(r[5]),buyPrice:Number(r[6]),cost:Number(r[7]),status:String(r[8])}));}
function readMutualFunds_(){const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.mutualFunds);if(!sh||sh.getLastRow()<2)return [];return sh.getRange(2,1,sh.getLastRow()-1,MUTUAL_FUND_HEADERS.length).getValues().filter(r=>r[0]).map(r=>({schemeName:String(r[0]),amc:String(r[1]||''),category:String(r[2]||''),subCategory:String(r[3]||''),folioNo:String(r[4]||''),source:String(r[5]||''),units:Number(r[6])||0,invested:Number(r[7])||0,value:Number(r[8])||0,returns:Number(r[9])||0,xirr:Number(r[10])||0}));}
function readImportedStocks_(){const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.stocks);if(!sh||sh.getLastRow()<2)return [];return sh.getRange(2,1,sh.getLastRow()-1,STOCK_HEADERS.length).getValues().filter(r=>r[0]).map(r=>({symbol:String(r[0]),isin:String(r[1]||''),qty:Number(r[2])||0,avg:Number(r[3])||0,invested:Number(r[4])||0,ltp:Number(r[5])||0,value:Number(r[6])||0,pnl:Number(r[7])||0}));}
function mutualFundHolding_(fund){const symbol=fund.schemeName||fund.folioNo||'Mutual fund';const qty=fund.units;const avg=qty?fund.invested/qty:0;const ltp=qty?fund.value/qty:0;const pnl=fund.value-fund.invested;return {symbol,displayName:fund.schemeName,qty,invested:fund.invested,value:fund.value,pnl,returnPct:fund.invested?pnl/fund.invested*100:0,avg,ltp,previousClose:0,dayPnl:0,assetType:'MF',lots:[],mutualFund:fund};}
function importedStockHolding_(stock){const pnl=stock.value-stock.invested;return {symbol:stock.symbol,displayName:stock.symbol,qty:stock.qty,invested:stock.invested,value:stock.value,pnl,returnPct:stock.invested?pnl/stock.invested*100:0,avg:stock.avg||0,ltp:stock.ltp,previousClose:0,dayPnl:0,assetType:'STOCK',lots:[],importedStock:stock};}
function importData_(kind, rows){if(kind==='stocks')return importStocks_(rows);if(kind==='mutualFunds')return importMutualFunds_(rows);throw new Error('Unsupported import type.');}
function updateHolding_(input){
  const kind=String(input.kind||'');
  if(kind==='stocks'){
    const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.stocks), values=sh&&sh.getDataRange().getValues();
    if(!sh||!input.symbol)throw new Error('Stock holding not found.');
    for(let row=1;row<values.length;row++)if(String(values[row][0])===String(input.symbol)){sh.getRange(row+1,1,1,STOCK_HEADERS.length).setValues([[String(input.symbol),String(input.isin||''),Number(input.qty)||0,Number(input.avg)||0,Number(input.invested)||0,Number(input.ltp)||0,Number(input.value)||0,Number(input.value||0)-Number(input.invested||0)]]);return {ok:true};}
  }
  if(kind==='mutualFunds'){
    const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.mutualFunds), values=sh&&sh.getDataRange().getValues();
    if(!sh||!input.schemeName)throw new Error('Mutual-fund holding not found.');
    for(let row=1;row<values.length;row++)if(String(values[row][0])===String(input.schemeName)&&String(values[row][4]||'')===String(input.folioNo||'')){sh.getRange(row+1,1,1,MUTUAL_FUND_HEADERS.length).setValues([[String(input.schemeName),String(input.amc||''),String(input.category||''),String(input.subCategory||''),String(input.folioNo||''),String(input.source||''),Number(input.units)||0,Number(input.invested)||0,Number(input.value)||0,Number(input.value||0)-Number(input.invested||0),Number(input.xirr)||0]]);return {ok:true};}
  }
  throw new Error('Holding not found or cannot be edited.');
}
function importStocks_(rows){
  if(!Array.isArray(rows))throw new Error('Stock rows must be an array.');
  const ss=SpreadsheetApp.getActive(),sh=ensureSheet_(ss,SHEETS.stocks,STOCK_HEADERS);
  const clean=rows.filter(r=>r&&String(r.symbol||'').trim()).map(r=>STOCK_HEADERS.map((_,i)=>[r.symbol,r.isin,r.qty,r.avg,r.invested,r.ltp,r.value,r.pnl][i]));
  if(!clean.length)throw new Error('No stock rows found. Check the header names and data.');
  if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,STOCK_HEADERS.length).clearContent();
  sh.getRange(2,1,clean.length,STOCK_HEADERS.length).setValues(clean);
  return {ok:true,count:clean.length,stocks:readImportedStocks_()};
}
function importMutualFunds_(rows){
  if(!Array.isArray(rows))throw new Error('Mutual fund rows must be an array.');
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.mutualFunds)||SpreadsheetApp.getActive().insertSheet(SHEETS.mutualFunds);
  ensureSheet_(SpreadsheetApp.getActive(),SHEETS.mutualFunds,MUTUAL_FUND_HEADERS);
  const required=['schemeName','amc','category','subCategory','folioNo','source','units','invested','value','returns','xirr'];
  const clean=rows.filter(r=>r&&String(r.schemeName||'').trim()).map(r=>required.map(key=>key==='schemeName'||key==='amc'||key==='category'||key==='subCategory'||key==='folioNo'||key==='source'?String(r[key]||'').trim():Number(r[key])||0));
  if(!clean.length)throw new Error('No mutual fund rows found. Check the header names and data.');
  if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,MUTUAL_FUND_HEADERS.length).clearContent();
  sh.getRange(2,1,clean.length,MUTUAL_FUND_HEADERS.length).setValues(clean);
  return {ok:true,count:clean.length,mutualFunds:readMutualFunds_()};
}
function upsertQuote_(input){
  const symbol=String(input.symbol||'').trim().toUpperCase(), price=Number(input.price); if(!symbol||!(price>0))throw new Error('Symbol and positive price are required.');
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.quotes),vals=sh.getDataRange().getValues();
  for(let r=1;r<vals.length;r++)if(String(vals[r][0]).toUpperCase()===symbol){sh.getRange(r+1,1,1,4).setValues([[symbol,price,Number(input.previousClose||vals[r][2]||0),new Date()]]);return {ok:true};}
  sh.appendRow([symbol,price,Number(input.previousClose||0),new Date()]); return {ok:true};
}
function saveSnapshot_(input){const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.snapshots);sh.appendRow([input.date||new Date(),Number(input.value||0),Number(input.invested||0),Number(input.pnl||0)]);return {ok:true};}

/** Refresh all currently held STOCK symbols using GoogleFinance formulas.
 * Google documents the price quote as real-time but delayed by up to 20 minutes.
 */
function refreshLiveQuotes_(){
  const status=getMarketStatus_(), txs=readTransactions_(), activeLots=calculateAllLots_();
  const stockSymbols=new Set(txs.filter(t=>t.assetType==='STOCK').map(t=>t.symbol));
  const symbols=[...new Set(activeLots.filter(l=>l.remainingQty>1e-9 && stockSymbols.has(l.symbol)).map(l=>l.symbol))];
  if(!symbols.length)return {ok:true,quotes:readQuotes_(),market:status,updated:0,message:'No stock holdings to refresh.'};
  const ss=SpreadsheetApp.getActive(), tempName='_LedgerlyLiveQuotes';
  let temp=ss.getSheetByName(tempName); if(!temp)temp=ss.insertSheet(tempName); temp.clear();
  temp.getRange(1,1,1,4).setValues([['Symbol','LTP','Previous Close','Trade Time']]);
  temp.getRange(2,1,symbols.length,1).setValues(symbols.map(s=>[s]));
  for(let i=0;i<symbols.length;i++){
    const r=i+2;
    temp.getRange(r,2).setFormula(`=IFERROR(GOOGLEFINANCE("NSE:${symbols[i]}","price"),"")`);
    temp.getRange(r,3).setFormula(`=IFERROR(GOOGLEFINANCE("NSE:${symbols[i]}","closeyest"),"")`);
    temp.getRange(r,4).setFormula(`=IFERROR(GOOGLEFINANCE("NSE:${symbols[i]}","tradetime"),"")`);
  }
  SpreadsheetApp.flush(); Utilities.sleep(1500);
  const rows=temp.getRange(2,1,symbols.length,4).getValues(), quoteSheet=ss.getSheetByName(SHEETS.quotes);
  const existing=quoteSheet.getDataRange().getValues(); let updated=0;
  rows.forEach(r=>{
    const symbol=String(r[0]).toUpperCase(), price=Number(r[1]), prev=Number(r[2]);
    if(!(price>0))return;
    let found=false;
    for(let i=1;i<existing.length;i++)if(String(existing[i][0]).toUpperCase()===symbol){quoteSheet.getRange(i+1,1,1,4).setValues([[symbol,price,prev||Number(existing[i][2]||0),new Date()]]);found=true;break;}
    if(!found)quoteSheet.appendRow([symbol,price,prev||0,new Date()]); updated++;
  });
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
function getHoliday_(date){const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.holidays);if(!sh||sh.getLastRow()<2)return '';const vals=sh.getRange(2,1,sh.getLastRow()-1,2).getValues();for(const r of vals){if(formatDate_(r[0])===date)return String(r[1]||'NSE holiday');}return '';}
function seed2026Holidays_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.holidays); if(sh.getLastRow()>1)return;
  const rows=[
    ['2026-01-15','Municipal Corporation Election in Maharashtra'],['2026-01-26','Republic Day'],['2026-02-19','Chhatrapati Shivaji Maharaj Jayanti'],['2026-03-03','Holi'],['2026-03-19','Gudhi Padwa'],['2026-03-26','Ram Navami'],['2026-03-31','Mahavir Jayanti'],['2026-04-01','Annual Bank Closing'],['2026-04-03','Good Friday'],['2026-04-14','Dr. Babasaheb Ambedkar Jayanti'],['2026-05-01','Maharashtra Din / Buddha Pournima'],['2026-05-28','Bakri Id'],['2026-06-26','Muharram'],['2026-08-26','Id-E-Milad'],['2026-09-14','Ganesh Chaturthi'],['2026-10-02','Mahatma Gandhi Jayanti'],['2026-10-20','Dussehra'],['2026-11-10','Diwali (Bali Pratipada)'],['2026-11-24','Guru Nanak Jayanti'],['2026-12-25','Christmas']
  ]; sh.getRange(2,1,rows.length,2).setValues(rows.map(r=>[r[0],r[1]]));
}
function ensureSheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);return sh;}
function formatDate_(v){if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');return String(v||'');}
function formatDateTime_(v){if(!v)return '';if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm');return String(v);}