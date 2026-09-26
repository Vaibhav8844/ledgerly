import React, {useEffect, useMemo, useState} from 'react';
import {createRoot} from 'react-dom/client';
import {
  AreaChart, Area, BarChart, Bar, LineChart, Line, PieChart as RePieChart,
  Pie, Cell, ResponsiveContainer, Tooltip, XAxis, YAxis, CartesianGrid, Legend
} from 'recharts';
import {
  LayoutDashboard, ArrowLeftRight, Plus, Search, WalletCards, PieChart,
  Settings, Trash2, X, TrendingUp, TrendingDown, CircleDollarSign,
  Layers3, Menu, ChevronRight, RefreshCw, Wifi, Upload, Pencil,
  Landmark, ReceiptText, BarChart3, Banknote, Target, Activity,
  CircleAlert, CheckCircle2, CalendarRange, SlidersHorizontal
} from 'lucide-react';
import * as XLSX from 'xlsx';
import './styles.css';

const API=(import.meta.env.VITE_API_URL||'').replace(/\/$/,'');
console.log('LEDGERLY API URL:', API);
const money=(n)=>new Intl.NumberFormat('en-IN',{style:'currency',currency:'INR',maximumFractionDigits:0}).format(Number(n||0));
const moneyCompact=(n)=>{
  const v=Number(n||0);
  if(Math.abs(v)>=10000000)return `₹${(v/10000000).toFixed(1)}Cr`;
  if(Math.abs(v)>=100000)return `₹${(v/100000).toFixed(1)}L`;
  if(Math.abs(v)>=1000)return `₹${(v/1000).toFixed(1)}k`;
  return money(v);
};
const pct=(n)=>`${Number(n||0)>=0?'+':''}${Number(n||0).toFixed(2)}%`;
const parseNumber=(value)=>Number(String(value??'').replace(/[₹,%]/g,'').replace(/,/g,'').trim())||0;
const safeArray=(v)=>Array.isArray(v)?v:[];
const safeFinance=(v)=>v&&typeof v==='object'?v:{};
const accountName=(account)=>String(account?.account||account?.name||account?.accountName||account?.bankName||account?.displayName||'Unnamed account').trim();
const amountFrom=(item,keys)=>{for(const key of keys){if(item?.[key]!==undefined&&item?.[key]!==null&&item?.[key]!=='')return Number(item[key])||0;}return 0;};
function normalizeMonthlyRows(finance){
  const source=safeArray(finance.monthly||finance.monthlyBreakdown||finance.monthlyCashFlow||finance.cashFlow||finance.monthlyTrend);
  if(source.length)return source.map(row=>({month:String(row.month||row.period||row.monthName||row.label||'—'),income:amountFrom(row,['income','totalIncome','incomeAmount','credits']),spending:amountFrom(row,['spending','expenses','expense','totalExpenses','expenseAmount','debits']),investments:amountFrom(row,['investments','investment','totalInvestments','investmentAmount'])}));
  const transactions=safeArray(finance.financeTransactions||finance.transactions||finance.recentTransactions), grouped={};
  transactions.forEach(transaction=>{
    const match=String(transaction.date||transaction.transactionDate||'').match(/^(\d{4})[-/]?(\d{2})/); if(!match)return;
    const month=`${match[1]}-${match[2]}`, type=String(transaction.type||transaction.transactionType||'').toLowerCase(), amount=Math.abs(amountFrom(transaction,['amount','total','value']));
    const row=grouped[month]||{month,income:0,spending:0,investments:0};
    if(type==='income')row.income+=amount; else if(type==='investment')row.investments+=amount; else row.spending+=amount;
    grouped[month]=row;
  });
  return Object.values(grouped).sort((a,b)=>b.month.localeCompare(a.month));
}
function cachedPortfolio_(period){
  try{
    const raw=localStorage.getItem('ledgerly:lastPortfolio');
    if(!raw)return null;
    const parsed=JSON.parse(raw);
    if(parsed?.period && parsed.period!==period)return null;
    return parsed?.data||null;
  }catch(_){return null}
}
function cachePortfolio_(period,value){
  try{localStorage.setItem('ledgerly:lastPortfolio',JSON.stringify({period,data:value,savedAt:Date.now()}));}catch(_){}
}
const periodLabels={all:'All time','this-month':'This month','last-month':'Last month','last-3-months':'Last 3 months','last-6-months':'Last 6 months','this-year':'This year'};

function parseDelimited(text){
  const delimiter=text.split(/\r?\n/)[0].includes('\t')?'\t':',';
  const rows=[]; let row=[],cell='',quoted=false;
  for(let i=0;i<text.length;i++){
    const ch=text[i],next=text[i+1];
    if(ch==='"'&&quoted&&next==='"'){cell+='"';i++;continue}
    if(ch==='"'){quoted=!quoted;continue}
    if(ch===delimiter&&!quoted){row.push(cell);cell='';continue}
    if((ch==='\n'||ch==='\r')&&!quoted){
      if(ch==='\r'&&next==='\n')i++;
      row.push(cell);if(row.some(v=>v.trim()))rows.push(row);row=[];cell='';continue
    }
    cell+=ch;
  }
  if(cell||row.length){row.push(cell);if(row.some(v=>v.trim()))rows.push(row)}
  return rows;
}
function normalizeHeader(value){
  return String(value||'').toLowerCase().replace(/[\s._\-\/&()]/g,'');
}
function firstIndex(headers, aliases){
  for(const alias of aliases){const i=headers.indexOf(normalizeHeader(alias));if(i>=0)return i;}
  return undefined;
}
function parseMutualFundRows(rows){
  if(rows.length<2)throw new Error('The file must contain a header row and at least one mutual fund.');
  const headers=rows[0].map(normalizeHeader);
  const indexes={
    schemeName:firstIndex(headers,['Scheme Name','Mutual Fund','Fund Name','Scheme']),
    amc:firstIndex(headers,['AMC','Fund House','Asset Management Company']),
    category:firstIndex(headers,['Category','Fund Category']),
    subCategory:firstIndex(headers,['Sub-category','Sub Category','Subcategory']),
    folioNo:firstIndex(headers,['Folio No.','Folio Number','Folio']),
    source:firstIndex(headers,['Source','Platform']),
    units:firstIndex(headers,['Units','Quantity']),
    invested:firstIndex(headers,['Invested Value','Investment Value','Invested Amount','Cost Value']),
    value:firstIndex(headers,['Current Value','Current Market Value','Market Value','Present Value','Value']),
    returns:firstIndex(headers,['Returns','Return','Profit/Loss','P&L','Unrealised P&L','Unrealized P&L']),
    xirr:firstIndex(headers,['XIRR','XIRR %'])
  };
  if(['schemeName','units','invested','value'].some(key=>indexes[key]===undefined))
    throw new Error('Required mutual-fund headers were not found. Expected Scheme Name, Units, Invested Value and Current Value/Market Value.');
  return rows.slice(1).map(row=>Object.fromEntries(Object.entries(indexes).map(([key,index])=>[
    key,['units','invested','value','returns','xirr'].includes(key)?parseNumber(row[index]):String(row[index]??'').trim()
  ]))).filter(row=>row.schemeName);
}
function parseStockRows(rows){
  if(rows.length<2)throw new Error('The file must contain a header row and at least one stock.');
  const headers=rows[0].map(normalizeHeader);
  const indexes={
    symbol:firstIndex(headers,['Stock Name','Stock','Company','Symbol','Scrip','Scrip Name']),
    isin:firstIndex(headers,['ISIN']),
    qty:firstIndex(headers,['Quantity','Qty','Shares','Units']),
    avg:firstIndex(headers,['Average buy price','Average Buy Price','Avg Buy Price','Average Price','Avg Price']),
    invested:firstIndex(headers,['Buy value','Buy Value','Invested Value','Investment Value','Cost Value','Buy Amount']),
    ltp:firstIndex(headers,['Closing price','Closing Price','Current Price','LTP','CMP','Market Price']),
    value:firstIndex(headers,['Closing value','Closing Value','Current Value','Current Market Value','Market Value','Market Value Amount']),
    pnl:firstIndex(headers,['Unrealised P&L','Unrealized P&L','Unrealised P/L','Unrealized P/L','P&L','Profit/Loss'])
  };
  if(['symbol','qty','invested','value'].some(key=>indexes[key]===undefined))
    throw new Error('Required stock headers were not found. Expected Stock Name/Symbol, Quantity, Buy Value/Invested Value and Closing/Current Value.');
  return rows.slice(1).map(row=>Object.fromEntries(Object.entries(indexes).map(([key,index])=>[
    key,['qty','avg','invested','ltp','value','pnl'].includes(key)?parseNumber(row[index]):String(row[index]??'').trim()
  ]))).filter(row=>row.symbol);
}
function detectImportKind(rows){
  if(!rows.length)return null;
  const headers=rows[0].map(normalizeHeader);
  const score=(groups)=>groups.reduce((n,aliases)=>n+(aliases.some(a=>headers.includes(normalizeHeader(a)))?1:0),0);
  const stockScore=score([
    ['Stock Name','Stock','Company','Symbol','Scrip','Scrip Name'],['Quantity','Qty','Shares'],
    ['Buy value','Invested Value','Investment Value','Cost Value'],['Closing value','Current Value','Market Value']
  ]);
  const mfScore=score([
    ['Scheme Name','Mutual Fund','Fund Name','Scheme'],['Units','Quantity'],
    ['Invested Value','Investment Value','Invested Amount'],['Current Value','Market Value','Present Value']
  ]);
  return stockScore>=3?'stocks':mfScore>=3?'mutualFunds':null;
}
function parseImportRows(rows){
  const kind=detectImportKind(rows);
  if(kind==='stocks')return {kind,rows:parseStockRows(rows)};
  if(kind==='mutualFunds')return {kind,rows:parseMutualFundRows(rows)};
  throw new Error('Could not identify this sheet as a stock or mutual-fund holdings sheet.');
}
async function parseMutualFundUpload(file){
  if(file.name.toLowerCase().endsWith('.xlsx')){
    const workbook=XLSX.read(await file.arrayBuffer(),{type:'array'});
    const combined={stocks:[],mutualFunds:[]};
    workbook.SheetNames.forEach(sheetName=>{
      const rows=XLSX.utils.sheet_to_json(workbook.Sheets[sheetName],{header:1,defval:''});
      if(!rows.length)return;
      try{const parsed=parseImportRows(rows);combined[parsed.kind].push(...parsed.rows);}catch(_){/* ignore unrelated sheets */}
    });
    if(!combined.stocks.length&&!combined.mutualFunds.length)throw new Error('No recognizable stock or mutual-fund holdings sheet was found in the workbook.');
    return combined;
  }
  return parseImportRows(parseDelimited(await file.text()));
}

function App(){
  const [tab,setTab]=useState(()=>sessionStorage.getItem('ledgerly.activeTab')||'dashboard');
  const [period,setPeriod]=useState('this-month');
  const cached=cachedPortfolio_(period);
  const [data,setData]=useState(cached||{
    transactions:[],quotes:{},snapshots:[],holdings:[],
    totals:{invested:0,value:0,pnl:0,returnPct:0,realized:0,todayPnl:0},
    market:{open:false,label:'Market closed'},finance:null
  });
  const [modal,setModal]=useState(false);
  const [editHolding,setEditHolding]=useState(null);
  const [editAccount,setEditAccount]=useState(null);
  const [loading,setLoading]=useState(!cached);
  const [refreshing,setRefreshing]=useState(false);
  const [mobileOpen,setMobileOpen]=useState(false);
  const [error,setError]=useState('');

  const hasUsableData=()=>!!(data&&(data.holdings?.length||data.finance||data.bankAccounts?.length));

  const load=async()=>{
    const hasData=hasUsableData();
    if(!hasData)setLoading(true);
    setError('');
    try{
      const r=await fetch(`${API}?action=portfolio&period=${encodeURIComponent(period)}`);
      if(!r.ok)throw new Error(r.status===404?'Ledgerly Apps Script deployment was not found. Deploy the current Code.gs as a web app and update client/.env with its /exec URL.':`Ledgerly API returned ${r.status}`);
      const next=await r.json();
      if(next.success===false||next.ok===false)throw new Error(next.error||next.message||'Could not load portfolio.');
      setData(next);
      cachePortfolio_(period,next);
    }catch(e){
      console.error(e);
      setError(e.message||'Could not load Ledgerly.');
    }finally{
      if(!hasData)setLoading(false);
    }
  };

  const refreshQuotes=async()=>{
    setRefreshing(true);
    try{
      const r=await fetch(`${API}?action=refreshQuotes`);
      const next=await r.json();
      if(!r.ok||next.success===false||next.ok===false)throw new Error(r.status===404?'Ledgerly Apps Script deployment was not found. Deploy the current Code.gs as a web app and update client/.env with its /exec URL.':next.error||'Could not refresh prices.');
      setData(d=>({...d,
        quotes:next.quotes||d.quotes,
        market:next.market||d.market,
        holdings:next.holdings||d.holdings,
        totals:next.totals||d.totals,
        importedHoldings:next.importedHoldings||d.importedHoldings,
        lastQuoteRefresh:Date.now()
      }));
    }catch(e){
      setError(e.message||'Could not refresh prices.');
    }finally{
      setRefreshing(false);
    }
  };

  useEffect(()=>{load()},[period]);
  useEffect(()=>{
    const timer=setInterval(()=>{
      if(!document.hidden)refreshQuotes();
    },60000);
    return()=>clearInterval(timer);
  },[]);
  const holdings=data.holdings||[];
  const totals=data.totals||{invested:0,value:0,pnl:0,returnPct:0,realized:0,todayPnl:0};
  const finance={...safeFinance(data.finance),bankAccounts:safeArray(data.finance?.bankAccounts??data.bankAccounts),bankTotal:data.finance?.bankTotal??data.bankTotal};
  const nav=[
    ['dashboard','Dashboard',LayoutDashboard],
    ['holdings','Investments',WalletCards],
    ['transactions','Investments',ArrowLeftRight],
    ['finance','Finance Activity',ReceiptText],
    ['analytics','Analytics',PieChart],
    ['accounts','Accounts',Landmark],
    ['settings','Settings',Settings]
  ];

  const importData=async e=>{
    const file=e.target.files?.[0];if(!file)return;
    try{
      const imported=await parseMutualFundUpload(file);
      let stockCount=0,mfCount=0;
      if(imported.stocks?.length){
        const r=await fetch(API,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'importData',kind:'stocks',rows:imported.stocks})});
        const next=await r.json();
        if(!r.ok||next.success===false||next.ok===false)throw new Error(next.error||'Stock import failed');
        stockCount=Number(next.count||imported.stocks.length);
      }
      if(imported.mutualFunds?.length){
        const r=await fetch(API,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'importData',kind:'mutualFunds',rows:imported.mutualFunds})});
        const next=await r.json();
        if(!r.ok||next.success===false||next.ok===false)throw new Error(next.error||'Mutual-fund import failed');
        mfCount=Number(next.count||imported.mutualFunds.length);
      }
      await load();
      alert(`Imported ${stockCount} stock rows and ${mfCount} mutual-fund rows.`);
    }catch(error){alert(error.message||'Could not import data.')}
    e.target.value='';
  };

  return <div className="app">
    <aside className={`sidebar ${mobileOpen?'open':''}`}>
      <div className="brand"><div className="brandmark">L</div><div><b>Ledgerly</b><span>Wealth dashboard</span></div></div>
      {nav.map(([id,label,Icon])=><button key={id} className={`nav ${tab===id?'active':''}`} onClick={()=>{sessionStorage.setItem('ledgerly.activeTab',id);setTab(id);setMobileOpen(false)}}><Icon size={19}/>{label}</button>)}
      <div className="side-bottom"><div className="secure"><CircleDollarSign size={18}/><div><b>Private by design</b><span>Ledgerly + Finance Assistant</span></div></div></div>
    </aside>
    {mobileOpen&&<div className="scrim" onClick={()=>setMobileOpen(false)}/>}
    <main className="main">
      <header>
        <div className="header-title"><button className="mobile-menu" onClick={()=>setMobileOpen(true)}><Menu/></button><div><span className="eyebrow">PERSONAL WEALTH</span><h1>{tab==='dashboard'?'Financial overview':tab[0].toUpperCase()+tab.slice(1)}</h1></div></div>
        <div className="header-actions">
          <PeriodSelect value={period} onChange={setPeriod}/>
          <MarketBadge market={data.market}/>
          <button className="icon-btn" title="Refresh prices" onClick={refreshQuotes} disabled={refreshing}><RefreshCw size={18} className={refreshing?'spin':''}/></button>
          <label className="secondary import-button" title="Import stock or mutual-fund Excel, CSV or TSV"><Upload size={17}/><span>Import</span><input type="file" accept=".xlsx,.csv,.tsv" onChange={importData}/></label>
          <button className="primary add-button" onClick={()=>setModal(true)}><Plus size={18}/><span>Add investment</span></button>
        </div>
      </header>
      {error&&<div className="global-error"><CircleAlert size={17}/><span>{error}</span><button onClick={load}>Retry</button></div>}
      {loading?<Loading/>:<>
        {tab==='dashboard'&&<Dashboard totals={totals} holdings={holdings} snapshots={data.snapshots||[]} market={data.market} quotes={data.quotes} finance={finance} period={period}/>}
        {tab==='holdings'&&<HoldingsPage holdings={holdings} totals={totals} onEdit={h=>setEditHolding(h)}/>}
        {tab==='transactions'&&<TransactionsPage investmentTxs={data.transactions||[]} financeTxs={safeArray(finance.financeTransactions||finance.transactions||finance.recentTransactions)} onRefresh={load}/>}
        {tab==='finance'&&<FinanceActivityPage transactions={safeArray(finance.financeTransactions||finance.transactions||finance.recentTransactions)} period={period}/>}
        {tab==='analytics'&&<AnalyticsPage holdings={holdings} totals={totals} finance={finance} snapshots={data.snapshots||[]}/>}
        {tab==='accounts'&&<AccountsPage finance={finance} onEdit={setEditAccount}/>}
        {tab==='settings'&&<SettingsPage finance={finance} onRefresh={load}/>}
      </>}
    </main>
    {modal&&<TransactionModal finance={finance} onClose={()=>setModal(false)} onSaved={async()=>{setModal(false);await load({silent:true});refreshQuotes()}}/>}
    {editHolding&&<HoldingEditModal holding={editHolding} onClose={()=>setEditHolding(null)} onSaved={()=>{setEditHolding(null);load({silent:true})}}/>}
    {editAccount&&<BankAccountEditModal account={editAccount} onClose={()=>setEditAccount(null)} onSaved={()=>{setEditAccount(null);load({silent:true})}}/>}
  </div>
}

function PeriodSelect({value,onChange}){
  return <label className="period-select"><CalendarRange size={15}/><select value={value} onChange={e=>onChange(e.target.value)}>{Object.entries(periodLabels).map(([k,v])=><option key={k} value={k}>{v}</option>)}</select></label>
}

function MarketBadge({market}){
  const open=!!market?.open;
  return <div className={`market-badge ${open?'open':'closed'}`} title={`${market?.reason||''}${market?.time?' • '+market.time+' IST':''}`}><span className="market-dot"/><span>{open?'Market open':'Market closed'}</span></div>
}

function Loading(){return <div className="content"><div className="loading-grid">{[1,2,3,4,5,6].map(i=><div className="skeleton" key={i}/>)}</div></div>}

function Dashboard({totals,holdings,snapshots,market,quotes,finance,period}){
  const hasFinance=finance&&Object.keys(finance).length>0;
  const bankTotal=Number(finance.bankTotal??finance.bank?.total??0);
  const income=Number(finance.income??finance.period?.income??0);
  const spending=Number(finance.spending??finance.period?.spending??0);
  const investments=Number(finance.investments??finance.period?.investments??0);
  const savings=income-spending;
  const savingsRate=income?savings/income*100:0;
  const netWorth=Number(finance.netWorth?.total??(bankTotal+Number(totals.value||0)+Number(finance.otherAssets||0)-Number(finance.liabilities||0)));
  const monthly=safeArray(finance.monthly||finance.cashFlow); const monthlyTrend=safeArray(finance.monthlyTrend||monthly);
  const categories=safeArray(finance.categorySpending||finance.categories);
  const allocation=[
    {name:'Stocks',value:holdings.filter(h=>h.assetType==='STOCK'||h.importedStock).reduce((s,h)=>s+Number(h.value||0),0)},
    {name:'Mutual funds',value:holdings.filter(h=>h.assetType==='MF'||h.mutualFund).reduce((s,h)=>s+Number(h.value||0),0)}
  ].filter(x=>x.value>0);

  return <div className="content">
    <section className="wealth-hero">
      <div><span className="muted-light">Total net worth</span><div className="hero-value">{hasFinance?money(netWorth):'Finance sync required'}</div><div className="hero-sub"><span>{periodLabels[period]||'Selected period'}</span><span>•</span><span>Bank {money(bankTotal)}</span><span>•</span><span>Investments {money(totals.value)}</span></div></div>
      <div className="hero-stat"><span>Portfolio value</span><b>{money(totals.value)}</b><small className={totals.pnl>=0?'gain-light':'loss-light'}>{money(totals.pnl)} · {pct(totals.returnPct)}</small></div>
    </section>

    <div className="metric-grid five">
      <Metric title="Bank balance" value={hasFinance?money(bankTotal):'—'} note="Finance Assistant" icon={<Landmark size={17}/>} />
      <Metric title="Income" value={hasFinance?money(income):'—'} note={periodLabels[period]} icon={<TrendingUp size={17}/>} positive/>
      <Metric title="Spending" value={hasFinance?money(spending):'—'} note={periodLabels[period]} icon={<ReceiptText size={17}/>} />
      <Metric title="Invested" value={hasFinance?money(investments):money(totals.invested)} note="Cash invested" icon={<Target size={17}/>} />
      <Metric title="Savings rate" value={hasFinance?`${savingsRate.toFixed(1)}%`:'—'} note="(Income − spending) / income" icon={<Banknote size={17}/>} positive={savingsRate>=0}/>
    </div>

    <div className="metric-grid five pnl-strip">
      <Metric title="Unrealized P&L" value={money(totals.pnl)} note="Current holdings" icon={totals.pnl>=0?<TrendingUp size={17}/>:<TrendingDown size={17}/>} positive={totals.pnl>=0}/>
      <Metric title="Realized P&L" value={money(totals.realized)} note="Closed FIFO lots" icon={totals.realized>=0?<TrendingUp size={17}/>:<TrendingDown size={17}/>} positive={totals.realized>=0}/>
      <Metric title="Today P&L" value={money(totals.todayPnl)} note="Quoted holdings" icon={totals.todayPnl>=0?<TrendingUp size={17}/>:<TrendingDown size={17}/>} positive={totals.todayPnl>=0}/>
      <Metric title="Invested capital" value={money(totals.invested)} note="Open holdings" icon={<WalletCards size={17}/>} />
      <Metric title="Holdings" value={String(holdings.length)} note="Stocks + mutual funds" icon={<Layers3 size={17}/>} />
    </div>

    {hasFinance&&safeArray(finance.bankAccounts).length>0&&<div className="dashboard-accounts">{safeArray(finance.bankAccounts).map((a,i)=><div className="dashboard-account" key={a.id||accountName(a)||i}><div className="account-icon"><Landmark size={17}/></div><div><span>{a.type||'Bank account'}</span><b>{accountName(a)}</b></div><strong>{money(a.balance)}</strong></div>)}</div>}

    {!hasFinance?<FinanceConnectionCard/>:<>
      <div className="chart-grid two">
        <ChartCard title="Income vs spending vs investments" subtitle={`Monthly comparison · ${periodLabels[period]||''}`}>
          {monthlyTrend.length?<ResponsiveContainer width="100%" height={310}><BarChart data={monthlyTrend}><CartesianGrid strokeDasharray="3 3" vertical={false}/><XAxis dataKey="month" tickLine={false}/><YAxis tickFormatter={moneyCompact} tickLine={false} axisLine={false}/><Tooltip formatter={v=>money(v)}/><Legend/><Bar dataKey="income" name="Income" stackId="cash" fill="#5b8def" radius={[4,4,0,0]}/><Bar dataKey="spending" name="Spending" stackId="out" fill="#e79a5a" radius={[4,4,0,0]}/><Bar dataKey="investments" name="Investments" stackId="out" fill="#7b68d9" radius={[4,4,0,0]}/></BarChart></ResponsiveContainer>:<ChartEmpty text="Add or sync transactions to see monthly cash flow."/>}
        </ChartCard>
        <ChartCard title="Spending trend" subtitle="Monthly expenses">
          {monthlyTrend.length?<ResponsiveContainer width="100%" height={310}><LineChart data={monthlyTrend}><CartesianGrid strokeDasharray="3 3" vertical={false}/><XAxis dataKey="month" tickLine={false}/><YAxis tickFormatter={moneyCompact} tickLine={false} axisLine={false}/><Tooltip formatter={v=>money(v)}/><Line type="monotone" dataKey="spending" name="Spending" stroke="#d86b61" strokeWidth={3} dot={{r:3}}/></LineChart></ResponsiveContainer>:<ChartEmpty text="No monthly spending history yet."/>}
        </ChartCard>
      </div>

      <div className="chart-grid two">
        <ChartCard title="Spending by category" subtitle="Where your money went">
          {categories.length?<CategoryPie data={categories}/>:<ChartEmpty text="No expense categories for this period."/>}
        </ChartCard>
        <ChartCard title="Investment allocation" subtitle="Current portfolio value">
          {allocation.length?<ResponsiveContainer width="100%" height={310}><RePieChart><Pie data={allocation} dataKey="value" nameKey="name" innerRadius={75} outerRadius={110} paddingAngle={3}>{allocation.map((x,i)=><Cell key={x.name} fill={['#6658dc','#5b8def'][i%2]}/>)}</Pie><Tooltip formatter={v=>money(v)}/><Legend/></RePieChart></ResponsiveContainer>:<ChartEmpty text="Add stock or mutual-fund holdings to see allocation."/>}
        </ChartCard>
      </div>

      <div className="chart-grid two">
        <ChartCard title="Bank balances" subtitle="Current balances from Finance Assistant">
          {safeArray(finance.bankAccounts).length?<ResponsiveContainer width="100%" height={300}><BarChart data={safeArray(finance.bankAccounts).map(account=>({...account,name:accountName(account)}))} layout="vertical" margin={{left:20,right:20}}><CartesianGrid strokeDasharray="3 3" horizontal={false}/><XAxis type="number" tickFormatter={moneyCompact}/><YAxis type="category" dataKey="name" width={105}/><Tooltip formatter={v=>money(v)}/><Bar dataKey="balance" name="Balance" fill="#4e9b83" radius={[0,6,6,0]}/></BarChart></ResponsiveContainer>:<ChartEmpty text="No bank accounts were returned by Finance Assistant."/>}
        </ChartCard>
        <ChartCard title="Top merchants" subtitle="Highest spending merchants">
          {safeArray(finance.merchantSpending).length?<ResponsiveContainer width="100%" height={300}><BarChart data={safeArray(finance.merchantSpending).slice(0,8)} layout="vertical" margin={{left:15,right:20}}><CartesianGrid strokeDasharray="3 3" horizontal={false}/><XAxis type="number" tickFormatter={moneyCompact}/><YAxis type="category" dataKey="merchant" width={105}/><Tooltip formatter={v=>money(v)}/><Bar dataKey="amount" name="Spending" fill="#d88b60" radius={[0,6,6,0]}/></BarChart></ResponsiveContainer>:<ChartEmpty text="No merchant data for this period."/>}
        </ChartCard>
      </div>

      <div className="section-head"><div><h2>Recent financial activity</h2><p>Latest income, spending and investment cash-flow entries.</p></div><span className="pill">{periodLabels[period]}</span></div>
      <RecentFinance transactions={safeArray(finance.recentTransactions||finance.financeTransactions||finance.transactions).slice(0,8)}/>
    </>}

    <div className="section-head"><div><h2>Portfolio performance</h2><p>Value history from saved daily snapshots.</p></div><span className="pill">Live valuation</span></div>
    <div className="chart-card">{snapshots.length?<ResponsiveContainer width="100%" height={300}><AreaChart data={snapshots}><defs><linearGradient id="portfolioFill" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#6d5ce7" stopOpacity={.24}/><stop offset="95%" stopColor="#6d5ce7" stopOpacity={0}/></linearGradient></defs><XAxis dataKey="date" tickLine={false} axisLine={false}/><YAxis tickFormatter={moneyCompact} tickLine={false} axisLine={false}/><Tooltip formatter={v=>money(v)}/><Area type="monotone" dataKey="value" stroke="#6d5ce7" fill="url(#portfolioFill)" strokeWidth={3}/></AreaChart></ResponsiveContainer>:<ChartEmpty text="Your investment performance chart starts after snapshots are recorded."/>}</div>
    <div className="feed-note"><Wifi size={15}/><span><b>Price feed:</b> GoogleFinance / NSE. Quotes can be delayed by up to 20 minutes; this dashboard is for portfolio tracking, not order execution.</span></div>
  </div>
}

function CategoryPie({data}){
  return <ResponsiveContainer width="100%" height={310}><RePieChart><Pie data={data.slice(0,10)} dataKey="amount" nameKey="category" innerRadius={72} outerRadius={108} paddingAngle={2}>{data.slice(0,10).map((x,i)=><Cell key={`${x.category}-${i}`} fill={['#6658dc','#5b8def','#4e9b83','#d88b60','#d86b61','#7f8aa6','#8a73c6','#6aa6a0','#c28a55','#9a7cbe'][i%10]}/>)}</Pie><Tooltip formatter={v=>money(v)}/><Legend/></RePieChart></ResponsiveContainer>
}

function ChartCard({title,subtitle,children}){return <div className="chart-card"><div className="chart-head"><div><h2>{title}</h2><p>{subtitle}</p></div><BarChart3 size={18}/></div>{children}</div>}
function ChartEmpty({text}){return <div className="chart-empty"><Activity size={28}/><span>{text}</span></div>}
function FinanceConnectionCard(){return <div className="connection-card"><CircleAlert size={22}/><div><b>Finance Assistant data is not available yet</b><p>Ledgerly is connected to its investment backend, but the portfolio response does not currently contain the Finance Assistant analytics payload. Once the backend bridge returns it, bank balances, spending and cash-flow charts will populate automatically.</p></div></div>}

function RecentFinance({transactions}){
  if(!transactions.length)return <div className="table-card"><ChartEmpty text="No financial transactions in this period."/></div>;
  return <div className="table-card responsive-table"><table><thead><tr><th>Date</th><th>Merchant / source</th><th>Category</th><th>Account</th><th>Type</th><th className="right">Amount</th></tr></thead><tbody>{transactions.map((t,i)=>{
    const type=String(t.type||t.transactionType||'').toLowerCase();
    const positive=['income','sell'].includes(type);
    return <tr key={t.id||i}><td>{t.date||'—'}</td><td><b>{t.merchant||t.sourceName||t.note||'—'}</b></td><td>{t.category||'—'}</td><td>{t.account||'—'}</td><td><span className={`type ${positive?'income':'expense'}`}>{t.type||'—'}</span></td><td className={`right amount ${positive?'gain':'loss'}`}>{positive?'+':'-'}{money(Math.abs(Number(t.amount??t.total??0)))}</td></tr>
  })}</tbody></table></div>
}

function Metric({title,value,note,positive,muted,icon}){return <div className="metric"><div className="metric-top"><span>{icon}{title}</span></div><b className={positive===true?'gain':positive===false?'loss':''}>{value}</b><small className={muted?'muted':''}>{note}</small></div>}

function HoldingsPage({holdings,totals,onEdit}){
  return <div className="content"><div className="page-intro"><div><span className="eyebrow">INVESTMENTS</span><h2>Stocks & mutual funds</h2><p>Live holdings, FIFO lots and imported positions.</p></div><div className="intro-stats"><span>Invested <b>{money(totals.invested)}</b></span><span>Value <b>{money(totals.value)}</b></span><span>P&L <b className={totals.pnl>=0?'gain':'loss'}>{money(totals.pnl)}</b></span></div></div><Holdings holdings={holdings} onEdit={onEdit}/></div>
}

function Holdings({holdings,onEdit,compact}){
  return <div className={`holdings-card ${compact?'compact':''}`}>{holdings.length===0?<Empty title="No holdings yet" text="Add your first BUY transaction or import a holdings file." icon={<Layers3 size={25}/>}/>:<div className="holdings-list">{holdings.map(h=><div className="holding-row" key={`${h.symbol}-${h.mutualFund?.folioNo||h.importedStock?.isin||h.assetType}`}><div className="asset"><div className="asset-icon">{String(h.symbol||'?').slice(0,1)}</div><div><b>{h.displayName||h.symbol}</b><span>{h.qty} units · avg {money(h.avg)} · LTP {money(h.ltp)}</span></div></div><div className="holding-price"><b>{money(h.value)}</b><span className={h.pnl>=0?'gain':'loss'}>{money(h.pnl)} · {pct(h.returnPct)}</span>{h.previousClose>0&&<small className={h.dayPnl>=0?'gain':'loss'}>Today {money(h.dayPnl||0)}</small>}</div>{onEdit&&(h.importedStock||h.mutualFund)&&<button className="icon-btn holding-edit" title="Edit holding" onClick={()=>onEdit(h)}><Pencil size={15}/></button>}<div className="mini-bar"><span style={{width:`${Math.min(100,Math.max(2,Math.abs(h.returnPct)))}%`}}/></div></div>)}</div>}</div>
}

function TransactionsPage({investmentTxs,financeTxs,onRefresh}){
  const [mode,setMode]=useState('investments');
  const [q,setQ]=useState('');
  const [type,setType]=useState('all');
  const list=(mode==='investments'?investmentTxs:financeTxs).filter(t=>{
    const text=`${t.symbol||''} ${t.note||''} ${t.merchant||''} ${t.category||''} ${t.account||''}`.toLowerCase();
    const typ=String(t.type||t.transactionType||'').toLowerCase();
    return text.includes(q.toLowerCase())&&(type==='all'||typ===type);
  });
  return <div className="content"><div className="page-intro"><div><span className="eyebrow">LEDGERS</span><h2>Transactions</h2><p>Investment trades and Finance Assistant cash-flow transactions stay separate.</p></div></div><div className="seg page-seg"><button className={mode==='investments'?'on':''} onClick={()=>{setMode('investments');setType('all')}}>Investments</button><button className={mode==='finance'?'on':''} onClick={()=>{setMode('finance');setType('all')}}>Income & expenses</button></div><div className="toolbar"><div className="search"><Search size={17}/><input placeholder={mode==='investments'?'Search symbol or note':'Search merchant, category or account'} value={q} onChange={e=>setQ(e.target.value)}/></div><select className="filter-select" value={type} onChange={e=>setType(e.target.value)}><option value="all">All types</option>{mode==='investments'?<><option value="buy">Buy</option><option value="sell">Sell</option></>:<><option value="income">Income</option><option value="expense">Expense</option><option value="investment">Investment</option><option value="credit card payment">Card payment</option></>}</select><span className="muted">{list.length} transactions</span></div>{mode==='investments'?<div className="table-card responsive-table">{list.length===0?<Empty title="No investment transactions" text="Add a BUY or SELL to begin." icon={<ArrowLeftRight size={25}/>}/>:<table><thead><tr><th>Date</th><th>Symbol</th><th>Type</th><th>Qty</th><th>Price</th><th>Charges</th><th>Total</th><th></th></tr></thead><tbody>{list.map(t=><tr key={t.id}><td>{t.date}</td><td><b>{t.symbol}</b></td><td><span className={`type ${String(t.type).toLowerCase()}`}>{t.type}</span></td><td>{t.quantity}</td><td>{money(t.price)}</td><td>{money(t.charges)}</td><td>{money(t.total)}</td><td><button className="danger-icon" onClick={async()=>{if(!confirm('Delete this investment transaction?'))return;await fetch(`${API}?action=deleteTransaction&id=${encodeURIComponent(t.id)}`);onRefresh()}}><Trash2 size={16}/></button></td></tr>)}</tbody></table>}</div>:<RecentFinance transactions={list}/>}</div>
}

function FinanceActivityPage({transactions,period}){
  const [q,setQ]=useState('');
  const [type,setType]=useState('all');
  const [account,setAccount]=useState('all');
  const accounts=[...new Set(transactions.map(t=>String(t.account||'').trim()).filter(Boolean))].sort();
  const filtered=transactions.slice().sort((a,b)=>{
    const first=Date.parse(String(a.date||''));
    const second=Date.parse(String(b.date||''));
    if(Number.isNaN(first)||Number.isNaN(second))return String(b.date||'').localeCompare(String(a.date||''));
    return second-first;
  }).filter(t=>{
    const text=`${t.merchant||''} ${t.category||''} ${t.subcategory||''} ${t.account||''} ${t.paymentMode||t.payment_mode||''} ${t.remarks||t.note||''}`.toLowerCase();
    const typ=String(t.type||t.transactionType||'').toLowerCase();
    return text.includes(q.toLowerCase()) &&
      (type==='all'||typ===type) &&
      (account==='all'||String(t.account||'')===account);
  });
  return <div className="content">
    <div className="page-intro"><div><span className="eyebrow">FINANCE ASSISTANT</span><h2>Finance activity</h2><p>Historical income, expenses, investments and card payments read directly from your month-wise Finance Assistant sheets.</p></div><div className="big-total"><span>Showing</span><b>{filtered.length}</b><small>{periodLabels[period]||'Selected period'}</small></div></div>
    <div className="toolbar finance-toolbar">
      <div className="search"><Search size={17}/><input placeholder="Search merchant, category, account or remarks" value={q} onChange={e=>setQ(e.target.value)}/></div>
      <select className="filter-select" value={type} onChange={e=>setType(e.target.value)}><option value="all">All types</option><option value="income">Income</option><option value="expense">Expense</option><option value="investment">Investment</option><option value="credit card payment">Card payment</option></select>
      <select className="filter-select" value={account} onChange={e=>setAccount(e.target.value)}><option value="all">All accounts</option>{accounts.map(a=><option key={a} value={a}>{a}</option>)}</select>
    </div>
    <div className="table-card responsive-table">
      {filtered.length?<table><thead><tr><th>Date</th><th>Merchant / source</th><th>Category</th><th>Subcategory</th><th>Account</th><th>Payment mode</th><th>Type</th><th>Remarks</th><th className="right">Amount</th></tr></thead><tbody>{filtered.map((t,i)=>{
        const typ=String(t.type||t.transactionType||'').toLowerCase();
        const positive=typ==='income';
        return <tr key={t.id||`${t.date}-${i}`}><td>{t.date||'—'}</td><td><b>{t.merchant||t.sourceName||'—'}</b></td><td>{t.category||'—'}</td><td>{t.subcategory||'—'}</td><td>{t.account||'—'}</td><td>{t.paymentMode||t.payment_mode||'—'}</td><td><span className={`type ${positive?'income':'expense'}`}>{t.type||'—'}</span></td><td className="remarks-cell">{t.remarks||t.note||'—'}</td><td className={`right amount ${positive?'gain':'loss'}`}>{positive?'+':'-'}{money(Math.abs(Number(t.amount??t.total??0)))}</td></tr>
      })}</tbody></table>:<ChartEmpty text="No Finance Assistant transactions match these filters."/>}
    </div>
    <div className="feed-note"><Wifi size={15}/><span><b>Source of truth:</b> These records are read from Finance Assistant. Ledgerly does not copy them into its investment ledger.</span></div>
  </div>
}

function AnalyticsPage({holdings,totals,finance,snapshots}){
  const monthly=normalizeMonthlyRows(finance); const monthlyTrend=monthly;
  const categories=safeArray(finance.categorySpending||finance.categories);
  const payments=safeArray(finance.paymentModeSpending);
  const merchants=safeArray(finance.merchantSpending);
  const income=Number(finance.income||0), spending=Number(finance.spending||0), investments=Number(finance.investments||0);
  const savings=income-spending;
  const savingsRate=income?savings/income*100:0;
  return <div className="content">
    <div className="page-intro"><div><span className="eyebrow">ANALYTICS</span><h2>Financial analytics</h2><p>Detailed cash-flow, spending and investment analysis.</p></div></div>
    <div className="analytics-summary"><Metric title="Income" value={money(income)} note="Selected period" positive icon={<TrendingUp size={17}/>}/><Metric title="Spending" value={money(spending)} note="Selected period" icon={<ReceiptText size={17}/>}/><Metric title="Investments" value={money(investments)} note="Selected period" icon={<Target size={17}/>}/><Metric title="Net savings" value={money(savings)} note={`${savingsRate.toFixed(1)}% savings rate`} positive={savings>=0} icon={<Banknote size={17}/>}/></div>
    <div className="chart-grid two">
      <ChartCard title="Monthly cash flow" subtitle="Income, spending and investments"><MonthlyCashFlow data={monthly}/></ChartCard>
      <ChartCard title="Category spending" subtitle="Selected period"><CategoryPie data={categories}/></ChartCard>
      <ChartCard title="Monthly spending" subtitle="Trend across available months"><MonthlyLine data={monthlyTrend} keyName="spending" label="Spending"/></ChartCard>
      <ChartCard title="Monthly income" subtitle="Trend across available months"><MonthlyLine data={monthlyTrend} keyName="income" label="Income"/></ChartCard>
      <ChartCard title="Payment mode spending" subtitle="How spending was paid"><GenericBar data={payments} labelKey="paymentMode" valueKey="amount" label="Spending"/></ChartCard>
      <ChartCard title="Top merchants" subtitle="Largest spending merchants"><GenericBar data={merchants.slice(0,10)} labelKey="merchant" valueKey="amount" label="Spending"/></ChartCard>
    </div>
    <div className="analytics-grid">
      <div className="dark-panel"><span>Portfolio P&L</span><strong>{money(totals.pnl)}</strong><div className={totals.pnl>=0?'gain-light':'loss-light'}>{pct(totals.returnPct)} on current investment holdings</div></div>
      <div className="insight-card"><span className="muted">Investment positions</span><strong>{holdings.length}</strong><p>FIFO lot accounting remains owned by the Ledgerly investment backend. Finance Assistant provides the cash-flow side.</p></div>
    </div>
    <div className="chart-grid two">
      <ChartCard title="Stocks vs mutual funds" subtitle="Current value, invested capital and P&L">
        <div className="asset-pnl-grid">{['STOCK','MF'].map(kind=>{
          const rows=holdings.filter(h=>kind==='STOCK'?(h.assetType==='STOCK'||h.assetType==='EQUITY'||h.assetType==='ETF'):(h.assetType==='MF'||h.assetType==='MUTUAL FUND'||h.assetType==='MUTUAL_FUND'));
          const invested=rows.reduce((a,h)=>a+Number(h.invested||0),0); const value=rows.reduce((a,h)=>a+Number(h.value||0),0); const pnl=value-invested;
          return <div className="asset-pnl-card" key={kind}><span>{kind==='STOCK'?'Stocks':'Mutual funds'}</span><b>{money(value)}</b><small>Invested {money(invested)}</small><strong className={pnl>=0?'gain':'loss'}>{money(pnl)} · {pct(invested?pnl/invested*100:0)}</strong></div>
        })}</div>
      </ChartCard>
      <ChartCard title="Top holdings by P&L" subtitle="Unrealized gain/loss">
        {holdings.length?<div className="pnl-list">{holdings.slice().sort((a,b)=>Number(b.pnl||0)-Number(a.pnl||0)).slice(0,10).map(h=><div className="pnl-row" key={`${h.symbol}-${h.assetType}`}><div><b>{h.displayName||h.symbol}</b><span>{h.assetType==='MF'?'Mutual fund':'Stock'} · {h.qty} units</span></div><strong className={h.pnl>=0?'gain':'loss'}>{money(h.pnl)}<small>{pct(h.returnPct)}</small></strong></div>)}</div>:<ChartEmpty text="No holdings available."/>}
      </ChartCard>
    </div>
    <div className="table-card"><div className="card-title"><div><h2>Monthly breakdown</h2><p>Use this to compare income, spending, investments and resulting cash flow.</p></div></div>{monthly.length?<table><thead><tr><th>Month</th><th>Income</th><th>Spending</th><th>Investments</th><th>Net cash flow</th><th>Savings rate</th></tr></thead><tbody>{monthly.map((m,i)=>{const inc=Number(m.income||0),sp=Number(m.spending||0),inv=Number(m.investments||0),net=inc-sp-inv;return <tr key={m.month||i}><td><b>{m.month}</b></td><td className="gain">{money(inc)}</td><td>{money(sp)}</td><td>{money(inv)}</td><td className={net>=0?'gain':'loss'}>{money(net)}</td><td>{inc?(sp/inc*100).toFixed(1)+'%':'—'}</td></tr>})}</tbody></table>:<ChartEmpty text="No monthly finance data returned yet."/>}</div>
  </div>
}

function MonthlyCashFlow({data}){
  if(!data.length)return <ChartEmpty text="No monthly cash-flow history."/>
  return <ResponsiveContainer width="100%" height={310}><BarChart data={data}><CartesianGrid strokeDasharray="3 3" vertical={false}/><XAxis dataKey="month" tickLine={false}/><YAxis tickFormatter={moneyCompact} tickLine={false} axisLine={false}/><Tooltip formatter={v=>money(v)}/><Legend/><Bar dataKey="income" name="Income" fill="#5b8def"/><Bar dataKey="spending" name="Spending" fill="#d88b60"/><Bar dataKey="investments" name="Investments" fill="#6658dc"/></BarChart></ResponsiveContainer>
}
function MonthlyLine({data,keyName,label}){if(!data.length)return <ChartEmpty text="No monthly history."/>;return <ResponsiveContainer width="100%" height={310}><LineChart data={data}><CartesianGrid strokeDasharray="3 3" vertical={false}/><XAxis dataKey="month" tickLine={false}/><YAxis tickFormatter={moneyCompact} tickLine={false} axisLine={false}/><Tooltip formatter={v=>money(v)}/><Line type="monotone" dataKey={keyName} name={label} stroke="#6658dc" strokeWidth={3} dot={{r:3}}/></LineChart></ResponsiveContainer>}
function GenericBar({data,labelKey,valueKey,label}){if(!data.length)return <ChartEmpty text="No data for this period."/>;return <ResponsiveContainer width="100%" height={310}><BarChart data={data.slice(0,10)} layout="vertical" margin={{left:10,right:20}}><CartesianGrid strokeDasharray="3 3" horizontal={false}/><XAxis type="number" tickFormatter={moneyCompact}/><YAxis type="category" dataKey={labelKey} width={110}/><Tooltip formatter={v=>money(v)}/><Bar dataKey={valueKey} name={label} fill="#6658dc" radius={[0,6,6,0]}/></BarChart></ResponsiveContainer>}

function AccountsPage({finance,onEdit}){
  const accounts=safeArray(finance.bankAccounts);
  const total=Number(finance.bankTotal||accounts.reduce((s,a)=>s+Number(a.balance||0),0));
  return <div className="content"><div className="page-intro"><div><span className="eyebrow">CASH</span><h2>Bank accounts</h2><p>Balances and names are stored in the connected account source.</p></div><div className="big-total"><span>Total cash</span><b>{money(total)}</b></div></div><div className="account-grid">{accounts.length?accounts.map((a,i)=><div className="account-card" key={a.id||accountName(a)||i}><div className="account-icon"><Landmark size={19}/></div><div className="account-card-main"><b>{accountName(a)}</b><span>{a.type||'Bank account'}</span></div><strong>{money(a.balance)}</strong><button className="icon-btn account-edit" title={`Edit ${accountName(a)}`} aria-label={`Edit ${accountName(a)}`} onClick={()=>onEdit(a)}><Pencil size={16}/></button></div>):<Empty title="No bank accounts returned" text="Add an account in the BankAccounts sheet, then refresh." icon={<Landmark size={25}/>}/>}</div><div className="connection-card success"><CheckCircle2 size={21}/><div><b>Bank account source connected</b><p>Changes saved here are returned through the same portfolio response used by the dashboard and account selectors.</p></div></div></div>
}

function SettingsPage({finance}){
  const connected=finance?.connected!==false&&Object.keys(finance||{}).length>0;
  return <div className="content"><div className="settings-card"><div><span className="eyebrow">SYSTEM</span><h2>Ledgerly settings</h2><p>Investment data lives in Ledgerly; cash-flow data lives in Finance Assistant.</p></div><div className="setting-row"><div><b>Finance Assistant connection</b><span>{connected?'Bank and cash-flow data are available to Ledgerly.':'No Finance Assistant analytics payload was returned.'}</span></div><span className={`status-pill ${connected?'connected':''}`}><span/> {connected?'Connected':'Not available'}</span></div><div className="setting-row"><div><b>Accounting method</b><span>FIFO is used for lot matching on investment sells.</span></div><span className="pill">FIFO</span></div><div className="setting-row"><div><b>Market feed</b><span>Held NSE stocks use the existing GoogleFinance-backed quote refresh.</span></div><span className="pill">GoogleFinance</span></div><div className="setting-row"><div><b>Data storage</b><span>Google Sheets via Apps Script. React is not the source of truth.</span></div><span className="pill">Sheets</span></div></div></div>
}

function HoldingEditModal({holding,onClose,onSaved}){
  const source=holding.importedStock?'stocks':holding.mutualFund?'mutualFunds':'';
  const original=holding.importedStock||holding.mutualFund||{};
  const [form,setForm]=useState({...original});const [saving,setSaving]=useState(false);
  const update=(key,value)=>setForm(current=>({...current,[key]:value}));
  const submit=async event=>{event.preventDefault();setSaving(true);try{const response=await fetch(API,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'updateHolding',kind:source,...form})});const result=await response.json();if(!response.ok||result.success===false||result.ok===false)throw new Error(result.error||result.message||'Could not update holding');onSaved()}catch(error){alert(error.message||'Could not update holding')}finally{setSaving(false)}};
  return <div className="modal-backdrop"><div className="modal"><div className="modal-head"><div><span className="eyebrow">MANUAL ADJUSTMENT</span><h2>Edit holding</h2></div><button className="icon-btn" onClick={onClose}><X/></button></div><form onSubmit={submit}>{source==='stocks'?<div className="form-grid"><label>Stock name<input required value={form.symbol||''} onChange={e=>update('symbol',e.target.value.toUpperCase())}/></label><label>ISIN<input value={form.isin||''} onChange={e=>update('isin',e.target.value)}/></label><label>Quantity<input required type="number" step="any" min="0" value={form.qty??0} onChange={e=>update('qty',e.target.value)}/></label><label>Average buy price<input type="number" step="any" min="0" value={form.avg??0} onChange={e=>update('avg',e.target.value)}/></label><label>Buy value<input required type="number" step="any" min="0" value={form.invested??0} onChange={e=>update('invested',e.target.value)}/></label><label>Closing price<input required type="number" step="any" min="0" value={form.ltp??0} onChange={e=>update('ltp',e.target.value)}/></label><label>Closing value<input required type="number" step="any" min="0" value={form.value??0} onChange={e=>update('value',e.target.value)}/></label></div>:<div className="form-grid"><label>Scheme name<input required value={form.schemeName||''} onChange={e=>update('schemeName',e.target.value)}/></label><label>Folio no.<input value={form.folioNo||''} onChange={e=>update('folioNo',e.target.value)}/></label><label>Units<input required type="number" step="any" min="0" value={form.units??0} onChange={e=>update('units',e.target.value)}/></label><label>Invested value<input required type="number" step="any" min="0" value={form.invested??0} onChange={e=>update('invested',e.target.value)}/></label><label>Current value<input required type="number" step="any" min="0" value={form.value??0} onChange={e=>update('value',e.target.value)}/></label><label>XIRR<input type="number" step="any" value={form.xirr??0} onChange={e=>update('xirr',e.target.value)}/></label></div>}<div className="modal-actions"><button type="button" className="secondary" onClick={onClose}>Cancel</button><button className="primary" disabled={saving}>{saving?'Saving...':'Save changes'}</button></div></form></div></div>
}

function BankAccountEditModal({account,onClose,onSaved}){
  const [form,setForm]=useState({
    account:accountName(account),
    name:accountName(account),
    type:account.type||'Bank Account',
    currentBalance:account.balance??0,
    openingBalance:account.openingBalance??0,
    openingDate:account.openingDate||'',
    manualAdjustment:account.manualAdjustment??0,
    active:account.active!==false,
    notes:account.notes||''
  });
  const [saving,setSaving]=useState(false);
  const update=(key,value)=>setForm(f=>({...f,[key]:value}));
  const submit=async e=>{
    e.preventDefault();
    if(!String(form.name||'').trim())return alert('Account name is required.');
    setSaving(true);
    try{
      const r=await fetch(API,{
        method:'POST',
        headers:{'Content-Type':'text/plain;charset=utf-8'},
        body:JSON.stringify({
          action:'updateBankAccount',
          account:accountName(account),
          name:String(form.name).trim(),
          type:String(form.type||'Bank Account').trim(),
          currentBalance:Number(form.currentBalance||0),
          openingBalance:Number(form.openingBalance||0),
          openingDate:String(form.openingDate||''),
          manualAdjustment:Number(form.manualAdjustment||0),
          active:!!form.active,
          notes:String(form.notes||'')
        })
      });
      const next=await r.json();
      if(!r.ok||next.success===false||next.ok===false)
        throw new Error(next.error||next.message||'Could not update bank account.');
      onSaved();
    }catch(err){
      alert(err.message||'Could not update bank account.');
    }finally{
      setSaving(false);
    }
  };
  return <div className="modal-backdrop">
    <div className="modal">
      <div className="modal-head">
        <div><span className="eyebrow">ACCOUNT DETAILS</span><h2>Edit bank account</h2></div>
        <button className="icon-btn" onClick={onClose} aria-label="Close"><X/></button>
      </div>
      <form onSubmit={submit}>
        <div className="form-grid">
          <label>Account name<input required value={form.name} onChange={e=>update('name',e.target.value)}/></label>
          <label>Account type<input value={form.type} onChange={e=>update('type',e.target.value)}/></label>
          <label>Current balance<input required type="number" step="any" value={form.currentBalance} onChange={e=>update('currentBalance',e.target.value)}/></label>
          <label>Opening balance<input required type="number" step="any" value={form.openingBalance} onChange={e=>update('openingBalance',e.target.value)}/></label>
          <label>Opening date<input type="date" value={form.openingDate} onChange={e=>update('openingDate',e.target.value)}/></label>
          <label>Manual adjustment (calculated)<input type="number" step="any" value={form.manualAdjustment} readOnly/></label>
          <label className="checkbox-label"><span>Active account</span><input type="checkbox" checked={form.active} onChange={e=>update('active',e.target.checked)}/></label>
          <label>Notes<input value={form.notes} onChange={e=>update('notes',e.target.value)}/></label>
        </div>
        <div className="feed-note">
          <Wifi size={15}/>
          <span>Changing <b>Current balance</b> adjusts the Finance Assistant manual adjustment so existing transactions are preserved. Renaming the account also updates existing Finance Assistant transaction account references.</span>
        </div>
        <div className="modal-actions">
          <button type="button" className="secondary" onClick={onClose}>Cancel</button>
          <button className="primary" disabled={saving}>{saving?'Saving...':'Save changes'}</button>
        </div>
      </form>
    </div>
  </div>
}
function TransactionModal({finance,onClose,onSaved}){
  const accounts=safeArray(finance.bankAccounts);
  const [form,setForm]=useState({type:'BUY',assetType:'STOCK',symbol:'',date:new Date().toISOString().slice(0,10),quantity:'',price:'',charges:'',note:'',fundingAccount:''});
  const [saving,setSaving]=useState(false);
  const update=(k,v)=>setForm(f=>({...f,[k]:v}));
  const submit=async e=>{e.preventDefault();if(!form.fundingAccount)return alert('Select the bank account used for this investment.');setSaving(true);try{const r=await fetch(API,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'transaction',...form})});const next=await r.json();if(!r.ok||next.success===false||next.ok===false)throw new Error(next.error||'Could not save transaction');onSaved()}catch(err){alert(err.message||'Could not save transaction')}finally{setSaving(false)}};
  const total=Number(form.quantity||0)*Number(form.price||0)+Number(form.charges||0);
  return <div className="modal-backdrop"><div className="modal"><div className="modal-head"><div><span className="eyebrow">INVESTMENT CASH FLOW</span><h2>Add investment transaction</h2></div><button className="icon-btn" onClick={onClose}><X/></button></div><form onSubmit={submit}><div className="seg"><button type="button" className={form.type==='BUY'?'on':''} onClick={()=>update('type','BUY')}>Buy</button><button type="button" className={form.type==='SELL'?'on':''} onClick={()=>update('type','SELL')}>Sell</button></div><div className="form-grid"><label>Asset type<select value={form.assetType} onChange={e=>update('assetType',e.target.value)}><option value="STOCK">Stock</option><option value="MF">Mutual fund</option></select></label><label>Symbol / scheme<input required placeholder="e.g. RELIANCE" value={form.symbol} onChange={e=>update('symbol',e.target.value.toUpperCase())}/></label><label>Date<input required type="date" value={form.date} onChange={e=>update('date',e.target.value)}/></label><label>Quantity<input required min="0.000001" step="any" type="number" value={form.quantity} onChange={e=>update('quantity',e.target.value)}/></label><label>Traded price / NAV<input required min="0" step="any" type="number" value={form.price} onChange={e=>update('price',e.target.value)}/></label><label>Brokerage + charges<input min="0" step="any" type="number" value={form.charges} onChange={e=>update('charges',e.target.value)}/></label><label>Pay from / receive into<select required value={form.fundingAccount} onChange={e=>update('fundingAccount',e.target.value)}><option value="">Select account</option>{accounts.map(a=><option key={accountName(a)} value={accountName(a)}>{accountName(a)}</option>)}</select></label><label>Note<input placeholder="Optional" value={form.note} onChange={e=>update('note',e.target.value)}/></label></div><div className="cash-impact"><span>{form.type==='BUY'?'Cash outflow':'Cash inflow'}</span><b>{money(total)}</b><small>{form.fundingAccount?'Finance Assistant will receive the cash-flow entry for this account.':'Select an account to link cash flow.'}</small></div><div className="modal-actions"><button type="button" className="secondary" onClick={onClose}>Cancel</button><button className="primary" disabled={saving}>{saving?'Saving…':'Save transaction'}</button></div></form></div></div>
}

function Empty({icon,title,text}){return <div className="empty">{icon&&<div className="empty-icon">{icon}</div>}<h3>{title}</h3><p>{text}</p></div>}

createRoot(document.getElementById('root')).render(<App/>);
