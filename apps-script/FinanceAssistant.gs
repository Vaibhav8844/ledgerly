/**
 * PERSONAL FINANCE API — V8.1 FINAL
 *
 * Design:
 *  - Configuration-as-data: Accounts, PaymentModes, Categories, MerchantRules, Settings
 *  - Strategy-style query routing
 *  - Deterministic normalization/classification after AI parsing
 *  - Multiple transactions per POST request
 *  - Monthly transaction tabs: YYYY-MM
 *  - Cross-month queries
 *  - Relative-date resolution: today/yesterday/day-before-yesterday/N days ago
 *  - Alias resolution: BOB CC, SBI CC, ICICI CC, HDFC CC, etc.
 *  - Credit-card purchase vs credit-card bill-payment separation
 *  - Investment classification: SIP/MF/IPO/Stocks/ETF/etc.
 *  - Duplicate detection with LockService
 *  - Minimal JSON responses for Apple Shortcuts/Siri
 */

const SPREADSHEET_ID = "1aNCzB6nNYPHcv0fyLOGiVi5H2-hujP-rUAJmIrU8U90";

const CONFIG = {
  sheets: {
    accounts: "Accounts",
    paymentModes: "PaymentModes",
    categories: "Categories",
    merchantRules: "MerchantRules",
    settings: "Settings",
    summary: "Monthly Summary",
    dashboard: "Dashboard"
  },

  transactionHeaders: [
    "ID", "Date", "Time", "Month", "Type", "Category", "Subcategory",
    "Amount", "Payment Mode", "Account", "Merchant", "Remarks",
    "Tags", "Source", "To Account"
  ],

  defaultSource: "Apple Shortcut",
  defaultDuplicateMinutes: 5,
  cacheSeconds: 21600
};

const CACHE_KEY = "finance_config_v81";

// Performance layer: month-level transaction cache + materialized bank balances.
// Google Sheets I/O is the expensive part of this API, so read-heavy dashboard
// requests should normally be served from CacheService rather than rescanning sheets.
const LEDGERLY_PERF = {
  transactionCacheSeconds: 300,
  bankCacheSeconds: 300,
  analyticsCacheSeconds: 120,
  txPrefix: "finance_tx_v81_",
  bankKey: "finance_bank_v81"
};

/* =========================================================
 * HTTP ENTRY POINTS
 * ========================================================= */

function doPost(e) {
  try {
    const body = parseRequestBody_(e);

    if (!body || typeof body !== "object" || Array.isArray(body)) {
      return jsonResponse_({ success: false, message: "Invalid JSON request" });
    }

    // Private server-to-server API used by the separate Ledgerly backend.
    const bridge = handleLedgerlyBridgePost_(body);
    if (bridge) return jsonResponse_(bridge);

    if (normalizeText_(body.query_type) === "configuration") {
      return jsonResponse_(getConfiguration_());
    }

    if (body.query_type && !Array.isArray(body.expenses) && body.amount === undefined) {
      return jsonResponse_(handleFinanceQuery_(body));
    }

    if (Array.isArray(body.expenses)) {
      return jsonResponse_(addTransactions_(body.expenses, body.source));
    }

    if (body.amount !== undefined) {
      return jsonResponse_(addTransactions_([body], body.source));
    }

    return jsonResponse_({
      success: false,
      message: "Request must contain expenses[] or query_type"
    });
  } catch (err) {
    return jsonResponse_({ success: false, message: safeErrorMessage_(err) });
  }
}

function doGet(e) {
  const action = e && e.parameter ? String(e.parameter.action || "") : "";
  if (action === "health") {
    return jsonResponse_({ success: true, api: "Personal Finance API", version: "8.1", isolatedLedgerly: true });
  }
  return jsonResponse_({ success: true, api: "Personal Finance API", version: "8.1" });
}

/* =========================================================
 * SETUP / ADMIN
 * ========================================================= */

function setupFinanceSheets() {
  const ss = getSpreadsheet_();

  ensureConfigSheet_(ss, CONFIG.sheets.accounts,
    ["Account", "Type", "Active", "Payment Mode", "Keywords"]);

  ensureConfigSheet_(ss, CONFIG.sheets.paymentModes,
    ["Payment Mode", "Type", "Account", "Keywords", "Active"]);

  ensureConfigSheet_(ss, CONFIG.sheets.categories,
    ["Category", "Subcategory", "Keywords", "Active"]);

  ensureConfigSheet_(ss, CONFIG.sheets.merchantRules,
    ["Merchant", "Category", "Subcategory", "Keywords", "Active"]);

  ensureConfigSheet_(ss, CONFIG.sheets.settings,
    ["Setting", "Value"]);

  ensureConfigSheet_(ss, CONFIG.sheets.summary,
    ["Month", "Expenses", "Investments", "Credit Card Payments", "Income", "Updated At"]);

  seedDefaults_();
  setupFinanceBankBalances_();
  invalidateConfigCache_();

  return "Finance sheets ready.";
}

function upgradeFinanceShorthandAliases() {
  const ss = getSpreadsheet_();

  const aliases = {
    "SBI Credit Card": [
      "sbi cc", "sbi card", "sbi credit card", "sbicard"
    ],
    "ICICI Credit Card": [
      "icici cc", "icici card", "icici credit card", "icici"
    ],
    "BOB Credit Card": [
      "bob cc", "bob card", "bob credit card",
      "bank of baroda cc", "bank of baroda card",
      "bank of baroda credit card", "baroda cc"
    ],
    "HDFC Credit Card": [
      "hdfc cc", "hdfc card", "hdfc credit card"
    ],
    "SBI Savings Account": [
      "sbi savings", "sbi savings account", "sbi bank", "savings account"
    ]
  };

  const accountSheet = ss.getSheetByName(CONFIG.sheets.accounts);
  if (accountSheet) {
    mergeAliasesIntoSheet_(accountSheet, "Account", "Keywords", aliases);
  }

  const pmSheet = ss.getSheetByName(CONFIG.sheets.paymentModes);
  if (pmSheet) {
    mergeAliasesIntoSheet_(pmSheet, "Payment Mode", "Keywords", aliases);
  }

  invalidateConfigCache_();
  return "Aliases upgraded.";
}

function migrateLegacyTransactions() {
  const ss = getSpreadsheet_();
  const legacy = ss.getSheetByName("Transactions");

  if (!legacy || legacy.getLastRow() < 2) {
    return "No legacy Transactions sheet found.";
  }

  const headers = legacy.getRange(
    1, 1, 1, legacy.getLastColumn()
  ).getValues()[0].map(h => String(h).trim());

  const map = {};
  headers.forEach((h, i) => map[h] = i);

  const values = legacy.getRange(
    2, 1, legacy.getLastRow() - 1, legacy.getLastColumn()
  ).getValues();

  const grouped = {};

  values.forEach(row => {
    const date = formatSheetDate_(getCell_(row, map, "Date"));
    if (!isValidDateString_(date)) return;

    const month = date.slice(0, 7);

    const tx = {
      id: clean_(getCell_(row, map, "ID")) || generateId_(),
      date: date,
      time: formatSheetTime_(getCell_(row, map, "Time")),
      month: month,
      type: clean_(getCell_(row, map, "Type")) || "Expense",
      category: clean_(getCell_(row, map, "Category")) || "Other",
      subcategory: clean_(getCell_(row, map, "Subcategory")),
      amount: parseAmount_(getCell_(row, map, "Amount")),
      paymentMode: clean_(getCell_(row, map, "Payment Mode")),
      account: clean_(getCell_(row, map, "Account")),
      merchant: clean_(getCell_(row, map, "Merchant")),
      remarks: clean_(getCell_(row, map, "Remarks")),
      tags: clean_(getCell_(row, map, "Tags")),
      source: clean_(getCell_(row, map, "Source")) || CONFIG.defaultSource,
      toAccount: clean_(getCell_(row, map, "To Account"))
    };

    if (!grouped[month]) grouped[month] = [];
    grouped[month].push(tx);
  });

  Object.keys(grouped).forEach(month => {
    appendTransactions_(month, grouped[month]);
  });

  updateMonthlySummary_(Object.keys(grouped));

  return "Migrated " + values.length + " legacy rows.";
}

function installFinanceConfigTriggers() {
  const ss = getSpreadsheet_();

  ScriptApp.getProjectTriggers().forEach(trigger => {
    const fn = trigger.getHandlerFunction();
    if (fn === "onFinanceConfigEdit" || fn === "onFinanceConfigChange") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("onFinanceConfigEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ScriptApp.newTrigger("onFinanceConfigChange")
    .forSpreadsheet(ss)
    .onChange()
    .create();

  return "Configuration triggers installed.";
}

function onFinanceConfigEdit(e) {
  if (!e || !e.range) return;
  const sheetName = e.range.getSheet().getName();
  if (isConfigSheet_(sheetName)) invalidateConfigCache_();
  if (sheetName === LEDGERLY_BRIDGE.bankSheet) invalidateFinanceAnalyticsCache_();
}

function onFinanceConfigChange() {
  invalidateConfigCache_();
}

function seedDefaults_() {
  const ss = getSpreadsheet_();

  seedIfEmpty_(ss.getSheetByName(CONFIG.sheets.settings), [
    ["Duplicate Window Minutes", "5"],
    ["Default Source", CONFIG.defaultSource]
  ]);

  seedIfEmpty_(ss.getSheetByName(CONFIG.sheets.accounts), [
    ["SBI Savings Account", "Bank Account", "TRUE", "UPI",
      "sbi savings,sbi savings account,sbi bank,savings account"],
    ["SBI Credit Card", "Credit Card", "TRUE", "SBI Credit Card",
      "sbi cc,sbi card,sbi credit card,sbicard"],
    ["ICICI Credit Card", "Credit Card", "TRUE", "ICICI Credit Card",
      "icici cc,icici card,icici credit card,icici"],
    ["BOB Credit Card", "Credit Card", "TRUE", "BOB Credit Card",
      "bob cc,bob card,bob credit card,bank of baroda cc,bank of baroda card,bank of baroda credit card,baroda cc"],
    ["HDFC Credit Card", "Credit Card", "TRUE", "HDFC Credit Card",
      "hdfc cc,hdfc card,hdfc credit card"],
    ["Cash", "Cash", "TRUE", "Cash", "cash"]
  ]);

  seedIfEmpty_(ss.getSheetByName(CONFIG.sheets.paymentModes), [
    ["UPI", "Digital", "SBI Savings Account", "upi,upi payment", "TRUE"],
    ["Cash", "Cash", "Cash", "cash", "TRUE"],
    ["SBI Credit Card", "Credit Card", "SBI Credit Card",
      "sbi cc,sbi card,sbi credit card,sbicard", "TRUE"],
    ["ICICI Credit Card", "Credit Card", "ICICI Credit Card",
      "icici cc,icici card,icici credit card,icici", "TRUE"],
    ["BOB Credit Card", "Credit Card", "BOB Credit Card",
      "bob cc,bob card,bob credit card,bank of baroda cc,bank of baroda card,bank of baroda credit card,baroda cc", "TRUE"],
    ["HDFC Credit Card", "Credit Card", "HDFC Credit Card",
      "hdfc cc,hdfc card,hdfc credit card", "TRUE"]
  ]);

  seedIfEmpty_(ss.getSheetByName(CONFIG.sheets.categories), [
    ["Food & Dining", "Snacks", "snack,snacks,chips,biscuit,biscuits", "TRUE"],
    ["Food & Dining", "Breakfast", "breakfast", "TRUE"],
    ["Food & Dining", "Lunch", "lunch", "TRUE"],
    ["Food & Dining", "Dinner", "dinner,diner", "TRUE"],
    ["Food & Dining", "Coffee & Beverages", "coffee,tea,cafe,café,beverage,beverages", "TRUE"],
    ["Food & Dining", "Food Delivery", "food delivery,delivery", "TRUE"],

    ["Groceries", "Fruits", "fruit,fruits", "TRUE"],
    ["Groceries", "Vegetables", "vegetable,vegetables", "TRUE"],
    ["Groceries", "Groceries", "grocery,groceries", "TRUE"],

    ["Transport", "Cab", "uber,ola,rapido,cab,taxi,auto,autorickshaw,auto rickshaw", "TRUE"],
    ["Transport", "Fuel", "petrol,diesel,fuel,gas station", "TRUE"],
    ["Transport", "Parking", "parking", "TRUE"],
    ["Transport", "Toll", "toll,fastag", "TRUE"],
    ["Transport", "Bus", "bus,bus ticket", "TRUE"],
    ["Transport", "Train", "train,railway,rail ticket,train ticket", "TRUE"],
    ["Transport", "Flight", "flight,airfare,air ticket,plane", "TRUE"],
    ["Transport", "Metro", "metro", "TRUE"],

    ["Travel", "Hotel", "hotel,stay,airbnb", "TRUE"],
    ["Travel", "Travel", "travel", "TRUE"],

    ["Shopping", "Online Shopping", "amazon,flipkart,myntra,online shopping", "TRUE"],
    ["Shopping", "Electronics", "croma,vijay sales,electronics", "TRUE"],
    ["Shopping", "Clothing", "clothes,clothing,shirt,shoes", "TRUE"],

    ["Bills & Utilities", "Rent", "rent,house rent", "TRUE"],
    ["Bills & Utilities", "Electricity", "electricity,electric bill", "TRUE"],
    ["Bills & Utilities", "Internet", "internet,broadband,wifi", "TRUE"],
    ["Bills & Utilities", "Mobile", "mobile recharge,recharge", "TRUE"],

    ["Entertainment", "Subscription", "netflix,spotify,prime,hotstar,subscription", "TRUE"],
    ["Health & Fitness", "Medicine", "medicine,medicines,pharmacy,doctor", "TRUE"],
    ["Health & Fitness", "Fitness", "gym,fitness", "TRUE"],
    ["Education", "Courses", "course,tuition,education", "TRUE"],

    ["Financial", "Credit Card Bill", "credit card bill,card bill", "TRUE"],

    ["Investments", "SIP", "sip,systematic investment plan", "TRUE"],
    ["Investments", "Mutual Fund", "mutual fund,mf", "TRUE"],
    ["Investments", "IPO", "ipo,ipos,initial public offering", "TRUE"],
    ["Investments", "Stocks", "stock,stocks,share,shares,equity", "TRUE"],
    ["Investments", "ETF", "etf", "TRUE"],
    ["Investments", "Bonds", "bond,bonds", "TRUE"],
    ["Investments", "Gold", "gold", "TRUE"],
    ["Investments", "PPF", "ppf", "TRUE"],
    ["Investments", "NPS", "nps", "TRUE"],
    ["Investments", "FD", "fd,fixed deposit", "TRUE"],
    ["Investments", "RD", "rd,recurring deposit", "TRUE"],

    ["Other", "", "", "TRUE"]
  ]);

  seedIfEmpty_(ss.getSheetByName(CONFIG.sheets.merchantRules), [
    ["Swiggy", "Food & Dining", "Food Delivery", "swiggy", "TRUE"],
    ["Zomato", "Food & Dining", "Food Delivery", "zomato", "TRUE"],
    ["Uber", "Transport", "Cab", "uber", "TRUE"],
    ["Ola", "Transport", "Cab", "ola", "TRUE"],
    ["Rapido", "Transport", "Cab", "rapido", "TRUE"],
    ["Amazon", "Shopping", "Online Shopping", "amazon", "TRUE"],
    ["Flipkart", "Shopping", "Online Shopping", "flipkart", "TRUE"],
    ["Netflix", "Entertainment", "Subscription", "netflix", "TRUE"],
    ["Spotify", "Entertainment", "Subscription", "spotify", "TRUE"]
  ]);
}

/* =========================================================
 * TRANSACTION WRITE PATH
 * ========================================================= */

function addTransactions_(expenses, source) {
  if (!Array.isArray(expenses) || expenses.length === 0) {
    return {
      success: false,
      added: 0,
      duplicates: 0,
      failed: 0,
      message: "No expenses supplied"
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const cfg = getConfig_();
    // Ensure the materialized bank-balance column exists before this write.
    // This prevents the first transaction after migration from being counted twice.
    try { ensureBankBalancesMaterialized_(); } catch (_) {}
    const now = new Date();
    const defaultSource =
      clean_(source) ||
      cfg.settings["Default Source"] ||
      CONFIG.defaultSource;

    const results = [];
    const accepted = [];
    let duplicates = 0;

    expenses.forEach((raw, index) => {
      try {
        const tx = normalizeTransaction_(raw, cfg, now, defaultSource);

        if (!isFinite(tx.amount) || tx.amount <= 0) {
          throw new Error("Invalid amount");
        }

        const duplicate = findDuplicate_(tx, cfg);

        if (duplicate) {
          duplicates++;
          results.push({
            index: index + 1,
            status: "duplicate",
            id: duplicate.id
          });
          return;
        }

        tx.id = generateId_();
        accepted.push(tx);

        results.push({
          index: index + 1,
          status: "added",
          id: tx.id
        });

      } catch (err) {
        results.push({
          index: index + 1,
          status: "failed",
          error: safeErrorMessage_(err)
        });
      }
    });

    // Group by month and write each month in one batch.
    const grouped = {};
    accepted.forEach(tx => {
      if (!grouped[tx.month]) grouped[tx.month] = [];
      grouped[tx.month].push(tx);
    });

    Object.keys(grouped).forEach(month => {
      appendTransactions_(month, grouped[month]);
    });

    updateMonthlySummaryIncremental_(accepted);
    applyBankBalanceDeltas_(accepted, 1);
    invalidateFinanceAnalyticsCache_();

    const added = accepted.length;
    const failed = results.filter(r => r.status === "failed").length;

    let message;

    if (added && duplicates && failed) {
      message = added + " added, " + duplicates +
        " duplicate(s) skipped, " + failed + " failed";
    } else if (added && duplicates) {
      message = added + " added, " + duplicates + " duplicate(s) skipped";
    } else if (added && failed) {
      message = added + " added, " + failed + " failed";
    } else if (added === 1) {
      message = "1 transaction added";
    } else if (added > 1) {
      message = added + " transactions added";
    } else if (duplicates) {
      message = duplicates + " duplicate(s) skipped";
    } else {
      message = "No transactions were added";
    }

    return {
      success: failed === 0,
      added: added,
      duplicates: duplicates,
      failed: failed,
      message: message,
      results: results
    };

  } finally {
    lock.releaseLock();
  }
}

function normalizeTransaction_(raw, cfg, now, source) {
  const input = raw || {};

  // source_text is strongly preferred. It lets the backend recover
  // information that the AI may have omitted from structured fields.
  const text = buildTransactionText_(input);
  const lower = normalizeText_(text);

  const amount = parseAmount_(input.amount);
  if (!amount || amount <= 0) throw new Error("Invalid amount");

  const date = resolveTransactionDate_(
    input.date,
    now,
    lower
  );

  let type = normalizeType_(input.type) || "Expense";

  let category = canonicalCategory_(input.category, cfg);
  let subcategory = canonicalSubcategory_(
    input.subcategory,
    category,
    cfg
  );

  let merchant = clean_(input.merchant);
  let remarks = clean_(input.remarks);
  let tags = clean_(input.tags);

  // Payment method and account are resolved independently.
  let paymentMode = resolvePaymentMode_(
    input.payment_mode,
    lower,
    cfg
  );

  let toAccount = resolveAccount_(
    input.to_account,
    lower,
    cfg
  );

  /*
   * 1. Explicit activity/category keywords.
   * These are stronger than generic merchant defaults.
   *
   * Example:
   * "lunch on Swiggy" -> Lunch, not Food Delivery.
   */
  const explicitHit = bestCategoryHit_(
    lower,
    cfg.categories.filter(c => c.active && !same_(c.category, "Other"))
  );

  /*
   * 2. Merchant identification.
   */
  let merchantRule = findMerchantRule_(
    merchant,
    lower,
    cfg.merchantRules
  );

  if (merchantRule && !merchant) {
    merchant = merchantRule.merchant;
  }

  /*
   * 3. Merchant rule supplies a fallback category/subcategory.
   * Explicit activity keywords win over this.
   */
  if (merchantRule) {
    if (!category || category === "Other") {
      category = merchantRule.category || "";
    }

    if (!subcategory) {
      subcategory = merchantRule.subcategory || "";
    }
  }

  /*
   * 4. Explicit category keyword wins over merchant generic category.
   */
  if (explicitHit) {
    if (!category || category === "Other") {
      category = explicitHit.category;
    }

    if (explicitHit.subcategory) {
      subcategory = explicitHit.subcategory;
    }
  }

  /*
   * 5. If the parser supplied a merchant but no rule was found,
   * use the category keyword engine.
   */
  if ((!category || category === "Other") && explicitHit) {
    category = explicitHit.category;
    subcategory = explicitHit.subcategory || subcategory;
  }

  /*
   * 6. Credit-card bill detection.
   *
   * "337 dinner using BOB CC" = normal Expense.
   * "13204 pay BOB CC bill using UPI" = Credit Card Payment.
   */
  if (looksLikeCreditCardBill_(lower)) {
    type = "Credit Card Payment";
    category = "Financial";
    subcategory = "Credit Card Bill";

    const destinationCard =
      resolveCreditCardDestination_(input.to_account, lower, cfg);

    if (destinationCard) {
      toAccount = destinationCard;
    }

    const sourceMethod =
      resolvePaymentMethodFromBillText_(lower, cfg);

    if (sourceMethod) {
      paymentMode = sourceMethod;
    } else {
      // Never guess the method for a bill payment.
      paymentMode = resolvePaymentMode_(
        input.payment_mode,
        lower,
        cfg
      );
    }

  } else {
    /*
     * Normal expense:
     * a credit card is a payment method, not To Account.
     */
    if (toAccount && isCreditCardAccount_(toAccount, cfg)) {
      if (!paymentMode) paymentMode = toAccount;
      toAccount = "";
    }

    const cardFromText = resolveCreditCardFromText_(lower, cfg);
    if (cardFromText) {
      paymentMode = cardFromText;
    }
  }

  /*
   * 7. Investment detection.
   * Investment keywords take precedence over ordinary Expense.
   */
  const investmentHit = classifyInvestment_(
    lower,
    cfg.categories
  );

  if (investmentHit) {
    type = "Investment";
    category = "Investments";
    subcategory = investmentHit.subcategory || subcategory || "";
  }

  /*
   * 8. Payment-mode correction.
   */
  if (paymentMode) {
    paymentMode = canonicalPaymentMode_(paymentMode, cfg);
  }

  /*
   * 9. Account comes from the payment mode.
   */
  let account = "";
  if (paymentMode) {
    account = accountForPaymentMode_(paymentMode, cfg);
  }

  // Explicit account supplied by an internal integration wins over
  // payment-mode inference. This is used by the isolated Ledgerly backend.
  const explicitAccount = clean_(input.account);
  if (explicitAccount) {
    const matchedAccount = cfg.accounts.find(a =>
      same_(a.name, explicitAccount) ||
      (a.keywords || []).some(k => same_(k, explicitAccount))
    );
    if (matchedAccount) account = matchedAccount.name;
  }

  /*
   * 10. If UPI was explicitly used but configuration does not contain
   * an account mapping, leave Account blank rather than inventing one.
   */

  /*
   * 11. Final validation.
   */
  if (!category || !isConfiguredCategory_(category, cfg)) {
    category = "Other";
  }

  if (
    subcategory &&
    !isConfiguredSubcategory_(category, subcategory, cfg)
  ) {
    subcategory = "";
  }

  if (isPaymentModeLike_(subcategory, cfg)) {
    subcategory = "";
  }

  if (!remarks) {
    remarks = buildDefaultRemarks_(
      lower,
      category,
      subcategory,
      type,
      merchant,
      toAccount
    );
  }

  return {
    id: "",
    date: date,
    time: resolveTime_(input.time, now),
    month: date.slice(0, 7),
    type: type,
    category: category,
    subcategory: subcategory || "",
    amount: amount,
    paymentMode: paymentMode || "",
    account: account || "",
    merchant: merchant || "",
    remarks: remarks || "",
    tags: tags || "",
    source: source || CONFIG.defaultSource,
    toAccount: toAccount || ""
  };
}

/* =========================================================
 * MONTHLY STORAGE
 * ========================================================= */

function getOrCreateMonthSheet_(month) {
  if (!/^\d{4}-\d{2}$/.test(month)) {
    throw new Error("Invalid month");
  }

  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(month);

  if (!sh) {
    sh = ss.insertSheet(month);
    sh.getRange(
      1, 1, 1, CONFIG.transactionHeaders.length
    ).setValues([CONFIG.transactionHeaders]);
    sh.setFrozenRows(1);
  }

  ensureHeaders_(sh, CONFIG.transactionHeaders);
  return sh;
}

function appendTransactions_(month, transactions) {
  if (!transactions || !transactions.length) return;

  const sh = getOrCreateMonthSheet_(month);
  const headerMap = getHeaderMap_(sh);
  const lastColumn = sh.getLastColumn();

  const rows = transactions.map(tx => {
    const row = new Array(lastColumn).fill("");

    setRowValue_(row, headerMap, "ID", tx.id);
    setRowValue_(row, headerMap, "Date", tx.date);
    setRowValue_(row, headerMap, "Time", tx.time);
    setRowValue_(row, headerMap, "Month", tx.month);
    setRowValue_(row, headerMap, "Type", tx.type);
    setRowValue_(row, headerMap, "Category", tx.category);
    setRowValue_(row, headerMap, "Subcategory", tx.subcategory);
    setRowValue_(row, headerMap, "Amount", tx.amount);
    setRowValue_(row, headerMap, "Payment Mode", tx.paymentMode);
    setRowValue_(row, headerMap, "Account", tx.account);
    setRowValue_(row, headerMap, "Merchant", tx.merchant);
    setRowValue_(row, headerMap, "Remarks", tx.remarks);
    setRowValue_(row, headerMap, "Tags", tx.tags);
    setRowValue_(row, headerMap, "Source", tx.source);
    setRowValue_(row, headerMap, "To Account", tx.toAccount);

    return row;
  });

  sh.getRange(
    sh.getLastRow() + 1,
    1,
    rows.length,
    lastColumn
  ).setValues(rows);

  invalidateFinanceMonthCache_(month);
}

/* =========================================================
 * DUPLICATE DETECTION
 * ========================================================= */

function findDuplicate_(tx, cfg) {
  const minutes =
    Number(cfg.settings["Duplicate Window Minutes"]) ||
    CONFIG.defaultDuplicateMinutes;

  const windowMs = minutes * 60 * 1000;
  const date = parseDateOnly_(tx.date);

  const start = addDays_(date, -1);
  const end = addDays_(date, 1);

  const candidates = readTransactions_(
    monthsBetween_(dateString_(start), dateString_(end))
  );

  const target = transactionTimestamp_(tx);

  return candidates.find(existing => {
    if (!existing.id) return false;

    return (
      Math.abs(transactionTimestamp_(existing) - target) <= windowMs &&
      Number(existing.amount) === Number(tx.amount) &&
      same_(existing.type, tx.type) &&
      same_(existing.category, tx.category) &&
      same_(existing.subcategory, tx.subcategory) &&
      same_(existing.paymentMode, tx.paymentMode) &&
      same_(existing.account, tx.account) &&
      same_(existing.merchant, tx.merchant) &&
      same_(existing.toAccount, tx.toAccount)
    );
  }) || null;
}

/* =========================================================
 * READ TRANSACTIONS
 * ========================================================= */

function readTransactions_(months) {
  const requested = Array.isArray(months) ? months.filter(Boolean).map(String) : [];
  if (!requested.length) return [];

  const cache = CacheService.getScriptCache();
  const output = [];
  const missing = [];

  requested.forEach(month => {
    const key = LEDGERLY_PERF.txPrefix + month;
    const cached = cache.get(key);
    if (cached !== null) {
      try {
        const rows = JSON.parse(cached);
        if (Array.isArray(rows)) output.push.apply(output, rows);
        return;
      } catch (_) {}
    }
    missing.push(month);
  });

  if (!missing.length) return output;

  const ss = getSpreadsheet_();
  missing.forEach(month => {
    const sh = ss.getSheetByName(month);
    if (!sh || sh.getLastRow() < 2) {
      cache.put(LEDGERLY_PERF.txPrefix + month, "[]", LEDGERLY_PERF.transactionCacheSeconds);
      return;
    }

    const lastColumn = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0].map(v => String(v).trim());
    const map = {};
    headers.forEach((h, i) => map[h] = i);
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, lastColumn).getValues();
    const monthRows = [];

    values.forEach(row => {
      const tx = {
        id: getCell_(row, map, "ID"),
        date: formatSheetDate_(getCell_(row, map, "Date")),
        time: formatSheetTime_(getCell_(row, map, "Time")),
        month: clean_(getCell_(row, map, "Month")) || month,
        type: clean_(getCell_(row, map, "Type")),
        category: clean_(getCell_(row, map, "Category")),
        subcategory: clean_(getCell_(row, map, "Subcategory")),
        amount: parseAmount_(getCell_(row, map, "Amount")),
        paymentMode: clean_(getCell_(row, map, "Payment Mode")),
        account: clean_(getCell_(row, map, "Account")),
        merchant: clean_(getCell_(row, map, "Merchant")),
        remarks: clean_(getCell_(row, map, "Remarks")),
        tags: clean_(getCell_(row, map, "Tags")),
        source: clean_(getCell_(row, map, "Source")),
        toAccount: clean_(getCell_(row, map, "To Account"))
      };
      if (tx.id || tx.amount) monthRows.push(tx);
    });

    // CacheService entries have a size limit; skip caching unusually large months.
    const serialized = JSON.stringify(monthRows);
    if (serialized.length < 95000) {
      try { cache.put(LEDGERLY_PERF.txPrefix + month, serialized, LEDGERLY_PERF.transactionCacheSeconds); } catch (_) {}
    }
    output.push.apply(output, monthRows);
  });

  return output;
}

function invalidateFinanceMonthCache_(month) {
  if (!month) return;
  try { CacheService.getScriptCache().remove(LEDGERLY_PERF.txPrefix + String(month)); } catch (_) {}
}

function invalidateFinanceAnalyticsCache_() {
  try { CacheService.getScriptCache().remove(LEDGERLY_PERF.bankKey); } catch (_) {}
}

/* =========================================================
 * FINANCE QUERY ENGINE
 * ========================================================= */

function handleFinanceQuery_(q) {
  const cfg = getConfig_();
  const range = resolveQueryRange_(q);

  const months = monthsBetween_(
    range.start,
    range.end
  );

  const txs = readTransactions_(months)
    .filter(t => withinDateRange_(t.date, range.start, range.end));

  const type = normalizeText_(q.query_type);

  switch (type) {
    case "total_spending":
      return queryTotalSpending_(txs, range);

    case "category_spending":
      return queryCategorySpending_(txs, range, q);

    case "merchant_spending":
      return queryMerchantSpending_(txs, range, q);

    case "payment_spending":
      return queryPaymentSpending_(txs, range, q, cfg);

    case "account_spending":
      return queryAccountSpending_(txs, range, q, cfg);

    case "spending_by_category":
      return querySpendingByCategory_(txs, range);

    case "transaction_count":
      return queryTransactionCount_(txs, range);

    case "largest_expenses":
      return queryLargest_(txs, range, q);

    case "recent_transactions":
      return queryRecent_(txs, range, q);

    case "credit_card_payments":
      return queryCreditCardPayments_(txs, range, q, cfg);

    case "investment_spending":
      return queryInvestmentSpending_(txs, range);

    case "investments_by_type":
      return queryInvestmentsByType_(txs, range, q, cfg);

    case "income":
      return queryIncome_(txs, range);

    default:
      return {
        success: false,
        message: "Unknown query_type"
      };
  }
}

function queryTotalSpending_(txs, range) {
  const list = txs.filter(t => same_(t.type, "Expense"));

  return queryResult_("total_spending", range, {
    total: roundMoney_(sum_(list)),
    transaction_count: list.length
  });
}

function queryCategorySpending_(txs, range, q) {
  const cfg = getConfig_();

  const category =
    canonicalCategory_(q.category, cfg);

  const subcategory =
    canonicalSubcategory_(q.subcategory, category, cfg);

  const list = txs.filter(t =>
    same_(t.type, "Expense") &&
    (!category || same_(t.category, category)) &&
    (!subcategory || same_(t.subcategory, subcategory))
  );

  return queryResult_("category_spending", range, {
    category: category || "",
    subcategory: subcategory || "",
    total: roundMoney_(sum_(list)),
    transaction_count: list.length
  });
}

function queryMerchantSpending_(txs, range, q) {
  const merchant = clean_(q.merchant);

  const list = txs.filter(t =>
    same_(t.type, "Expense") &&
    (!merchant || same_(t.merchant, merchant))
  );

  return queryResult_("merchant_spending", range, {
    merchant: merchant || "",
    total: roundMoney_(sum_(list)),
    transaction_count: list.length
  });
}

function queryPaymentSpending_(txs, range, q, cfg) {
  const pm = resolvePaymentMode_(
    q.payment_mode,
    q.payment_mode,
    cfg
  );

  const list = txs.filter(t =>
    same_(t.type, "Expense") &&
    (!pm || same_(t.paymentMode, pm))
  );

  return queryResult_("payment_spending", range, {
    payment_mode: pm || "",
    total: roundMoney_(sum_(list)),
    transaction_count: list.length
  });
}

function queryAccountSpending_(txs, range, q, cfg) {
  const account = resolveAccount_(
    q.account,
    q.account,
    cfg
  );

  const list = txs.filter(t =>
    same_(t.type, "Expense") &&
    (!account || same_(t.account, account))
  );

  return queryResult_("account_spending", range, {
    account: account || "",
    total: roundMoney_(sum_(list)),
    transaction_count: list.length
  });
}

function querySpendingByCategory_(txs, range) {
  const map = {};

  txs
    .filter(t => same_(t.type, "Expense"))
    .forEach(t => {
      const key = t.category || "Other";
      map[key] = (map[key] || 0) + Number(t.amount || 0);
    });

  const breakdown = Object.keys(map)
    .sort((a, b) => map[b] - map[a])
    .map(k => ({
      category: k,
      total: roundMoney_(map[k])
    }));

  return queryResult_("spending_by_category", range, {
    total: roundMoney_(
      sum_(txs.filter(t => same_(t.type, "Expense")))
    ),
    breakdown: breakdown
  });
}

function queryTransactionCount_(txs, range) {
  return queryResult_("transaction_count", range, {
    count: txs.length
  });
}

function queryLargest_(txs, range, q) {
  const limit = getLimit_(q.limit, 5);

  const list = txs
    .filter(t => same_(t.type, "Expense"))
    .sort((a, b) => Number(b.amount) - Number(a.amount))
    .slice(0, limit)
    .map(simplifyTransaction_);

  return queryResult_("largest_expenses", range, {
    transactions: list
  });
}

function queryRecent_(txs, range, q) {
  const limit = getLimit_(q.limit, 5);

  const list = txs
    .slice()
    .sort((a, b) =>
      transactionTimestamp_(b) - transactionTimestamp_(a)
    )
    .slice(0, limit)
    .map(simplifyTransaction_);

  return queryResult_("recent_transactions", range, {
    transactions: list
  });
}

function queryCreditCardPayments_(txs, range, q, cfg) {
  const account = resolveAccount_(
    q.account,
    q.account,
    cfg
  );

  const list = txs.filter(t =>
    same_(t.type, "Credit Card Payment") &&
    (!account || same_(t.toAccount, account))
  );

  return queryResult_("credit_card_payments", range, {
    account: account || "",
    total: roundMoney_(sum_(list)),
    transaction_count: list.length,
    transactions: list.map(simplifyTransaction_)
  });
}

function queryInvestmentSpending_(txs, range) {
  const list = txs.filter(t =>
    same_(t.type, "Investment") ||
    same_(t.category, "Investments")
  );

  return queryResult_("investment_spending", range, {
    total: roundMoney_(sum_(list)),
    transaction_count: list.length
  });
}

function queryInvestmentsByType_(txs, range, q, cfg) {
  const requested =
    canonicalSubcategory_(q.subcategory, "Investments", cfg);

  const list = txs.filter(t =>
    (
      same_(t.type, "Investment") ||
      same_(t.category, "Investments")
    ) &&
    (!requested || same_(t.subcategory, requested))
  );

  const map = {};

  list.forEach(t => {
    const key = t.subcategory || "Other";
    map[key] = (map[key] || 0) + Number(t.amount || 0);
  });

  const breakdown = Object.keys(map)
    .sort((a, b) => map[b] - map[a])
    .map(k => ({
      subcategory: k,
      total: roundMoney_(map[k])
    }));

  return queryResult_("investments_by_type", range, {
    subcategory: requested || "",
    total: roundMoney_(sum_(list)),
    transaction_count: list.length,
    breakdown: breakdown
  });
}

function queryIncome_(txs, range) {
  const list = txs.filter(t => same_(t.type, "Income"));

  return queryResult_("income", range, {
    total: roundMoney_(sum_(list)),
    transaction_count: list.length
  });
}

function queryResult_(queryType, range, data) {
  return Object.assign({
    success: true,
    query_type: queryType,
    start_date: range.start,
    end_date: range.end
  }, data);
}

/* =========================================================
 * DATE / PERIOD ENGINE
 * ========================================================= */

function resolveQueryRange_(q) {
  const tz =
    getSpreadsheet_().getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    "Asia/Kolkata";

  const now = new Date();
  const todayStr =
    Utilities.formatDate(now, tz, "yyyy-MM-dd");

  const period = normalizePeriod_(q.period);

  if (period) {
    return periodToRange_(period, now, tz);
  }

  // Only use explicit start/end when the assistant intentionally
  // supplied them. Empty strings must NOT become today's date.
  const start = parseDateOnly_(q.start_date);
  const end = parseDateOnly_(q.end_date);

  if (start || end) {
    return {
      start: dateString_(start || end),
      end: dateString_(end || start)
    };
  }

  // Default date range for spending queries = current month.
  const qt = normalizeText_(q.query_type);

  const monthQueries = [
    "total_spending",
    "category_spending",
    "merchant_spending",
    "payment_spending",
    "account_spending",
    "spending_by_category",
    "transaction_count",
    "largest_expenses",
    "recent_transactions",
    "credit_card_payments",
    "investment_spending",
    "investments_by_type",
    "income"
  ];

  if (monthQueries.indexOf(qt) >= 0) {
    return periodToRange_("this_month", now, tz);
  }

  return {
    start: "1900-01-01",
    end: todayStr
  };
}

function normalizePeriod_(p) {
  const s = normalizeText_(p)
    .replace(/\s+/g, "_");

  const aliases = {
    "today": "today",
    "yesterday": "yesterday",
    "tomorrow": "tomorrow",
    "day_before_yesterday": "day_before_yesterday",
    "this_week": "this_week",
    "last_week": "last_week",
    "this_month": "this_month",
    "last_month": "last_month",
    "this_year": "this_year",
    "last_year": "last_year",
    "last_7_days": "last_7_days",
    "last_30_days": "last_30_days",
    "thismonth": "this_month",
    "lastmonth": "last_month",
    "thisweek": "this_week",
    "lastweek": "last_week"
  };

  return aliases[s] || "";
}

function formatDateInTimezone_(date, timezone) {
  return Utilities.formatDate(
    date,
    timezone,
    "yyyy-MM-dd"
  );
}

function periodToRange_(period, now, tz) {
  const today = parseDateOnly_(
    Utilities.formatDate(now, tz, "yyyy-MM-dd")
  );

  let start;
  let end;

  switch (period) {
    case "today":
      start = today;
      end = today;
      break;

    case "yesterday":
      start = addDays_(today, -1);
      end = start;
      break;

    case "day_before_yesterday":
      start = addDays_(today, -2);
      end = start;
      break;

    case "tomorrow":
      start = addDays_(today, 1);
      end = start;
      break;

    case "this_week": {
      const day = today.getDay(); // Sunday=0
      const mondayOffset = -((day + 6) % 7);
      start = addDays_(today, mondayOffset);
      end = today;
      break;
    }

    case "last_week": {
      const day = today.getDay();
      const thisMonday = addDays_(
        today,
        -((day + 6) % 7)
      );
      start = addDays_(thisMonday, -7);
      end = addDays_(thisMonday, -1);
      break;
    }

    case "this_month":
      start = new Date(
        today.getFullYear(),
        today.getMonth(),
        1
      );
      end = today;
      break;

    case "last_month":
      start = new Date(
        today.getFullYear(),
        today.getMonth() - 1,
        1
      );
      end = new Date(
        today.getFullYear(),
        today.getMonth(),
        0
      );
      break;

    case "this_year":
      start = new Date(today.getFullYear(), 0, 1);
      end = today;
      break;

    case "last_year":
      start = new Date(today.getFullYear() - 1, 0, 1);
      end = new Date(today.getFullYear() - 1, 11, 31);
      break;

    case "last_7_days":
      start = addDays_(today, -6);
      end = today;
      break;

    case "last_30_days":
      start = addDays_(today, -29);
      end = today;
      break;

    default:
      start = today;
      end = today;
  }

  return {
    start: dateString_(start),
    end: dateString_(end)
  };
}

function resolveTransactionDate_(value, now, contextText) {
  const tz =
    getSpreadsheet_().getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    "Asia/Kolkata";

  // IMPORTANT:
  // Get today's calendar date from the spreadsheet timezone.
  const todayString = Utilities.formatDate(
    now,
    tz,
    "yyyy-MM-dd"
  );

  const today = parseDateOnly_(todayString);
  const context = normalizeText_(contextText || "");

  /*
   * RELATIVE DATE HAS ABSOLUTE PRIORITY.
   *
   * Never trust the AI-generated `date` if the original
   * transaction text contains a relative date.
   */
  const relativeDays = relativeDaysFromText_(context);

  if (relativeDays !== null) {
    return formatDateInTimezone_(
      addDays_(today, relativeDays),
      tz
    );
  }

  /*
   * If there is no relative-date expression, use the
   * explicitly supplied date.
   */
  const explicitValue = clean_(value);

  if (!explicitValue) {
    return todayString;
  }

  const explicitNormalized =
    normalizeText_(explicitValue);

  if (explicitNormalized === "today") {
    return todayString;
  }

  if (explicitNormalized === "yesterday") {
    return formatDateInTimezone_(
      addDays_(today, -1),
      tz
    );
  }

  if (
    explicitNormalized === "day before yesterday" ||
    explicitNormalized === "day_before_yesterday" ||
    explicitNormalized === "two days ago" ||
    explicitNormalized === "2 days ago"
  ) {
    return formatDateInTimezone_(
      addDays_(today, -2),
      tz
    );
  }

  

  const agoMatch = explicitNormalized.match(
    /^(\d+)\s+days?\s+ago$/
  );

  if (agoMatch) {
    return formatDateInTimezone_(
      addDays_(today, -Number(agoMatch[1])),
      tz
    );
  }

  const explicitDate =
    parseDateOnly_(explicitValue);

  if (explicitDate) {
    return formatDateInTimezone_(
      explicitDate,
      tz
    );
  }

  // Invalid/missing date → today.
  return todayString;
}

function relativeDaysFromText_(text) {
  const t = normalizeText_(text);

  if (!t) return null;

  // Most specific phrases first.
  if (
    /\bday before yesterday\b/.test(t) ||
    /\b2 days ago\b/.test(t) ||
    /\btwo days ago\b/.test(t)
  ) {
    return -2;
  }

  const wordNumbers = {
    one: 1,
    two: 2,
    three: 3,
    four: 4,
    five: 5,
    six: 6,
    seven: 7,
    eight: 8,
    nine: 9,
    ten: 10,
    eleven: 11,
    twelve: 12,
    thirteen: 13,
    fourteen: 14,
    fifteen: 15,
    sixteen: 16,
    seventeen: 17,
    eighteen: 18,
    nineteen: 19,
    twenty: 20
  };

  const wordMatch = t.match(
    /\b(one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty)\s+days?\s+ago\b/
  );

  if (wordMatch) {
    return -wordNumbers[wordMatch[1]];
  }

  const numberMatch = t.match(
    /\b(\d+)\s+days?\s+(?:ago|back)\b/
  );

  if (numberMatch) {
    return -Number(numberMatch[1]);
  }

  if (/\byesterday\b/.test(t)) return -1;
  if (/\btomorrow\b/.test(t)) return 1;
  if (/\btoday\b/.test(t)) return 0;

  return null;
}

function monthsBetween_(startStr, endStr) {
  const start = parseDateOnly_(startStr);
  const end = parseDateOnly_(endStr);

  if (!start || !end) return [];

  const output = [];

  let cursor = new Date(
    start.getFullYear(),
    start.getMonth(),
    1
  );

  const stop = new Date(
    end.getFullYear(),
    end.getMonth(),
    1
  );

  while (cursor <= stop) {
    output.push(
      Utilities.formatDate(
        cursor,
        Session.getScriptTimeZone() || "Asia/Kolkata",
        "yyyy-MM"
      )
    );

    cursor = new Date(
      cursor.getFullYear(),
      cursor.getMonth() + 1,
      1
    );
  }

  return output;
}

/* =========================================================
 * CLASSIFICATION ENGINE
 * ========================================================= */

function looksLikeCreditCardBill_(text) {
  const t = normalizeText_(text);

  const hasCard = /\b(?:[a-z]+\s+)*(?:cc|card|credit card)\b/.test(t);

  const billWords =
    /\b(?:bill|bill payment|credit card bill|card bill|card payment)\b/.test(t);

  const paymentAction =
    /\b(?:pay|paid|payment|paying|clear|cleared|settle|settled|repay|repaid)\b/.test(t);

  const purchaseContext =
    /\b(?:using|with|on|via|through|by)\s+(?:[a-z]+\s+)*(?:cc|card|credit card)\b/.test(t);

  // Explicit "pay X CC bill" is a bill payment.
  if (billWords && hasCard) return true;

  // "pay ICICI CC using UPI"
  if (
    paymentAction &&
    hasCard &&
    !purchaseContext
  ) {
    return true;
  }

  return false;
}

function resolveCreditCardDestination_(hint, text, cfg) {
  const fromHint = resolveAccount_(
    hint,
    "",
    cfg
  );

  if (
    fromHint &&
    isCreditCardAccount_(fromHint, cfg)
  ) {
    return fromHint;
  }

  return resolveCreditCardFromText_(text, cfg);
}

function resolvePaymentMethodFromBillText_(text, cfg) {
  const t = normalizeText_(text);

  // Prefer explicit source/payment wording.
  const patterns = [
    /\b(?:using|via|through|from|with|by)\s+(.+)$/,
    /\busing\s+(.+?)(?:\s+(?:yesterday|today|tomorrow|day before yesterday|\d+\s+days?\s+ago))?$/
  ];

  for (let i = 0; i < patterns.length; i++) {
    const match = t.match(patterns[i]);

    if (match) {
      const result = resolvePaymentMode_(
        "",
        match[1],
        cfg
      );

      if (result) return result;
    }
  }

  if (/\bupi\b/.test(t)) {
    return canonicalPaymentMode_("UPI", cfg);
  }

  if (/\bcash\b/.test(t)) {
    return canonicalPaymentMode_("Cash", cfg);
  }

  return "";
}

function resolveCreditCardFromText_(text, cfg) {
  const cards = cfg.accounts.filter(a =>
    a.active &&
    /credit\s*card/i.test(
      a.type + " " + a.name
    )
  );

  return bestAliasMatch_(
    text,
    cards.map(a => ({
      canonical: a.name,
      keywords: a.keywords
    })),
    true
  );
}

function classifyInvestment_(text, categories) {
  const rows = categories.filter(c =>
    c.active &&
    same_(c.category, "Investments")
  );

  return bestCategoryHit_(text, rows);
}

function bestCategoryHit_(text, rows) {
  let best = null;
  let bestScore = -1;

  rows.forEach(row => {
    const terms = (row.keywords || [])
      .map(normalizeText_)
      .filter(Boolean);

    terms.forEach(term => {
      if (!textMatchesTerm_(text, term)) return;

      // Longer/specific terms win.
      // Exact keyword presence is strongly preferred.
      let score = term.length;

      if (text.indexOf(term) >= 0) {
        score += 100;
      }

      if (term.indexOf(" ") >= 0) {
        score += 25;
      }

      if (score > bestScore) {
        bestScore = score;
        best = {
          category: row.category,
          subcategory: row.subcategory
        };
      }
    });
  });

  return best;
}

function findMerchantRule_(merchant, text, rules) {
  let best = null;
  let bestScore = -1;

  const merchantText = normalizeText_(merchant);

  rules
    .filter(r => r.active)
    .forEach(rule => {
      const terms = [rule.merchant]
        .concat(rule.keywords || [])
        .map(normalizeText_)
        .filter(Boolean);

      terms.forEach(term => {
        const merchantHit =
          merchantText &&
          textMatchesTerm_(merchantText, term);

        const textHit =
          textMatchesTerm_(text, term);

        if (!merchantHit && !textHit) return;

        let score = term.length;

        if (merchantHit) score += 200;
        if (textHit) score += 50;

        if (score > bestScore) {
          bestScore = score;
          best = rule;
        }
      });
    });

  return best;
}

function resolvePaymentMode_(hint, text, cfg) {
  const h = normalizeText_(hint);

  if (h) {
    const exact = bestAliasMatch_(
      h,
      cfg.paymentModes
        .filter(p => p.active)
        .map(p => ({
          canonical: p.name,
          keywords: p.keywords
        })),
      false
    );

    if (exact) return exact;

    const accountAlias = bestAliasMatch_(
      h,
      cfg.accounts
        .filter(a => a.active)
        .map(a => ({
          canonical: a.paymentMode || a.name,
          keywords: a.keywords
        })),
      false
    );

    if (accountAlias) return accountAlias;
  }

  return (
    bestAliasMatch_(
      text,
      cfg.paymentModes
        .filter(p => p.active)
        .map(p => ({
          canonical: p.name,
          keywords: p.keywords
        })),
      true
    ) ||
    bestAliasMatch_(
      text,
      cfg.accounts
        .filter(a => a.active)
        .map(a => ({
          canonical: a.paymentMode || a.name,
          keywords: a.keywords
        })),
      true
    ) ||
    ""
  );
}

function resolveAccount_(hint, text, cfg) {
  const h = normalizeText_(hint);

  if (h) {
    const exact = bestAliasMatch_(
      h,
      cfg.accounts
        .filter(a => a.active)
        .map(a => ({
          canonical: a.name,
          keywords: a.keywords
        })),
      false
    );

    if (exact) return exact;
  }

  return bestAliasMatch_(
    text,
    cfg.accounts
      .filter(a => a.active)
      .map(a => ({
        canonical: a.name,
        keywords: a.keywords
      })),
    true
  ) || "";
}

function bestAliasMatch_(text, items, longText) {
  const t = normalizeText_(text);

  if (!t) return "";

  let best = "";
  let bestScore = -1;

  items.forEach(item => {
    const aliases = [item.canonical]
      .concat(item.keywords || [])
      .map(normalizeText_)
      .filter(Boolean);

    aliases.forEach(alias => {
      const hit = longText
        ? textMatchesTerm_(t, alias)
        : (
            t === alias ||
            t.indexOf(alias) >= 0 ||
            alias.indexOf(t) >= 0
          );

      if (!hit) return;

      let score = alias.length;

      if (t === alias) score += 10000;
      if (alias.indexOf(" ") >= 0) score += 100;

      if (score > bestScore) {
        bestScore = score;
        best = item.canonical;
      }
    });
  });

  return best;
}

/* =========================================================
 * CONFIGURATION
 * ========================================================= */

function getConfig_() {
  const cached =
    CacheService.getScriptCache().get(CACHE_KEY);

  if (cached) {
    return JSON.parse(cached);
  }

  const ss = getSpreadsheet_();

  const cfg = {
    accounts: readConfigRows_(
      ss.getSheetByName(CONFIG.sheets.accounts),
      ["Account", "Type", "Active", "Payment Mode", "Keywords"]
    ).map(r => ({
      name: clean_(r.Account),
      type: clean_(r.Type),
      active: isTrue_(r.Active),
      paymentMode: clean_(r["Payment Mode"]),
      keywords: splitKeywords_(r.Keywords)
    })),

    paymentModes: readConfigRows_(
      ss.getSheetByName(CONFIG.sheets.paymentModes),
      ["Payment Mode", "Type", "Account", "Keywords", "Active"]
    ).map(r => ({
      name: clean_(r["Payment Mode"]),
      type: clean_(r.Type),
      account: clean_(r.Account),
      keywords: splitKeywords_(r.Keywords),
      active: isTrue_(r.Active)
    })),

    categories: readConfigRows_(
      ss.getSheetByName(CONFIG.sheets.categories),
      ["Category", "Subcategory", "Keywords", "Active"]
    ).map(r => ({
      category: clean_(r.Category),
      subcategory: clean_(r.Subcategory),
      keywords: splitKeywords_(r.Keywords),
      active: isTrue_(r.Active)
    })),

    merchantRules: readConfigRows_(
      ss.getSheetByName(CONFIG.sheets.merchantRules),
      ["Merchant", "Category", "Subcategory", "Keywords", "Active"]
    ).map(r => ({
      merchant: clean_(r.Merchant),
      category: clean_(r.Category),
      subcategory: clean_(r.Subcategory),
      keywords: splitKeywords_(r.Keywords),
      active: isTrue_(r.Active)
    })),

    settings: readSettings_(
      ss.getSheetByName(CONFIG.sheets.settings)
    )
  };

  CacheService.getScriptCache().put(
    CACHE_KEY,
    JSON.stringify(cfg),
    CONFIG.cacheSeconds
  );

  return cfg;
}

function getConfiguration_() {
  const cfg = getConfig_();

  return {
    success: true,

    accounts: cfg.accounts
      .filter(a => a.active)
      .map(a => ({
        name: a.name,
        type: a.type,
        payment_mode: a.paymentMode,
        keywords: a.keywords
      })),

    payment_modes: cfg.paymentModes
      .filter(p => p.active)
      .map(p => ({
        name: p.name,
        type: p.type,
        account: p.account,
        keywords: p.keywords
      })),

    categories: cfg.categories
      .filter(c => c.active)
      .map(c => ({
        category: c.category,
        subcategory: c.subcategory,
        keywords: c.keywords
      }))
  };
}

function invalidateConfigCache_() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

/* =========================================================
 * MONTHLY SUMMARY
 * ========================================================= */

function updateMonthlySummaryIncremental_(transactions) {
  if (!transactions || !transactions.length) return;
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CONFIG.sheets.summary);
  if (!sh) return;

  const last = sh.getLastRow();
  const width = 6;
  const rows = last >= 2 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
  const rowMap = {};
  rows.forEach((r, i) => rowMap[String(r[0])] = i + 2);

  const deltas = {};
  transactions.forEach(t => {
    const m = String(t.month || t.date || "").slice(0, 7);
    if (!m) return;
    if (!deltas[m]) deltas[m] = {expenses:0, investments:0, cc:0, income:0};
    const amount = Number(t.amount || 0);
    if (same_(t.type, "Expense")) deltas[m].expenses += amount;
    if (same_(t.type, "Investment") || same_(t.category, "Investments")) deltas[m].investments += amount;
    if (same_(t.type, "Credit Card Payment")) deltas[m].cc += amount;
    if (same_(t.type, "Income") || same_(t.type, "Refund")) deltas[m].income += amount;
  });

  Object.keys(deltas).forEach(month => {
    const d = deltas[month];
    if (rowMap[month]) {
      const r = rowMap[month];
      const current = sh.getRange(r, 2, 1, 4).getValues()[0];
      sh.getRange(r, 2, 1, 5).setValues([[
        Number(current[0]||0)+d.expenses,
        Number(current[1]||0)+d.investments,
        Number(current[2]||0)+d.cc,
        Number(current[3]||0)+d.income,
        new Date()
      ]]);
    } else {
      sh.getRange(sh.getLastRow()+1, 1, 1, 6).setValues([[
        month, roundMoney_(d.expenses), roundMoney_(d.investments),
        roundMoney_(d.cc), roundMoney_(d.income), new Date()
      ]]);
    }
  });
}

function ensureBankBalancesMaterialized_() {
  const ss=getSpreadsheet_();
  const sh=ss.getSheetByName(LEDGERLY_BRIDGE.bankSheet)||setupFinanceBankBalances_();
  const values=sh.getDataRange().getValues();
  if(values.length<2)return sh;
  const headers=values[0].map(v=>String(v).trim());
  const before=headers.indexOf("Current Balance");
  const currentIdx=ensureBankCurrentBalanceColumn_(sh);
  if(before>=0)return sh;
  const accountIdx=headers.indexOf("Account"), openingIdx=headers.indexOf("Opening Balance"), adjustIdx=headers.indexOf("Manual Adjustment"), dateIdx=headers.indexOf("Opening Date"), activeIdx=headers.indexOf("Active");
  const txs=readAllFinanceTransactions_();
  values.slice(1).forEach((row,i)=>{
    const name=String(row[accountIdx]||"").trim();if(!name)return;
    const opening=Number(row[openingIdx]||0), adjustment=Number(row[adjustIdx]||0), openingDate=formatSheetDate_(row[dateIdx])||"1900-01-01";
    let movement=0;txs.forEach(t=>{if(t.date<openingDate)return;bankDeltaForTransaction_(t).forEach(pair=>{if(same_(pair[0],name))movement+=Number(pair[1]||0);});});
    sh.getRange(i+2,currentIdx+1).setValue(roundMoney_(opening+adjustment+movement));
  });
  return sh;
}

function ensureBankCurrentBalanceColumn_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v).trim());
  let idx = headers.indexOf("Current Balance");
  if (idx >= 0) return idx;
  idx = headers.length;
  sh.getRange(1, idx + 1).setValue("Current Balance");
  return idx;
}

function bankDeltaForTransaction_(t) {
  const amount = Number(t.amount || 0);
  if (!(amount > 0)) return [];
  if (same_(t.type, "Income") || same_(t.type, "Refund")) return [[t.account, amount]];
  if (same_(t.type, "Expense") || same_(t.type, "Investment") || same_(t.type, "Credit Card Payment")) return [[t.account, -amount]];
  if (same_(t.type, "Transfer")) return [[t.account, -amount], [t.toAccount, amount]];
  return [];
}

function applyBankBalanceDeltas_(transactions, direction) {
  if (!transactions || !transactions.length) return;
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(LEDGERLY_BRIDGE.bankSheet) || setupFinanceBankBalances_();
  const currentIdx = ensureBankCurrentBalanceColumn_(sh);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(v => String(v).trim());
  const accountIdx = headers.indexOf("Account");
  const rowMap = {};
  for (let i=1;i<values.length;i++) {
    const name=String(values[i][accountIdx]||"").trim().toLowerCase();
    if(name) rowMap[name]=i+1;
  }
  const deltas = {};
  transactions.forEach(t => bankDeltaForTransaction_(t).forEach(pair => {
    const name=String(pair[0]||"").trim().toLowerCase();
    if(name) deltas[name]=(deltas[name]||0)+Number(pair[1]||0)*Number(direction||1);
  }));
  Object.keys(deltas).forEach(name => {
    const row=rowMap[name];
    if(!row) return;
    const cell=sh.getRange(row,currentIdx+1);
    const current=Number(cell.getValue()||0);
    cell.setValue(roundMoney_(current+deltas[name]));
  });
}

function updateMonthlySummary_(months) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CONFIG.sheets.summary);

  if (!sh || !months || !months.length) return;

  const existing = {};

  if (sh.getLastRow() >= 2) {
    sh.getRange(
      2, 1, sh.getLastRow() - 1, 1
    ).getValues().forEach((r, i) => {
      existing[String(r[0])] = i + 2;
    });
  }

  months.forEach(month => {
    const txs = readTransactions_([month]);

    const expenses = sum_(
      txs.filter(t => same_(t.type, "Expense"))
    );

    const investments = sum_(
      txs.filter(t =>
        same_(t.type, "Investment") ||
        same_(t.category, "Investments")
      )
    );

    const ccPayments = sum_(
      txs.filter(t => same_(t.type, "Credit Card Payment"))
    );

    const income = sum_(
      txs.filter(t => same_(t.type, "Income"))
    );

    const row = [
      month,
      roundMoney_(expenses),
      roundMoney_(investments),
      roundMoney_(ccPayments),
      roundMoney_(income),
      new Date()
    ];

    if (existing[month]) {
      sh.getRange(
        existing[month],
        1,
        1,
        row.length
      ).setValues([row]);
    } else {
      sh.getRange(
        sh.getLastRow() + 1,
        1,
        1,
        row.length
      ).setValues([row]);
    }
  });
  refreshDashboardV84_();
}

/* =========================================================
 * HELPERS
 * ========================================================= */

function parseRequestBody_(e) {
  if (!e || !e.postData) return null;

  const raw = String(
    e.postData.contents || ""
  ).trim();

  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (err) {
    throw new Error("Invalid JSON request");
  }
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function ensureConfigSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);

  if (!sh) sh = ss.insertSheet(name);

  ensureHeaders_(sh, headers);
  sh.setFrozenRows(1);

  return sh;
}

function ensureHeaders_(sh, headers) {
  const lastColumn = sh.getLastColumn();

  if (!lastColumn) {
    sh.getRange(
      1, 1, 1, headers.length
    ).setValues([headers]);
    return;
  }

  const current = sh.getRange(
    1, 1, 1, lastColumn
  ).getValues()[0].map(v => String(v).trim());

  const missing = headers.filter(
    h => current.indexOf(h) < 0
  );

  if (missing.length) {
    sh.getRange(
      1,
      current.length + 1,
      1,
      missing.length
    ).setValues([missing]);
  }
}

function readConfigRows_(sh, headers) {
  if (!sh || sh.getLastRow() < 2) return [];

  const sheetHeaders = sh.getRange(
    1, 1, 1, sh.getLastColumn()
  ).getValues()[0].map(v => String(v).trim());

  const rows = sh.getRange(
    2, 1, sh.getLastRow() - 1, sh.getLastColumn()
  ).getValues();

  return rows
    .map(row => {
      const output = {};

      headers.forEach(header => {
        const index = sheetHeaders.indexOf(header);
        output[header] =
          index >= 0 ? row[index] : "";
      });

      return output;
    })
    .filter(row =>
      Object.keys(row).some(
        k => String(row[k]).trim() !== ""
      )
    );
}

function readSettings_(sh) {
  const output = {};

  readConfigRows_(
    sh,
    ["Setting", "Value"]
  ).forEach(row => {
    output[
      String(row.Setting).trim()
    ] = row.Value;
  });

  return output;
}

function seedIfEmpty_(sh, rows) {
  if (!sh || sh.getLastRow() > 1) return;

  sh.getRange(
    2, 1, rows.length, rows[0].length
  ).setValues(rows);
}

function mergeAliasesIntoSheet_(
  sh,
  keyHeader,
  keywordHeader,
  aliasMap
) {
  if (!sh || sh.getLastRow() < 2) return;

  const headers = sh.getRange(
    1, 1, 1, sh.getLastColumn()
  ).getValues()[0].map(String);

  const keyIndex = headers.indexOf(keyHeader);
  const keywordIndex = headers.indexOf(keywordHeader);

  if (keyIndex < 0 || keywordIndex < 0) return;

  const values = sh.getRange(
    2, 1, sh.getLastRow() - 1, sh.getLastColumn()
  ).getValues();

  values.forEach(row => {
    const name = String(
      row[keyIndex] || ""
    ).trim();

    if (!aliasMap[name]) return;

    const oldKeywords =
      splitKeywords_(row[keywordIndex]);

    row[keywordIndex] =
      Array.from(
        new Set(
          oldKeywords.concat(aliasMap[name])
        )
      ).join(",");
  });

  sh.getRange(
    2, 1, values.length, sh.getLastColumn()
  ).setValues(values);
}

function isConfigSheet_(name) {
  return [
    CONFIG.sheets.accounts,
    CONFIG.sheets.paymentModes,
    CONFIG.sheets.categories,
    CONFIG.sheets.merchantRules,
    CONFIG.sheets.settings,
    LEDGERLY_BRIDGE.bankSheet
  ].indexOf(name) >= 0;
}

function getHeaderMap_(sh) {
  const headers = sh.getRange(
    1, 1, 1, sh.getLastColumn()
  ).getValues()[0].map(v => String(v).trim());

  const map = {};
  headers.forEach((h, i) => map[h] = i);

  return map;
}

function setRowValue_(row, map, key, value) {
  if (map[key] !== undefined) {
    row[map[key]] = value;
  }
}

function getCell_(row, map, key) {
  return map[key] === undefined
    ? ""
    : row[map[key]];
}

function canonicalCategory_(value, cfg) {
  const v = clean_(value);
  if (!v) return "";

  const normalized = normalizeText_(v);

  const row = cfg.categories.find(c =>
    c.active &&
    normalizeText_(c.category) === normalized
  );

  return row ? row.category : v;
}

function canonicalSubcategory_(value, category, cfg) {
  const v = clean_(value);
  if (!v || !category) return "";

  const row = cfg.categories.find(c =>
    c.active &&
    same_(c.category, category) &&
    same_(c.subcategory, v)
  );

  return row ? row.subcategory : v;
}

function isConfiguredCategory_(value, cfg) {
  return cfg.categories.some(c =>
    c.active &&
    same_(c.category, value)
  );
}

function isConfiguredSubcategory_(
  category,
  subcategory,
  cfg
) {
  return cfg.categories.some(c =>
    c.active &&
    same_(c.category, category) &&
    same_(c.subcategory, subcategory)
  );
}

function isPaymentModeLike_(subcategory, cfg) {
  return cfg.paymentModes.some(p =>
    p.active &&
    same_(p.name, subcategory)
  );
}

function isCreditCardAccount_(name, cfg) {
  const account = cfg.accounts.find(
    a => same_(a.name, name)
  );

  return !!account &&
    /credit\s*card/i.test(
      account.type + " " + account.name
    );
}

function canonicalPaymentMode_(value, cfg) {
  return resolvePaymentMode_(
    value,
    "",
    cfg
  );
}

function accountForPaymentMode_(paymentMode, cfg) {
  const pm = cfg.paymentModes.find(
    p => same_(p.name, paymentMode)
  );

  return pm ? pm.account : "";
}

function textMatchesTerm_(text, term) {
  if (!term) return false;

  const t = normalizeText_(text);
  const x = normalizeText_(term);

  if (!x) return false;

  // Short terms need boundaries to avoid:
  // "mf" matching unrelated words.
  if (x.length <= 3) {
    const regex = new RegExp(
      "(^|\\s|[^a-z0-9])" +
      escapeRegex_(x) +
      "($|\\s|[^a-z0-9])",
      "i"
    );

    return regex.test(t);
  }

  return t.indexOf(x) >= 0;
}

function escapeRegex_(s) {
  return String(s).replace(
    /[.*+?^${}()|[\]\\]/g,
    "\\$&"
  );
}

function buildTransactionText_(x) {
  return [
    x.source_text,
    x.raw_text,
    x.original_text,
    x.merchant,
    x.remarks,
    x.category,
    x.subcategory,
    x.payment_mode,
    x.paymentMode,
    x.account,
    x.to_account,
    x.toAccount,
    x.type
  ]
    .filter(v =>
      v !== undefined &&
      v !== null &&
      String(v).trim() !== ""
    )
    .join(" ");
}

function buildDefaultRemarks_(
  text,
  category,
  subcategory,
  type,
  merchant,
  toAccount
) {
  if (
    same_(type, "Credit Card Payment") &&
    toAccount
  ) {
    return toAccount + " bill payment";
  }

  if (merchant && subcategory) {
    return subcategory + " at " + merchant;
  }

  if (subcategory) return subcategory;
  if (merchant) return merchant;
  return category || "";
}

function normalizeType_(value) {
  const s = normalizeText_(value);

  if (!s) return "";

  if (
    s === "investment" ||
    s === "invest" ||
    s === "investments"
  ) {
    return "Investment";
  }

  if (s === "income") return "Income";

  if (
    s.indexOf("credit card payment") >= 0 ||
    s.indexOf("credit card bill") >= 0
  ) {
    return "Credit Card Payment";
  }

  return "Expense";
}

function parseAmount_(value) {
  if (typeof value === "number") {
    return isFinite(value) ? value : 0;
  }

  const s = String(value || "")
    .replace(/[₹,\s]/g, "")
    .replace(/rs\.?/gi, "")
    .replace(/inr/gi, "");

  const n = Number(s);

  return isFinite(n) ? n : 0;
}

function clean_(value) {
  return (
    value === null ||
    value === undefined
  )
    ? ""
    : String(value).trim();
}

function normalizeText_(value) {
  return clean_(value)
    .toLowerCase()
    .replace(/[’']/g, "'")
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function splitKeywords_(value) {
  return clean_(value)
    .split(",")
    .map(normalizeText_)
    .filter(Boolean);
}

function isTrue_(value) {
  return [
    true,
    "true",
    "1",
    "yes",
    "y"
  ].indexOf(
    typeof value === "string"
      ? normalizeText_(value)
      : value
  ) >= 0;
}

function same_(a, b) {
  return normalizeText_(a) ===
    normalizeText_(b);
}

function sum_(array) {
  return array.reduce(
    (total, item) =>
      total + Number(item.amount || 0),
    0
  );
}

function roundMoney_(value) {
  return Math.round(
    Number(value) * 100
  ) / 100;
}

function getLimit_(value, defaultValue) {
  const n = Number(value);

  if (!isFinite(n) || n <= 0) {
    return defaultValue;
  }

  return Math.min(
    Math.floor(n),
    100
  );
}

function transactionTimestamp_(tx) {
  const date = parseDateOnly_(tx.date);

  if (!date) return 0;

  const parts = clean_(tx.time)
    .split(":")
    .map(Number);

  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    parts[0] || 0,
    parts[1] || 0,
    parts[2] || 0
  ).getTime();
}

function withinDateRange_(date, start, end) {
  return (
    isValidDateString_(date) &&
    date >= start &&
    date <= end
  );
}

function simplifyTransaction_(tx) {
  return {
    id: tx.id,
    date: tx.date,
    time: tx.time,
    type: tx.type,
    category: tx.category,
    subcategory: tx.subcategory,
    amount: Number(tx.amount || 0),
    payment_mode: tx.paymentMode,
    account: tx.account,
    merchant: tx.merchant,
    remarks: tx.remarks,
    to_account: tx.toAccount
  };
}

function generateId_() {
  return (
    "TXN-" +
    Date.now() +
    "-" +
    Math.floor(Math.random() * 1000)
  );
}

function addDays_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function dateString_(date) {
  const tz =
    getSpreadsheet_().getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    "Asia/Kolkata";

  return Utilities.formatDate(
    date,
    tz,
    "yyyy-MM-dd"
  );
}

function parseDateOnly_(value) {
  if (!value) return null;

  if (
    Object.prototype.toString.call(value) ===
      "[object Date]" &&
    !isNaN(value)
  ) {
    return new Date(
      value.getFullYear(),
      value.getMonth(),
      value.getDate()
    );
  }

  const s = String(value).trim();

  let match = s.match(
    /^(\d{4})-(\d{1,2})-(\d{1,2})$/
  );

  if (match) {
    return new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3])
    );
  }

  match = s.match(
    /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/
  );

  if (match) {
    return new Date(
      Number(match[3]),
      Number(match[2]) - 1,
      Number(match[1])
    );
  }

  const d = new Date(s);

  return isNaN(d)
    ? null
    : new Date(
        d.getFullYear(),
        d.getMonth(),
        d.getDate()
      );
}

function isValidDateString_(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(
    clean_(value)
  ) && !!parseDateOnly_(value);
}

function formatSheetDate_(value) {
  const d = parseDateOnly_(value);
  return d ? dateString_(d) : clean_(value);
}

function formatSheetTime_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      getSpreadsheet_().getSpreadsheetTimeZone() ||
        "Asia/Kolkata",
      "HH:mm:ss"
    );
  }

  return clean_(value);
}

function resolveTime_(value, now) {
  if (clean_(value)) return clean_(value);

  return Utilities.formatDate(
    now,
    getSpreadsheet_().getSpreadsheetTimeZone() ||
      "Asia/Kolkata",
    "HH:mm:ss"
  );
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(
      JSON.stringify(obj)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}

function safeErrorMessage_(err) {
  return String(
    err && err.message
      ? err.message
      : err || "Unknown error"
  ).slice(0, 500);
}






/* =========================================================
 * PUBLIC FUNCTIONS
 * ========================================================= */

function setupDashboardV84() {
  refreshDashboardV84_();
  installDashboardEditTriggerV84_();
  return "Dashboard V8.4 created/refreshed successfully. Month selector trigger installed.";
}

function refreshDashboardV84() {
  refreshDashboardV84_();
  return "Dashboard V8.4 refreshed successfully.";
}

/**
 * Use this function from an onEdit trigger if you want the charts
 * to change immediately when the month selector is changed.
 *
 * It is safe to use as a simple trigger because it only reacts
 * to the Dashboard!B2 selector.
 */
function onDashboardEditV84(e) {
  if (!e || !e.range) return;

  const range = e.range;

  if (
    range.getSheet().getName() !== "Dashboard" ||
    range.getA1Notation() !== "B2"
  ) {
    return;
  }

  refreshDashboardV84_();
}

/**
 * Installs the spreadsheet on-edit trigger required for the
 * month selector.
 *
 * A named function such as onDashboardEditV84 is NOT a native
 * simple trigger, so without this installable trigger changing
 * B2 would only change the cell and not rebuild the charts.
 */
function installDashboardEditTriggerV84_() {
  const ss = getSpreadsheet_();
  const triggers = ScriptApp.getProjectTriggers();

  const alreadyInstalled =
    triggers.some(
      trigger =>
        trigger.getHandlerFunction() ===
        "onDashboardEditV84"
    );

  if (!alreadyInstalled) {
    ScriptApp.newTrigger(
      "onDashboardEditV84"
    )
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  }
}

/**
 * Run this manually if the trigger was deleted:
 *
 *   installDashboardEditTriggerV84()
 */
function installDashboardEditTriggerV84() {
  installDashboardEditTriggerV84_();
  return "Dashboard month-selector trigger installed.";
}

/* =========================================================
 * MAIN DASHBOARD
 * ========================================================= */

function refreshDashboardV84_() {
  const ss = getSpreadsheet_();

  let dashboard = ss.getSheetByName("Dashboard");

  if (!dashboard) {
    dashboard = ss.insertSheet("Dashboard", 0);
  }

  let data = ss.getSheetByName("Dashboard Data");

  if (!data) {
    data = ss.insertSheet("Dashboard Data");
  }

  // Keep data visible while chart source ranges are rebuilt.
  data.showSheet();

  // Find actual monthly transaction tabs.
  const sheetMonths =
    ss.getSheets()
      .map(s => s.getName())
      .filter(n => /^\d{4}-\d{2}$/.test(n))
      .sort();

  const transactions =
    sheetMonths.length
      ? readTransactions_(sheetMonths)
      : [];

  // Only months containing real transaction records are selectable.
  const monthsWithData =
    dashboardMonthsWithDataV84_(transactions);

  // Determine selected month from Dashboard!B2.
  const existingSelection =
    String(
      dashboard.getRange("B2").getDisplayValue() || ""
    ).trim();

  let selectedMonth = "";

  if (
    existingSelection &&
    monthsWithData.indexOf(existingSelection) >= 0
  ) {
    selectedMonth = existingSelection;
  } else if (monthsWithData.length) {
    // Default to latest month with actual transaction data.
    selectedMonth =
      monthsWithData[monthsWithData.length - 1];
  }

  // Transactions for selected month.
  const selectedTransactions =
    selectedMonth
      ? transactions.filter(
          t =>
            dashboardMonthV84_(t) ===
            selectedMonth
        )
      : [];

  const selectedExpenses =
    selectedTransactions.filter(
      t =>
        dashboardTypeV84_(t) === "Expense"
    );

  const selectedInvestments =
    selectedTransactions.filter(
      t =>
        dashboardTypeV84_(t) === "Investment" ||
        dashboardTextV84_(t.category) === "investments"
    );

  const selectedIncome =
    selectedTransactions.filter(
      t =>
        dashboardTypeV84_(t) === "Income"
    );

  const currentSpending =
    dashboardSumV84_(selectedExpenses);

  const currentInvestments =
    dashboardSumV84_(selectedInvestments);

  const currentIncome =
    dashboardSumV84_(selectedIncome);

  // Previous calendar month for MoM.
  const previousMonth =
    selectedMonth
      ? dashboardPreviousMonthV84_(selectedMonth)
      : "";

  const previousTransactions =
    previousMonth
      ? transactions.filter(
          t =>
            dashboardMonthV84_(t) ===
            previousMonth
        )
      : [];

  const previousSpending =
    dashboardSumV84_(
      previousTransactions.filter(
        t =>
          dashboardTypeV84_(t) === "Expense"
      )
    );

  const mom =
    previousSpending > 0
      ? (
          (currentSpending - previousSpending) /
          previousSpending
        ) * 100
      : null;

  // Selected-month breakdowns.
  const categoryMap =
    dashboardAggregateV84_(
      selectedExpenses,
      "category"
    );

  const paymentMap =
    dashboardAggregateV84_(
      selectedExpenses,
      "paymentMode"
    );

  const investmentMap =
    dashboardAggregateV84_(
      selectedInvestments,
      "subcategory"
    );

  const merchantMap =
    dashboardAggregateV84_(
      selectedExpenses,
      "merchant"
    );

  // -------------------------------------------------------
  // BUILD CHART DATA
  // -------------------------------------------------------

  data.clear();
  data.getCharts().forEach(
    c => data.removeChart(c)
  );
  data.setHiddenGridlines(true);

  // A. Spending trend up to selected month.
  const selectedIndex =
    monthsWithData.indexOf(selectedMonth);

  const trendMonths =
    selectedIndex >= 0
      ? monthsWithData.slice(
          0,
          selectedIndex + 1
        )
      : [];

  data.getRange("A1:B1")
    .setValues([
      ["Month", "Spending"]
    ]);

  const trendRows =
    trendMonths.map(month => {
      const mt =
        transactions.filter(
          t =>
            dashboardMonthV84_(t) ===
            month
        );

      return [
        month,
        dashboardRoundV84_(
          dashboardSumV84_(
            mt.filter(
              t =>
                dashboardTypeV84_(t) === "Expense"
            )
          )
        )
      ];
    });

  if (trendRows.length) {
    data.getRange(
      2,
      1,
      trendRows.length,
      2
    ).setValues(trendRows);
  }

  // B. Monthly income/spending/investments up to selected month.
  data.getRange("D1:G1")
    .setValues([
      [
        "Month",
        "Spending",
        "Investments",
        "Income"
      ]
    ]);

  const comparisonRows =
    trendMonths.map(month => {
      const mt =
        transactions.filter(
          t =>
            dashboardMonthV84_(t) ===
            month
        );

      return [
        month,
        dashboardRoundV84_(
          dashboardSumV84_(
            mt.filter(
              t =>
                dashboardTypeV84_(t) === "Expense"
            )
          )
        ),
        dashboardRoundV84_(
          dashboardSumV84_(
            mt.filter(
              t =>
                dashboardTypeV84_(t) === "Investment" ||
                dashboardTextV84_(t.category) === "investments"
            )
          )
        ),
        dashboardRoundV84_(
          dashboardSumV84_(
            mt.filter(
              t =>
                dashboardTypeV84_(t) === "Income"
            )
          )
        )
      ];
    });

  if (comparisonRows.length) {
    data.getRange(
      2,
      4,
      comparisonRows.length,
      4
    ).setValues(comparisonRows);
  }

  // C. Selected-month category breakdown.
  data.getRange("I1:J1")
    .setValues([
      ["Category", "Amount"]
    ]);

  const categoryRows =
    dashboardSortedRowsV84_(categoryMap);

  if (categoryRows.length) {
    data.getRange(
      2,
      9,
      categoryRows.length,
      2
    ).setValues(categoryRows);
  }

  // D. Selected-month payment mode.
  data.getRange("L1:M1")
    .setValues([
      ["Payment Mode", "Amount"]
    ]);

  const paymentRows =
    dashboardSortedRowsV84_(paymentMap);

  if (paymentRows.length) {
    data.getRange(
      2,
      12,
      paymentRows.length,
      2
    ).setValues(paymentRows);
  }

  // E. Selected-month investment mix.
  data.getRange("A20:B20")
    .setValues([
      ["Investment Type", "Amount"]
    ]);

  const investmentRows =
    dashboardSortedRowsV84_(investmentMap);

  if (investmentRows.length) {
    data.getRange(
      21,
      1,
      investmentRows.length,
      2
    ).setValues(investmentRows);
  }

  // F. Selected-month top merchants.
  data.getRange("D20:E20")
    .setValues([
      ["Merchant", "Amount"]
    ]);

  const merchantRows =
    dashboardSortedRowsV84_(merchantMap)
      .filter(
        row => row[0] !== ""
      )
      .slice(0, 10);

  if (merchantRows.length) {
    data.getRange(
      21,
      4,
      merchantRows.length,
      2
    ).setValues(merchantRows);
  }

  // -------------------------------------------------------
  // DATA FORMATTING
  // -------------------------------------------------------

  if (trendRows.length) {
    data.getRange(
      2,
      2,
      trendRows.length,
      1
    ).setNumberFormat("₹#,##0");
  }

  if (comparisonRows.length) {
    data.getRange(
      2,
      5,
      comparisonRows.length,
      3
    ).setNumberFormat("₹#,##0");
  }

  if (categoryRows.length) {
    data.getRange(
      2,
      10,
      categoryRows.length,
      1
    ).setNumberFormat("₹#,##0");
  }

  if (paymentRows.length) {
    data.getRange(
      2,
      13,
      paymentRows.length,
      1
    ).setNumberFormat("₹#,##0");
  }

  if (investmentRows.length) {
    data.getRange(
      21,
      2,
      investmentRows.length,
      1
    ).setNumberFormat("₹#,##0");
  }

  if (merchantRows.length) {
    data.getRange(
      21,
      5,
      merchantRows.length,
      1
    ).setNumberFormat("₹#,##0");
  }

  // -------------------------------------------------------
  // REBUILD VISIBLE DASHBOARD
  // -------------------------------------------------------

  dashboard.clear();
  dashboard.getCharts()
    .forEach(
      c => dashboard.removeChart(c)
    );

  dashboard.clearConditionalFormatRules();
  dashboard.setHiddenGridlines(true);
  dashboard.setFrozenRows(2);

  for (let c = 1; c <= 12; c++) {
    dashboard.setColumnWidth(
      c,
      125
    );
  }

  dashboard.setColumnWidth(
    1,
    150
  );

  // Title.
  dashboard.getRange("A1:L1")
    .merge();

  dashboard.getRange("A1")
    .setValue(
      "PERSONAL FINANCE DASHBOARD"
    )
    .setFontSize(20)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  dashboard.setRowHeight(
    1,
    40
  );

  // Month selector label.
  dashboard.getRange("A2")
    .setValue("SELECT MONTH")
    .setFontWeight("bold")
    .setHorizontalAlignment("right");

  // Selector itself.
  dashboard.getRange("B2")
    .setValue(
      selectedMonth || ""
    )
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setFontSize(12);

  if (monthsWithData.length) {
    const validation =
      SpreadsheetApp.newDataValidation()
        .requireValueInList(
          monthsWithData,
          true
        )
        .setAllowInvalid(false)
        .build();

    dashboard.getRange("B2")
      .setDataValidation(
        validation
      );
  }

  dashboard.getRange("C2:L2")
    .merge();

  dashboard.getRange("C2")
    .setValue(
      selectedMonth
        ? "Showing financial analysis for " +
          selectedMonth
        : "No transaction data available"
    )
    .setHorizontalAlignment("center")
    .setFontSize(11);

  // -------------------------------------------------------
  // FOUR KPI CARDS
  // -------------------------------------------------------

  dashboardCardV84_(
    dashboard,
    "A4:C4",
    "A5:C6",
    "SPENDING",
    currentSpending,
    true
  );

  dashboardCardV84_(
    dashboard,
    "D4:F4",
    "D5:F6",
    "INVESTMENTS",
    currentInvestments,
    true
  );

  dashboardCardV84_(
    dashboard,
    "G4:I4",
    "G5:I6",
    "INCOME",
    currentIncome,
    true
  );

  dashboardCardV84_(
    dashboard,
    "J4:L4",
    "J5:L6",
    "MOM SPENDING",
    mom === null
      ? "N/A"
      : (mom > 0 ? "+" : "") +
        dashboardRoundV84_(mom) +
        "%",
    false
  );

  dashboard.getRange("A8:L8")
    .merge();

  dashboard.getRange("A8")
    .setValue(
      selectedMonth
        ? "Selected month: " +
          selectedMonth +
          " • Bill payments excluded from spending"
        : "Select a month after transactions are added"
    )
    .setHorizontalAlignment("center")
    .setFontSize(10);

  // -------------------------------------------------------
  // CHART 1 — SPENDING TREND
  // -------------------------------------------------------

  if (trendRows.length) {
    const chart =
      dashboard.newChart()
        .setChartType(
          Charts.ChartType.LINE
        )
        .addRange(
          data.getRange(
            1,
            1,
            trendRows.length + 1,
            2
          )
        )
        .setPosition(
          10,
          1,
          0,
          0
        )
        .setOption(
          "title",
          "Spending Trend Through Selected Month"
        )
        .setOption(
          "legend",
          {position: "none"}
        )
        .setOption(
          "height",
          320
        )
        .setOption(
          "width",
          650
        )
        .build();

    dashboard.insertChart(
      chart
    );
  }

  // -------------------------------------------------------
  // CHART 2 — CATEGORY
  // -------------------------------------------------------

  if (categoryRows.length) {
    const chart =
      dashboard.newChart()
        .setChartType(
          Charts.ChartType.PIE
        )
        .addRange(
          data.getRange(
            1,
            9,
            categoryRows.length + 1,
            2
          )
        )
        .setPosition(
          10,
          7,
          0,
          0
        )
        .setOption(
          "title",
          "Spending by Category • " +
            selectedMonth
        )
        .setOption(
          "pieHole",
          0.4
        )
        .setOption(
          "height",
          320
        )
        .setOption(
          "width",
          650
        )
        .build();

    dashboard.insertChart(
      chart
    );
  }

  // -------------------------------------------------------
  // CHART 3 — INCOME VS SPENDING VS INVESTMENTS
  // -------------------------------------------------------

  if (comparisonRows.length) {
    const chart =
      dashboard.newChart()
        .setChartType(
          Charts.ChartType.COLUMN
        )
        .addRange(
          data.getRange(
            1,
            4,
            comparisonRows.length + 1,
            4
          )
        )
        .setPosition(
          27,
          1,
          0,
          0
        )
        .setOption(
          "title",
          "Income vs Spending vs Investments"
        )
        .setOption(
          "height",
          320
        )
        .setOption(
          "width",
          650
        )
        .build();

    dashboard.insertChart(
      chart
    );
  }

  // -------------------------------------------------------
  // CHART 4 — PAYMENT MODE
  // -------------------------------------------------------

  if (paymentRows.length) {
    const chart =
      dashboard.newChart()
        .setChartType(
          Charts.ChartType.BAR
        )
        .addRange(
          data.getRange(
            1,
            12,
            paymentRows.length + 1,
            2
          )
        )
        .setPosition(
          27,
          7,
          0,
          0
        )
        .setOption(
          "title",
          "Spending by Payment Mode • " +
            selectedMonth
        )
        .setOption(
          "legend",
          {position: "none"}
        )
        .setOption(
          "height",
          320
        )
        .setOption(
          "width",
          650
        )
        .build();

    dashboard.insertChart(
      chart
    );
  }

  // -------------------------------------------------------
  // CHART 5 — INVESTMENT MIX
  // -------------------------------------------------------

  if (investmentRows.length) {
    const chart =
      dashboard.newChart()
        .setChartType(
          Charts.ChartType.PIE
        )
        .addRange(
          data.getRange(
            20,
            1,
            investmentRows.length + 1,
            2
          )
        )
        .setPosition(
          44,
          1,
          0,
          0
        )
        .setOption(
          "title",
          "Investment Mix • " +
            selectedMonth
        )
        .setOption(
          "pieHole",
          0.4
        )
        .setOption(
          "height",
          320
        )
        .setOption(
          "width",
          650
        )
        .build();

    dashboard.insertChart(
      chart
    );
  }

  // -------------------------------------------------------
  // CHART 6 — TOP MERCHANTS
  // -------------------------------------------------------

  if (merchantRows.length) {
    const chart =
      dashboard.newChart()
        .setChartType(
          Charts.ChartType.BAR
        )
        .addRange(
          data.getRange(
            20,
            4,
            merchantRows.length + 1,
            2
          )
        )
        .setPosition(
          44,
          7,
          0,
          0
        )
        .setOption(
          "title",
          "Top Merchants • " +
            selectedMonth
        )
        .setOption(
          "legend",
          {position: "none"}
        )
        .setOption(
          "height",
          320
        )
        .setOption(
          "width",
          650
        )
        .build();

    dashboard.insertChart(
      chart
    );
  }

  dashboard.getRange("A61:L61")
    .merge();

  dashboard.getRange("A61")
    .setValue(
      "Select another month from the dropdown to update the dashboard"
    )
    .setHorizontalAlignment("center")
    .setFontSize(9);

  // Hide data source only after charts exist.
  data.hideSheet();

  SpreadsheetApp.flush();
}

/* =========================================================
 * HELPERS
 * ========================================================= */

function dashboardMonthsWithDataV84_(transactions) {
  const found = {};

  transactions.forEach(t => {
    const month =
      dashboardMonthV84_(t);

    if (!month) return;

    if (
      t.amount === "" ||
      t.amount === null ||
      typeof t.amount === "undefined"
    ) {
      return;
    }

    const amount =
      Number(t.amount);

    if (!isFinite(amount)) return;

    found[month] = true;
  });

  return Object.keys(found)
    .sort();
}

function dashboardMonthV84_(t) {
  const month =
    String(t.month || "").trim();

  if (/^\d{4}-\d{2}$/.test(month)) {
    return month;
  }

  if (!t.date) return "";

  const d =
    t.date instanceof Date
      ? t.date
      : new Date(t.date);

  if (isNaN(d.getTime())) {
    return "";
  }

  return Utilities.formatDate(
    d,
    getSpreadsheet_().getSpreadsheetTimeZone() ||
      "Asia/Kolkata",
    "yyyy-MM"
  );
}

function dashboardTypeV84_(t) {
  return String(
    t.type || ""
  ).trim();
}

function dashboardTextV84_(value) {
  return String(
    value || ""
  )
    .trim()
    .toLowerCase();
}

function dashboardSumV84_(rows) {
  return rows.reduce(
    (sum, t) =>
      sum +
      Number(t.amount || 0),
    0
  );
}

function dashboardAggregateV84_(
  rows,
  field
) {
  const map = {};

  rows.forEach(t => {
    const key =
      String(t[field] || "")
        .trim() ||
      "Other";

    map[key] =
      (map[key] || 0) +
      Number(t.amount || 0);
  });

  return map;
}

function dashboardSortedRowsV84_(map) {
  return Object.keys(map)
    .sort(
      (a, b) =>
        Number(map[b]) -
        Number(map[a])
    )
    .map(k => [
      k,
      dashboardRoundV84_(map[k])
    ]);
}

function dashboardRoundV84_(value) {
  return Math.round(
    Number(value) * 100
  ) / 100;
}

function dashboardPreviousMonthV84_(
  month
) {
  const parts =
    month.split("-");

  const d =
    new Date(
      Number(parts[0]),
      Number(parts[1]) - 2,
      1
    );

  return Utilities.formatDate(
    d,
    getSpreadsheet_().getSpreadsheetTimeZone() ||
      "Asia/Kolkata",
    "yyyy-MM"
  );
}

function dashboardCardV84_(
  sh,
  labelRange,
  valueRange,
  label,
  value,
  currency
) {
  sh.getRange(
    labelRange
  ).merge();

  sh.getRange(
    labelRange.split(":")[0]
  )
    .setValue(label)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sh.getRange(
    valueRange
  ).merge();

  sh.getRange(
    valueRange.split(":")[0]
  )
    .setValue(value)
    .setFontSize(17)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  if (currency) {
    sh.getRange(
      valueRange.split(":")[0]
    ).setNumberFormat(
      "₹#,##0"
    );
  }
}

/* =========================================================
 * ISOLATED LEDGERLY BRIDGE
 *
 * Ledgerly is a separate Apps Script project/database. It never reads
 * this spreadsheet directly. It calls these small server-to-server APIs.
 * ========================================================= */

const LEDGERLY_BRIDGE = {
  secretProperty: "LEDGERLY_BRIDGE_SECRET",
  legacySecretProperty: "1nw8QpT85epAlmAmtNf3yLZ8JrM5uOb3LqHSTgukcf2Q",
  bankSheet: "BankBalances"
};

function setLedgerlyBridgeSecret(secret) {
  secret = String(secret || "").trim();
  if (!secret || secret.length < 24) {
    throw new Error("Use a random secret of at least 24 characters.");
  }
  PropertiesService.getScriptProperties().setProperty(
    LEDGERLY_BRIDGE.secretProperty,
    secret
  );
  return "Ledgerly bridge secret saved.";
}

function getLedgerlyBridgeSecret_() {
  const p=PropertiesService.getScriptProperties();
  return String(
    p.getProperty(LEDGERLY_BRIDGE.secretProperty) ||
    p.getProperty(LEDGERLY_BRIDGE.legacySecretProperty) ||
    ""
  );
}

function requireLedgerlyBridgeAuth_(body) {
  const expected = getLedgerlyBridgeSecret_();
  const supplied = String(body && body.apiKey || "").trim();
  if (!expected) throw new Error("Ledgerly bridge secret is not configured in Finance Assistant.");
  if (!supplied || supplied !== expected) throw new Error("Unauthorized Ledgerly bridge request.");
}

function handleLedgerlyBridgePost_(body) {
  const action = String(body && body.action || "");
  const allowed = [
    "ledgerlyDashboardData",
    "ledgerlyCashFlow",
    "ledgerlyDeleteFinanceTransaction",
    "ledgerlyUpdateBankAccount"
  ];

  if (allowed.indexOf(action) < 0) return null;

  requireLedgerlyBridgeAuth_(body);

  if (action === "ledgerlyDashboardData") {
    return getLedgerlyDashboardData_(body);
  }

  if (action === "ledgerlyCashFlow") {
    return addLedgerlyCashFlow_(body);
  }

  if (action === "ledgerlyDeleteFinanceTransaction") {
    return { success: deleteFinanceTransactionById_(body.financeTransactionId) };
  }

  if (action === "ledgerlyUpdateBankAccount") {
    return updateLedgerlyBankAccount_(body);
  }

  return null;
}

function ensureFinanceBankAccountConfig_(ss, account, paymentMode, keywords) {
  const sh = ss.getSheetByName(CONFIG.sheets.accounts);
  if (!sh) return;
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(v => String(v).trim());
  const idx = {}; headers.forEach((h, i) => idx[h] = i);
  for (let i = 1; i < values.length; i++) {
    if (same_(values[i][idx["Account"]], account)) return;
  }
  sh.appendRow([account, "Bank Account", "TRUE", paymentMode || "", keywords || ""]);
}

function setupFinanceBankBalances_() {
  const ss = getSpreadsheet_();
  ensureFinanceBankAccountConfig_(ss, "SBI Savings Account", "UPI", "sbi savings,sbi savings account,sbi bank");
  ensureFinanceBankAccountConfig_(ss, "HSBC Savings Account", "", "hsbc,hsbc savings,hsbc savings account");
  const sh = ensureConfigSheet_(ss, LEDGERLY_BRIDGE.bankSheet, [
    "Account", "Opening Balance", "Opening Date", "Manual Adjustment", "Active", "Notes"
  ]);

  const existing = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach((r, i) => {
      const name = String(r[0] || "").trim();
      if (name) existing[name.toLowerCase()] = i + 2;
    });
  }

  // Keep these as account-picker/configuration rows only. They do not create money.
  ["SBI Savings Account", "HSBC Savings Account"].forEach(name => {
    if (!existing[name.toLowerCase()]) {
      sh.appendRow([name, 0, new Date(), 0, true, "Set the opening balance once; transactions then drive the balance."]);
    }
  });

  return sh;
}

function getFinanceBankAccounts_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(LEDGERLY_PERF.bankKey);
  if (cached) { try { return JSON.parse(cached); } catch (_) {} }

  const sh = ensureBankBalancesMaterialized_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return {accounts: [], total: 0};
  const headers = values[0].map(v=>String(v).trim());
  const idx={}; headers.forEach((h,i)=>idx[h]=i);
  const currentIdx=idx["Current Balance"];
  const accounts=[];
  values.slice(1).forEach(row=>{
    const name=String(row[idx["Account"]]||"").trim();
    if(!name||String(row[idx["Active"]]).toUpperCase()==="FALSE")return;
    accounts.push({account:name,name:name,type:String(row[idx["Type"]]||"Bank Account"),balance:roundMoney_(Number(row[currentIdx]||0)),openingBalance:Number(row[idx["Opening Balance"]]||0),manualAdjustment:Number(row[idx["Manual Adjustment"]]||0)});
  });
  const result={accounts,total:roundMoney_(accounts.reduce((s,a)=>s+a.balance,0))};
  try{cache.put(LEDGERLY_PERF.bankKey,JSON.stringify(result),LEDGERLY_PERF.bankCacheSeconds);}catch(_){}
  return result;
}

function readAllFinanceTransactions_() {
  const ss = getSpreadsheet_();
  const months = ss.getSheets().map(s => s.getName()).filter(n => /^\d{4}-\d{2}$/.test(n));
  return readTransactions_(months);
}

function getFinanceCurrentMonthSummary_() {
  const now = new Date();
  const month = Utilities.formatDate(now, getSpreadsheet_().getSpreadsheetTimeZone() || "Asia/Kolkata", "yyyy-MM");
  const txs = readTransactions_([month]);
  const expenses = txs.filter(t => same_(t.type, "Expense")).reduce((s, t) => s + Number(t.amount || 0), 0);
  const investments = txs.filter(t => same_(t.type, "Investment") || same_(t.category, "Investments")).reduce((s, t) => s + Number(t.amount || 0), 0);
  const income = txs.filter(t => same_(t.type, "Income") || same_(t.type, "Refund")).reduce((s, t) => s + Number(t.amount || 0), 0);
  return { month, spending: roundMoney_(expenses), investment: roundMoney_(investments), income: roundMoney_(income) };
}



function addLedgerlyCashFlow_(body) {
  const type = String(body.financeType || "").trim();
  const account = String(body.account || "").trim();
  const amount = Number(body.amount);
  const date = String(body.date || "").trim();
  const symbol = String(body.symbol || "").trim();
  const assetType = String(body.assetType || "STOCK").trim().toUpperCase();

  if (!(amount > 0) || !date || !account || !symbol) {
    throw new Error("Ledgerly cash-flow request requires positive amount, date, account and symbol.");
  }
  const cfg = getConfig_();
  const accountConfig = cfg.accounts.find(a => same_(a.name, account));
  if (!accountConfig || !same_(accountConfig.type, "Bank Account")) {
    throw new Error("Ledgerly funding account is not configured as an active Finance Assistant bank account: " + account);
  }
  if (!["Investment", "Income"].includes(type)) {
    throw new Error("Ledgerly cash flow type must be Investment or Income.");
  }

  const result = addTransactions_([{
    date,
    amount,
    type,
    category: "Investments",
    subcategory: assetType === "MF" ? "Mutual Fund" : "Stocks",
    account,
    payment_mode: "",
    merchant: symbol,
    remarks: String(body.remarks || "Ledgerly investment cash flow"),
    tags: "Ledgerly",
    source: "Ledgerly"
  }], "Ledgerly");

  const added = result && result.results ? result.results.find(r => r.status === "added") : null;
  if (!added) {
    throw new Error(result && result.message ? result.message : "Finance transaction was not created.");
  }

  return { success: true, financeTransactionId: added.id };
}

function deleteFinanceTransactionById_(id) {
  id = String(id || "").trim();
  if (!id) throw new Error("Finance transaction ID is required.");
  const ss = getSpreadsheet_();
  const monthSheets = ss.getSheets().map(s => s.getName()).filter(n => /^\d{4}-\d{2}$/.test(n));

  for (const name of monthSheets) {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) continue;
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v).trim());
    const idx={}; headers.forEach((h,i)=>idx[h]=i);
    const values=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
    for(let i=0;i<values.length;i++){
      if(String(values[i][idx["ID"]])!==id)continue;
      const r=values[i];
      const tx={
        id, date:formatSheetDate_(r[idx["Date"]]), month:clean_(r[idx["Month"]])||name,
        type:clean_(r[idx["Type"]]), category:clean_(r[idx["Category"]]), amount:parseAmount_(r[idx["Amount"]]),
        account:clean_(r[idx["Account"]]), toAccount:clean_(r[idx["To Account"]])
      };
      sh.deleteRow(i+2);
      invalidateFinanceMonthCache_(name);
      applyBankBalanceDeltas_([tx],-1);
      updateMonthlySummaryIncremental_([Object.assign({},tx,{amount:-Number(tx.amount||0)})]);
      invalidateFinanceAnalyticsCache_();
      return true;
    }
  }
  return false;
}

function updateLedgerlyBankAccount_(body) {
  const oldName=String(body.account||body.accountName||body.id||body.accountId||"").trim();
  const newName=String(body.name||body.accountName||oldName).trim();
  if(!oldName||!newName)throw new Error("Bank account name is required.");
  const ss=getSpreadsheet_();
  const sh=ss.getSheetByName(LEDGERLY_BRIDGE.bankSheet)||setupFinanceBankBalances_();
  const currentIdx=ensureBankCurrentBalanceColumn_(sh);
  const values=sh.getDataRange().getValues();
  const headers=values[0].map(v=>String(v).trim()); const idx={}; headers.forEach((h,i)=>idx[h]=i);
  let rowNo=0;
  for(let i=1;i<values.length;i++)if(same_(values[i][idx["Account"]],oldName)){rowNo=i+1;break;}
  if(!rowNo)throw new Error("Bank account not found: "+oldName);
  for(let i=1;i<values.length;i++)if(i+1!==rowNo&&same_(values[i][idx["Account"]],newName))throw new Error("An account with this name already exists: "+newName);

  const row=sh.getRange(rowNo,1,1,sh.getLastColumn()).getValues()[0];
  const opening=Number(body.openingBalance!==undefined?body.openingBalance:row[idx["Opening Balance"]]||0);
  const openingDate=String(body.openingDate||formatSheetDate_(row[idx["Opening Date"]])||"").trim();
  const active=body.active===undefined?String(row[idx["Active"]]).toUpperCase()!=="FALSE":body.active!==false;
  const notes=body.notes!==undefined?String(body.notes||""):String(row[idx["Notes"]]||"");
  const type=String(body.type||"Bank Account").trim();

  const txs=readAllFinanceTransactions_();
  let movement=0;
  txs.forEach(t=>{
    if(openingDate&&t.date<openingDate)return;
    bankDeltaForTransaction_(t).forEach(pair=>{if(same_(pair[0],oldName))movement+=Number(pair[1]||0);});
  });
  const currentBalance=Number(body.currentBalance!==undefined?body.currentBalance:body.balance!==undefined?body.balance:row[currentIdx]||0);
  const manualAdjustment=roundMoney_(currentBalance-opening-movement);

  row[idx["Account"]]=newName;
  if(idx["Opening Balance"]>=0)row[idx["Opening Balance"]]=opening;
  if(idx["Opening Date"]>=0)row[idx["Opening Date"]]=openingDate||row[idx["Opening Date"]];
  if(idx["Manual Adjustment"]>=0)row[idx["Manual Adjustment"]]=manualAdjustment;
  if(idx["Active"]>=0)row[idx["Active"]]=active;
  if(idx["Notes"]>=0)row[idx["Notes"]]=notes;
  row[currentIdx]=currentBalance;
  sh.getRange(rowNo,1,1,sh.getLastColumn()).setValues([row]);

  // Keep historical transaction account references consistent after a rename.
  if(!same_(oldName,newName)){
    const months=ss.getSheets().map(s=>s.getName()).filter(n=>/^\\d{4}-\\d{2}$/.test(n));
    months.forEach(month=>{
      const txsh=ss.getSheetByName(month); if(!txsh||txsh.getLastRow()<2)return;
      const hdr=txsh.getRange(1,1,1,txsh.getLastColumn()).getValues()[0].map(v=>String(v).trim());
      const a=hdr.indexOf("Account"),to=hdr.indexOf("To Account");
      if(a<0)return;
      const vals=txsh.getRange(2,1,txsh.getLastRow()-1,txsh.getLastColumn()).getValues(); let changed=false;
      vals.forEach(r=>{if(same_(r[a],oldName)){r[a]=newName;changed=true;}if(to>=0&&same_(r[to],oldName)){r[to]=newName;changed=true;}});
      if(changed)txsh.getRange(2,1,vals.length,txsh.getLastColumn()).setValues(vals);
      invalidateFinanceMonthCache_(month);
    });
  }
  invalidateFinanceAnalyticsCache_();
  return {success:true,account:{account:newName,name:newName,type,balance:roundMoney_(currentBalance),openingBalance:opening,manualAdjustment,openingDate,active,notes},bankTotal:getFinanceBankAccounts_().total};
}


function getLedgerlyDashboardData_(payload) {
  payload = payload || {};
  const period = String(payload.period || "this-month");
  const range = resolveLedgerlyPeriod_(period);
  const selectedMonths = monthsBetween_(range.start, range.end);
  const txs = readTransactions_(selectedMonths).filter(t => withinDateRange_(t.date, range.start, range.end));
  const expenses=txs.filter(t=>same_(t.type,"Expense"));
  const income=txs.filter(t=>same_(t.type,"Income")||same_(t.type,"Refund"));
  const investments=txs.filter(t=>same_(t.type,"Investment")||same_(t.category,"Investments"));
  const ccPayments=txs.filter(t=>same_(t.type,"Credit Card Payment"));
  const sum=rows=>roundMoney_(rows.reduce((s,t)=>s+Number(t.amount||0),0));
  const incomeTotal=sum(income), spendingTotal=sum(expenses), investmentTotal=sum(investments), ccPaymentTotal=sum(ccPayments);

  // Reuse the single selected-period read for all selected-period analytics.
  const grouped={};
  txs.forEach(t=>{const m=String(t.month||t.date||"").slice(0,7);if(m)(grouped[m]||(grouped[m]=[])).push(t);});
  const monthlyFromMap=months=>months.map(month=>{
    const rows=grouped[month]||[];
    const inc=sum(rows.filter(t=>same_(t.type,"Income")||same_(t.type,"Refund")));
    const sp=sum(rows.filter(t=>same_(t.type,"Expense")));
    const inv=sum(rows.filter(t=>same_(t.type,"Investment")||same_(t.category,"Investments")));
    return {month,income:inc,spending:sp,investments:inv,netCashFlow:roundMoney_(inc-sp-inv)};
  });
  const recent=txs.slice().sort((a,b)=>String(b.date).localeCompare(String(a.date))||String(b.time||"").localeCompare(String(a.time||""))).slice(0,100);
  const largest=expenses.slice().sort((a,b)=>Number(b.amount||0)-Number(a.amount||0)).slice(0,10);
  const bank=getFinanceBankAccounts_();
  return {
    connected:true, period:{key:period,start:range.start,end:range.end}, income:incomeTotal, spending:spendingTotal,
    investments:investmentTotal, creditCardPayments:ccPaymentTotal,
    netSavings:roundMoney_(incomeTotal-spendingTotal), cashFlowAfterInvestments:roundMoney_(incomeTotal-spendingTotal-investmentTotal),
    savingsRate:incomeTotal?roundMoney_((incomeTotal-spendingTotal)/incomeTotal*100):0,
    transactionCount:txs.length, financeTransactions:txs,
    bankAccounts:bank.accounts, bankTotal:bank.total,
    categorySpending:ledgerlyGroupExpenses_(expenses,"category","amount"),
    merchantSpending:ledgerlyGroupExpenses_(expenses,"merchant","amount"),
    paymentModeSpending:ledgerlyGroupExpenses_(expenses,"paymentMode","amount"),
    monthly:monthlyFromMap(selectedMonths), monthlyTrend:ledgerlyMonthlyTrend_(),
    recentTransactions:recent, largestExpenses:largest, lastUpdatedAt:new Date().toISOString()
  };
}

function ledgerlyMonthlyTrend_(){
  const ss=getSpreadsheet_();
  const sh=ss.getSheetByName(CONFIG.sheets.summary);
  if(!sh||sh.getLastRow()<2){
    const tz=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone()||"Asia/Kolkata",today=new Date(),months=[];
    for(let i=11;i>=0;i--){const d=new Date(today.getFullYear(),today.getMonth()-i,1);months.push(Utilities.formatDate(d,tz,"yyyy-MM"));}
    return ledgerlyMonthlyBreakdown_(months);
  }
  const rows=sh.getRange(2,1,sh.getLastRow()-1,6).getValues();
  const map={};
  rows.forEach(r=>{const month=String(r[0]||"").slice(0,7);if(month)map[month]={month,spending:Number(r[1]||0),investments:Number(r[2]||0),income:Number(r[4]||0)};});
  const tz=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone()||"Asia/Kolkata",today=new Date(),out=[];
  for(let i=11;i>=0;i--){const d=new Date(today.getFullYear(),today.getMonth()-i,1),month=Utilities.formatDate(d,tz,"yyyy-MM"),r=map[month]||{month,spending:0,investments:0,income:0};out.push({month,income:roundMoney_(r.income),spending:roundMoney_(r.spending),investments:roundMoney_(r.investments),netCashFlow:roundMoney_(r.income-r.spending-r.investments)});}
  return out;
}

function resolveLedgerlyPeriod_(period) {
  const today = new Date();
  const end = new Date(
    today.getFullYear(),
    today.getMonth(),
    today.getDate()
  );

  let start;

  switch (String(period || "this-month")) {
    case "all":
    case "all-time":
      start = new Date(2000, 0, 1);
      break;

    case "last-month":
      start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      end.setTime(
        new Date(today.getFullYear(), today.getMonth(), 0).getTime()
      );
      break;

    case "last-3-months":
      start = new Date(today.getFullYear(), today.getMonth() - 2, 1);
      break;

    case "last-6-months":
      start = new Date(today.getFullYear(), today.getMonth() - 5, 1);
      break;

    case "this-year":
      start = new Date(today.getFullYear(), 0, 1);
      break;

    case "this-month":
    default:
      start = new Date(today.getFullYear(), today.getMonth(), 1);
      break;
  }

  return {
    start: dateString_(start),
    end: dateString_(end)
  };
}

function ledgerlyGroupExpenses_(rows, key, valueKey) {
  const totals = {};

  rows.forEach(t => {
    const name = String(t[key] || "Uncategorized").trim() || "Uncategorized";
    totals[name] = (totals[name] || 0) + Number(t[valueKey] || 0);
  });

  return Object.entries(totals)
    .map(([name, amount]) => ({
      [key === "paymentMode" ? "paymentMode" : key]: name,
      amount: roundMoney_(amount)
    }))
    .sort((a, b) => b.amount - a.amount);
}

function ledgerlyMonthlyBreakdown_(months) {
  return months.map(month => {
    const txs = readTransactions_([month]);

    const income = txs
      .filter(t => same_(t.type, "Income"))
      .reduce((s, t) => s + Number(t.amount || 0), 0);

    const spending = txs
      .filter(t => same_(t.type, "Expense"))
      .reduce((s, t) => s + Number(t.amount || 0), 0);

    const investments = txs
      .filter(t =>
        same_(t.type, "Investment") ||
        same_(t.category, "Investments")
      )
      .reduce((s, t) => s + Number(t.amount || 0), 0);

    return {
      month,
      income: roundMoney_(income),
      spending: roundMoney_(spending),
      investments: roundMoney_(investments),
      netCashFlow: roundMoney_(
        income - spending - investments
      )
    };
  });
}

/**
 * Returns the latest 12 calendar months for trend charts.
 * This is intentionally independent of the selected dashboard period.
 */
function ledgerlyMonthlyTrend_() {
  const tz =
    getSpreadsheet_().getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    "Asia/Kolkata";

  const today = new Date();
  const months = [];

  for (let i = 11; i >= 0; i--) {
    const d = new Date(
      today.getFullYear(),
      today.getMonth() - i,
      1
    );

    months.push(
      Utilities.formatDate(d, tz, "yyyy-MM")
    );
  }

  return ledgerlyMonthlyBreakdown_(months);
}

/*
 * Reads the durable BankBalances sheet created by the isolated
 * Finance Assistant setup.
 *
 * Expected headers:
 *   Account
 *   Opening Balance
 *   Current Balance
 *
 * If Current Balance is not present, this function falls back to
 * Opening Balance. The later balance-cache optimization can update
 * Current Balance without changing the frontend contract.
 */





/* =========================================================
 * LEDGERLY ANALYTICS TEST
 * ========================================================= */

function testLedgerlyDashboardConnection() {
  const result = getLedgerlyDashboardData_({
    period: "this-month"
  });
  Logger.log(JSON.stringify(result, null, 2));
}
