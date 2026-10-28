const TZ = 'Europe/London';
const SH = { SETTINGS: 'Settings', AD: 'Ad Spend (Manual)', DAILY: 'Daily Data', DASH: 'Dashboard' };


// Settings tab cells (column B)
const ROW = { START: 2, SHOP: 3, PATH: 4, TAG: 5, API: 6, REFRESH: 7, ERROR: 8 };


const DEFAULTS = {
  START: '2026-10-25',
  SHOP: '4b0f4b-2.myshopify.com',
  PATH: '/products/the-phoenix-gravity-water-filter',
  TAG: 'appstle_subscription_recurring_order',
  API: '2026-07'
};


/* ---------------------------------------------------------------- menu */


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Black Friday')
    .addItem('1. Set up sheets', 'setupSheets')
    .addItem('2. Save Shopify keys', 'saveShopifyKeys')
    .addItem('3. Refresh now', 'refreshData')
    .addItem('4. Start 15-min auto-refresh', 'installTrigger')
    .addSeparator()
    .addItem('Stop auto-refresh', 'removeTriggers')
    .addToUi();
}


function saveShopifyKeys() {
  const ui = SpreadsheetApp.getUi();
  const id = ui.prompt('Shopify Client ID', 'Paste the Client ID (stored in Script Properties only):', ui.ButtonSet.OK_CANCEL);
  if (id.getSelectedButton() !== ui.Button.OK || !id.getResponseText().trim()) return;
  const secret = ui.prompt('Shopify Client Secret', 'Paste the Client Secret (stored in Script Properties only):', ui.ButtonSet.OK_CANCEL);
  if (secret.getSelectedButton() !== ui.Button.OK || !secret.getResponseText().trim()) return;
  PropertiesService.getScriptProperties().setProperties({
    SHOPIFY_CLIENT_ID: id.getResponseText().trim(),
    SHOPIFY_CLIENT_SECRET: secret.getResponseText().trim()
  });
  CacheService.getScriptCache().remove('shopify_token');
  ui.alert('Keys saved. Now run "3. Refresh now".');
}


function installTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('refreshData').timeBased().everyMinutes(15).create();
  SpreadsheetApp.getActive().toast('Auto-refresh every 15 minutes is on.', 'Black Friday', 5);
}


function removeTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshData')
    .forEach(t => ScriptApp.deleteTrigger(t));
}


/* ---------------------------------------------------------------- setup */


function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);


  const dash = getOrCreate_(ss, SH.DASH);
  const daily = getOrCreate_(ss, SH.DAILY);
  const ad = getOrCreate_(ss, SH.AD);
  const settings = getOrCreate_(ss, SH.SETTINGS);


  // ---- Settings
  settings.getRange('A1:C1').setValues([['Setting', 'Value', 'Notes']]).setFontWeight('bold');
  const labels = [
    ['Sale start date', null, 'First day included in every view.'],
    ['Shopify store domain', DEFAULTS.SHOP, 'The .myshopify.com domain.'],
    ['Product page path', DEFAULTS.PATH, 'Landing page used for the product page conversion rate.'],
    ['Recurring order tag', DEFAULTS.TAG, 'Appstle subscription renewals carry this order tag.'],
    ['Shopify API version', DEFAULTS.API, 'Change if Shopify retires this version.'],
    ['Last refresh', '', 'Written by the script (UK time).'],
    ['Last error', '', 'Blank when the last refresh succeeded.']
  ];
  labels.forEach((r, i) => {
    const row = i + 2;
    settings.getRange(row, 1).setValue(r[0]);
    settings.getRange(row, 3).setValue(r[2]);
    const cell = settings.getRange(row, 2);
    if (cell.isBlank() && r[1] !== null && r[1] !== '') cell.setValue(r[1]);
  });
  if (settings.getRange(ROW.START, 2).isBlank()) {
    settings.getRange(ROW.START, 2).setValue(parseDay_(DEFAULTS.START));
  }
  settings.getRange(ROW.START, 2).setNumberFormat('dd mmm yyyy');
  settings.getRange(ROW.REFRESH, 2).setNumberFormat('dd mmm yyyy hh:mm:ss');
  settings.setColumnWidth(1, 190).setColumnWidth(2, 330).setColumnWidth(3, 420);


  // ---- Ad Spend (Manual)
  ad.getRange('A1:C1').setValues([['Date', 'Meta Ads spend (£)', 'Google Ads spend (£)']])
    .setFontWeight('bold').setBackground('#e8eaf6');
  ad.setFrozenRows(1);
  const startStr = fmt_(readSettings_().start);
  ensureAdSpendRows_(ad, addDays_(startStr, 44));
  ad.getRange('A2:A').setNumberFormat('ddd dd mmm yyyy');
  ad.getRange('B2:C').setNumberFormat('£#,##0.00');
  ad.setColumnWidth(1, 150).setColumnWidth(2, 150).setColumnWidth(3, 160);


  // ---- Daily Data
  const dailyHeaders = [
    'Date', 'Total sales', 'Orders', 'Recurring sales', 'Recurring orders',
    'Sessions', 'Sessions completed checkout', 'Product page sessions', 'Product page completed checkout',
    'Sales excl. recurring', 'Orders excl. recurring', 'Meta Ads spend', 'Google Ads spend', 'Total ad spend',
    'Blended ROAS', 'Conversion rate (excl. recurring)', 'Shopify tracked conv. rate', 'Product page conv. rate'
  ];
  daily.getRange(1, 1, 1, dailyHeaders.length).setValues([dailyHeaders])
    .setFontWeight('bold').setBackground('#e8eaf6').setWrap(true);
  daily.setFrozenRows(1);
  daily.getRange('A2:A').setNumberFormat('ddd dd mmm yyyy');
  ['B2:B', 'D2:D', 'J2:J', 'L2:N'].forEach(a => daily.getRange(a).setNumberFormat('£#,##0.00'));
  ['C2:C', 'E2:I', 'K2:K'].forEach(a => daily.getRange(a).setNumberFormat('#,##0'));
  daily.getRange('O2:O').setNumberFormat('0.00');
  daily.getRange('P2:R').setNumberFormat('0.00%');
  daily.setColumnWidths(1, dailyHeaders.length, 120);
  daily.setColumnWidth(1, 150);


  buildDashboard_(dash);


  ss.setActiveSheet(dash);
  ss.moveActiveSheet(1);
  SpreadsheetApp.getActive().toast('Sheets are set up. Next: Save Shopify keys.', 'Black Friday', 6);
}


function buildDashboard_(sh) {
  sh.clear();
  const headers = [
    'Period', 'From', 'To', 'Total sales', 'Sales excl. recurring', 'Ad spend (Meta + Google)',
    'Blended ROAS', 'Sessions', 'Orders excl. recurring', 'Conversion rate (excl. recurring)',
    'Shopify tracked conv. rate', 'Product page sessions', 'Product page conv. rate'
  ];


  sh.getRange('A1').setValue('Black Friday Live Dashboard: Phoenix Gravity Water Filters UK')
    .setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setFormula('="Last refresh: "&IF(Settings!B7="","never",TEXT(Settings!B7,"ddd dd mmm yyyy hh:mm"))&IF(Settings!B8="",""," | ERROR: "&Settings!B8)');
  sh.getRange('A3').setValue('Conversion rate = non-recurring orders ÷ store sessions. Shopify tracked rate only counts sessions from visitors who accepted cookies, so it runs lower.')
    .setFontStyle('italic').setFontColor('#5f6368');


  const S = 'Settings!$B$2';
  const blocks = [
    {
      title: 'Current periods', top: 5,
      rows: [
        ['Today', '=TODAY()', '=TODAY()'],
        ['Yesterday', '=TODAY()-1', '=TODAY()-1'],
        ['This week (Mon-Sun)', '=TODAY()-WEEKDAY(TODAY(),3)', '=B9+6'],
        ['This month', '=DATE(YEAR(TODAY()),MONTH(TODAY()),1)', '=EOMONTH(B10,0)'],
        ['Sale to date', '=' + S, '=TODAY()']
      ]
    },
    {
      title: 'Weekly (weeks start Monday)', top: 13,
      rows: Array.from({ length: 10 }, (_, i) => [
        null, `=${S}-WEEKDAY(${S},3)+${7 * i}`, `=B${15 + i}+6`
      ])
    },
    {
      title: 'Monthly', top: 27,
      rows: Array.from({ length: 4 }, (_, i) => [
        null, `=DATE(YEAR(${S}),MONTH(${S})+${i},1)`, `=EOMONTH(B${29 + i},0)`
      ])
    }
  ];


  blocks.forEach(b => {
    sh.getRange(b.top, 1).setValue(b.title).setFontWeight('bold').setFontSize(12);
    const hRow = b.top + 1;
    sh.getRange(hRow, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#e8eaf6').setWrap(true).setVerticalAlignment('middle');
    b.rows.forEach((r, i) => {
      const row = hRow + 1 + i;
      const label = r[0] !== null ? r[0]
        : (b.title.indexOf('Weekly') === 0 ? `="Week of "&TEXT(B${row},"d mmm")` : `=TEXT(B${row},"mmmm yyyy")`);
      if (label.charAt(0) === '=') sh.getRange(row, 1).setFormula(label);
      else sh.getRange(row, 1).setValue(label);
      sh.getRange(row, 2).setFormula(r[1]);
      sh.getRange(row, 3).setFormula(r[2]);
      sh.getRange(row, 4, 1, 10).setFormulas([metricFormulas_(row)]);
    });
    const n = b.rows.length;
    const first = hRow + 1;
    sh.getRange(first, 2, n, 2).setNumberFormat('dd mmm yyyy');
    sh.getRange(first, 4, n, 3).setNumberFormat('£#,##0.00');
    sh.getRange(first, 7, n, 1).setNumberFormat('0.00');
    sh.getRange(first, 8, n, 2).setNumberFormat('#,##0');
    sh.getRange(first, 10, n, 2).setNumberFormat('0.00%');
    sh.getRange(first, 12, n, 1).setNumberFormat('#,##0');
    sh.getRange(first, 13, n, 1).setNumberFormat('0.00%');
    sh.getRange(first, 1, n, 1).setFontWeight('bold');
  });


  sh.setColumnWidth(1, 170);
  sh.setColumnWidths(2, 2, 105);
  sh.setColumnWidths(4, 10, 130);
  sh.setFrozenRows(3);
}


// Returns formulas for columns D..M of a dashboard row (row holds From in B, To in C).
function metricFormulas_(r) {
  const rng = col =>
    `SUMIFS('${SH.DAILY}'!$${col}:$${col},'${SH.DAILY}'!$A:$A,">="&$B${r},'${SH.DAILY}'!$A:$A,"<="&$C${r})`;
  return [
    '=' + rng('B'),                                   // D Total sales
    '=' + rng('J'),                                   // E Sales excl. recurring
    '=' + rng('N'),                                   // F Ad spend
    `=IF(F${r}>0,E${r}/F${r},"")`,                  // G Blended ROAS
    '=' + rng('F'),                                   // H Sessions
    '=' + rng('K'),                                   // I Orders excl. recurring
    `=IF(H${r}>0,I${r}/H${r},"")`,                  // J Conversion rate
    `=IF(H${r}>0,${rng('G')}/H${r},"")`,            // K Shopify tracked conv. rate
    '=' + rng('H'),                                   // L Product page sessions
    `=IF(L${r}>0,${rng('I')}/L${r},"")`             // M Product page conv. rate
  ];
}


/* ---------------------------------------------------------------- refresh */


function refreshData() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  const ss = SpreadsheetApp.getActive();
  const settingsSheet = ss.getSheetByName(SH.SETTINGS);
  try {
    const cfg = readSettings_();
    const startStr = fmt_(cfg.start);
    const todayStr = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    if (todayStr < startStr) {
      settingsSheet.getRange(ROW.ERROR, 2).setValue('Sale has not started yet (starts ' + startStr + ').');
      return;
    }


    const since = `SINCE ${startStr} UNTIL today`;
    const tag = cfg.tag.replace(/'/g, "\\'");
    const path = cfg.path.replace(/'/g, "\\'");


    const total = byDay_(shopifyql_(cfg, `FROM sales SHOW total_sales, orders TIMESERIES day ${since}`));
    const recurring = byDay_(shopifyql_(cfg,
      `FROM sales SHOW total_sales, orders WHERE order_tags CONTAINS '${tag}' TIMESERIES day ${since}`));
    const sessions = byDay_(shopifyql_(cfg,
      `FROM sessions SHOW sessions, sessions_that_completed_checkout TIMESERIES day ${since}`));
    const product = byDay_(shopifyql_(cfg,
      `FROM sessions SHOW sessions, sessions_that_completed_checkout WHERE landing_page_path = '${path}' TIMESERIES day ${since}`));


    const days = dayList_(startStr, todayStr);
    const num = (m, d, k) => (m[d] && isFinite(parseFloat(m[d][k]))) ? parseFloat(m[d][k]) : 0;


    const values = days.map(d => [
      parseDay_(d),
      num(total, d, 'total_sales'), num(total, d, 'orders'),
      num(recurring, d, 'total_sales'), num(recurring, d, 'orders'),
      num(sessions, d, 'sessions'), num(sessions, d, 'sessions_that_completed_checkout'),
      num(product, d, 'sessions'), num(product, d, 'sessions_that_completed_checkout')
    ]);


    const adSheet = ss.getSheetByName(SH.AD);
    ensureAdSpendRows_(adSheet, todayStr);


    const daily = ss.getSheetByName(SH.DAILY);
    const n = values.length;
    daily.getRange(2, 1, Math.max(daily.getMaxRows() - 1, 1), 18).clearContent();
    daily.getRange(2, 1, n, 9).setValues(values);
    const formulas = [];
    for (let i = 0; i < n; i++) {
      const r = i + 2;
      formulas.push([
        `=B${r}-D${r}`,                                                          // J sales excl. recurring
        `=C${r}-E${r}`,                                                          // K orders excl. recurring
        `=SUMIFS('${SH.AD}'!$B:$B,'${SH.AD}'!$A:$A,$A${r})`,                    // L Meta
        `=SUMIFS('${SH.AD}'!$C:$C,'${SH.AD}'!$A:$A,$A${r})`,                    // M Google
        `=L${r}+M${r}`,                                                          // N total ad spend
        `=IF(N${r}>0,J${r}/N${r},"")`,                                         // O blended ROAS
        `=IF(F${r}>0,K${r}/F${r},"")`,                                         // P conversion (excl. recurring)
        `=IF(F${r}>0,G${r}/F${r},"")`,                                         // Q Shopify tracked
        `=IF(H${r}>0,I${r}/H${r},"")`                                          // R product page
      ]);
    }
    daily.getRange(2, 10, n, 9).setFormulas(formulas);


    settingsSheet.getRange(ROW.REFRESH, 2).setValue(new Date());
    settingsSheet.getRange(ROW.ERROR, 2).setValue('');
  } catch (e) {
    settingsSheet.getRange(ROW.ERROR, 2).setValue(Utilities.formatDate(new Date(), TZ, 'dd MMM HH:mm') + ': ' + e.message);
    console.error(e);
  } finally {
    lock.releaseLock();
  }
}


/* ---------------------------------------------------------------- Shopify */


function getToken_(cfg, forceNew) {
  const cache = CacheService.getScriptCache();
  if (!forceNew) {
    const cached = cache.get('shopify_token');
    if (cached) return cached;
  }
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHOPIFY_CLIENT_ID');
  const secret = props.getProperty('SHOPIFY_CLIENT_SECRET');
  if (!id || !secret) throw new Error('Shopify keys missing. Use Black Friday > Save Shopify keys.');


  const res = UrlFetchApp.fetch(
    `https://${cfg.shop}/admin/oauth/access_token`,
    {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      headers: {
        Accept: 'application/json'
      },
      payload: {
        grant_type: 'client_credentials',
        client_id: id,
        client_secret: secret
      },
      muteHttpExceptions: true
    }
  );
  if (res.getResponseCode() !== 200) {
    throw new Error(
      'Shopify token request failed (' +
      res.getResponseCode() +
      '): ' +
      res.getContentText().slice(0, 1000)
    );
  }
  const body = JSON.parse(res.getContentText());
  if (!body.access_token) throw new Error('Shopify token response had no access_token.');
  cache.put('shopify_token', body.access_token, 60 * 60 * 5); // token lasts ~24h; refresh well before
  return body.access_token;
}


// Runs a ShopifyQL query and returns an array of row objects keyed by column name.
function shopifyql_(cfg, query, isRetry) {
  const token = getToken_(cfg, isRetry);
  const gql = 'query($q:String!){shopifyqlQuery(query:$q){tableData{columns{name dataType displayName} rows} parseErrors}}';
  const res = UrlFetchApp.fetch(`https://${cfg.shop}/admin/api/${cfg.api}/graphql.json`, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-Shopify-Access-Token': token },
    payload: JSON.stringify({ query: gql, variables: { q: query } }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code === 401 && !isRetry) return shopifyql_(cfg, query, true);
  if (code === 429 || code >= 500) {
    if (!isRetry) { Utilities.sleep(3000); return shopifyql_(cfg, query, true); }
  }
  if (code !== 200) throw new Error('Shopify API ' + code + ': ' + res.getContentText().slice(0, 200));


  const body = JSON.parse(res.getContentText());
  if (body.errors && body.errors.length) {
    throw new Error('Shopify GraphQL error: ' + body.errors.map(e => e.message).join('; '));
  }
  const out = body.data && body.data.shopifyqlQuery;
  if (!out) throw new Error('Empty ShopifyQL response.');
  if (out.parseErrors && out.parseErrors.length) {
    const msg = out.parseErrors.map(e => (typeof e === 'string' ? e : (e.message || JSON.stringify(e)))).join('; ');
    throw new Error('ShopifyQL error: ' + msg + ' [query: ' + query + ']');
  }
  if (!out.tableData) return [];


  const cols = out.tableData.columns.map(c => c.name);
  let rows = out.tableData.rows;
  if (typeof rows === 'string') rows = JSON.parse(rows);
  return (rows || []).map(r => {
    if (Array.isArray(r)) {
      const o = {};
      cols.forEach((c, i) => { o[c] = r[i]; });
      return o;
    }
    return r;
  });
}


function byDay_(rows) {
  const m = {};
  rows.forEach(r => {
    const d = String(r.day || r.date || '').substring(0, 10);
    if (d) m[d] = r;
  });
  return m;
}


/* ---------------------------------------------------------------- helpers */


function readSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.SETTINGS);
  if (!sh) throw new Error('Run Black Friday > Set up sheets first.');
  const v = sh.getRange(2, 2, 7, 1).getValues().map(r => r[0]);
  const start = v[0] instanceof Date ? v[0] : parseDay_(String(v[0]));
  return {
    start: start,
    shop: String(v[1]).trim(),
    path: String(v[2]).trim(),
    tag: String(v[3]).trim(),
    api: String(v[4]).trim()
  };
}


function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}


function fmt_(date) {
  return Utilities.formatDate(date, TZ, 'yyyy-MM-dd');
}


function parseDay_(str) {
  return Utilities.parseDate(str, TZ, 'yyyy-MM-dd');
}


function addDays_(str, n) {
  const p = str.split('-').map(Number);
  return Utilities.formatDate(new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)), 'UTC', 'yyyy-MM-dd');
}


function dayList_(fromStr, toStr) {
  const out = [];
  for (let d = fromStr, guard = 0; d <= toStr && guard < 800; d = addDays_(d, 1), guard++) out.push(d);
  return out;
}


// Makes sure the manual ad spend tab has one row per day up to (and including) lastStr.
function ensureAdSpendRows_(sh, lastStr) {
  const startStr = fmt_(readSettings_().start);
  const lastRow = sh.getLastRow();
  let next = startStr;
  if (lastRow >= 2) {
    const dates = sh.getRange(2, 1, lastRow - 1, 1).getValues().map(r => r[0]).filter(d => d instanceof Date);
    if (dates.length) next = addDays_(fmt_(dates[dates.length - 1]), 1);
  }
  if (next > lastStr) return;
  const rows = dayList_(next, lastStr).map(d => [parseDay_(d), '', '']);
  sh.getRange(Math.max(lastRow, 1) + 1, 1, rows.length, 3).setValues(rows);
  sh.getRange(Math.max(lastRow, 1) + 1, 1, rows.length, 1).setNumberFormat('ddd dd mmm yyyy');
  sh.getRange(Math.max(lastRow, 1) + 1, 2, rows.length, 2).setNumberFormat('£#,##0.00');
}
function testShopifySettings() {
  const cfg = readSettings_();

  console.log('Shopify domain: ' + cfg.shop);
  console.log('Product path: ' + cfg.path);
  console.log('API version: ' + cfg.api);

  if (!cfg.shop || !cfg.shop.endsWith('.myshopify.com')) {
    throw new Error(
      'Invalid Shopify domain. Check Settings!B3.'
    );
  }

  console.log(
    'OAuth URL: https://' +
    cfg.shop +
    '/admin/oauth/access_token'
  );
}
