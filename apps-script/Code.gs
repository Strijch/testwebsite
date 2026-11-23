/**
 * Plantae website backend (Google Apps Script, bound to "Plantae_Website_Content_Manager").
 *
 *  1. doGet   -> JSON feed the website reads (products, categories, events, home, policies)
 *               plus drop sales (drops, dropItems, dropSettings, serverNow).
 *               ?view=drop returns only the small, cached drop feed used for live stock.
 *  2. doPost  -> places a drop order. Checks the sale window and stock inside a lock, so two
 *               people can never buy the last plant, then logs the order and its lines.
 *  3. setupDropSales() / the "Plantae drops" menu -> creates the drop tabs in this sheet.
 *
 * Shipping, invoicing and payment stay manual: the Orders tab tracks them.
 */

var TZ = 'Africa/Johannesburg';
var TAB = {
  products: 'Products', cats: 'Categories', events: 'Open Days & Events', home: 'Home Page',
  policies: 'Dispatch & Policies', drops: 'Drop Events', dropPlants: 'Drop Plants',
  orders: 'Orders', lines: 'Order Lines', settings: 'Drop Settings', dispatch: 'Dispatch'
};
var FIRST = 5;            // data starts on row 5 (rows 1-3 are title/notes, row 4 is the header)
var ROWS = { drops: 200, dropPlants: 600, orders: 1000, lines: 3000, dispatch: 300 };
var DELIVERY = ['Courier door-to-door', 'PostNet counter', 'Collect at open day or show'];
var CACHE_KEY = 'plantae_drop_feed_v1';

/* ------------------------------------------------------------------ web app */

function doGet(e) {
  var view = e && e.parameter && e.parameter.view;
  var out;
  try {
    if (view === 'drop') {
      var cache = CacheService.getScriptCache();
      var hit = cache.get(CACHE_KEY);
      if (hit) return jsonOut_(hit, true);
      out = dropFeed_(loadDropData_());
      var s = JSON.stringify(out);
      if (s.length < 90000) cache.put(CACHE_KEY, s, 8);
      return jsonOut_(s, true);
    }
    out = {
      products: tabObjects_(TAB.products),
      categories: tabObjects_(TAB.cats),
      events: tabObjects_(TAB.events),
      home: tabObjects_(TAB.home),
      policies: tabObjects_(TAB.policies)
    };
    var df = dropFeed_(loadDropData_());
    for (var k in df) out[k] = df[k];
  } catch (err) {
    out = { error: String(err) };
  }
  return jsonOut_(JSON.stringify(out), true);
}

function doPost(e) {
  var out;
  try {
    var body = JSON.parse(e.postData.contents);
    if (body && body.action === 'order') out = placeOrder_(body);
    else out = { ok: false, code: 'bad_request', message: 'Unknown request.' };
  } catch (err) {
    console.error(err);
    out = { ok: false, code: 'error', message: 'Something went wrong on our side. Please WhatsApp us your order.' };
  }
  return jsonOut_(JSON.stringify(out), false);
}

function jsonOut_(text, isString) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------ reading tabs */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

/** Returns {sheet, headerRow, headers[], values[][] (data rows only)} or null. */
function readTab_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) return null;
  var all = sh.getDataRange().getValues();
  var h = -1;
  for (var r = 0; r < Math.min(10, all.length); r++) {
    var filled = 0;
    for (var c = 0; c < all[r].length; c++) if (String(all[r][c]).trim() !== '') filled++;
    if (filled >= 3) { h = r; break; }
  }
  if (h < 0) return { sheet: sh, headerRow: 0, headers: [], rows: [] };
  var headers = all[h].map(function (x) { return String(x).trim(); });
  var rows = [];
  for (var i = h + 1; i < all.length; i++) {
    var first = all[i][0];
    if (first === '' || first === null) continue;          // formula rows that are still blank
    rows.push({ row: i + 1, v: all[i] });
  }
  return { sheet: sh, headerRow: h + 1, headers: headers, rows: rows };
}

function col_(t, startsWith) {
  var s = startsWith.toLowerCase();
  for (var i = 0; i < t.headers.length; i++) if (t.headers[i].toLowerCase().indexOf(s) === 0) return i;
  return -1;
}

function cell_(t, rowObj, startsWith) {
  var i = col_(t, startsWith);
  return i < 0 ? '' : rowObj.v[i];
}

function str_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function num_(v) { var n = parseFloat(String(v).replace(/[^\d.\-]/g, '')); return isNaN(n) ? 0 : n; }
function yes_(v) { return /^y/i.test(str_(v)); }

function iso_(d, withTime) {
  return Utilities.formatDate(d, TZ, withTime ? "yyyy-MM-dd'T'HH:mm:ssXXX" : 'yyyy-MM-dd');
}

function toDate_(v) {
  if (v instanceof Date) return v;
  var s = str_(v);
  if (!s) return null;
  var d = new Date(s);
  return isNaN(d) ? null : d;
}

/** Whole tab as objects keyed by header text, the shape the website already reads. */
function tabObjects_(name) {
  var t = readTab_(name);
  if (!t) return [];
  return t.rows.map(function (r) {
    var o = {};
    t.headers.forEach(function (h, i) {
      if (!h) return;
      var v = r.v[i];
      if (v instanceof Date) {
        o[h] = iso_(v, /^(opens|closes|start|end|received)/i.test(h));
      } else if (typeof v === 'string' && /^\d{1,2} [A-Za-z]{3} \d{4}$/.test(v.trim()) && /date/i.test(h)) {
        var d = new Date(v.trim() + ' 12:00:00');
        o[h] = isNaN(d) ? v : iso_(d, false);
      } else {
        o[h] = v;
      }
    });
    return o;
  });
}

/* --------------------------------------------------------------- drop data */

function loadDropData_() {
  var data = {
    now: new Date(),
    settings: {},
    drops: [],
    plants: [],
    products: {},
    sold: {},          // dropId|sku -> units on active orders
    t: {}
  };
  // settings
  var st = readTab_(TAB.settings);
  if (st) st.rows.forEach(function (r) { data.settings[str_(r.v[0]).toLowerCase()] = r.v[1]; });

  // products by SKU
  var pt = readTab_(TAB.products);
  if (pt) {
    var iSku = col_(pt, 'SKU'), iCard = col_(pt, 'Card name'), iPhoto = col_(pt, 'Photo 1'),
        iReg = col_(pt, 'Regular price'), iSale = col_(pt, 'Sale price'), iMeta = col_(pt, 'Meta line'),
        iCat = col_(pt, 'Category');
    pt.rows.forEach(function (r) {
      var sku = str_(r.v[iSku]);
      if (!sku) return;
      var reg = num_(r.v[iReg]), sale = num_(r.v[iSale]);
      data.products[sku] = {
        name: str_(r.v[iCard]), photo: str_(r.v[iPhoto]), meta: str_(r.v[iMeta]), cat: str_(r.v[iCat]),
        price: sale > 0 && sale < reg ? sale : reg, was: sale > 0 && sale < reg ? reg : 0
      };
    });
  }

  // drops
  var dt = readTab_(TAB.drops);
  data.t.drops = dt;
  if (dt) dt.rows.forEach(function (r) {
    var id = str_(cell_(dt, r, 'Drop ID'));
    var start = toDate_(cell_(dt, r, 'Opens')), end = toDate_(cell_(dt, r, 'Closes'));
    if (!id || !start || !end) return;
    data.drops.push({
      id: id, name: str_(cell_(dt, r, 'Drop name')) || id, start: start, end: end,
      enabled: !/^n/i.test(str_(cell_(dt, r, 'Open for orders'))),
      show: !/^n/i.test(str_(cell_(dt, r, 'Show on site'))),
      note: str_(cell_(dt, r, 'Message'))
    });
  });

  // active order numbers (anything not cancelled)
  var ot = readTab_(TAB.orders);
  data.t.orders = ot;
  var active = {};
  if (ot) {
    var iNo = col_(ot, 'Order no'), iStat = col_(ot, 'Order status');
    ot.rows.forEach(function (r) {
      active[str_(r.v[iNo])] = str_(r.v[iStat]).toLowerCase() !== 'cancelled';
    });
  }

  // units sold per drop+sku
  var lt = readTab_(TAB.lines);
  if (lt) {
    var lNo = col_(lt, 'Order no'), lDrop = col_(lt, 'Drop ID'), lSku = col_(lt, 'SKU'), lQty = col_(lt, 'Qty');
    lt.rows.forEach(function (r) {
      if (!active[str_(r.v[lNo])]) return;
      var k = str_(r.v[lDrop]) + '|' + str_(r.v[lSku]);
      data.sold[k] = (data.sold[k] || 0) + num_(r.v[lQty]);
    });
  }

  // plants in each drop
  var ppt = readTab_(TAB.dropPlants);
  if (ppt) {
    var iD = col_(ppt, 'Drop ID'), iS = col_(ppt, 'SKU'), iPrice = col_(ppt, 'Drop price override'),
        iQty = col_(ppt, 'Qty for this drop'), iMax = col_(ppt, 'Max per order');
    ppt.rows.forEach(function (r) {
      var dropId = str_(r.v[iD]), sku = str_(r.v[iS]);
      if (!dropId || !sku) return;
      var p = data.products[sku] || { name: sku, photo: '', meta: '', cat: '', price: 0, was: 0 };
      var override = num_(r.v[iPrice]);
      var price = override > 0 ? override : p.price;
      var total = num_(r.v[iQty]);
      var sold = data.sold[dropId + '|' + sku] || 0;
      data.plants.push({
        dropId: dropId, sku: sku, name: p.name || sku, cat: p.cat, meta: p.meta, photo: p.photo,
        price: price, was: override > 0 ? (p.price > override ? p.price : 0) : p.was,
        total: total, sold: sold, left: Math.max(0, total - sold), max: num_(r.v[iMax])
      });
    });
  }
  return data;
}

function dropFeed_(data) {
  var shown = {};
  var drops = data.drops.filter(function (d) { return d.show; }).map(function (d) {
    shown[d.id] = true;
    return { id: d.id, name: d.name, start: iso_(d.start, true), end: iso_(d.end, true), enabled: d.enabled, note: d.note };
  });
  var items = data.plants.filter(function (p) { return shown[p.dropId]; }).map(function (p) {
    return { dropId: p.dropId, sku: p.sku, name: p.name, cat: p.cat, meta: p.meta, photo: p.photo,
             price: p.price, was: p.was, total: p.total, left: p.left, max: p.max };
  });
  var s = data.settings;
  return {
    serverNow: iso_(data.now, true),
    drops: drops,
    dropItems: items,
    dropSettings: { whatsapp: str_(s['whatsapp number shown on the page']), maxPerOrder: num_(s['max plants per order (0 = no limit)']) }
  };
}

/* ------------------------------------------------------------ placing orders */

function cleanText_(v, max) {
  return str_(v).replace(/[\u0000-\u001f<>]/g, ' ').replace(/\s+/g, ' ').slice(0, max);
}

function placeOrder_(body) {
  if (str_(body.website)) return { ok: false, code: 'bad_request', message: 'Could not place order.' };   // honeypot

  var c = body.customer || {};
  var cust = {
    name: cleanText_(c.name, 80), phone: String(c.phone || '').replace(/[^\d+]/g, '').slice(0, 16),
    email: cleanText_(c.email, 120), method: cleanText_(c.method, 40), address: cleanText_(c.address, 120),
    suburb: cleanText_(c.suburb, 60), city: cleanText_(c.city, 60), province: cleanText_(c.province, 40),
    postal: cleanText_(c.postal, 8), notes: cleanText_(c.notes, 300)
  };
  var errors = {};
  if (cust.name.length < 2) errors.name = 'Please enter your name.';
  if (cust.phone.replace(/\D/g, '').length < 9) errors.phone = 'Please enter a WhatsApp or mobile number.';
  if (cust.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(cust.email)) errors.email = 'That email address does not look right.';
  if (DELIVERY.indexOf(cust.method) < 0) errors.method = 'Please choose how you would like your plants.';
  if (cust.method && cust.method !== DELIVERY[2]) {
    if (!cust.address) errors.address = 'Please enter your street address.';
    if (!cust.city) errors.city = 'Please enter your city or town.';
    if (!/^\d{4}$/.test(cust.postal)) errors.postal = 'Please enter a 4-digit postal code.';
  }
  var dropId = cleanText_(body.dropId, 60);
  var token = cleanText_(body.token, 60);
  var req = {};
  (Array.isArray(body.items) ? body.items : []).slice(0, 60).forEach(function (it) {
    var sku = cleanText_(it && it.sku, 40), q = Math.floor(Number(it && it.qty));
    if (sku && q >= 1 && q <= 99) req[sku] = (req[sku] || 0) + q;
  });
  if (!dropId || !Object.keys(req).length) return { ok: false, code: 'empty', message: 'Your basket is empty.' };
  if (Object.keys(errors).length) return { ok: false, code: 'invalid', errors: errors, message: 'Please check the highlighted details.' };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (e) {
    return { ok: false, code: 'busy', message: 'We are very busy right now. Please try again in a few seconds.' };
  }
  var result, mail = null;
  try {
    var data = loadDropData_();
    var ot = data.t.orders;
    if (!ot) return { ok: false, code: 'error', message: 'Orders tab is missing. Please WhatsApp us your order.' };

    // the same submission twice (e.g. a retry after a dropped connection) returns the first order
    var iTok = col_(ot, 'Submission ref');
    if (token && iTok >= 0) {
      for (var k = 0; k < ot.rows.length; k++) {
        if (str_(ot.rows[k].v[iTok]) === token) {
          return { ok: true, duplicate: true, orderNo: str_(cell_(ot, ot.rows[k], 'Order no')),
                   total: num_(cell_(ot, ot.rows[k], 'Items total')), summary: str_(cell_(ot, ot.rows[k], 'Items')) };
        }
      }
    }

    var drop = null;
    data.drops.forEach(function (d) { if (d.id === dropId) drop = d; });
    if (!drop || !drop.enabled) return { ok: false, code: 'closed', message: 'This drop is not taking orders.' };
    var now = data.now.getTime();
    if (now < drop.start.getTime()) return { ok: false, code: 'not_started', message: 'This drop has not opened yet.' };
    if (now >= drop.end.getTime()) return { ok: false, code: 'closed', message: 'This drop has closed.' };

    var byKey = {};
    data.plants.forEach(function (p) { if (p.dropId === dropId) byKey[p.sku] = p; });

    var bad = [], lines = [], totalUnits = 0, itemsTotal = 0;
    Object.keys(req).forEach(function (sku) {
      var p = byKey[sku], q = req[sku];
      if (!p) { bad.push({ sku: sku, name: sku, left: 0, requested: q, reason: 'gone' }); return; }
      var cap = p.max > 0 ? Math.min(p.left, p.max) : p.left;
      if (q > cap) { bad.push({ sku: sku, name: p.name, left: p.left, requested: q, reason: p.left <= 0 ? 'sold_out' : (p.max > 0 && q > p.max ? 'max' : 'short') }); return; }
      lines.push({ p: p, q: q });
      totalUnits += q; itemsTotal += q * p.price;
    });
    var cap2 = num_(data.settings['max plants per order (0 = no limit)']);
    if (!bad.length && cap2 > 0 && totalUnits > cap2) {
      return { ok: false, code: 'too_many', max: cap2, message: 'Please order no more than ' + cap2 + ' plants at a time.' };
    }
    if (bad.length) {
      var fresh = dropFeed_(data);
      return { ok: false, code: 'stock', unavailable: bad,
               dropItems: fresh.dropItems.filter(function (i) { return i.dropId === dropId; }),
               message: 'Some plants were just claimed by someone else. Nothing has been ordered yet.' };
    }

    // order number: prefix + running number (highest existing + 1)
    var prefix = str_(data.settings['order number prefix']) || 'PLD';
    var iNo = col_(ot, 'Order no'), max = 0;
    ot.rows.forEach(function (r) {
      var m = String(r.v[iNo]).match(/(\d+)$/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
    var orderNo = prefix + '-' + ('0000' + (max + 1)).slice(-4);
    var summary = lines.map(function (l) { return l.q + '× ' + l.p.name; }).join('; ');

    // Orders row: write only the script-owned columns so the sheet's formulas stay intact
    var os = ot.sheet, orow = firstEmptyRow_(os, 1, ROWS.orders);
    if (!orow) return { ok: false, code: 'error', message: 'The order register is full. Please WhatsApp us your order.' };
    setByHeader_(ot, orow, {
      'Order no': orderNo, 'Received': data.now, 'Drop ID': dropId, 'Customer name': cust.name,
      'WhatsApp / mobile': "'" + cust.phone, 'Email': cust.email, 'Delivery method': cust.method,
      'Address': cust.address, 'Suburb': cust.suburb, 'City': cust.city, 'Province': cust.province,
      'Postal code': cust.postal, 'Customer notes': cust.notes, 'Items': summary, 'Items total': itemsTotal,
      'Order status': 'Active', 'Submission ref': token
    });

    // Order Lines
    var lt = readTab_(TAB.lines), ls = lt.sheet, lrow = firstEmptyRow_(ls, 1, ROWS.lines);
    if (!lrow || lrow + lines.length - 1 > FIRST + ROWS.lines - 1) return { ok: false, code: 'error', message: 'The order register is full. Please WhatsApp us your order.' };
    lines.forEach(function (l, i) {
      setByHeader_(lt, lrow + i, { 'Order no': orderNo, 'Drop ID': dropId, 'SKU': l.p.sku, 'Plant': l.p.name, 'Qty': l.q, 'Unit price': l.p.price });
    });
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove(CACHE_KEY);

    result = { ok: true, orderNo: orderNo, total: itemsTotal, summary: summary,
               lines: lines.map(function (l) { return { sku: l.p.sku, name: l.p.name, qty: l.q, price: l.p.price }; }) };
    mail = { orderNo: orderNo, cust: cust, summary: summary, total: itemsTotal, drop: drop, settings: data.settings };
  } finally {
    lock.releaseLock();
  }
  if (mail) notify_(mail);
  return result;
}

function firstEmptyRow_(sheet, colIndex, rows) {
  var vals = sheet.getRange(FIRST, colIndex, rows, 1).getValues();
  for (var i = 0; i < vals.length; i++) if (str_(vals[i][0]) === '') return FIRST + i;
  return 0;
}

function setByHeader_(t, row, obj) {
  Object.keys(obj).forEach(function (h) {
    var i = col_(t, h);
    if (i >= 0) t.sheet.getRange(row, i + 1).setValue(obj[h]);
  });
}

function notify_(m) {
  try {
    var owner = str_(m.settings['email me when an order arrives']);
    var money = 'R' + m.total.toLocaleString('en-ZA');
    if (owner) {
      MailApp.sendEmail(owner, 'New drop order ' + m.orderNo + ' – ' + m.cust.name + ' – ' + money,
        m.drop.name + '\n' + m.orderNo + '\n\n' + m.cust.name + '\n' + m.cust.phone + '\n' + m.cust.email + '\n' +
        m.cust.method + '\n' + [m.cust.address, m.cust.suburb, m.cust.city, m.cust.province, m.cust.postal].filter(String).join(', ') +
        '\n\n' + m.summary + '\nPlants total: ' + money + ' (shipping to be added)\n\n' + m.cust.notes);
    }
    if (m.cust.email && /^y/i.test(str_(m.settings['send customers a confirmation email']) || 'Yes')) {
      MailApp.sendEmail(m.cust.email, 'Your Plantae order ' + m.orderNo,
        'Hi ' + m.cust.name.split(' ')[0] + ',\n\nThank you – your plants are reserved.\n\nOrder ' + m.orderNo + '\n' + m.summary +
        '\nPlants total: ' + money + '\n\nWe will send you an invoice (with shipping added) shortly. Please use ' + m.orderNo +
        ' as your EFT payment reference. Plants are dispatched once payment is received.\n\nPlantae');
    }
  } catch (err) {
    console.warn('Notification failed', err);
  }
}

/* ------------------------------------------------------------------ set-up */

var COLOR = { head: '#1f4a36', calcHead: '#6b7a70', input: '#fff8e7', calc: '#e3ebe1' };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Plantae drops')
    .addItem('Set up the drop tabs', 'setupDropSales')
    .addItem('Add an example drop (for a demo)', 'addExampleDrop')
    .addToUi();
}

function money_(r, a, b) { return '=IF(' + a + r + '="","",' + b + ')'; }

function specs_() {
  var P = "Products!$A$5:$A$700";
  var lk = function (colLetter, r) {
    return '=IF($B' + r + '="","",IFERROR(INDEX(Products!$' + colLetter + '$5:$' + colLetter + '$700,MATCH($B' + r + ',' + P + ',0)),""))';
  };
  return [
    {
      name: TAB.settings, rows: 12, title: 'Drop settings',
      subtitle: 'Small switches for drop sales. Change the cream cells only.',
      cols: [{ h: 'Setting', w: 300 }, { h: 'Value', w: 260 }, { h: 'What it does', w: 460 }],
      seed: [
        ['Order number prefix', 'PLD', 'Order numbers look like PLD-0001. The customer uses this as the EFT payment reference.'],
        ['Email me when an order arrives', '', 'Type an email address to get a message for every new drop order. Leave blank for none.'],
        ['Send customers a confirmation email', 'Yes', 'Yes or No. Only sent when the customer typed an email address.'],
        ['WhatsApp number shown on the page', '', 'Shown next to the order form for questions, for example +27 84 752 6823.'],
        ['Max plants per order (0 = no limit)', 0, 'Stops one person from taking everything. 0 means no limit.']
      ]
    },
    {
      name: TAB.drops, rows: ROWS.drops, title: 'Drop events',
      subtitle: 'One row per WhatsApp drop. Buying is only possible between Opens and Closes. Set Open for orders to No to stop sales straight away.',
      cols: [
        { h: 'Drop ID', w: 130 }, { h: 'Drop name', w: 240 },
        { h: 'Opens (date and time)', w: 160, fmt: 'yyyy-mm-dd hh:mm' }, { h: 'Closes (date and time)', w: 160, fmt: 'yyyy-mm-dd hh:mm' },
        { h: 'Open for orders', w: 110, list: ['Yes', 'No'] }, { h: 'Show on site', w: 100, list: ['Yes', 'No'] },
        { h: 'Message on the drop page', w: 320 },
        { h: 'Status now', w: 100, f: function (r) { return '=IF(A' + r + '="","",IF(E' + r + '="No","Paused",IF(NOW()<C' + r + ',"Upcoming",IF(NOW()>=D' + r + ',"Closed","Live"))))'; } },
        { h: 'Plants listed', w: 90, f: function (r) { return '=IF(A' + r + '="","",COUNTIF(\'Drop Plants\'!$A$5:$A$' + (FIRST + ROWS.dropPlants - 1) + ',A' + r + '))'; } },
        { h: 'Orders', w: 80, f: function (r) { return '=IF(A' + r + '="","",COUNTIFS(Orders!$C$5:$C$' + (FIRST + ROWS.orders - 1) + ',A' + r + ',Orders!$X$5:$X$' + (FIRST + ROWS.orders - 1) + ',"<>Cancelled"))'; } },
        { h: 'Plants sold', w: 90, f: function (r) { return '=IF(A' + r + '="","",SUMIFS(\'Drop Plants\'!$J$5:$J$' + (FIRST + ROWS.dropPlants - 1) + ',\'Drop Plants\'!$A$5:$A$' + (FIRST + ROWS.dropPlants - 1) + ',A' + r + '))'; } },
        { h: 'Plants total (R)', w: 120, fmt: '"R"#,##0', f: function (r) { return '=IF(A' + r + '="","",SUMIFS(Orders!$O$5:$O$' + (FIRST + ROWS.orders - 1) + ',Orders!$C$5:$C$' + (FIRST + ROWS.orders - 1) + ',A' + r + ',Orders!$X$5:$X$' + (FIRST + ROWS.orders - 1) + ',"<>Cancelled"))'; } }
      ]
    },
    {
      name: TAB.dropPlants, rows: ROWS.dropPlants, title: 'Drop plants',
      subtitle: 'The plants in each drop. Type the Drop ID, the SKU from the Products tab and how many are available. One-offs are a quantity of 1. Name, photo and price fill in by themselves.',
      cols: [
        { h: 'Drop ID', w: 130 }, { h: 'SKU', w: 100 },
        { h: 'Plant', w: 280, f: function (r) { return lk('V', r); } },
        { h: 'Photo file', w: 220, f: function (r) { return lk('P', r); } },
        { h: 'Product price (R)', w: 120, fmt: '"R"#,##0', f: function (r) {
            return '=IF($B' + r + '="","",IFERROR(IF(INDEX(Products!$K$5:$K$700,MATCH($B' + r + ',' + P + ',0))>0,INDEX(Products!$K$5:$K$700,MATCH($B' + r + ',' + P + ',0)),INDEX(Products!$J$5:$J$700,MATCH($B' + r + ',' + P + ',0))),""))'; } },
        { h: 'Drop price override (R)', w: 140, fmt: '"R"#,##0' },
        { h: 'Price in this drop (R)', w: 130, fmt: '"R"#,##0', f: function (r) { return '=IF($B' + r + '="","",IF(N(F' + r + ')>0,F' + r + ',E' + r + '))'; } },
        { h: 'Qty for this drop', w: 110 }, { h: 'Max per order (blank = none)', w: 150 },
        { h: 'Qty sold', w: 80, f: function (r) { return '=IF($B' + r + '="","",SUMIFS(\'Order Lines\'!$E$5:$E$' + (FIRST + ROWS.lines - 1) + ',\'Order Lines\'!$B$5:$B$' + (FIRST + ROWS.lines - 1) + ',$A' + r + ',\'Order Lines\'!$C$5:$C$' + (FIRST + ROWS.lines - 1) + ',$B' + r + ',\'Order Lines\'!$H$5:$H$' + (FIRST + ROWS.lines - 1) + ',"Yes"))'; } },
        { h: 'Qty left', w: 80, f: function (r) { return '=IF($B' + r + '="","",N(H' + r + ')-N(J' + r + '))'; } },
        { h: 'Status', w: 100, f: function (r) { return '=IF($B' + r + '="","",IF(K' + r + '<=0,"Sold out",IF(K' + r + '<=2,"Low","Available")))'; } },
        { h: 'Notes', w: 240 }
      ]
    },
    {
      name: TAB.orders, rows: ROWS.orders, title: 'Orders',
      subtitle: 'Every drop order lands here by itself (columns A to O). You fill in Shipping, Invoice sent, Paid, Dispatched and Tracking. The Next step column tells you what to do. Set Order status to Cancelled to put the plants back on sale.',
      cols: [
        { h: 'Order no', w: 100 }, { h: 'Received', w: 140, fmt: 'yyyy-mm-dd hh:mm' }, { h: 'Drop ID', w: 110 },
        { h: 'Customer name', w: 160 }, { h: 'WhatsApp / mobile', w: 130 }, { h: 'Email', w: 190 },
        { h: 'Delivery method', w: 190 }, { h: 'Address', w: 200 }, { h: 'Suburb', w: 120 }, { h: 'City', w: 120 },
        { h: 'Province', w: 120 }, { h: 'Postal code', w: 80 }, { h: 'Customer notes', w: 200 },
        { h: 'Items', w: 360 }, { h: 'Items total (R)', w: 110, fmt: '"R"#,##0' },
        { h: 'Shipping (R)', w: 100, fmt: '"R"#,##0' },
        { h: 'Order total (R)', w: 110, fmt: '"R"#,##0', f: function (r) { return '=IF(A' + r + '="","",O' + r + '+N(P' + r + '))'; } },
        { h: 'Invoice sent?', w: 100, list: ['Yes', 'No'] }, { h: 'Paid?', w: 80, list: ['Yes', 'No'] },
        { h: 'Date paid', w: 110, fmt: 'yyyy-mm-dd' }, { h: 'Dispatched?', w: 100, list: ['Yes', 'No'] },
        { h: 'Courier', w: 130 }, { h: 'Tracking / waybill', w: 150 },
        { h: 'Order status', w: 100, list: ['Active', 'Cancelled'] },
        { h: 'Next step', w: 170, f: function (r) {
            return '=IF(A' + r + '="","",IF(X' + r + '="Cancelled","Cancelled",IF(U' + r + '="Yes","Dispatched",IF(S' + r + '="Yes","Ready to pack",IF(R' + r + '="Yes","Awaiting payment",IF(P' + r + '="","Add shipping and invoice","Send invoice"))))))'; } },
        { h: 'Pack order', w: 80, f: function (r) { return '=IF(Y' + r + '="Ready to pack",COUNTIF(Y$5:Y' + r + ',"Ready to pack"),"")'; } },
        { h: 'Submission ref', w: 120 }
      ]
    },
    {
      name: TAB.lines, rows: ROWS.lines, title: 'Order lines',
      subtitle: 'One row per plant per order, written by the website. Do not edit. This is the register of which plant went to whom.',
      cols: [
        { h: 'Order no', w: 100 }, { h: 'Drop ID', w: 110 }, { h: 'SKU', w: 100 }, { h: 'Plant', w: 300 },
        { h: 'Qty', w: 60 }, { h: 'Unit price (R)', w: 110, fmt: '"R"#,##0' },
        { h: 'Line total (R)', w: 110, fmt: '"R"#,##0', f: function (r) { return '=IF(A' + r + '="","",E' + r + '*F' + r + ')'; } },
        { h: 'Counts toward stock?', w: 150, f: function (r) {
            return '=IF(A' + r + '="","",IF(IFERROR(INDEX(Orders!$X$5:$X$' + (FIRST + ROWS.orders - 1) + ',MATCH(A' + r + ',Orders!$A$5:$A$' + (FIRST + ROWS.orders - 1) + ',0)),"")="Cancelled","No","Yes"))'; } }
      ]
    },
    {
      name: TAB.dispatch, rows: ROWS.dispatch, title: 'Dispatch: ready to pack',
      subtitle: 'Fills in by itself with every order that is paid and not yet dispatched, oldest first. When you ship an order, set Dispatched? to Yes on the Orders tab and it drops off this list.',
      cols: [
        { h: 'Row ref', w: 60, f: function (r) { return '=IFERROR(MATCH(ROWS(A$5:A' + r + '),Orders!$Z$5:$Z$' + (FIRST + ROWS.orders - 1) + ',0),"")'; } },
        { h: 'Order no', w: 100, f: function (r) { return disp_('A', r); } },
        { h: 'Customer', w: 160, f: function (r) { return disp_('D', r); } },
        { h: 'WhatsApp / mobile', w: 130, f: function (r) { return disp_('E', r); } },
        { h: 'Delivery method', w: 190, f: function (r) { return disp_('G', r); } },
        { h: 'Deliver to', w: 320, f: function (r) {
            return '=IF($A' + r + '="","",TRIM(INDEX(Orders!$H$5:$H$' + (FIRST + ROWS.orders - 1) + ',$A' + r + ')&", "&INDEX(Orders!$I$5:$I$' + (FIRST + ROWS.orders - 1) + ',$A' + r + ')&", "&INDEX(Orders!$J$5:$J$' + (FIRST + ROWS.orders - 1) + ',$A' + r + ')&", "&INDEX(Orders!$K$5:$K$' + (FIRST + ROWS.orders - 1) + ',$A' + r + ')&" "&INDEX(Orders!$L$5:$L$' + (FIRST + ROWS.orders - 1) + ',$A' + r + ')))'; } },
        { h: 'Plants to pack', w: 360, f: function (r) { return disp_('N', r); } },
        { h: 'Customer notes', w: 200, f: function (r) { return disp_('M', r); } },
        { h: 'Packed', w: 80, list: ['Yes'] }
      ]
    }
  ];
}

function disp_(letter, r) {
  return '=IF($A' + r + '="","",INDEX(Orders!$' + letter + '$5:$' + letter + '$' + (FIRST + ROWS.orders - 1) + ',$A' + r + '))';
}

function setupDropSales() {
  var ss = ss_(), ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  var made = [], kept = [];
  specs_().forEach(function (spec) {
    var sh = ss.getSheetByName(spec.name);
    if (sh) { kept.push(spec.name); return; }
    sh = ss.insertSheet(spec.name);
    sh.getRange(1, 1).setValue(spec.title).setFontSize(16).setFontWeight('bold');
    sh.getRange(2, 1).setValue(spec.subtitle).setFontColor('#55625a');
    var n = spec.cols.length;
    sh.getRange(4, 1, 1, n).setValues([spec.cols.map(function (c) { return c.h; })])
      .setFontWeight('bold').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
    spec.cols.forEach(function (c, i) {
      sh.getRange(4, i + 1).setBackground(c.f ? COLOR.calcHead : COLOR.head);
      sh.setColumnWidth(i + 1, c.w || 120);
      var body = sh.getRange(FIRST, i + 1, spec.rows, 1);
      body.setBackground(c.f ? COLOR.calc : COLOR.input);
      if (c.f) {
        var fs = [];
        for (var r = FIRST; r < FIRST + spec.rows; r++) fs.push([c.f(r)]);
        body.setFormulas(fs);
      }
      if (c.fmt) body.setNumberFormat(c.fmt);
      if (c.list) body.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(c.list, true).setAllowInvalid(false).build());
    });
    if (spec.seed) sh.getRange(FIRST, 1, spec.seed.length, spec.seed[0].length).setValues(spec.seed);
    sh.setFrozenRows(4);
    made.push(spec.name);
  });

  // Order status defaults to Active only when the website writes a row; nothing to prefill.
  // Make the Orders and Dispatch tabs easy to scan.
  var os = ss.getSheetByName(TAB.orders);
  if (os && made.indexOf(TAB.orders) >= 0) {
    os.getRange('Y5:Y' + (FIRST + ROWS.orders - 1)).setFontWeight('bold');
    os.hideColumns(26, 2);   // Pack order helper and Submission ref
    var rules = [
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Ready to pack').setBackground('#b7e1c6').setRanges([os.getRange('Y5:Y' + (FIRST + ROWS.orders - 1))]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Cancelled').setBackground('#e6c3bd').setRanges([os.getRange('Y5:Y' + (FIRST + ROWS.orders - 1))]).build()
    ];
    os.setConditionalFormatRules(rules);
  }
  var ds = ss.getSheetByName(TAB.dispatch);
  if (ds && made.indexOf(TAB.dispatch) >= 0) ds.hideColumns(1, 1);

  // Read Me notes
  var rm = ss.getSheetByName('Read Me');
  if (rm && made.length) {
    var last = rm.getLastRow() + 2;
    rm.getRange(last, 1, 7, 2).setValues([
      ['Drop sales tabs', ''],
      ['Drop Events', 'One row per WhatsApp drop: name, opens, closes. Buying is only possible between the two times.'],
      ['Drop Plants', 'The plants in each drop and how many are available. One-offs are a quantity of 1.'],
      ['Orders', 'Every drop order arrives here. Add shipping, mark invoice sent, paid and dispatched. Next step tells you what to do.'],
      ['Order Lines', 'Written by the website: which plant went to which order. Do not edit.'],
      ['Dispatch', 'Fills in by itself: paid orders that still need packing and shipping.'],
      ['Drop Settings', 'Order number prefix, email alerts, and the WhatsApp number shown on the drop page.']
    ]);
    rm.getRange(last, 1).setFontWeight('bold');
  }
  var msg = (made.length ? 'Created: ' + made.join(', ') + '. ' : '') + (kept.length ? 'Already there, left alone: ' + kept.join(', ') + '.' : '');
  if (ss.getSpreadsheetTimeZone() !== TZ) {
    msg += ' Note: this spreadsheet is set to ' + ss.getSpreadsheetTimeZone() + '. Set File > Settings > Time zone to Johannesburg so drop times are read correctly.';
  }
  if (ui) ui.alert('Plantae drops', msg, ui.ButtonSet.OK); else console.log(msg);
  return msg;
}

/** Adds a clearly labelled example drop that opens in 2 minutes and runs for 2 hours. */
function addExampleDrop() {
  var ss = ss_();
  if (!ss.getSheetByName(TAB.drops)) setupDropSales();
  var dt = readTab_(TAB.drops), dsh = dt.sheet;
  var row = firstEmptyRow_(dsh, 1, ROWS.drops);
  var start = new Date(Date.now() + 2 * 60000), end = new Date(start.getTime() + 2 * 3600000);
  dsh.getRange(row, 1, 1, 7).setValues([['D-EXAMPLE', 'EXAMPLE drop: delete this row after the demo', start, end, 'Yes', 'Yes', 'Example only. Orders placed here are real rows in the Orders tab, so cancel them afterwards.']]);

  var pt = readTab_(TAB.products), skus = [];
  pt.rows.forEach(function (r) { var s = str_(r.v[col_(pt, 'SKU')]); if (s && skus.length < 6) skus.push(s); });
  var qty = [1, 1, 3, 5, 2, 1];
  var psh = ss.getSheetByName(TAB.dropPlants), prow = firstEmptyRow_(psh, 1, ROWS.dropPlants);
  skus.forEach(function (s, i) { psh.getRange(prow + i, 1, 1, 2).setValues([['D-EXAMPLE', s]]); psh.getRange(prow + i, 8).setValue(qty[i]); });
  CacheService.getScriptCache().remove(CACHE_KEY);
  var msg = 'Example drop added. It opens in 2 minutes and closes 2 hours after that.';
  try { SpreadsheetApp.getUi().alert('Plantae drops', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { console.log(msg); }
}
