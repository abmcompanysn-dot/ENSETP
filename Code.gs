// ============================================================
//  ENSETP GALA 2026 — Google Apps Script Backend v3.0
//  Script lié à la feuille Google Sheets via Extensions → Apps Script
//  Déployer comme "Application Web" — Accès : Tout le monde
// ============================================================

// ── CONFIG PAR DÉFAUT (remplacés par la feuille "Configuration") ──
var CFG = {
  COMMANDES_SHEET:  'Commandes',
  PRESENCES_SHEET:  'Présences',
  ADMINS_SHEET:     'Sous-Admins',
  CONFIG_SHEET:     'Configuration',
  SHEET_ID:         '',             // Vide = feuille liée au script

  // Ces valeurs sont lues depuis l'onglet "Configuration" en priorité
  ADMIN_EMAIL:      'contact@mahu.cards',
  EMAIL_FROM:       'contact@mahu.cards',
  EVENT_DATE:       'Samedi 20 Juin 2026',
  EVENT_LIEU:       'Au Magic Land',

  PAYDUNIA_KEY:     '',   // Clé privée (PAYDUNYA-PRIVATE-KEY)
  PAYDUNIA_SECRET:  '',   // Non utilisé — vérification via confirm endpoint
  PAYDUNIA_URL:     'https://paydunya.com/api/v1/checkout-invoice/create',

  // URL de CE script Apps Script (webhook Paydunia → doPost)
  APPS_SCRIPT_URL:  'https://script.google.com/macros/s/AKfycby18WcAD27h5r6Tv3uGHfFKTYo6EOd_iYCE42iuw5XA64JH5Amhul0Nuz1WO1mb9-op/exec',

  // URL du site public (où l'utilisateur est redirigé après paiement)
  SITE_URL:         'https://ensetp.mahu.cards'
};

// ── LECTURE DE LA CONFIG DEPUIS LE SHEET ──────────────────
function getSheetConfig() {
  try {
    var ss = CFG.SHEET_ID
      ? SpreadsheetApp.openById(CFG.SHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CFG.CONFIG_SHEET);
    if (!sh) return {};
    var data = sh.getDataRange().getValues();
    var cfg = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1] !== '') cfg[data[i][0]] = String(data[i][1]);
    }
    return cfg;
  } catch (e) {
    Logger.log('getSheetConfig error: ' + e.message);
    return {};
  }
}

// Fusionne CFG + valeurs du sheet (sheet prioritaire)
function getCFG() {
  var sc = getSheetConfig();
  return {
    ADMIN_EMAIL:          sc.ADMIN_EMAIL          || CFG.ADMIN_EMAIL,
    EMAIL_FROM:           sc.EMAIL_FROM           || CFG.EMAIL_FROM,
    EVENT_DATE:           sc.EVENT_DATE           || CFG.EVENT_DATE,
    EVENT_LIEU:           sc.EVENT_LIEU           || CFG.EVENT_LIEU,
    PAYDUNIA_KEY:         sc.PAYDUNIA_KEY         || CFG.PAYDUNIA_KEY,
    PAYDUNIA_MASTER_KEY:  sc.PAYDUNIA_MASTER_KEY  || '',
    PAYDUNIA_TOKEN:       sc.PAYDUNIA_TOKEN        || '',
    PAYDUNYA_MODE:        sc.PAYDUNYA_MODE         || 'live',
    PAYDUNIA_URL:         CFG.PAYDUNIA_URL,
    APPS_SCRIPT_URL:      sc.APPS_SCRIPT_URL      || CFG.APPS_SCRIPT_URL,
    SITE_URL:             sc.SITE_URL             || CFG.SITE_URL
  };
}

// ── MENU GOOGLE SHEETS ─────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎩 Gala ENSETP 2026')
    .addItem('⚙️  Initialiser toutes les feuilles',   'setupSheet')
    .addSeparator()
    .addItem('📊 Rapport quotidien (e-mail)',          'sendDailyReport')
    .addItem('⏰ Activer rappel quotidien à 8h',       'setupDailyReminder')
    .addSeparator()
    .addItem('🧪 Test — Envoyer un ticket bidon',     'testSendTicket')
    .addItem('🧹 Supprimer les triggers',             'clearTriggers')
    .addToUi();
}

// ── INITIALISATION DES FEUILLES ────────────────────────────
function setupSheet() {
  var ss = CFG.SHEET_ID
    ? SpreadsheetApp.openById(CFG.SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  _initConfig(ss);
  _initCommandes(ss);
  _initPresences(ss);
  _initAdmins(ss);

  SpreadsheetApp.getUi().alert(
    '✅ Initialisation terminée !\n\n' +
    '4 feuilles créées :\n' +
    '• Configuration — clés API et paramètres\n' +
    '• Commandes    — tickets achetés\n' +
    '• Présences    — confirmations\n' +
    '• Sous-Admins  — équipe admin\n\n' +
    '👉 Remplissez vos clés Paydunia dans l\'onglet "Configuration"\n' +
    '   avant d\'activer les paiements en ligne.'
  );
}

function _initConfig(ss) {
  var sh = ss.getSheetByName(CFG.CONFIG_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.CONFIG_SHEET, 0);

  var headers = ['Clé', 'Valeur', 'Description'];
  sh.getRange(1, 1, 1, 3).setValues([headers])
    .setFontWeight('bold').setBackground('#2C2C54').setFontColor('#FFFFFF').setFontSize(10);
  sh.setFrozenRows(1);

  if (sh.getLastRow() <= 1) {
    var rows = [
      ['PAYDUNIA_KEY',        '',                      '🔑 PAYDUNYA-PRIVATE-KEY — Clé privée (tableau de bord PayDunya)'],
      ['PAYDUNIA_MASTER_KEY', '',                      '👑 PAYDUNYA-MASTER-KEY — Clé maîtresse PayDunya'],
      ['PAYDUNIA_TOKEN',      '',                      '🎫 PAYDUNYA-TOKEN — Jeton d\'accès PayDunya'],
      ['PAYDUNYA_MODE',       'live',                  '⚙️ Mode PayDunya : "test" (sandbox) ou "live" (production)'],
      ['APPS_SCRIPT_URL',     CFG.APPS_SCRIPT_URL,    '🔗 URL de ce script — PayDunya envoie les confirmations ici'],
      ['SITE_URL',            CFG.SITE_URL,            '🌐 URL du site public (redirection après paiement)'],
      ['ADMIN_EMAIL',         CFG.ADMIN_EMAIL,         '📧 E-mail recevant toutes les notifications admin'],
      ['EMAIL_FROM',          CFG.EMAIL_FROM,          '📤 E-mail expéditeur des tickets'],
      ['EVENT_DATE',          CFG.EVENT_DATE,          '📅 Date de l\'événement (affiché sur les tickets)'],
      ['EVENT_LIEU',          CFG.EVENT_LIEU,          '📍 Lieu de l\'événement (affiché sur les tickets)']
    ];
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 300);
  sh.setColumnWidth(3, 340);

  // Surligner les 4 clés Paydunia vides en orange
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty().setBackground('#3d2000').setFontColor('#ffaa00')
    .setRanges([sh.getRange('B2:B5')]).build();
  sh.setConditionalFormatRules([rule]);

  Logger.log('Feuille Configuration prête.');
}

function _initCommandes(ss) {
  var sh = ss.getSheetByName(CFG.COMMANDES_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.COMMANDES_SHEET);
  var headers = ['ID Ticket','Prénom','Nom','E-mail','Téléphone','Type','Qté','Prix Unit.','Total','Paiement','Date','Statut','Email Envoyé'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#D4AF37').setFontColor('#000000').setFontSize(10);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 200); sh.setColumnWidth(4, 200); sh.setColumnWidth(11, 160);
  sh.setColumnWidths(2, 2, 120);
  Logger.log('Feuille Commandes prête.');
}

function _initPresences(ss) {
  var sh = ss.getSheetByName(CFG.PRESENCES_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.PRESENCES_SHEET);
  var headers = ['ID Ticket','Prénom','Nom','E-mail','Confirmé le','Source'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1A6B35').setFontColor('#FFFFFF').setFontSize(10);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 200); sh.setColumnWidth(4, 200); sh.setColumnWidth(5, 160);
  Logger.log('Feuille Présences prête.');
}

function _initAdmins(ss) {
  var sh = ss.getSheetByName(CFG.ADMINS_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.ADMINS_SHEET);
  var headers = ['Nom','E-mail','Rôle','Date Ajout'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#2C2C54').setFontColor('#FFFFFF').setFontSize(10);
  sh.setFrozenRows(1); sh.setColumnWidth(2, 200);
  Logger.log('Feuille Sous-Admins prête.');
}

// ── SÉCURITÉ — RATE LIMITING ───────────────────────────────
function checkRateLimit(requestId) {
  try {
    var cache = CacheService.getScriptCache();
    var key = 'rl_' + (requestId || 'global');
    var count = parseInt(cache.get(key) || '0');
    if (count >= 60) return false;  // max 60 requêtes / minute
    cache.put(key, String(count + 1), 60);
    return true;
  } catch (e) {
    return true;  // En cas d'erreur cache, on laisse passer
  }
}

// ── SÉCURITÉ — VALIDATION SIGNATURE PAYDUNIA ──────────────
function validatePayduniaSignature(rawBody, receivedSig) {
  var cfg = getCFG();
  if (!cfg.PAYDUNIA_SECRET || !receivedSig) return true;  // Pas de secret = skip
  try {
    var sig = Utilities.computeHmacSha256Signature(rawBody, cfg.PAYDUNIA_SECRET, Utilities.Charset.UTF_8);
    var sigHex = sig.map(function(b){ return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
    return sigHex === receivedSig.toLowerCase().replace(/^sha256=/,'');
  } catch (e) {
    Logger.log('Signature validation error: ' + e.message);
    return false;
  }
}

// ── POINT D'ENTRÉE HTTP ────────────────────────────────────
function doPost(e) {
  try {
    // Rate limiting
    var clientId = e.parameter ? (e.parameter.client || 'api') : 'api';
    if (!checkRateLimit(clientId)) {
      return _jsonOut({ error: 'Trop de requêtes. Réessayez dans une minute.', code: 429 });
    }

    var rawBody = e.postData ? e.postData.contents : '';
    var payload = JSON.parse(rawBody);
    var action  = payload.action;
    var result  = {};

    // PayDunya webhook : pas de champ 'action' — structure { data: { invoice, custom_data } }
    // ou ancienne structure { order_id, status, ... }
    var isPaydunyaWebhook = !action && (
      (payload.data && payload.data.invoice) ||
      payload.order_id || payload.transaction_ref || payload.reference
    );
    if (isPaydunyaWebhook) {
      return _jsonOut(handlePayduniaWebhook(payload));
    }

    if      (action === 'createOrder')     result = handleCreateOrder(payload.order);
    else if (action === 'payduniaWebhook') result = handlePayduniaWebhook(payload);
    else if (action === 'confirmPresence') result = handleConfirmPresence(payload);
    else if (action === 'addSubAdmin')     result = handleAddSubAdmin(payload);
    else if (action === 'getStats')        result = getStats();
    else if (action === 'sendOTP')         result = handleSendOTP(payload);
    else if (action === 'verifyOTP')       result = handleVerifyOTP(payload);
    else if (action === 'verifyTicket')    result = handleVerifyTicket(payload);
    else                                   result = { error: 'Action inconnue: ' + action };

    return _jsonOut(result);
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return _jsonOut({ error: err.message, code: 500 });
  }
}

function doGet(e) {
  // Callback Paydunia GET (redirection utilisateur après paiement)
  if (e.parameter && e.parameter.payment_status) {
    handlePayduniaCallback(e.parameter);
    var cfg = getCFG();
    var url = cfg.SITE_URL + '?payment_status=' + e.parameter.payment_status + '&order_id=' + (e.parameter.order_id || '');
    return HtmlService.createHtmlOutput(
      '<html><head><meta http-equiv="refresh" content="0;url=' + url + '"></head>' +
      '<body>Redirection en cours...</body></html>'
    );
  }
  return _jsonOut({ status: 'ENSETP Gala API active', version: '3.0' });
}

function _jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── CRÉATION DE COMMANDE ───────────────────────────────────
function handleCreateOrder(order) {
  var cfg = getCFG();

  // 1. Enregistrement dans la feuille (toujours pending jusqu'à confirmation Paydunia)
  order.status = 'pending';
  logToSheet(order);

  // 2. Si clé Paydunia configurée → créer lien de paiement
  if (cfg.PAYDUNIA_KEY) {
    var payLink = createPayduniaPayment(order, cfg);
    if (payLink) {
      Logger.log('Lien Paydunia créé: ' + payLink);
      return { success: true, paymentUrl: payLink, orderId: order.id };
    }
    Logger.log('⚠️ Paydunia configuré mais aucun lien retourné — bascule mode démo');
  }

  // 3. Mode démo (sans Paydunia ou si Paydunia échoue) → ticket envoyé directement
  order.status = 'paid';
  updateOrderStatus(order.id, 'paid');
  sendTicketEmail(order, cfg);
  sendAdminNotification(order, cfg);
  return { success: true, orderId: order.id, mode: 'demo' };
}

// ── CONFIRMATION DE PRÉSENCE ───────────────────────────────
function handleConfirmPresence(payload) {
  try {
    var cfg = getCFG();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CFG.PRESENCES_SHEET);
    if (!sh) { _initPresences(ss); sh = ss.getSheetByName(CFG.PRESENCES_SHEET); }

    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === payload.orderId) {
        return { success: true, already: true, message: 'Déjà confirmé.' };
      }
    }

    sh.appendRow([
      payload.orderId,
      (payload.name || '').split(' ')[0] || '',
      (payload.name || '').split(' ').slice(1).join(' ') || '',
      payload.email || '',
      new Date(),
      'Web'
    ]);

    GmailApp.sendEmail(
      cfg.ADMIN_EMAIL,
      '[GALA ENSETP] ✅ Présence confirmée – ' + payload.orderId,
      'Nouvelle confirmation de présence :\n\n'
        + 'Ticket : ' + payload.orderId + '\n'
        + 'Nom    : ' + (payload.name  || '–') + '\n'
        + 'E-mail : ' + (payload.email || '–') + '\n'
        + 'Le     : ' + new Date().toLocaleString('fr-FR') + '\n\n'
        + '—\nMahu Events · mahu.cards'
    );
    return { success: true, message: 'Présence enregistrée.' };
  } catch (err) {
    Logger.log('handleConfirmPresence error: ' + err.message);
    return { error: err.message };
  }
}

// ── SOUS-ADMIN ─────────────────────────────────────────────
function handleAddSubAdmin(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CFG.ADMINS_SHEET);
    if (!sh) { _initAdmins(ss); sh = ss.getSheetByName(CFG.ADMINS_SHEET); }
    sh.appendRow([payload.name||'', payload.email||'', payload.role||'readonly', new Date()]);
    return { success: true };
  } catch (err) { return { error: err.message }; }
}

// ── PAYDUNYA ───────────────────────────────────────────────
function createPayduniaPayment(order, cfg) {
  try {
    // Headers d'authentification PayDunya (clés dans le tableau de bord PayDunya)
    var headers = {
      'PAYDUNYA-MASTER-KEY':  cfg.PAYDUNIA_MASTER_KEY,
      'PAYDUNYA-PRIVATE-KEY': cfg.PAYDUNIA_KEY,
      'PAYDUNYA-TOKEN':       cfg.PAYDUNIA_TOKEN,
      'PAYDUNYA-MODE':        cfg.PAYDUNYA_MODE || 'live'
    };

    // Structure payload attendue par l'API PayDunya
    var body = {
      invoice: {
        total_amount: order.total,
        description:  'Ticket Gala ENSETP 2026 – ' + order.type.toUpperCase()
      },
      store: {
        name:        'Gala ENSETP 2026 · ENSETP',
        website_url: cfg.SITE_URL
      },
      actions: {
        cancel_url:   cfg.SITE_URL + '?payment_status=cancel&order_id='  + order.id,
        return_url:   cfg.SITE_URL + '?payment_status=success&order_id=' + order.id,
        callback_url: cfg.APPS_SCRIPT_URL
      },
      custom_data: {
        order_id:       order.id,
        customer_name:  order.prenom + ' ' + order.nom,
        customer_email: order.email,
        customer_phone: order.tel || ''
      }
    };

    var resp = UrlFetchApp.fetch(cfg.PAYDUNIA_URL, {
      method: 'post', contentType: 'application/json',
      headers: headers, payload: JSON.stringify(body), muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    Logger.log('PayDunya init (' + resp.getResponseCode() + '): ' + JSON.stringify(data));

    // response_code "00" = succès, response_text = checkout URL
    if (data.response_code === '00') {
      Logger.log('✅ PayDunya checkout URL: ' + data.response_text);
      return data.response_text;
    }
    Logger.log('⚠️ PayDunya erreur: ' + (data.response_text || JSON.stringify(data)));
    return null;
  } catch (e) {
    Logger.log('PayDunya error: ' + e.message);
    return null;
  }
}

function handlePayduniaWebhook(payload) {
  var cfg = getCFG();

  // PayDunya webhook structure: { data: { invoice: { token, status }, custom_data: { order_id } } }
  var invoiceData  = (payload.data && payload.data.invoice)     ? payload.data.invoice     : payload;
  var customData   = (payload.data && payload.data.custom_data) ? payload.data.custom_data : payload;
  var token   = invoiceData.token   || payload.token;
  var orderId = customData.order_id || payload.order_id || payload.transaction_ref || payload.reference;

  Logger.log('Webhook PayDunya reçu — token: ' + token + ' order: ' + orderId);

  // Vérification sécurisée via l'endpoint confirm (recommandé par PayDunya)
  var status = '';
  if (token) {
    try {
      var confirmResp = UrlFetchApp.fetch(
        'https://paydunya.com/api/v1/checkout-invoice/confirm/' + token,
        {
          method: 'get',
          headers: {
            'PAYDUNYA-MASTER-KEY':  cfg.PAYDUNIA_MASTER_KEY,
            'PAYDUNYA-PRIVATE-KEY': cfg.PAYDUNIA_KEY,
            'PAYDUNYA-TOKEN':       cfg.PAYDUNIA_TOKEN,
            'PAYDUNYA-MODE':        cfg.PAYDUNYA_MODE || 'live'
          },
          muteHttpExceptions: true
        }
      );
      var confirmData = JSON.parse(confirmResp.getContentText());
      Logger.log('PayDunya confirm: ' + JSON.stringify(confirmData));
      status = String(confirmData.status || confirmData.invoice_status || '').toLowerCase();
      if (!orderId && confirmData.custom_data) orderId = confirmData.custom_data.order_id;
    } catch (e) {
      Logger.log('PayDunya confirm error: ' + e.message);
    }
  }

  // Fallback statut depuis le webhook direct
  if (!status) {
    status = String(invoiceData.status || payload.status || payload.payment_status || '').toLowerCase();
  }

  Logger.log('PayDunya — order: ' + orderId + ' status: ' + status);

  if (status === 'completed' || status === 'success' || status === 'successful') {
    if (!orderId) { Logger.log('⚠️ order_id introuvable dans le webhook'); return { error: 'order_id manquant' }; }
    updateOrderStatus(orderId, 'paid');
    var order = getOrderById(orderId);
    if (order) {
      sendTicketEmail(order, cfg);
      sendAdminNotification(order, cfg);
      Logger.log('✅ Ticket envoyé après confirmation PayDunya: ' + orderId);
    } else {
      Logger.log('⚠️ Commande introuvable pour: ' + orderId);
    }
    return { success: true, orderId: orderId };
  }

  return { success: false, reason: 'Statut non confirmé: ' + status };
}

function handlePayduniaCallback(params) {
  // Redirection GET après paiement (côté client) — ne pas livrer sans confirm webhook
  var orderId = params.order_id;
  var status  = String(params.payment_status || '').toLowerCase();
  if (status === 'success') {
    // Le webhook aura déjà traité le paiement ; ceci est juste un fallback de redirection
    Logger.log('Callback GET PayDunya — order: ' + orderId + ' (webhook devrait avoir déjà confirmé)');
  }
}

// ── GOOGLE SHEETS — COMMANDES ──────────────────────────────
function getSheet() {
  var ss = CFG.SHEET_ID
    ? SpreadsheetApp.openById(CFG.SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.COMMANDES_SHEET);
  if (!sh) { _initCommandes(ss); sh = ss.getSheetByName(CFG.COMMANDES_SHEET); }
  return sh;
}

function logToSheet(order) {
  var sh = getSheet();
  sh.appendRow([
    order.id, order.prenom, order.nom, order.email, order.tel||'',
    order.type.toUpperCase(), order.qty||1, order.price||0, order.total||0,
    order.payment||'', new Date(order.date||new Date()), order.status||'pending', 'Non'
  ]);
}

function updateOrderStatus(orderId, status) {
  var sh = getSheet(), data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      sh.getRange(i+1, 12).setValue(status==='paid'?'PAYÉ':status);
      if (status==='paid') sh.getRange(i+1, 13).setValue('Oui');
      return;
    }
  }
}

function getOrderById(orderId) {
  var sh = getSheet(), data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      return {
        id:data[i][0], prenom:data[i][1], nom:data[i][2], email:data[i][3],
        tel:data[i][4], type:data[i][5], qty:data[i][6], price:data[i][7],
        total:data[i][8], payment:data[i][9], date:data[i][10], status:data[i][11]
      };
    }
  }
  return null;
}

function getStats() {
  var sh = getSheet(), data = sh.getDataRange().getValues();
  var total=0, solo=0, couple=0, recettes=0;
  for (var i=1; i<data.length; i++) {
    total++;
    var t = String(data[i][5]).toLowerCase();
    if (t==='solo') solo++; else if (t==='couple') couple++;
    recettes += Number(data[i][8])||0;
  }
  return { total:total, solo:solo, couple:couple, recettes:recettes };
}

// ── EMAILS ─────────────────────────────────────────────────
function sendTicketEmail(order, cfg) {
  cfg = cfg || getCFG();
  var name  = order.prenom + ' ' + order.nom;
  var total = Number(order.total).toLocaleString('fr-FR') + ' FCFA';
  var type  = String(order.type).toUpperCase();
  var body  = buildEmailBody(name, type, total, order.id, buildTicketHtml(order, cfg), cfg);
  try {
    GmailApp.sendEmail(order.email,
      '🎉 Votre Ticket – Dîner de Gala ENSETP 2026 [' + order.id + ']',
      stripHtml(body),
      { name: 'Gala ENSETP 2026', htmlBody: body, replyTo: cfg.ADMIN_EMAIL }
    );
    Logger.log('✅ Ticket envoyé à ' + order.email);
  } catch (e) {
    Logger.log('❌ Erreur envoi email ticket: ' + e.message);
  }
}

function sendAdminNotification(order, cfg) {
  cfg = cfg || getCFG();
  var body = '🎫 Nouvelle commande — Gala ENSETP 2026\n\n'
    + 'ID       : ' + order.id + '\n'
    + 'Nom      : ' + order.prenom + ' ' + order.nom + '\n'
    + 'E-mail   : ' + order.email + '\n'
    + 'Tél.     : ' + (order.tel||'–') + '\n'
    + 'Type     : ' + order.type.toUpperCase() + '\n'
    + 'Qté      : ' + (order.qty||1) + '\n'
    + 'Total    : ' + Number(order.total).toLocaleString('fr-FR') + ' FCFA\n'
    + 'Paiement : ' + (order.payment||'–') + '\n'
    + 'Date     : ' + new Date(order.date).toLocaleString('fr-FR') + '\n\n'
    + '—\nMahu Events · mahu.cards';
  try {
    GmailApp.sendEmail(cfg.ADMIN_EMAIL, '[GALA ENSETP] Nouvelle commande – ' + order.id, body);
  } catch (e) {
    Logger.log('❌ Erreur notification admin: ' + e.message);
  }
}

// ── TICKET HTML (pour e-mail) ──────────────────────────────
function buildTicketHtml(order, cfg) {
  cfg = cfg || getCFG();
  return '<div style="background:linear-gradient(135deg,#1a1a1a,#0d0d0d);border:2px solid #D4AF37;border-radius:14px;overflow:hidden;font-family:Arial,sans-serif;max-width:460px;">'
    + '<div style="background:#0A0A0A;padding:16px 22px;border-bottom:1px solid rgba(212,175,55,0.3);display:flex;justify-content:space-between;align-items:center;">'
    +   '<span style="font-size:1.1rem;font-weight:bold;color:#D4AF37;letter-spacing:3px;">ENSETP</span>'
    +   '<span style="font-size:0.6rem;letter-spacing:2px;color:#D4AF37;text-transform:uppercase;">Dîner de Gala · 2026</span>'
    + '</div>'
    + '<div style="padding:20px 22px;">'
    +   '<div style="font-size:1.8rem;font-style:italic;color:#D4AF37;margin-bottom:4px;">Gala</div>'
    +   '<div style="font-size:0.6rem;letter-spacing:4px;text-transform:uppercase;color:#aaa;margin-bottom:16px;">De Fin d\'Année · ENSETP 2026</div>'
    +   '<table width="100%" style="color:#fff;border-collapse:collapse;">'
    +     '<tr><td style="padding:6px 0;"><div style="font-size:0.55rem;letter-spacing:2px;text-transform:uppercase;color:#888">Titulaire</div><div style="font-size:0.9rem;font-weight:600">' + order.prenom + ' ' + order.nom + '</div></td>'
    +         '<td style="padding:6px 0;"><div style="font-size:0.55rem;letter-spacing:2px;text-transform:uppercase;color:#888">Type</div><div><span style="background:linear-gradient(135deg,#D4AF37,#A07820);color:#000;font-size:0.6rem;font-weight:700;letter-spacing:2px;padding:3px 10px;border-radius:18px;">' + order.type.toUpperCase() + '</span></div></td></tr>'
    +     '<tr><td style="padding:6px 0;"><div style="font-size:0.55rem;letter-spacing:2px;text-transform:uppercase;color:#888">E-mail</div><div style="font-size:0.78rem;color:#ccc">' + order.email + '</div></td>'
    +         '<td style="padding:6px 0;"><div style="font-size:0.55rem;letter-spacing:2px;text-transform:uppercase;color:#888">Montant</div><div style="color:#D4AF37;font-weight:bold">' + Number(order.total).toLocaleString('fr-FR') + ' FCFA</div></td></tr>'
    +   '</table>'
    +   '<hr style="border:none;border-top:1px dashed rgba(212,175,55,0.3);margin:16px 0;" />'
    +   '<table width="100%" style="color:#fff;border-collapse:collapse;">'
    +     '<tr><td><div style="font-size:0.55rem;letter-spacing:2px;text-transform:uppercase;color:#888">Date</div><div style="font-weight:600">' + cfg.EVENT_DATE + '</div></td>'
    +         '<td><div style="font-size:0.55rem;letter-spacing:2px;text-transform:uppercase;color:#888">Lieu</div><div style="font-weight:600">' + cfg.EVENT_LIEU + '</div></td></tr>'
    +   '</table>'
    + '</div>'
    + '<div style="background:rgba(212,175,55,0.06);padding:12px 22px;border-top:1px dashed rgba(212,175,55,0.3);">'
    +   '<div style="font-family:monospace;font-size:0.72rem;color:#D4AF37;letter-spacing:2px;">' + order.id + '</div>'
    +   '<div style="font-size:0.6rem;color:#888;margin-top:3px;">Émis le ' + new Date(order.date||new Date()).toLocaleDateString('fr-FR') + '</div>'
    + '</div>'
    + '</div>';
}

function buildEmailBody(name, type, total, id, ticketHtml, cfg) {
  cfg = cfg || getCFG();
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="background:#0A0A0A;margin:0;padding:30px;font-family:Arial,sans-serif;">'
    + '<div style="max-width:560px;margin:0 auto;">'
    +   '<div style="text-align:center;margin-bottom:30px;">'
    +     '<div style="font-size:1.4rem;font-weight:bold;color:#D4AF37;letter-spacing:4px;margin-bottom:6px;">ENSETP</div>'
    +     '<div style="font-size:0.7rem;color:#888;letter-spacing:2px;text-transform:uppercase;">École Normale Supérieure d\'Enseignement Technique et Professionnel</div>'
    +   '</div>'
    +   '<div style="background:#1A1A1A;border:1px solid rgba(212,175,55,0.2);border-radius:16px;padding:32px;margin-bottom:24px;">'
    +     '<h2 style="color:#D4AF37;font-size:1.2rem;margin:0 0 14px">Bonjour ' + name + ' ! 🎉</h2>'
    +     '<p style="color:#ccc;font-size:0.9rem;line-height:1.7;margin:0 0 20px">Votre ticket pour le <strong style="color:#F5D66A">Dîner de Gala de Fin d\'Année ENSETP 2026</strong> a bien été reçu et votre paiement est confirmé.</p>'
    +     '<table style="width:100%;border-collapse:collapse;margin-bottom:24px;">'
    +       '<tr><td style="padding:8px 14px;background:#222;color:#888;font-size:0.72rem;text-transform:uppercase;letter-spacing:1px;">Type de ticket</td><td style="padding:8px 14px;background:#1A1A1A;color:#fff;font-weight:600;border-left:2px solid #D4AF37">' + type + '</td></tr>'
    +       '<tr><td style="padding:8px 14px;background:#222;color:#888;font-size:0.72rem;text-transform:uppercase;letter-spacing:1px;">Montant payé</td><td style="padding:8px 14px;background:#1A1A1A;color:#D4AF37;font-weight:bold;border-left:2px solid #D4AF37">' + total + '</td></tr>'
    +       '<tr><td style="padding:8px 14px;background:#222;color:#888;font-size:0.72rem;text-transform:uppercase;letter-spacing:1px;">Date</td><td style="padding:8px 14px;background:#1A1A1A;color:#fff;border-left:2px solid #D4AF37">' + cfg.EVENT_DATE + '</td></tr>'
    +       '<tr><td style="padding:8px 14px;background:#222;color:#888;font-size:0.72rem;text-transform:uppercase;letter-spacing:1px;">Lieu</td><td style="padding:8px 14px;background:#1A1A1A;color:#fff;border-left:2px solid #D4AF37">' + cfg.EVENT_LIEU + '</td></tr>'
    +     '</table>'
    +     '<p style="color:#888;font-size:0.8rem;margin:0">Référence : <strong style="color:#D4AF37">' + id + '</strong></p>'
    +   '</div>'
    +   '<div style="margin-bottom:24px;"><h3 style="color:#D4AF37;font-size:0.88rem;letter-spacing:2px;text-transform:uppercase;margin-bottom:16px;">Votre E-Ticket</h3>' + ticketHtml + '</div>'
    +   '<div style="background:#1A1A1A;border:1px solid rgba(212,175,55,0.15);border-radius:12px;padding:22px;margin-bottom:24px;">'
    +     '<h3 style="color:#D4AF37;font-size:0.82rem;letter-spacing:2px;margin:0 0 12px">Artistes Invités</h3>'
    +     '<p style="color:#ccc;font-size:0.85rem;line-height:1.7;margin:0">ZOU NAME &bull; MIST CASH &bull; BLM PRO &bull; ME MAN &bull; NIKO<br>'
    +     '<span style="color:#888;font-size:0.78rem;">Ambiance · Surprises · Prestations Live · Partage</span></p>'
    +   '</div>'
    +   '<p style="color:#888;font-size:0.75rem;text-align:center;line-height:1.7;margin:0">'
    +     'Ce ticket est personnel et non remboursable.<br>'
    +     'Présentez-le (imprimé ou sur écran) à l\'entrée de l\'événement.<br>'
    +     '<span style="color:#D4AF37">ENSETP · Former aujourd\'hui, réussir demain !</span><br>'
    +     '<span style="color:#555;font-size:0.68rem;">Créé avec Mahu Events · mahu.cards</span>'
    +   '</p>'
    + '</div></body></html>';
}

function stripHtml(html) {
  return html.replace(/<[^>]*>/g,' ').replace(/\s+/g,' ').trim();
}

// ── RAPPORT QUOTIDIEN ─────────────────────────────────────
function setupDailyReminder() {
  clearTriggers();
  ScriptApp.newTrigger('sendDailyReport').timeBased().atHour(8).everyDays(1).create();
  SpreadsheetApp.getUi().alert('✅ Rappel quotidien activé (tous les jours à 8h).');
}

function clearTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction()==='sendDailyReport') ScriptApp.deleteTrigger(t);
  });
}

function sendDailyReport() {
  var cfg   = getCFG();
  var stats = getStats();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ps    = ss.getSheetByName(CFG.PRESENCES_SHEET);
  var presences = ps ? Math.max(0, ps.getLastRow()-1) : 0;
  var body = '📊 Rapport quotidien — Gala ENSETP 2026\n\n'
    + '📅 ' + new Date().toLocaleString('fr-FR') + '\n\n'
    + '🎫 TICKETS\n  Total : ' + stats.total + '\n  Solo  : ' + stats.solo + '\n  Couple: ' + stats.couple + '\n\n'
    + '✅ PRÉSENCES\n  Confirmées : ' + presences + '\n\n'
    + '💰 RECETTES\n  Total : ' + stats.recettes.toLocaleString('fr-FR') + ' FCFA\n\n'
    + '—\nMahu Events · mahu.cards';
  GmailApp.sendEmail(cfg.ADMIN_EMAIL, '[GALA ENSETP] Rapport du ' + new Date().toLocaleDateString('fr-FR'), body);
}

// ── TEST ───────────────────────────────────────────────────
function testSendTicket() {
  var cfg = getCFG();
  var testOrder = {
    id:'TK-ENSETP-2026-TEST-00001', prenom:'Cheikh', nom:'Diallo',
    email: Session.getActiveUser().getEmail(),
    tel:'+221 77 000 00 00', type:'solo', qty:1, price:1000, total:1000,
    payment:'wave', date:new Date().toISOString(), status:'paid'
  };
  logToSheet(testOrder);
  sendTicketEmail(testOrder, cfg);
  sendAdminNotification(testOrder, cfg);
  Logger.log('✅ Test envoyé à ' + testOrder.email + ' + ' + cfg.ADMIN_EMAIL);
  SpreadsheetApp.getUi().alert(
    '✅ Test envoyé !\n\n' +
    '• Ticket → ' + testOrder.email + '\n' +
    '• Notification admin → ' + cfg.ADMIN_EMAIL + '\n\n' +
    'Vérifiez vos boîtes mail.'
  );
}

// ── VÉRIFICATION BILLET À L'ENTRÉE ────────────────────────
function handleVerifyTicket(payload) {
  try {
    var ticketId = String(payload.ticketId || '').trim().toUpperCase();
    if (!ticketId) return { valid: false, error: 'ID ticket requis.' };

    var order = getOrderById(ticketId);
    if (!order) return { valid: false, error: 'Ticket introuvable. Vérifiez l\'ID.' };

    var paid = String(order.status).toLowerCase();
    if (paid !== 'payé' && paid !== 'paid') {
      return {
        valid: false,
        error: 'Ticket non payé.',
        order: { id: order.id, nom: order.prenom + ' ' + order.nom, type: order.type }
      };
    }

    // Vérifier si déjà scanné à l'entrée (cache 24h)
    var cache = CacheService.getScriptCache();
    var scanKey = 'scan_' + ticketId;
    var alreadyScanned = cache.get(scanKey);
    if (!alreadyScanned) cache.put(scanKey, '1', 86400);

    Logger.log((alreadyScanned ? '⚠️ Re-scan' : '✅ Scan') + ' billet: ' + ticketId);
    return {
      valid: true,
      alreadyScanned: !!alreadyScanned,
      order: {
        id: order.id,
        nom: order.prenom + ' ' + order.nom,
        type: String(order.type).toUpperCase(),
        qty: order.qty || 1,
        email: order.email
      }
    };
  } catch (err) {
    Logger.log('handleVerifyTicket error: ' + err.message);
    return { valid: false, error: err.message };
  }
}

// ── AUTHENTIFICATION ADMIN PAR OTP ────────────────────────
function handleSendOTP(payload) {
  try {
    var cfg   = getCFG();
    var email = String(payload.email || '').toLowerCase().trim();
    if (!email) return { error: 'Email requis.' };

    // Emails autorisés : admin principal + sous-admins du sheet
    var authorized = [cfg.ADMIN_EMAIL.toLowerCase()];
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName(CFG.ADMINS_SHEET);
      if (sh) {
        var data = sh.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (data[i][1]) authorized.push(String(data[i][1]).toLowerCase().trim());
        }
      }
    } catch(e) { Logger.log('Sous-admins non lus: ' + e.message); }

    if (authorized.indexOf(email) === -1) {
      Logger.log('OTP refusé — email non autorisé: ' + email);
      return { error: 'Cet e-mail n\'est pas autorisé à accéder au panneau d\'administration.' };
    }

    // Rate limiting : max 3 OTP par email par heure
    var cache  = CacheService.getScriptCache();
    var rlKey  = 'otp_rl_' + email.replace(/[^a-z0-9]/g, '_');
    var rlCount = parseInt(cache.get(rlKey) || '0');
    if (rlCount >= 3) return { error: 'Trop de tentatives. Attendez 1 heure avant de réessayer.' };
    cache.put(rlKey, String(rlCount + 1), 3600);

    // Génération du code OTP à 6 chiffres
    var otp    = String(Math.floor(100000 + Math.random() * 900000));
    var otpKey = 'otp_' + email.replace(/[^a-z0-9]/g, '_');
    cache.put(otpKey, otp, 300); // expire en 5 minutes

    // Envoi de l'e-mail
    var html = '<!DOCTYPE html><html><body style="background:#0A0A0A;margin:0;padding:30px;font-family:Arial,sans-serif;">'
      + '<div style="max-width:420px;margin:0 auto;background:#1A1A1A;border:1px solid rgba(212,175,55,0.25);border-radius:16px;padding:36px;text-align:center;">'
      + '<div style="font-size:1.2rem;font-weight:bold;color:#D4AF37;letter-spacing:4px;margin-bottom:6px;">ENSETP</div>'
      + '<div style="font-size:0.7rem;color:#888;letter-spacing:1px;margin-bottom:28px;">GALA DE FIN D\'ANNÉE 2026 · ADMIN</div>'
      + '<div style="font-size:0.85rem;color:#ccc;margin-bottom:22px;line-height:1.6;">Votre code de connexion à l\'interface d\'administration :</div>'
      + '<div style="font-size:2.6rem;font-weight:bold;color:#D4AF37;letter-spacing:14px;'
      +   'background:rgba(212,175,55,0.08);border:1px solid rgba(212,175,55,0.3);'
      +   'border-radius:12px;padding:22px 16px;margin-bottom:24px;">' + otp + '</div>'
      + '<div style="font-size:0.75rem;color:#888;line-height:1.8;">'
      +   'Ce code expire dans <strong style="color:#F5D66A">5 minutes</strong>.<br>'
      +   'Si vous n\'avez pas demandé ce code, ignorez cet e-mail.'
      + '</div>'
      + '<div style="margin-top:28px;padding-top:20px;border-top:1px solid rgba(212,175,55,0.1);font-size:0.65rem;color:#555;">Mahu Events · mahu.cards</div>'
      + '</div></body></html>';

    GmailApp.sendEmail(email,
      '[Gala ENSETP 2026] Code de connexion admin : ' + otp,
      'Votre code de connexion admin est : ' + otp + '\n\nCe code expire dans 5 minutes.\nSi vous n\'avez pas demandé ce code, ignorez cet e-mail.\n\n—\nMahu Events · mahu.cards',
      { name: 'Gala ENSETP 2026', htmlBody: html, replyTo: cfg.ADMIN_EMAIL }
    );

    Logger.log('✅ OTP envoyé à ' + email);
    return { success: true, message: 'Code envoyé à ' + email };
  } catch (err) {
    Logger.log('handleSendOTP error: ' + err.message);
    return { error: 'Erreur lors de l\'envoi: ' + err.message };
  }
}

function handleVerifyOTP(payload) {
  try {
    var email = String(payload.email || '').toLowerCase().trim();
    var code  = String(payload.code  || '').trim();
    if (!email || !code) return { error: 'Email et code requis.' };

    var cache  = CacheService.getScriptCache();
    var otpKey = 'otp_' + email.replace(/[^a-z0-9]/g, '_');
    var stored = cache.get(otpKey);

    if (!stored) return { error: 'Code expiré. Demandez un nouveau code.' };
    if (stored !== code) {
      Logger.log('OTP invalide pour ' + email + ' — reçu: ' + code);
      return { error: 'Code incorrect. Vérifiez le code reçu par e-mail.' };
    }

    cache.remove(otpKey); // code à usage unique
    Logger.log('✅ Connexion admin via OTP: ' + email);
    return { success: true, email: email };
  } catch (err) {
    Logger.log('handleVerifyOTP error: ' + err.message);
    return { error: err.message };
  }
}
