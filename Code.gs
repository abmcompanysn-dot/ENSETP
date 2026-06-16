// ============================================================
//  ENSETP GALA 2026 — Google Apps Script Backend
//  Déployer comme "Application Web" (accès : Tout le monde)
//  URL obtenue → coller dans le panel Admin du site
// ============================================================

// ── CONFIGURATION ─────────────────────────────────────────
var CFG = {
  COMMANDES_SHEET:  'Commandes',
  PRESENCES_SHEET:  'Présences',
  ADMINS_SHEET:     'Sous-Admins',
  SHEET_ID:         '',                     // Laisser vide = feuille liée au script
  EMAIL_FROM:       'gala@ensetp.edu.sn',   // Adresse expéditeur (alias Gmail)
  ADMIN_EMAIL:      'contact@mahu.cards',   // Reçoit toutes les notifications
  EVENT_NAME:       'Dîner de Gala de Fin d\'Année – ENSETP 2026',
  EVENT_DATE:       'Samedi 20 Juin 2026',
  EVENT_LIEU:       'Au Magic Land',
  PAYDUNIA_KEY:     '',                     // Clé API Paydunia (depuis dashboard)
  PAYDUNIA_SECRET:  '',
  PAYDUNIA_URL:     'https://paydunia.com/api/v1/payment/init',
  CALLBACK_URL:     ''                      // URL de votre site déployé
};

// ── MENU GOOGLE SHEETS ─────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎩 Gala ENSETP 2026')
    .addItem('⚙️  Initialiser les feuilles',        'setupSheet')
    .addSeparator()
    .addItem('📊 Rapport quotidien (e-mail)',        'sendDailyReport')
    .addItem('⏰ Activer rappel quotidien (8h)',     'setupDailyReminder')
    .addSeparator()
    .addItem('🧪 Test — Envoyer un ticket bidon',   'testSendTicket')
    .addItem('🧹 Supprimer les triggers existants', 'clearTriggers')
    .addToUi();
}

// ── INITIALISATION DES FEUILLES ────────────────────────────
function setupSheet() {
  var ss = CFG.SHEET_ID
    ? SpreadsheetApp.openById(CFG.SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  _initCommandes(ss);
  _initPresences(ss);
  _initAdmins(ss);

  SpreadsheetApp.getUi().alert(
    '✅ Initialisation terminée !\n\n' +
    'Trois feuilles créées :\n' +
    '• Commandes — tickets achetés\n' +
    '• Présences — confirmations de présence\n' +
    '• Sous-Admins — admins délégués\n\n' +
    'Vous pouvez maintenant déployer ce script\net coller l\'URL dans le panel Admin du site.'
  );
}

function _initCommandes(ss) {
  var sh = ss.getSheetByName(CFG.COMMANDES_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.COMMANDES_SHEET);

  // En-têtes
  var headers = ['ID Ticket','Prénom','Nom','E-mail','Téléphone','Type','Qté','Prix Unit.','Total','Paiement','Date','Statut','Email Envoyé'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#D4AF37')
    .setFontColor('#000000')
    .setFontSize(10);
  sh.setFrozenRows(1);

  // Largeurs colonnes
  sh.setColumnWidth(1, 200);  // ID
  sh.setColumnWidth(4, 200);  // E-mail
  sh.setColumnWidth(11, 160); // Date
  sh.setColumnWidths(2, 2, 120); // Prénom / Nom

  // Alternance couleur — appliquée aux 500 premières lignes
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=MOD(ROW(),2)=0')
    .setBackground('#1a1a1a')
    .setRanges([sh.getRange('A2:M500')])
    .build();
  sh.setConditionalFormatRules([rule]);

  Logger.log('Feuille Commandes prête.');
}

function _initPresences(ss) {
  var sh = ss.getSheetByName(CFG.PRESENCES_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.PRESENCES_SHEET);

  var headers = ['ID Ticket','Prénom','Nom','E-mail','Confirmé le','IP / Agent'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1A6B35')
    .setFontColor('#FFFFFF')
    .setFontSize(10);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(4, 200);
  sh.setColumnWidth(5, 160);

  Logger.log('Feuille Présences prête.');
}

function _initAdmins(ss) {
  var sh = ss.getSheetByName(CFG.ADMINS_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.ADMINS_SHEET);

  var headers = ['Nom','E-mail','Rôle','Date Ajout'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#2C2C54')
    .setFontColor('#FFFFFF')
    .setFontSize(10);
  sh.setFrozenRows(1);
  sh.setColumnWidth(2, 200);

  Logger.log('Feuille Sous-Admins prête.');
}

// ── POINT D'ENTRÉE HTTP ────────────────────────────────────
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;
    var result  = {};

    if      (action === 'createOrder')       result = handleCreateOrder(payload.order);
    else if (action === 'payduniaWebhook')   result = handlePayduniaWebhook(payload);
    else if (action === 'confirmPresence')   result = handleConfirmPresence(payload);
    else if (action === 'addSubAdmin')       result = handleAddSubAdmin(payload);
    else if (action === 'getStats')          result = getStats();
    else                                     result = { error: 'Action inconnue: ' + action };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  if (e.parameter && e.parameter.payment_ref) {
    handlePayduniaCallback(e.parameter);
    return HtmlService.createHtmlOutput('<p>Paiement traité. Fermez cette fenêtre.</p>');
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'ENSETP Gala API active', version: '2.0' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── CRÉATION DE COMMANDE ───────────────────────────────────
function handleCreateOrder(order) {
  logToSheet(order);

  if (CFG.PAYDUNIA_KEY) {
    var payLink = createPayduniaPayment(order);
    if (payLink) return { success: true, paymentUrl: payLink, orderId: order.id };
  }

  order.status = 'paid';
  updateOrderStatus(order.id, 'paid');
  sendTicketEmail(order);
  sendAdminNotification(order);

  return { success: true, orderId: order.id, mode: 'demo' };
}

// ── CONFIRMATION DE PRÉSENCE ───────────────────────────────
function handleConfirmPresence(payload) {
  try {
    var ss = CFG.SHEET_ID
      ? SpreadsheetApp.openById(CFG.SHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();

    var sh = ss.getSheetByName(CFG.PRESENCES_SHEET);
    if (!sh) { _initPresences(ss); sh = ss.getSheetByName(CFG.PRESENCES_SHEET); }

    // Vérifier si déjà confirmé
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

    // Notifier contact@mahu.cards
    GmailApp.sendEmail(
      CFG.ADMIN_EMAIL,
      '[GALA ENSETP] Présence confirmée – ' + payload.orderId,
      'Nouvelle confirmation de présence :\n\n'
        + 'Ticket : ' + payload.orderId + '\n'
        + 'Nom    : ' + (payload.name || '–') + '\n'
        + 'E-mail : ' + (payload.email || '–') + '\n'
        + 'Le    : ' + new Date().toLocaleString('fr-FR')
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
    var ss = CFG.SHEET_ID
      ? SpreadsheetApp.openById(CFG.SHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CFG.ADMINS_SHEET);
    if (!sh) { _initAdmins(ss); sh = ss.getSheetByName(CFG.ADMINS_SHEET); }
    sh.appendRow([payload.name || '', payload.email || '', payload.role || 'readonly', new Date()]);
    return { success: true };
  } catch (err) {
    return { error: err.message };
  }
}

// ── PAYDUNIA ───────────────────────────────────────────────
function createPayduniaPayment(order) {
  try {
    var payload = {
      merchant_key:   CFG.PAYDUNIA_KEY,
      amount:         order.total,
      currency:       'XOF',
      order_id:       order.id,
      customer_name:  order.prenom + ' ' + order.nom,
      customer_email: order.email,
      customer_phone: order.tel,
      description:    'Ticket Gala ENSETP 2026 – ' + order.type.toUpperCase(),
      callback_url:   CFG.CALLBACK_URL + '?order_id=' + order.id,
      return_url:     CFG.CALLBACK_URL + '?payment_status=success&order_id=' + order.id,
      cancel_url:     CFG.CALLBACK_URL + '?payment_status=cancel&order_id=' + order.id
    };
    var resp = UrlFetchApp.fetch(CFG.PAYDUNIA_URL, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    return data.payment_url || data.url || null;
  } catch (e) {
    Logger.log('Paydunia error: ' + e.message);
    return null;
  }
}

function handlePayduniaWebhook(data) {
  var orderId = data.order_id || data.transaction_ref;
  var status  = data.status;
  if (status === 'SUCCESS' || status === 'success') {
    updateOrderStatus(orderId, 'paid');
    var order = getOrderById(orderId);
    if (order) { sendTicketEmail(order); sendAdminNotification(order); }
    return { success: true };
  }
  return { success: false, reason: 'Statut non confirmé: ' + status };
}

function handlePayduniaCallback(params) {
  var orderId = params.order_id;
  if (params.payment_status === 'success') {
    updateOrderStatus(orderId, 'paid');
    var order = getOrderById(orderId);
    if (order) { sendTicketEmail(order); sendAdminNotification(order); }
  }
}

// ── GOOGLE SHEETS — COMMANDES ──────────────────────────────
function getSheet() {
  var ss = CFG.SHEET_ID
    ? SpreadsheetApp.openById(CFG.SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.COMMANDES_SHEET);
  if (!sh) {
    _initCommandes(ss);
    sh = ss.getSheetByName(CFG.COMMANDES_SHEET);
  }
  return sh;
}

function logToSheet(order) {
  var sh = getSheet();
  sh.appendRow([
    order.id,
    order.prenom,
    order.nom,
    order.email,
    order.tel || '',
    order.type.toUpperCase(),
    order.qty  || 1,
    order.price || 0,
    order.total || 0,
    order.payment || '',
    new Date(order.date || new Date()),
    order.status || 'pending',
    'Non'
  ]);
}

function updateOrderStatus(orderId, status) {
  var sh   = getSheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      sh.getRange(i + 1, 12).setValue(status === 'paid' ? 'PAYÉ' : status);
      if (status === 'paid') sh.getRange(i + 1, 13).setValue('Oui');
      return;
    }
  }
}

function getOrderById(orderId) {
  var sh   = getSheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      return {
        id: data[i][0], prenom: data[i][1], nom: data[i][2],
        email: data[i][3], tel: data[i][4], type: data[i][5],
        qty: data[i][6], price: data[i][7], total: data[i][8],
        payment: data[i][9], date: data[i][10], status: data[i][11]
      };
    }
  }
  return null;
}

function getStats() {
  var sh   = getSheet();
  var data = sh.getDataRange().getValues();
  var total = 0, std = 0, vip = 0, table = 0, recettes = 0;
  for (var i = 1; i < data.length; i++) {
    total++;
    var t = String(data[i][5]).toLowerCase();
    if (t === 'solo')   std++;
    else if (t === 'couple') vip++;
    recettes += Number(data[i][8]) || 0;
  }
  return { total: total, solo: std, couple: vip, recettes: recettes };
}

// ── EMAILS ─────────────────────────────────────────────────
function sendTicketEmail(order) {
  var name  = order.prenom + ' ' + order.nom;
  var total = Number(order.total).toLocaleString() + ' FCFA';
  var type  = String(order.type).toUpperCase();
  var body  = buildEmailBody(name, type, total, order.id, buildTicketHtml(order));

  GmailApp.sendEmail(
    order.email,
    '🎉 Votre Ticket – Dîner de Gala ENSETP 2026 [' + order.id + ']',
    stripHtml(body),
    { name: 'Gala ENSETP 2026', htmlBody: body, replyTo: CFG.ADMIN_EMAIL }
  );
  Logger.log('Ticket envoyé à ' + order.email);
}

function sendAdminNotification(order) {
  var body = 'Nouvelle commande reçue :\n\n'
    + 'ID      : ' + order.id + '\n'
    + 'Nom     : ' + order.prenom + ' ' + order.nom + '\n'
    + 'E-mail  : ' + order.email + '\n'
    + 'Tél.    : ' + (order.tel || '–') + '\n'
    + 'Type    : ' + order.type.toUpperCase() + '\n'
    + 'Qté     : ' + (order.qty || 1) + '\n'
    + 'Total   : ' + Number(order.total).toLocaleString() + ' FCFA\n'
    + 'Paiement: ' + (order.payment || '–') + '\n'
    + 'Date    : ' + new Date(order.date).toLocaleString('fr-FR');
  GmailApp.sendEmail(CFG.ADMIN_EMAIL, '[GALA ENSETP] Nouvelle commande – ' + order.id, body);
}

// ── TICKET HTML (e-mail) ───────────────────────────────────
function buildTicketHtml(order) {
  return '<div style="background:linear-gradient(135deg,#1a1a1a,#0d0d0d);border:2px solid #D4AF37;border-radius:14px;overflow:hidden;font-family:Arial,sans-serif;max-width:460px;">'
    + '<div style="background:#0A0A0A;padding:16px 22px;border-bottom:1px solid rgba(212,175,55,0.3);display:flex;justify-content:space-between;align-items:center;">'
    +   '<span style="font-size:1.1rem;font-weight:bold;color:#D4AF37;letter-spacing:3px;">ENSETP</span>'
    +   '<span style="font-size:0.6rem;letter-spacing:2px;color:#D4AF37;text-transform:uppercase;">Dîner de Gala · 2026</span>'
    + '</div>'
    + '<div style="padding:20px 22px;">'
    +   '<div style="font-size:1.8rem;font-style:italic;color:#D4AF37;margin-bottom:4px;">Gala</div>'
    +   '<div style="font-size:0.6rem;letter-spacing:4px;text-transform:uppercase;color:#aaa;margin-bottom:16px;">De Fin d\'Année · ENSETP 2026</div>'
    +   '<table width="100%" style="color:#fff;">'
    +     '<tr>'
    +       '<td style="padding:6px 0;"><div style="font-size:0.55rem;letter-spacing:2px;text-transform:uppercase;color:#888">Titulaire</div><div style="font-size:0.9rem;font-weight:600">' + order.prenom + ' ' + order.nom + '</div></td>'
    +       '<td style="padding:6px 0;"><div style="font-size:0.55rem;letter-spacing:2px;text-transform:uppercase;color:#888">Type</div><div><span style="background:linear-gradient(135deg,#D4AF37,#A07820);color:#000;font-size:0.6rem;font-weight:700;letter-spacing:2px;padding:3px 10px;border-radius:18px;">' + order.type.toUpperCase() + '</span></div></td>'
    +     '</tr>'
    +     '<tr>'
    +       '<td style="padding:6px 0;"><div style="font-size:0.55rem;letter-spacing:2px;text-transform:uppercase;color:#888">E-mail</div><div style="font-size:0.78rem;color:#ccc">' + order.email + '</div></td>'
    +       '<td style="padding:6px 0;"><div style="font-size:0.55rem;letter-spacing:2px;text-transform:uppercase;color:#888">Montant</div><div style="color:#D4AF37;font-weight:bold">' + Number(order.total).toLocaleString() + ' FCFA</div></td>'
    +     '</tr>'
    +   '</table>'
    +   '<hr style="border:none;border-top:1px dashed rgba(212,175,55,0.3);margin:16px 0;" />'
    +   '<table width="100%" style="color:#fff;">'
    +     '<tr>'
    +       '<td><div style="font-size:0.55rem;letter-spacing:2px;text-transform:uppercase;color:#888">Date</div><div style="font-weight:600">' + CFG.EVENT_DATE + '</div></td>'
    +       '<td><div style="font-size:0.55rem;letter-spacing:2px;text-transform:uppercase;color:#888">Lieu</div><div style="font-weight:600">' + CFG.EVENT_LIEU + '</div></td>'
    +     '</tr>'
    +   '</table>'
    + '</div>'
    + '<div style="background:rgba(212,175,55,0.06);padding:12px 22px;border-top:1px dashed rgba(212,175,55,0.3);">'
    +   '<div style="font-family:monospace;font-size:0.72rem;color:#D4AF37;letter-spacing:2px;">' + order.id + '</div>'
    +   '<div style="font-size:0.6rem;color:#888;margin-top:3px;">Émis le ' + new Date(order.date || new Date()).toLocaleDateString('fr-FR') + '</div>'
    + '</div>'
    + '</div>';
}

function buildEmailBody(name, type, total, id, ticketHtml) {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="background:#0A0A0A;margin:0;padding:30px;font-family:Arial,sans-serif;">'
    + '<div style="max-width:560px;margin:0 auto;">'
    +   '<div style="text-align:center;margin-bottom:30px;">'
    +     '<div style="font-size:1.4rem;font-weight:bold;color:#D4AF37;letter-spacing:4px;margin-bottom:6px;">ENSETP</div>'
    +     '<div style="font-size:0.7rem;color:#888;letter-spacing:2px;text-transform:uppercase;">École Normale Supérieure d\'Enseignement Technique et Professionnel</div>'
    +   '</div>'
    +   '<div style="background:#1A1A1A;border:1px solid rgba(212,175,55,0.2);border-radius:16px;padding:32px;margin-bottom:24px;">'
    +     '<h2 style="color:#D4AF37;font-size:1.2rem;margin:0 0 14px">Bonjour ' + name + ' ! 🎉</h2>'
    +     '<p style="color:#ccc;font-size:0.9rem;line-height:1.7;margin:0 0 20px">Votre ticket pour le <strong style="color:#F5D66A">Dîner de Gala de Fin d\'Année ENSETP 2026</strong> a bien été reçu.</p>'
    +     '<table style="width:100%;border-collapse:collapse;margin-bottom:24px;">'
    +       '<tr><td style="padding:8px 14px;background:#222;color:#888;font-size:0.72rem;text-transform:uppercase;letter-spacing:1px;">Type de ticket</td><td style="padding:8px 14px;background:#1A1A1A;color:#fff;font-weight:600;border-left:2px solid #D4AF37">' + type + '</td></tr>'
    +       '<tr><td style="padding:8px 14px;background:#222;color:#888;font-size:0.72rem;text-transform:uppercase;letter-spacing:1px;">Montant payé</td><td style="padding:8px 14px;background:#1A1A1A;color:#D4AF37;font-weight:bold;border-left:2px solid #D4AF37">' + total + '</td></tr>'
    +       '<tr><td style="padding:8px 14px;background:#222;color:#888;font-size:0.72rem;text-transform:uppercase;letter-spacing:1px;">Date</td><td style="padding:8px 14px;background:#1A1A1A;color:#fff;border-left:2px solid #D4AF37">' + CFG.EVENT_DATE + '</td></tr>'
    +       '<tr><td style="padding:8px 14px;background:#222;color:#888;font-size:0.72rem;text-transform:uppercase;letter-spacing:1px;">Lieu</td><td style="padding:8px 14px;background:#1A1A1A;color:#fff;border-left:2px solid #D4AF37">' + CFG.EVENT_LIEU + '</td></tr>'
    +     '</table>'
    +     '<p style="color:#888;font-size:0.8rem;margin:0">Référence : <strong style="color:#D4AF37">' + id + '</strong></p>'
    +   '</div>'
    +   '<div style="margin-bottom:24px;">'
    +     '<h3 style="color:#D4AF37;font-size:0.88rem;letter-spacing:2px;text-transform:uppercase;margin-bottom:16px;">Votre E-Ticket</h3>'
    +     ticketHtml
    +   '</div>'
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
    + '</div>'
    + '</body></html>';
}

function stripHtml(html) {
  return html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
}

// ── RAPPORT QUOTIDIEN ─────────────────────────────────────
function setupDailyReminder() {
  clearTriggers();
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased().atHour(8).everyDays(1).create();
  SpreadsheetApp.getUi().alert('✅ Rappel quotidien activé (tous les jours à 8h).');
}

function clearTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendDailyReport') ScriptApp.deleteTrigger(t);
  });
}

function sendDailyReport() {
  var stats = getStats();

  // Compter les présences
  var ss = CFG.SHEET_ID
    ? SpreadsheetApp.openById(CFG.SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  var ps = ss.getSheetByName(CFG.PRESENCES_SHEET);
  var presences = ps ? Math.max(0, ps.getLastRow() - 1) : 0;

  var body = '📊 Rapport quotidien — Gala ENSETP 2026\n\n'
    + '📅 Date : ' + new Date().toLocaleString('fr-FR') + '\n\n'
    + '🎫 TICKETS\n'
    + '  Total vendus : ' + stats.total + '\n'
    + '  Standard     : ' + stats.standard + '\n'
    + '  VIP          : ' + stats.vip + '\n'
    + '  Table entière: ' + stats.table + '\n\n'
    + '✅ PRÉSENCES\n'
    + '  Confirmées   : ' + presences + '\n\n'
    + '💰 RECETTES\n'
    + '  Total        : ' + stats.recettes.toLocaleString() + ' FCFA\n\n'
    + '—\nEnvoyé par Mahu Events · mahu.cards';

  GmailApp.sendEmail(
    CFG.ADMIN_EMAIL,
    '[GALA ENSETP] Rapport du ' + new Date().toLocaleDateString('fr-FR'),
    body
  );
}

// ── TEST ───────────────────────────────────────────────────
function testSendTicket() {
  var testOrder = {
    id:       'TK-ENSETP-2026-TEST-ABCDE',
    prenom:   'Cheikh',
    nom:      'Diallo',
    email:    Session.getActiveUser().getEmail(),
    tel:      '+221 77 000 00 00',
    type:     'standard',
    qty:      1,
    price:    1000,
    total:    1000,
    type:     'solo',
    payment:  'wave',
    date:     new Date().toISOString(),
    status:   'paid'
  };
  logToSheet(testOrder);
  sendTicketEmail(testOrder);
  sendAdminNotification(testOrder);
  Logger.log('✅ Test ticket envoyé à ' + testOrder.email + ' et copie à ' + CFG.ADMIN_EMAIL);
  SpreadsheetApp.getUi().alert('✅ Test envoyé !\nVérifiez votre boîte mail (' + testOrder.email + ') et contact@mahu.cards.');
}
