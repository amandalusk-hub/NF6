// ============================================================
// PLAID UI HELPERS — Sidebar, credential setup, removal
// API logic (getPlaidConfig_, getPlaidLinkToken, exchangePlaidToken,
// syncPlaidAccounts) lives in Code.gs.
// ============================================================

// ===== CREDENTIALS SETUP =====

function setPlaidCredentials() {
  var ui = SpreadsheetApp.getUi();
  var clientId = ui.prompt('Plaid Setup', 'Enter your Plaid Client ID:', ui.ButtonSet.OK_CANCEL);
  if (clientId.getSelectedButton() !== ui.Button.OK) return;
  var secret = ui.prompt('Plaid Setup', 'Enter your Plaid Secret:', ui.ButtonSet.OK_CANCEL);
  if (secret.getSelectedButton() !== ui.Button.OK) return;
  var env = ui.prompt('Plaid Setup', 'Environment (sandbox / development / production):', ui.ButtonSet.OK_CANCEL);
  if (env.getSelectedButton() !== ui.Button.OK) return;
  var props = PropertiesService.getScriptProperties();
  props.setProperty('PLAID_CLIENT_ID', clientId.getResponseText().trim());
  props.setProperty('PLAID_SECRET', secret.getResponseText().trim());
  props.setProperty('PLAID_ENV', env.getResponseText().trim() || 'sandbox');
  ui.alert('Plaid credentials saved. You can now connect bank accounts.');
}

// ===== PLAID LINK SIDEBAR =====

function openPlaidLink() {
  var cfg = getPlaidConfig_();
  if (!cfg.clientId || !cfg.secret) {
    var ui   = SpreadsheetApp.getUi();
    var resp = ui.alert(
      'Plaid Not Configured',
      'Plaid credentials are not set. Would you like to set them now?',
      ui.ButtonSet.YES_NO
    );
    if (resp === ui.Button.YES) setPlaidCredentials();
    return;
  }
  // Token is fetched client-side via google.script.run inside the sidebar
  var html = HtmlService.createHtmlOutputFromFile('PlaidLink')
    .setTitle('Connect Bank Account')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function handlePlaidSuccess(publicToken) {
  try {
    var exchResult = exchangePlaidToken(publicToken);
    if (!exchResult.success) return { success: false, message: exchResult.error };
    var syncResult = syncPlaidAccounts();
    var count = syncResult.synced || 0;
    return { success: true, message: 'Bank connected! ' + count + ' account(s) synced to Assets tab.' };
  } catch(e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ===== STATEMENTS (PDF documents) =====

// Menu wrapper: download new statement PDFs to Drive and report the result.
function syncPlaidStatementsMenu() {
  var ui = SpreadsheetApp.getUi();
  var cfg = getPlaidConfig_();
  if (!cfg.clientId || !cfg.secret) {
    ui.alert('Plaid is not configured. Run "Configure Plaid Credentials" first.');
    return;
  }
  var res = syncPlaidStatements();
  if (!res.success && res.error) { ui.alert('Statements', res.error, ui.ButtonSet.OK); return; }
  var msg = res.message || '';
  if (res.errors && res.errors.length) msg += '\n\nFailed:\n' + res.errors.join('\n');
  if (res.folderUrl) msg += '\n\nFolder: ' + res.folderUrl;
  ui.alert('Plaid Statements', msg, ui.ButtonSet.OK);
}

// Set the end-user name the institution sees for statement requests.
function setPlaidStatementsEndUser() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Plaid Statements',
    'Enter the account holder name for statement requests (e.g. the JP Morgan account owner):',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var name = resp.getResponseText().trim();
  if (!name) { ui.alert('No name entered.'); return; }
  PropertiesService.getScriptProperties().setProperty('PLAID_STATEMENTS_END_USER', name);
  ui.alert('Saved. Reconnect the institution so statements are enabled on the connection.');
}

// ===== REMOVE CONNECTION =====

function removePlaidConnection() {
  var ui    = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var tokens = JSON.parse(props.getProperty('PLAID_TOKENS') || '[]');
  if (!tokens.length) {
    ui.alert('No Plaid connections to remove.');
    return;
  }
  var resp = ui.alert(
    'Remove Plaid Connections',
    'This will disconnect all ' + tokens.length + ' bank connection(s). Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp === ui.Button.YES) {
    props.deleteProperty('PLAID_TOKENS');
    ui.alert('All Plaid connections removed.');
  }
}
