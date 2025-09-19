/**
 * Module: webapp.gs
 * Purpose: Entry point and routing for reconciliation system.
 * Public API: doGet(), include(), onOpen()
 * Dependencies: reconciliation.gs, matching.gs
 * Notes: Adds viewport meta for responsive UI.
 */

/**
 * Serve the main reconciliation UI.
 * @returns {HtmlOutput}
 */
function doGet() {
  const tpl = HtmlService.createTemplateFromFile('index');
  const out = tpl.evaluate()
    .setTitle('Bank Reconciliation System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return out;
}

/**
 * Include an HTML partial by name.
 * @param {string} name
 * @returns {string}
 */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/**
 * Add reconciliation menu when spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Reconciliation')
    .addItem('Open Reconciliation System', 'openReconciliationDialog')
    .addToUi();
}

/**
 * Open reconciliation dialog from spreadsheet menu.
 * @returns {{ok: boolean, data?: string}}
 */
function openReconciliationDialog() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('index')
      .setWidth(1400)
      .setHeight(900);
    SpreadsheetApp.getUi().showModalDialog(html, 'Bank Reconciliation System');
    return { ok: true, data: 'dialog-opened' };
  } catch (err) {
    return { 
      ok: false, 
      error: { code: 'dialog-failed', message: String(err) } 
    };
  }
}

/**
 * Build reconciliation worksheet with complete structure.
 * @returns {{ok: boolean, data?: LayoutDTO, error?: {code: string, message: string}}}
 * @example
 *   const res = buildReconciliationWorksheet();
 *   if (res.ok) console.log('Built at:', res.data);
 */
function buildReconciliationWorksheet() {
  try {
    return createReconciliationSheet();
  } catch (err) {
    return { 
      ok: false, 
      error: { code: 'build-failed', message: String(err) } 
    };
  }
}

/**
 * Execute three-tier automated matching process.
 * @param {{enableTier1?: boolean, enableTier2?: boolean, enableTier3?: boolean}} criteria
 * @returns {{ok: boolean, data?: MatchResultDTO, error?: {code: string, message: string}}}
 * @example
 *   const res = executeAutoMatch({enableTier1: true});
 */
function executeAutoMatch(criteria = {}) {
  try {
    return runAutoMatchingProcess(criteria);
  } catch (err) {
    return { 
      ok: false, 
      error: { code: 'match-failed', message: String(err) } 
    };
  }
}

/**
 * Process manual match validation and application.
 * @param {string} transactionId
 * @param {number} checkRegisterRow
 * @returns {{ok: boolean, data?: string, error?: {code: string, message: string}}}
 */
function processManualMatch(transactionId, checkRegisterRow) {
  try {
    return validateAndApplyManualMatch(transactionId, checkRegisterRow);
  } catch (err) {
    return { 
      ok: false, 
      error: { code: 'manual-match-failed', message: String(err) } 
    };
  }
}

/**
 * Get current reconciliation status and data.
 * @returns {{ok: boolean, data?: ReconciliationStatusDTO, error?: {code: string, message: string}}}
 */
function getReconciliationStatus() {
  try {
    return getCurrentReconciliationData();
  } catch (err) {
    return { 
      ok: false, 
      error: { code: 'status-failed', message: String(err) } 
    };
  }
}

/**
 * Finalize month-end reconciliation process.
 * @returns {{ok: boolean, data?: CloseResultDTO, error?: {code: string, message: string}}}
 */
function closeMonth() {
  try {
    return processMonthEndClose();
  } catch (err) {
    return { 
      ok: false, 
      error: { code: 'close-failed', message: String(err) } 
    };
  }
}