/**
 * Module: sheets.gs
 * Purpose: Sheet data access and batch operations.
 * Public API:
 *   - getCheckRegisterUnmatched() -> {ok, data: CheckRegisterItemDTO[]}
 *   - getBankCsvUnmatched() -> {ok, data: BankCsvItemDTO[]}
 *   - getAvailableTransactionIds() -> {ok, data: string[]}
 * Dependencies: util.gs
 * Invariants: All reads are batched; sheet structure is validated.
 */

/**
 * Get unmatched items from Check Register sheet.
 * @returns {{ok: boolean, data?: CheckRegisterItemDTO[], error?: {code: string, message: string}}}
 * @example
 *   const res = getCheckRegisterUnmatched();
 *   if (res.ok) console.log('Found', res.data.length, 'unmatched items');
 */
function getCheckRegisterUnmatched() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const checkSheet = sheets.find(s => s.getName().includes('Check Register'));
    
    if (!checkSheet) {
      return { 
        ok: false, 
        error: { code: 'sheet-not-found', message: 'Check Register sheet not found' } 
      };
    }
    
    const lastRow = checkSheet.getLastRow();
    if (lastRow < 2) return { ok: true, data: [] };
    
    const data = checkSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const unmatched = [];
    
    data.forEach((row, index) => {
      const [date, checkNumber, description, withdrawal, deposit, balance, status] = row;
      
      // Skip if already matched or cleared
      if (status && (status === 'Cleared' || status === 'Matched' || status.toString().match(/^\d+$/))) {
        return;
      }
      
      // Skip beginning balance entries
      if (description && description.toString().toLowerCase().includes('beginning balance')) {
        return;
      }
      
      unmatched.push({
        row: index + 2,
        date: new Date(date),
        checkNumber: checkNumber || '',
        description: (description || '').toString().trim(),
        withdrawal: validateNumber(withdrawal),
        deposit: validateNumber(deposit),
        balance: validateNumber(balance),
        status: status || '',
        aging: calculateItemAging(new Date(date))
      });
    });
    
    return { ok: true, data: unmatched };
    
  } catch (err) {
    return { 
      ok: false, 
      error: { code: 'read-failed', message: String(err) } 
    };
  }
}

/**
 * Get unmatched items from Bank CSV sheet.
 * @returns {{ok: boolean, data?: BankCsvItemDTO[], error?: {code: string, message: string}}}
 * @example
 *   const res = getBankCsvUnmatched();
 *   if (res.ok) console.log('Found', res.data.length, 'bank items');
 */
function getBankCsvUnmatched() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Bank CSV');
    
    if (!sheet) {
      return { 
        ok: false, 
        error: { code: 'sheet-not-found', message: 'Bank CSV sheet not found' } 
      };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, data: [] };
    
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const matchedIds = _getMatchedTransactionIds();
    const unmatched = [];
    
    data.forEach((row, index) => {
      const [transactionId, date, description, amount, runningBalance] = row;
      
      if (!transactionId || matchedIds.includes(transactionId.toString())) {
        return;
      }
      
      unmatched.push({
        row: index + 2,
        transactionId: transactionId.toString(),
        date: new Date(date),
        description: (description || '').toString().trim(),
        amount: validateNumber(amount),
        runningBalance: validateNumber(runningBalance)
      });
    });
    
    return { ok: true, data: unmatched };
    
  } catch (err) {
    return { 
      ok: false, 
      error: { code: 'read-failed', message: String(err) } 
    };
  }
}

/**
 * Get available Transaction IDs for dropdown validation.
 * @returns {{ok: boolean, data?: string[], error?: {code: string, message: string}}}
 */
function getAvailableTransactionIds() {
  try {
    const bankData = getBankCsvUnmatched();
    if (!bankData.ok) return bankData;
    
    const ids = bankData.data.map(item => item.transactionId).sort();
    return { ok: true, data: ids };
    
  } catch (err) {
    return { 
      ok: false, 
      error: { code: 'ids-fetch-failed', message: String(err) } 
    };
  }
}

/**
 * Get Check Register data including matched items.
 * @returns {{ok: boolean, data?: CheckRegisterItemDTO[], error?: {code: string, message: string}}}
 */
function getCheckRegisterData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const checkSheet = sheets.find(s => s.getName().includes('Check Register'));
    
    if (!checkSheet) {
      return { 
        ok: false, 
        error: { code: 'sheet-not-found', message: 'Check Register sheet not found' } 
      };
    }
    
    const lastRow = checkSheet.getLastRow();
    if (lastRow < 2) return { ok: true, data: [] };
    
    const data = checkSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const items = [];
    
    data.forEach((row, index) => {
      const [date, checkNumber, description, withdrawal, deposit, balance, status] = row;
      
      items.push({
        row: index + 2,
        date: new Date(date),
        checkNumber: checkNumber || '',
        description: (description || '').toString().trim(),
        withdrawal: validateNumber(withdrawal),
        deposit: validateNumber(deposit),
        balance: validateNumber(balance),
        status: status || ''
      });
    });
    
    return { ok: true, data: items };
    
  } catch (err) {
    return { 
      ok: false, 
      error: { code: 'read-failed', message: String(err) } 
    };
  }
}

/**
 * Get sheet by name with error handling.
 * @param {string} sheetName
 * @returns {{ok: boolean, data?: Sheet, error?: {code: string, message: string}}}
 */
function getSheetByName(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet;
    
    if (sheetName === 'Check Register') {
      const sheets = ss.getSheets();
      sheet = sheets.find(s => s.getName().includes('Check Register'));
    } else {
      sheet = ss.getSheetByName(sheetName);
    }
    
    if (!sheet) {
      return { 
        ok: false, 
        error: { code: 'sheet-not-found', message: `Sheet "${sheetName}" not found` } 
      };
    }
    
    return { ok: true, data: sheet };
    
  } catch (err) {
    return { 
      ok: false, 
      error: { code: 'access-failed', message: String(err) } 
    };
  }
}

/**
 * Find Bank CSV row by Transaction ID.
 * @param {string} transactionId
 * @returns {number} Row number or 0 if not found
 */
function findBankCsvRowByTransactionId(transactionId) {
  try {
    const sheet = getSheetByName('Bank CSV');
    if (!sheet.ok) return 0;
    
    const data = sheet.data.getRange('A:A').getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === transactionId) {
        return i + 1;
      }
    }
    return 0;
    
  } catch (err) {
    return 0;
  }
}

// === Private Implementation ===

/**
 * Get list of matched Transaction IDs from Check Register.
 * @returns {Array<string>}
 */
function _getMatchedTransactionIds() {
  try {
    const checkData = getCheckRegisterData();
    if (!checkData.ok) return [];
    
    return checkData.data
      .filter(item => item.status && item.status.toString().match(/^\d+$/))
      .map(item => item.status.toString());
      
  } catch (err) {
    return [];
  }
}

/**
 * Calculate aging for an item based on date.
 * @param {Date} itemDate
 * @returns {string}
 */
function calculateItemAging(itemDate) {
  const today = new Date();
  const diffTime = today - itemDate;
  const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
  
  if (diffDays < 30) return 'Current';
  if (diffDays < 60) return '1 month';
  if (diffDays < 90) return '2 months';
  return '3+ months ⚠️';
}