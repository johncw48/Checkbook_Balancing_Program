/**
 * Module: matching.gs
 * Purpose: Auto-matching algorithms and validation logic.
 * Public API:
 *   - runAutoMatchingProcess(criteria) -> {ok, data: MatchResultDTO}
 *   - validateAndApplyManualMatch(transactionId, row) -> {ok, data: string}
 * Dependencies: sheets.gs, util.gs
 * Invariants: All matching preserves data integrity; uses Jaro-Winkler similarity.
 */

// ===== PROTECTED COMPONENT: THREE-TIER AUTO-MATCHING SYSTEM =====
// STATUS: BATTLE-TESTED - DO NOT MODIFY WITHOUT EXPLICIT PERMISSION
// PROBLEM SOLVED: Real bank data has timing mismatches, description variations, and processing delays
// WHY COMPLEX: Simple exact matching fails due to:
//   - Check dates vs bank posting dates differ by days/weeks
//   - Transaction descriptions vary between systems ("ABC SUPPLY" vs "ABC Supply Co")
//   - Processing delays require progressive date tolerance
//   - Must prevent duplicate matches across tiers
//   - Amount direction validation (withdrawals vs deposits vs negative amounts)
// ALGORITHM: Tier 1 (exact date) → Tier 2 (±7 days) → Tier 3 (±30 days)
// ===== END PROTECTED COMPONENT =====

/**
 * Execute three-tier auto-matching process.
 * @param {{enableTier1?: boolean, enableTier2?: boolean, enableTier3?: boolean}} criteria
 * @returns {{ok: boolean, data?: MatchResultDTO, error?: {code: string, message: string}}}
 * @example
 *   const res = runAutoMatchingProcess({enableTier1: true, enableTier2: false});
 */
function runAutoMatchingProcess(criteria = {}) {
  const lock = LockService.getDocumentLock();
  
  try {
    lock.waitLock(10000);
    
    const checkData = getCheckRegisterUnmatched();
    const bankData = getBankCsvUnmatched();
    
    if (!checkData.ok) return checkData;
    if (!bankData.ok) return bankData;
    
    const matches = {
      tier1: criteria.enableTier1 !== false ? _executeTier1Matching(checkData.data, bankData.data) : [],
      tier2: criteria.enableTier2 !== false ? _executeTier2Matching(checkData.data, bankData.data) : [],
      tier3: criteria.enableTier3 !== false ? _executeTier3Matching(checkData.data, bankData.data) : []
    };
    
    _applyMatches(matches);
    
    return {
      ok: true,
      data: {
        matches,
        summary: {
          tier1Count: matches.tier1.length,
          tier2Count: matches.tier2.length,
          tier3Count: matches.tier3.length,
          totalMatches: matches.tier1.length + matches.tier2.length + matches.tier3.length,
          executedAt: new Date().toISOString()
        }
      }
    };
    
  } catch (err) {
    return { 
      ok: false, 
      error: { code: 'auto-match-failed', message: String(err) } 
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Validate and apply manual match entry.
 * @param {string} transactionId
 * @param {number} checkRegisterRow
 * @returns {{ok: boolean, data?: string, error?: {code: string, message: string}}}
 */
function validateAndApplyManualMatch(transactionId, checkRegisterRow) {
  try {
    const validation = _validateManualMatch(transactionId);
    if (!validation.ok) return validation;
    
    const result = _applyManualMatch(transactionId, checkRegisterRow);
    if (!result.ok) return result;
    
    return { ok: true, data: 'Manual match applied successfully' };
    
  } catch (err) {
    return { 
      ok: false, 
      error: { code: 'manual-validation-failed', message: String(err) } 
    };
  }
}

// === Private Implementation ===

/**
 * Execute Tier 1 matching: exact date + amount + description similarity > 0.8.
 * @param {Array} checkItems
 * @param {Array} bankItems
 * @returns {Array<MatchDTO>}
 */
function _executeTier1Matching(checkItems, bankItems) {
  const matches = [];
  const usedBankItems = new Set();
  
  checkItems.forEach(checkItem => {
    if (checkItem.status) return; // Already matched
    
    bankItems.forEach(bankItem => {
      if (usedBankItems.has(bankItem.transactionId)) return;
      
      if (_isExactDateMatch(checkItem.date, bankItem.date) &&
          _isExactAmountMatch(checkItem, bankItem) &&
          _validateAmountTypes(checkItem, bankItem) &&
          jaroWinklerSimilarity(checkItem.description, bankItem.description) > 0.8) {
        
        matches.push({
          tier: 1,
          checkItem,
          bankItem,
          similarity: jaroWinklerSimilarity(checkItem.description, bankItem.description)
        });
        usedBankItems.add(bankItem.transactionId);
      }
    });
  });
  
  return matches;
}

/**
 * Execute Tier 2 matching: date Â±7 days + amount + description similarity > 0.8.
 * @param {Array} checkItems
 * @param {Array} bankItems
 * @returns {Array<MatchDTO>}
 */
function _executeTier2Matching(checkItems, bankItems) {
  const matches = [];
  const usedBankItems = new Set();
  
  // Get already matched bank items from Tier 1
  const tier1BankIds = _getMatchedBankTransactionIds();
  tier1BankIds.forEach(id => usedBankItems.add(id));
  
  checkItems.forEach(checkItem => {
    if (checkItem.status) return; // Already matched
    
    bankItems.forEach(bankItem => {
      if (usedBankItems.has(bankItem.transactionId)) return;
      
      if (_isDateWithinRange(checkItem.date, bankItem.date, 7) &&
          _isExactAmountMatch(checkItem, bankItem) &&
          _validateAmountTypes(checkItem, bankItem) &&
          jaroWinklerSimilarity(checkItem.description, bankItem.description) > 0.8) {
        
        matches.push({
          tier: 2,
          checkItem,
          bankItem,
          similarity: jaroWinklerSimilarity(checkItem.description, bankItem.description)
        });
        usedBankItems.add(bankItem.transactionId);
      }
    });
  });
  
  return matches;
}

/**
 * Execute Tier 3 matching: date Â±30 days + amount + description similarity > 0.8.
 * @param {Array} checkItems
 * @param {Array} bankItems
 * @returns {Array<MatchDTO>}
 */
function _executeTier3Matching(checkItems, bankItems) {
  const matches = [];
  const usedBankItems = new Set();
  
  // Get already matched bank items from Tier 1 and 2
  const matchedBankIds = _getMatchedBankTransactionIds();
  matchedBankIds.forEach(id => usedBankItems.add(id));
  
  checkItems.forEach(checkItem => {
    if (checkItem.status) return; // Already matched
    
    bankItems.forEach(bankItem => {
      if (usedBankItems.has(bankItem.transactionId)) return;
      
      if (_isDateWithinRange(checkItem.date, bankItem.date, 30) &&
          _isExactAmountMatch(checkItem, bankItem) &&
          _validateAmountTypes(checkItem, bankItem) &&
          jaroWinklerSimilarity(checkItem.description, bankItem.description) > 0.8) {
        
        matches.push({
          tier: 3,
          checkItem,
          bankItem,
          similarity: jaroWinklerSimilarity(checkItem.description, bankItem.description)
        });
        usedBankItems.add(bankItem.transactionId);
      }
    });
  });
  
  return matches;
}

/**
 * Apply matches to both Check Register and reconciliation sheet.
 * @param {Object} matches
 */
function _applyMatches(matches) {
  const checkSheet = getSheetByName('Check Register');
  const reconSheet = getSheetByName('Reconciliation Worksheet');
  
  if (!checkSheet.ok || !reconSheet.ok) return;
  
  const layout = getCurrentLayout();
  
  // Apply Tier 1 matches
  matches.tier1.forEach((match, index) => {
    _updateCheckRegisterStatus(match.checkItem.row, match.bankItem.transactionId);
    _addMatchToReconSheet(reconSheet.data, layout, match, index + 5); // Start after headers
  });
  
  // Apply Tier 2 matches
  matches.tier2.forEach((match, index) => {
    _updateCheckRegisterStatus(match.checkItem.row, match.bankItem.transactionId);
    _addMatchToReconSheet(reconSheet.data, layout, match, index + 5 + matches.tier1.length);
  });
  
  // Apply Tier 3 matches
  matches.tier3.forEach((match, index) => {
    _updateCheckRegisterStatus(match.checkItem.row, match.bankItem.transactionId);
    _addMatchToReconSheet(reconSheet.data, layout, match, index + 5 + matches.tier1.length + matches.tier2.length);
  });
}

/**
 * Validate manual match transaction ID.
 * @param {string} transactionId
 * @returns {{ok: boolean, error?: {code: string, message: string}}}
 */
function _validateManualMatch(transactionId) {
  if (!transactionId || typeof transactionId !== 'string') {
    return { ok: false, error: { code: 'invalid-id', message: 'Transaction ID required' } };
  }
  
  const bankData = getBankCsvUnmatched();
  if (!bankData.ok) return bankData;
  
  const bankItem = bankData.data.find(item => item.transactionId === transactionId);
  if (!bankItem) {
    return { 
      ok: false, 
      error: { code: 'id-not-found', message: `Transaction ID ${transactionId} not found in unmatched items` } 
    };
  }
  
  const usedIds = _getMatchedBankTransactionIds();
  if (usedIds.includes(transactionId)) {
    return { 
      ok: false, 
      error: { code: 'id-already-used', message: `Transaction ID ${transactionId} already matched` } 
    };
  }
  
  return { ok: true };
}

/**
 * Apply manual match to sheets.
 * @param {string} transactionId
 * @param {number} checkRegisterRow
 * @returns {{ok: boolean, error?: {code: string, message: string}}}
 */
function _applyManualMatch(transactionId, checkRegisterRow) {
  try {
    _updateCheckRegisterStatus(checkRegisterRow, transactionId);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: { code: 'apply-failed', message: String(err) } };
  }
}

/**
 * Check if dates match exactly.
 * @param {Date} date1
 * @param {Date} date2
 * @returns {boolean}
 */
function _isExactDateMatch(date1, date2) {
  return formatDate(date1) === formatDate(date2);
}

/**
 * Check if date is within range.
 * @param {Date} date1
 * @param {Date} date2
 * @param {number} daysTolerance
 * @returns {boolean}
 */
function _isDateWithinRange(date1, date2, daysTolerance) {
  const diffDays = Math.abs((date1 - date2) / (1000 * 60 * 60 * 24));
  return diffDays <= daysTolerance;
}

/**
 * Check if amounts match exactly.
 * @param {Object} checkItem
 * @param {Object} bankItem
 * @returns {boolean}
 */
function _isExactAmountMatch(checkItem, bankItem) {
  const checkAmount = checkItem.withdrawal || checkItem.deposit || 0;
  return Math.abs(Math.abs(checkAmount) - Math.abs(bankItem.amount)) < 0.01;
}

/**
 * Validate amount types align properly.
 * @param {Object} checkItem
 * @param {Object} bankItem
 * @returns {boolean}
 */
function _validateAmountTypes(checkItem, bankItem) {
  const isCheckWithdrawal = checkItem.withdrawal && checkItem.withdrawal > 0;
  const isCheckDeposit = checkItem.deposit && checkItem.deposit > 0;
  const isBankNegative = bankItem.amount < 0;
  const isBankPositive = bankItem.amount > 0;
  
  return (isCheckWithdrawal && isBankNegative) || (isCheckDeposit && isBankPositive);
}

/**
 * Get list of already matched bank transaction IDs.
 * @returns {Array<string>}
 */
function _getMatchedBankTransactionIds() {
  const checkData = getCheckRegisterData();
  if (!checkData.ok) return [];
  
  return checkData.data
    .filter(item => item.status && item.status.match(/^\d+$/))
    .map(item => item.status);
}

/**
 * Update Check Register status column.
 * @param {number} row
 * @param {string} transactionId
 */
function _updateCheckRegisterStatus(row, transactionId) {
  const sheet = getSheetByName('Check Register');
  if (sheet.ok && row > 0) {
    sheet.data.getRange(row, 7).setValue(transactionId); // Column G (Status)
  }
}

/**
 * Add match to reconciliation sheet Section 1.
 * @param {Sheet} sheet
 * @param {Object} layout
 * @param {Object} match
 * @param {number} targetRow
 */
function _addMatchToReconSheet(sheet, layout, match, targetRow) {
  const row = layout.section1.start + targetRow;
  const data = [
    match.tier,
    formatDate(match.checkItem.date),
    match.checkItem.checkNumber || '',
    match.checkItem.description,
    match.checkItem.withdrawal || '',
    match.checkItem.deposit || '',
    'Matched',
    match.bankItem.transactionId,
    formatDate(match.bankItem.date),
    match.bankItem.description,
    match.bankItem.amount < 0 ? Math.abs(match.bankItem.amount) : '',
    match.bankItem.amount > 0 ? match.bankItem.amount : '',
    'Clear'
  ];
  
  sheet.getRange(row, 1, 1, 13).setValues([data]);
  
  // Apply tier-based color coding
  const bgColor = match.tier === 1 ? '#E8F5E8' : '#E3F2FD';
  sheet.getRange(row, 1, 1, 13).setBackground(bgColor);
}

/**
 * Get current layout from reconciliation sheet.
 * @returns {Object}
 */
function getCurrentLayout() {
  return calculateLayout();
}