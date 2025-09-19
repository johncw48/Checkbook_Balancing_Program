/**
 * Module: util.gs
 * Purpose: Utility functions for reconciliation operations.
 * Public API:
 *   - jaroWinklerSimilarity(s1, s2) -> number
 *   - formatCurrency(amount) -> string
 *   - formatDate(date) -> string
 *   - validateNumber(value) -> number
 * Dependencies: None
 * Invariants: All functions are pure and deterministic.
 */

// ===== PROTECTED COMPONENT: JARO-WINKLER SIMILARITY ALGORITHM =====
// STATUS: BATTLE-TESTED - DO NOT MODIFY WITHOUT EXPLICIT PERMISSION
// PROBLEM SOLVED: Transaction descriptions never match exactly between check register and bank
// WHY COMPLEX: Simple string comparison fails because:
//   - Case differences ("ABC Supply" vs "ABC SUPPLY")
//   - Abbreviations vs full names ("WALMART" vs "Wal-Mart Stores")
//   - Extra characters ("CHECK #1234 TO ABC" vs "ABC SUPPLY")
//   - Requires similarity threshold tuning (0.8 works best after testing)
// ALGORITHM: Jaro distance + common prefix weighting for fuzzy string matching
// ===== END PROTECTED COMPONENT =====

/**
 * Calculate Jaro-Winkler similarity for description matching.
 * @param {string} s1
 * @param {string} s2
 * @returns {number} Similarity score between 0 and 1
 * @example
 *   const similarity = jaroWinklerSimilarity("ABC SUPPLY", "ABC Supply Co");
 *   if (similarity > 0.8) console.log('Good match');
 */
function jaroWinklerSimilarity(s1, s2) {
  if (!s1 || !s2) return 0;
  if (s1 === s2) return 1;
  
  s1 = s1.toUpperCase().trim();
  s2 = s2.toUpperCase().trim();
  
  const jaro = _calculateJaro(s1, s2);
  const prefix = _commonPrefixLength(s1, s2, 4);
  
  return jaro + (0.1 * prefix * (1 - jaro));
}

/**
 * Format currency consistently.
 * @param {number} amount
 * @returns {string}
 * @example
 *   formatCurrency(1234.56) // returns "$1,234.56"
 */
function formatCurrency(amount) {
  if (typeof amount !== 'number' || !isFinite(amount)) return '$0.00';
  return '$' + amount.toLocaleString('en-US', { 
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

/**
 * Format date consistently.
 * @param {Date|string} date
 * @returns {string}
 * @example
 *   formatDate(new Date()) // returns "01/15/2025"
 */
function formatDate(date) {
  const d = new Date(date);
  if (isNaN(d.getTime())) return '';
  
  const month = (d.getMonth() + 1).toString().padStart(2, '0');
  const day = d.getDate().toString().padStart(2, '0');
  const year = d.getFullYear();
  
  return `${month}/${day}/${year}`;
}

/**
 * Validate numeric input.
 * @param {*} value
 * @returns {number}
 */
function validateNumber(value) {
  const num = parseFloat(value);
  return isFinite(num) ? num : 0;
}

// === Private Implementation ===

/**
 * Calculate Jaro similarity.
 * @param {string} s1
 * @param {string} s2
 * @returns {number}
 */
function _calculateJaro(s1, s2) {
  const len1 = s1.length;
  const len2 = s2.length;
  
  if (len1 === 0 && len2 === 0) return 1;
  if (len1 === 0 || len2 === 0) return 0;
  
  const matchWindow = Math.floor(Math.max(len1, len2) / 2) - 1;
  if (matchWindow < 0) return 0;
  
  const s1Matches = new Array(len1).fill(false);
  const s2Matches = new Array(len2).fill(false);
  let matches = 0;
  
  // Find matches
  for (let i = 0; i < len1; i++) {
    const start = Math.max(0, i - matchWindow);
    const end = Math.min(i + matchWindow + 1, len2);
    
    for (let j = start; j < end; j++) {
      if (s2Matches[j] || s1[i] !== s2[j]) continue;
      s1Matches[i] = true;
      s2Matches[j] = true;
      matches++;
      break;
    }
  }
  
  if (matches === 0) return 0;
  
  // Count transpositions
  let transpositions = 0;
  let k = 0;
  
  for (let i = 0; i < len1; i++) {
    if (!s1Matches[i]) continue;
    while (!s2Matches[k]) k++;
    if (s1[i] !== s2[k]) transpositions++;
    k++;
  }
  
  return (matches / len1 + matches / len2 + (matches - transpositions / 2) / matches) / 3;
}

/**
 * Calculate common prefix length.
 * @param {string} s1
 * @param {string} s2
 * @param {number} maxLength
 * @returns {number}
 */
function _commonPrefixLength(s1, s2, maxLength) {
  let prefix = 0;
  const minLength = Math.min(s1.length, s2.length, maxLength);
  
  for (let i = 0; i < minLength; i++) {
    if (s1[i] === s2[i]) {
      prefix++;
    } else {
      break;
    }
  }
  
  return prefix;
}