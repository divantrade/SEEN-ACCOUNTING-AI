// =============================================
// Fuzzy Search - Arabic Text Matching
// Migrated from FuzzySearch.js
// =============================================

// Arabic text normalization
function normalizeArabic(text) {
  if (!text) return '';
  return text
    .replace(/[إأآا]/g, 'ا')
    .replace(/ة/g, 'ه')
    .replace(/ى/g, 'ي')
    .replace(/ؤ/g, 'و')
    .replace(/ئ/g, 'ي')
    .replace(/[َُِّْـًٌٍ]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

// Calculate similarity score between two strings (0-1)
function similarity(str1, str2) {
  const s1 = normalizeArabic(str1);
  const s2 = normalizeArabic(str2);

  if (s1 === s2) return 1;
  if (!s1 || !s2) return 0;

  // Exact contains
  if (s1.includes(s2) || s2.includes(s1)) return 0.9;

  // Levenshtein distance
  const maxLen = Math.max(s1.length, s2.length);
  if (maxLen === 0) return 1;

  const distance = levenshtein(s1, s2);
  return 1 - distance / maxLen;
}

function levenshtein(a, b) {
  const matrix = [];
  for (let i = 0; i <= b.length; i++) matrix[i] = [i];
  for (let j = 0; j <= a.length; j++) matrix[0][j] = j;

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b[i - 1] === a[j - 1]) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }
  return matrix[b.length][a.length];
}

/**
 * Search items by fuzzy matching
 * @param {string} query - Search query
 * @param {Array} items - Array of items to search
 * @param {string} key - Property name to search in
 * @param {number} threshold - Minimum similarity score (0-1)
 * @returns {Array} Sorted results with scores
 */
export function fuzzySearch(query, items, key = 'name', threshold = 0.4) {
  if (!query || !items?.length) return [];

  const results = items
    .map((item) => {
      const value = typeof item === 'string' ? item : item[key];
      const score = similarity(query, value);
      return { item, score };
    })
    .filter((r) => r.score >= threshold)
    .sort((a, b) => b.score - a.score);

  return results;
}

export { normalizeArabic, similarity };
