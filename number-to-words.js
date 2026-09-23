/**
 * Converts a number to words using the Indian numbering system (Lakh/Crore).
 * Output format: "Twenty Three Thousand Six Hundred Only"
 */

const ones = [
  '', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine',
  'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen',
  'Seventeen', 'Eighteen', 'Nineteen'
];

const tens = [
  '', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'
];

function convertBelowHundred(n) {
  if (n < 20) return ones[n];
  const t = Math.floor(n / 10);
  const o = n % 10;
  return tens[t] + (o ? ' ' + ones[o] : '');
}

function convertBelowThousand(n) {
  if (n < 100) return convertBelowHundred(n);
  const h = Math.floor(n / 100);
  const rem = n % 100;
  return ones[h] + ' Hundred' + (rem ? ' ' + convertBelowHundred(rem) : '');
}

/**
 * Convert an integer to Indian-system words.
 * Indian grouping: ones → thousands → lakhs → crores
 *   1,00,00,000 = 1 Crore
 *   1,00,000    = 1 Lakh
 *   1,000       = 1 Thousand
 *
 * @param {number} num - A non-negative integer
 * @returns {string} e.g. "Twenty Three Thousand Six Hundred Only"
 */
function numberToWordsIndian(num) {
  if (num === 0) return 'Zero Only';

  num = Math.floor(num);

  let words = '';

  // Crores (1,00,00,000+)
  const crores = Math.floor(num / 10000000);
  if (crores > 0) {
    words += convertBelowHundred(crores) + ' Crore ';
    num %= 10000000;
  }

  // Lakhs (1,00,000 – 99,99,999)
  const lakhs = Math.floor(num / 100000);
  if (lakhs > 0) {
    words += convertBelowHundred(lakhs) + ' Lakh ';
    num %= 100000;
  }

  // Thousands (1,000 – 99,999)
  const thousands = Math.floor(num / 1000);
  if (thousands > 0) {
    words += convertBelowHundred(thousands) + ' Thousand ';
    num %= 1000;
  }

  // Hundreds and below
  if (num > 0) {
    words += convertBelowThousand(num);
  }

  return words.trim() + ' Only';
}
