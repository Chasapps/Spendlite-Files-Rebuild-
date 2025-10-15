// ============================================================================
// SPENDLITE V6.6.27 - Personal Expense Tracker
// ============================================================================
// This application helps you track and categorize your spending from bank CSV files.
// Main features:
// - Import CSV transactions from your bank
// - Automatically categorize expenses using custom rules
// - Filter by month and category
// - Export totals and rules for backup
// ============================================================================

// ============================================================================
// SECTION 1: CONSTANTS AND CONFIGURATION
// ============================================================================
// These are values that don't change during the app's runtime

// CSV Column mapping - tells us where to find data in the bank's CSV file
// These are 0-based indices (counting starts from 0, not 1)
const COL = { 
  DATE: 2,      // Column 3 contains the transaction date
  DEBIT: 5,     // Column 6 contains the amount (money spent or received)
  LONGDESC: 9   // Column 10 contains the description of the transaction
};

// Pagination settings - how many items to show per page
const PAGE_SIZE = 10;              // Number of transactions per page
const CATEGORY_PAGE_SIZE = 10;     // Number of categories per page (if used)

// LocalStorage keys - these are like labels for saving data in the browser
// localStorage is a browser feature that saves data even after you close the tab
const LS_KEYS = { 
  RULES: 'spendlite_rules_v6626',           // Key for saving categorization rules
  FILTER: 'spendlite_filter_v6626',         // Key for saving active category filter
  MONTH: 'spendlite_month_v6627',           // Key for saving selected month
  TXNS_COLLAPSED: 'spendlite_txns_collapsed_v7',  // Key for saving show/hide state
  TXNS_JSON: 'spendlite_txns_json_v7'       // Key for saving all transactions
};

// Sample rules shown when user first loads the app
const SAMPLE_RULES = `# Rules format: KEYWORD => CATEGORY
`;

// ============================================================================
// SECTION 2: APPLICATION STATE (GLOBAL VARIABLES)
// ============================================================================
// These variables hold the current state of the application
// "let" means the value can change, unlike "const" which is fixed

let CURRENT_TXNS = [];        // Array holding all loaded transactions
let CURRENT_RULES = [];       // Array holding all categorization rules
let CURRENT_FILTER = null;    // Currently active category filter (null = show all)
let MONTH_FILTER = "";        // Currently selected month ('YYYY-MM' format or empty)
let CURRENT_PAGE = 1;         // Current page number for transaction pagination
let CATEGORY_PAGE = 1;        // Current page for category display (if used)

// ============================================================================
// SECTION 3: DATE FORMATTING FUNCTIONS
// ============================================================================
// These functions help display dates in a user-friendly way

/**
 * Converts a year-month string to a friendly display format
 * Example: "2025-06" becomes "June 2025"
 * @param {string} ym - Year-month in 'YYYY-MM' format
 * @returns {string} Friendly month name and year
 */
function formatMonthLabel(ym) {
  if (!ym) return 'All months';
  
  // Split "2025-06" into year and month parts
  const [y, m] = ym.split('-').map(Number);
  
  // Create a date object (month is 0-based, so subtract 1)
  const date = new Date(y, m - 1, 1);
  
  // Use browser's built-in formatting to get "June 2025"
  return date.toLocaleString(undefined, { month: 'long', year: 'numeric' });
}

/**
 * Returns a friendly label, handling both month strings and edge cases
 * @param {string} label - The month label to format
 * @returns {string} User-friendly label
 */
function friendlyMonthOrAll(label) {
  if (!label) return 'All months';
  
  // Check if it matches YYYY-MM format using regex
  if (/^\d{4}-\d{2}$/.test(label)) return formatMonthLabel(label);
  
  return String(label);
}

/**
 * Converts a friendly label to a filename-safe version
 * Example: "June 2025" becomes "June_2025"
 * @param {string} label - The label to convert
 * @returns {string} Filename-safe string
 */
function forFilename(label) {
  // Replace all whitespace with underscores
  return String(label).replace(/\s+/g, '_');
}

// ============================================================================
// SECTION 4: TEXT PROCESSING UTILITIES
// ============================================================================

/**
 * Converts text to Title Case (First Letter Of Each Word Capitalized)
 * Example: "COFFEE SHOP" becomes "Coffee Shop"
 * @param {string} str - The string to convert
 * @returns {string} Title-cased string
 */
function toTitleCase(str) {
  if (!str) return '';
  
  return String(str)
    .toLowerCase()                    // First make everything lowercase
    .replace(/[_-]+/g, ' ')          // Replace underscores and dashes with spaces
    .replace(/\s+/g, ' ')            // Replace multiple spaces with single space
    .trim()                           // Remove leading/trailing spaces
    .replace(/\b([a-z])/g, (m, p1) => p1.toUpperCase());  // Capitalize first letter of each word
}

/**
 * Safely converts a string to a number (handles currency symbols, commas, etc.)
 * Example: "$1,234.56" becomes 1234.56
 * @param {string|number} s - The value to parse
 * @returns {number} Parsed number or 0 if invalid
 */
function parseAmount(s) {
  if (s == null) return 0;
  
  // Remove everything except digits, minus sign, and decimal point
  // Then remove commas (which are used as thousands separators)
  s = String(s).replace(/[^\d\-,.]/g, '').replace(/,/g, '');
  
  // Convert to number, or return 0 if it fails
  return Number(s) || 0;
}

/**
 * Escapes HTML special characters to prevent display issues and security problems
 * Example: "<script>" becomes "&lt;script&gt;"
 * @param {string} s - The string to escape
 * @returns {string} HTML-safe string
 */
function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')    // & must be first
    .replace(/</g, '&lt;')     // < becomes &lt;
    .replace(/>/g, '&gt;')     // > becomes &gt;
    .replace(/"/g, '&quot;')   // " becomes &quot;
    .replace(/'/g, '&#039;');  // ' becomes &#039;
}

// ============================================================================
// SECTION 5: DATE PARSING (AUSTRALIAN FORMAT SUPPORT)
// ============================================================================

/**
 * Intelligently parses various date formats, with Australian DD/MM/YYYY support
 * This is important because Australian banks use DD/MM/YYYY while US uses MM/DD/YYYY
 * @param {string} s - The date string to parse
 * @returns {Date|null} Date object or null if parsing fails
 */
function parseDateSmart(s) {
  if (!s) return null;
  const str = String(s).trim();
  let m;  // Will hold regex match results

  // Pattern 1: ISO format (unambiguous): YYYY-MM-DD or YYYY/MM/DD
  m = str.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (m) {
    // m[1] = year, m[2] = month, m[3] = day
    // Note: JavaScript months are 0-based (0=January, 11=December)
    return new Date(+m[1], +m[2]-1, +m[3]);
  }

  // Pattern 2: Australian format DD/MM/YYYY (e.g., 1/6/2025 = 1 June 2025)
  m = str.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) {
    // m[1] = day, m[2] = month, m[3] = year (Australian order)
    return new Date(+m[3], +m[2]-1, +m[1]);
  }

  // Pattern 3: Month name format (e.g., "Mon 1 September, 2025")
  // First strip any leading time like "3:45pm "
  const s2 = str.replace(/^\d{1,2}:\d{2}\s*(am|pm)\s*/i, '');
  
  // Match: optional weekday, day number, month name, year
  m = s2.match(/^(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)?\s*(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December),?\s+(\d{4})/i);
  
  if (m) {
    const day = +m[1];
    const monthName = m[2].toLowerCase();
    const y = +m[3];
    
    // Map month names to numbers (0-based for JavaScript)
    const monthMap = {
      january: 0, february: 1, march: 2, april: 3, 
      may: 4, june: 5, july: 6, august: 7, 
      september: 8, october: 9, november: 10, december: 11
    };
    
    const monthIndex = monthMap[monthName];
    if (monthIndex != null) return new Date(y, monthIndex, day);
  }

  // Pattern 4: Couldn't parse - give up
  // Note: We don't use JavaScript's native Date() parser because it assumes
  // US format (MM/DD/YYYY) which would incorrectly parse Australian dates
  return null;
}

/**
 * Converts a Date object to YYYY-MM format
 * Example: June 1, 2025 becomes "2025-06"
 * @param {Date} d - The date to convert
 * @returns {string} Date in YYYY-MM format
 */
function yyyymm(d) { 
  // padStart(2,'0') ensures month is always 2 digits (e.g., "06" not "6")
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`; 
}

/**
 * Gets the month of the first transaction in the list
 * @param {Array} txns - Array of transactions (defaults to CURRENT_TXNS)
 * @returns {string|null} Month in YYYY-MM format or null if no transactions
 */
function getFirstTxnMonth(txns = CURRENT_TXNS) {
  if (!txns.length) return null;
  
  const d = parseDateSmart(txns[0].date);
  if (!d || isNaN(d)) return null;
  
  return yyyymm(d);
}

// ============================================================================
// SECTION 6: CSV LOADING AND TRANSACTION PARSING
// ============================================================================

/**
 * Parses CSV text and loads transactions into memory
 * This is the main entry point when a user uploads a CSV file
 * @param {string} csvText - Raw CSV file content as text
 * @returns {Array} Array of parsed transactions
 */
function loadCsvText(csvText) {
  // Use PapaParse library to convert CSV text into array of arrays
  // skipEmptyLines: true removes blank rows
  const rows = Papa.parse(csvText.trim(), { skipEmptyLines: true }).data;
  
  // Determine if first row is a header by checking if the amount column contains a number
  // If it's NOT a number (like "Amount"), it's a header row - skip it
  const startIdx = rows.length && isNaN(parseAmount(rows[0][COL.DEBIT])) ? 1 : 0;
  
  const txns = [];
  
  // Loop through each row starting from startIdx
  for (let i = startIdx; i < rows.length; i++) {
    const r = rows[i];
    
    // Skip if row is missing or doesn't have enough columns
    if (!r || r.length < 10) continue;
    
    // Extract data from specific columns
    const effectiveDate = r[COL.DATE] || '';
    const debit = parseAmount(r[COL.DEBIT]);
    const longDesc = (r[COL.LONGDESC] || '').trim();
    
    // Only include rows with valid data (has date or description, and non-zero amount)
    if ((effectiveDate || longDesc) && Number.isFinite(debit) && debit !== 0) {
       txns.push({ 
         date: effectiveDate, 
         amount: debit, 
         description: longDesc 
       });
    }
  }
  
  // Store transactions in global state
  CURRENT_TXNS = txns;
  
  // Save to browser's localStorage for persistence
  saveTxnsToLocalStorage();
  
  // Update UI elements
  try { updateMonthBanner(); } catch {}
  rebuildMonthDropdown();
  applyRulesAndRender();
  
  return txns;
}

// ============================================================================
// SECTION 7: MONTH FILTERING
// ============================================================================

/**
 * Rebuilds the month dropdown with all unique months from transactions
 * This runs after CSV is loaded to populate the month filter options
 */
function rebuildMonthDropdown() {
  const sel = document.getElementById('monthFilter');
  
  // Use a Set to collect unique months (Sets automatically remove duplicates)
  const months = new Set();
  
  // Loop through all transactions and extract their month
  for (const t of CURRENT_TXNS) {
    const d = parseDateSmart(t.date);
    if (d) months.add(yyyymm(d));  // Add month in YYYY-MM format
  }
  
  // Convert Set to Array and sort chronologically
  const list = Array.from(months).sort();
  
  const current = MONTH_FILTER;
  
  // Build dropdown HTML
  sel.innerHTML = `<option value="">All months</option>` + 
    list.map(m => `<option value="${m}">${formatMonthLabel(m)}</option>`).join('');
  
  // Restore previously selected month (if it still exists)
  sel.value = current && list.includes(current) ? current : "";
  
  updateMonthBanner();
}

/**
 * Returns transactions filtered by the selected month
 * @returns {Array} Filtered array of transactions
 */
function monthFilteredTxns() {
  // If no month filter is active, return all transactions
  if (!MONTH_FILTER) return CURRENT_TXNS;
  
  // Filter transactions to only those matching the selected month
  return CURRENT_TXNS.filter(t => {
    const d = parseDateSmart(t.date);
    return d && yyyymm(d) === MONTH_FILTER;
  });
}

// ============================================================================
// SECTION 8: CATEGORIZATION RULES
// ============================================================================

/**
 * Parses rules text into an array of rule objects
 * Rule format: "KEYWORD => CATEGORY" (one per line)
 * @param {string} text - Raw rules text from textarea
 * @returns {Array} Array of {keyword, category} objects
 */
function parseRules(text) {
  // Split text into lines (handles both Windows \r\n and Unix \n line endings)
  const lines = String(text || "").split(/\r?\n/);
  const rules = [];
  
  for (const line of lines) {
    const trimmed = line.trim();
    
    // Skip empty lines and comments (lines starting with #)
    if (!trimmed || trimmed.startsWith('#')) continue;
    
    // Split on "=>" to separate keyword from category
    const parts = trimmed.split(/=>/i);  // 'i' makes it case-insensitive
    
    if (parts.length >= 2) {
      const keyword = parts[0].trim().toLowerCase();   // Keywords are lowercase for matching
      const category = parts[1].trim().toUpperCase();  // Categories are uppercase for consistency
      
      if (keyword && category) {
        rules.push({ keyword, category });
      }
    }
  }
  
  return rules;
}

/**
 * Checks if a transaction description matches a keyword
 * Supports multi-word keywords (e.g., "paypal pypl" matches both "paypal" AND "pypl")
 * @param {string} descLower - Transaction description in lowercase
 * @param {string} keywordLower - Keyword to match (in lowercase)
 * @returns {boolean} True if description matches keyword
 */
function matchesKeyword(descLower, keywordLower) {
  if (!keywordLower) return false;
  
  const text = String(descLower || '').toLowerCase();
  
  // Split keyword into individual tokens (words)
  const tokens = String(keywordLower).toLowerCase().split(/\s+/).filter(Boolean);
  if (!tokens.length) return false;
  
  // Define what counts as a word boundary
  // Letters, digits, &, ., and _ are "word characters"
  // Everything else is a boundary
  const delim = '[^A-Za-z0-9&._]';
  
  // Check that ALL tokens appear in the description with word boundaries
  return tokens.every(tok => {
    // Escape special regex characters in the token
    const safe = tok.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    
    // Create regex: token must appear at word boundary
    // (?:^|delim) = start of string OR delimiter before
    // (?:delim|$) = delimiter after OR end of string
    const re = new RegExp(`(?:^|${delim})${safe}(?:${delim}|$)`, 'i');
    
    return re.test(text);
  });
}

/**
 * Applies categorization rules to transactions
 * Modifies the transactions in place, adding a 'category' property
 * @param {Array} txns - Array of transactions to categorize
 * @param {Array} rules - Array of categorization rules
 */
function categorise(txns, rules) {
  for (const t of txns) {
    // Get description (handles both 'desc' and 'description' property names)
    const descLower = String(t.desc || t.description || "").toLowerCase();
    const amount = Math.abs(Number(t.amount || t.debit || 0));

    // Step 1: Find first matching rule
    let matched = null;
    for (const r of rules) {
      if (matchesKeyword(descLower, r.keyword)) {
        matched = r.category;
        break;  // First match wins (stop looking)
      }
    }

    // Step 2: Special case - small purchases at petrol stations are likely coffee
    // If categorized as PETROL but amount is $2 or less, recategorize as COFFEE
    if (matched && String(matched).toUpperCase() === "PETROL" && amount <= 2) {
      matched = "COFFEE";
    }

    // Assign category (or "UNCATEGORISED" if no rules matched)
    t.category = matched || "UNCATEGORISED";
  }
}

// ============================================================================
// SECTION 9: CATEGORY TOTALS CALCULATION AND DISPLAY
// ============================================================================

/**
 * Calculates total spending for each category
 * @param {Array} txns - Array of categorized transactions
 * @returns {Object} Object with {rows, grand} - rows are [category, total] pairs
 */
function computeCategoryTotals(txns) {
  // Use a Map to accumulate totals by category
  // Map is better than Object for this because keys can be any value
  const byCat = new Map();
  
  for (const t of txns) {
    const cat = (t.category || 'UNCATEGORISED').toUpperCase();
    
    // Add this transaction's amount to the category total
    // If category doesn't exist yet, start at 0
    byCat.set(cat, (byCat.get(cat) || 0) + t.amount);
  }
  
  // Convert Map to array of [category, total] pairs and sort by total (highest first)
  const rows = [...byCat.entries()].sort((a, b) => b[1] - a[1]);
  
  // Calculate grand total across all categories
  const grand = rows.reduce((acc, [, v]) => acc + v, 0);
  
  return { rows, grand };
}

/**
 * Renders the category totals table in the UI
 * @param {Array} txns - Array of transactions to summarize
 */
function renderCategoryTotals(txns) {
  const { rows, grand } = computeCategoryTotals(txns);
  const totalsDiv = document.getElementById('categoryTotals');
  
  // Build HTML table
  let html = '<table class="cats">';
  html += '<colgroup><col class="col-cat"><col class="col-total"><col class="col-pct"></colgroup>';
  html += '<thead><tr><th>Category</th><th class="num">Total</th><th class="num">%</th></tr></thead>';
  html += '<tbody>';
  
  // Add a row for each category
  for (const [cat, total] of rows) {
    // Calculate percentage of grand total
    const pct = grand ? (total / grand * 100) : 0;
    
    html += `<tr>
      <td><a class="catlink" data-cat="${escapeHtml(cat)}"><span class="category-name">${escapeHtml(toTitleCase(cat))}</span></a></td>
      <td class="num">${total.toFixed(2)}</td>
      <td class="num">${pct.toFixed(1)}%</td>
    </tr>`;
  }
  
  html += `</tbody>`;
  
  // Footer with grand total
  html += `<tfoot><tr><td>Total</td><td class="num">${grand.toFixed(2)}</td><td class="num">100%</td></tr></tfoot>`;
  html += '</table>';
  
  totalsDiv.innerHTML = html;

  // Add click handlers to category links (for filtering)
  totalsDiv.querySelectorAll('a.catlink').forEach(a => {
    a.addEventListener('click', () => {
      // Set the clicked category as the active filter
      CURRENT_FILTER = a.getAttribute('data-cat');
      
      // Save filter to localStorage
      try { localStorage.setItem(LS_KEYS.FILTER, CURRENT_FILTER || ''); } catch {}
      
      // Update UI to show active filter
      updateFilterUI();
      CURRENT_PAGE = 1;  // Reset to first page
      renderTransactionsTable();
    });
  });
}

/**
 * Renders the month summary (count, debit, credit, net)
 */
function renderMonthTotals() {
  // Get transactions (filtered by both month and category)
  const txns = getFilteredTxns(monthFilteredTxns());
  
  let debit = 0, credit = 0, count = 0;
  
  for (const t of txns) {
    const amt = Number(t.amount) || 0;
    
    if (amt > 0) {
      debit += amt;  // Positive = money spent
    } else {
      credit += Math.abs(amt);  // Negative = money received
    }
    
    count++;
  }
  
  const net = debit - credit;  // Net spending
  
  const el = document.getElementById('monthTotals');
  if (el) {
    const label = friendlyMonthOrAll(MONTH_FILTER);
    const cat = CURRENT_FILTER ? ` + category "${CURRENT_FILTER}"` : "";
    
    el.innerHTML = `Showing <span class="badge">${count}</span> transactions for <strong>${label}${cat}</strong> · ` +
                   `Debit: <strong>$${debit.toFixed(2)}</strong> · ` +
                   `Credit: <strong>$${credit.toFixed(2)}</strong> · ` +
                   `Net: <strong>$${net.toFixed(2)}</strong>`;
  }
}

// ============================================================================
// SECTION 10: MAIN RENDER FUNCTION
// ============================================================================

/**
 * Applies rules to transactions and re-renders all UI elements
 * This is the main "refresh" function called when data changes
 * @param {Object} options - Options object
 * @param {boolean} options.keepPage - If true, stay on current page (don't reset to page 1)
 */
function applyRulesAndRender({keepPage = false} = {}) { 
  if (!keepPage) {
    CURRENT_PAGE = 1;  // Reset to first page
  }
  
  // Parse rules from textarea
  CURRENT_RULES = parseRules(document.getElementById('rulesBox').value);
  
  // Save rules to localStorage
  try { localStorage.setItem(LS_KEYS.RULES, document.getElementById('rulesBox').value); } catch {}
  
  // Get month-filtered transactions
  const txns = monthFilteredTxns();
  
  // Apply categorization rules
  categorise(txns, CURRENT_RULES);
  
  // Re-render all UI sections
  renderMonthTotals();
  renderCategoryTotals(txns);
  renderTransactionsTable(txns);
  
  // Save updated transactions
  saveTxnsToLocalStorage();
  
  try { updateMonthBanner(); } catch {}
}

// ============================================================================
// SECTION 11: TRANSACTION TABLE RENDERING
// ============================================================================

/**
 * Filters transactions by active category filter
 * @param {Array} txns - Transactions to filter
 * @returns {Array} Filtered transactions
 */
function getFilteredTxns(txns) {
  if (!CURRENT_FILTER) return txns;
  
  return txns.filter(t => 
    (t.category || 'UNCATEGORISED').toUpperCase() === CURRENT_FILTER
  );
}

/**
 * Updates the filter UI to show active filter or hide it
 */
function updateFilterUI() {
  const label = document.getElementById('activeFilter');
  const btn = document.getElementById('clearFilterBtn');
  
  if (CURRENT_FILTER) {
    label.textContent = `— filtered by "${CURRENT_FILTER}"`;
    btn.style.display = '';  // Show "clear filter" button
  } else {
    label.textContent = '';
    btn.style.display = 'none';  // Hide button
  }
}

/**
 * Updates the month banner text
 */
function updateMonthBanner() {
  const banner = document.getElementById('monthBanner');
  const label = friendlyMonthOrAll(MONTH_FILTER);
  banner.textContent = `— ${label}`;
}

/**
 * Renders the transactions table with pagination
 * @param {Array} txns - Transactions to display (defaults to month-filtered)
 */
function renderTransactionsTable(txns = monthFilteredTxns()) {
  const filtered = getFilteredTxns(txns);
  
  // Calculate total pages
  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  
  // Ensure current page is valid
  if (CURRENT_PAGE > totalPages) CURRENT_PAGE = totalPages;
  if (CURRENT_PAGE < 1) CURRENT_PAGE = 1;
  
  // Calculate which transactions to show on this page
  const start = (CURRENT_PAGE - 1) * PAGE_SIZE;
  const pageItems = filtered.slice(start, start + PAGE_SIZE);
  
  const table = document.getElementById('transactionsTable');
  
  // Build table HTML
  let html = '<tr><th>Date</th><th>Amount</th><th>Category</th><th>Description</th><th></th></tr>';
  
  pageItems.forEach((t) => {
    // Get the original index in CURRENT_TXNS (needed for the + button)
    const idx = CURRENT_TXNS.indexOf(t);
    const cat = (t.category || 'UNCATEGORISED').toUpperCase();
    const displayCat = toTitleCase(cat);
    
    html += `<tr>
      <td>${escapeHtml(t.date)}</td>
      <td>${t.amount.toFixed(2)}</td>
      <td><span class="category-name">${escapeHtml(displayCat)}</span></td>
      <td>${escapeHtml(t.description)}</td>
      <td><button class="rule-btn" onclick="assignCategory(${idx})">+</button></td>
    </tr>`;
  });
  
  table.innerHTML = html;
  
  // Render pagination controls
  renderPager(totalPages);
}

/**
 * Renders the pagination controls
 * @param {number} totalPages - Total number of pages
 */
function renderPager(totalPages) {
  const pager = document.getElementById('pager');
  if (!pager) return;
  
  const pages = totalPages || 1;
  const cur = CURRENT_PAGE;

  /**
   * Helper function to create a page button
   * @param {string} label - Button text
   * @param {number} page - Page number to navigate to
   * @param {boolean} disabled - Whether button should be disabled
   * @param {boolean} isActive - Whether this is the current page
   * @returns {string} HTML for button
   */
  function pageButton(label, page, disabled = false, isActive = false) {
    const disAttr = disabled ? ' disabled' : '';
    const activeClass = isActive ? ' active' : '';
    return `<button class="page-btn${activeClass}" data-page="${page}"${disAttr}>${label}</button>`;
  }

  // Calculate which page numbers to show
  // Show 5 page numbers centered around current page
  const windowSize = 5;
  let start = Math.max(1, cur - Math.floor(windowSize / 2));
  let end = Math.min(pages, start + windowSize - 1);
  start = Math.max(1, Math.min(start, end - windowSize + 1));

  let html = '';
  
  // First and Previous buttons
  html += pageButton('First', 1, cur === 1);
  html += pageButton('Prev', Math.max(1, cur - 1), cur === 1);

  // Page number buttons
  for (let p = start; p <= end; p++) {
    html += pageButton(String(p), p, false, p === cur);
  }

  // Next and Last buttons
  html += pageButton('Next', Math.min(pages, cur + 1), cur === pages);
  html += pageButton('Last', pages, cur === pages);
  
  // Page indicator
  html += `<span style="margin-left:8px">Page ${cur} / ${pages}</span>`;

  pager.innerHTML = html;
  
  // Add click handlers to all buttons
  pager.querySelectorAll('button.page-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const page = Number(e.currentTarget.getAttribute('data-page'));
      if (!page || page === CURRENT_PAGE) return;
      
      CURRENT_PAGE = page;
      renderTransactionsTable();
    });
  });

  // Bonus feature: Mouse wheel to flip pages
  const table = document.getElementById('transactionsTable');
  if (table && !table._wheelBound) {
    table.addEventListener('wheel', (e) => {
      if (pages <= 1) return;
      
      // Scroll down = next page, scroll up = previous page
      if (e.deltaY > 0 && CURRENT_PAGE < pages) {
        CURRENT_PAGE++;
        renderTransactionsTable();
      } else if (e.deltaY < 0 && CURRENT_PAGE > 1) {
        CURRENT_PAGE--;
        renderTransactionsTable();
      }
    }, { passive: true });
    
    table._wheelBound = true;  // Flag to prevent adding multiple listeners
  }
}

// ============================================================================
// SECTION 12: EXPORT FUNCTIONS
// ============================================================================

/**
 * Exports category totals as a formatted text file
 */
function exportTotals() {
  const txns = monthFilteredTxns();
  const { rows, grand } = computeCategoryTotals(txns);

  const label = friendlyMonthOrAll(MONTH_FILTER || getFirstTxnMonth(txns) || new Date());
  const header = `SpendLite Category Totals (${label})`;

  // Calculate column widths for nice alignment
  const catWidth = Math.max(8, ...rows.map(([cat]) => toTitleCase(cat).length), 'Category'.length);
  const amtWidth = 12;
  const pctWidth = 6;

  const lines = [];
  lines.push(header);
  lines.push('='.repeat(header.length));  // Underline
  
  // Header row
  lines.push(
    'Category'.padEnd(catWidth) + ' ' +
    'Amount'.padStart(amtWidth) + ' ' +
    '%'.padStart(pctWidth)
  );

  // Data rows
  for (const [cat, total] of rows) {
    const pct = grand ? (total / grand * 100) : 0;
    lines.push(
      toTitleCase(cat).padEnd(catWidth) + ' ' +
      total.toFixed(2).padStart(amtWidth) + ' ' +
      (pct.toFixed(1) + '%').padStart(pctWidth)
    );
  }

  // Total row
  lines.push('');
  lines.push(
    'TOTAL'.padEnd(catWidth) + ' ' +
    grand.toFixed(2).padStart(amtWidth) + ' ' +
    '100%'.padStart(pctWidth)
  );

  // Create downloadable file
  const blob = new Blob([lines.join('\n')], { type: 'text/plain' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `category_totals_${forFilename(label)}.txt`;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

/**
 * Exports categorization rules to a text file
 */
function exportRules() {
  const text = document.getElementById('rulesBox').value || '';
  const blob = new Blob([text], {type: 'text/plain'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'rules_export.txt';
  document.body.appendChild(a);
  a.click();
  a.remove();
}

/**
 * Imports rules from a text file
 * @param {File} file - The file to import
 */
function importRulesFromFile(file) {
  const reader = new FileReader();
  
  reader.onload = () => {
    const text = reader.result || '';
    document.getElementById('rulesBox').value = text;
    applyRulesAndRender();
  };
  
  reader.readAsText(file);
}

// ============================================================================
// SECTION 13: CATEGORY ASSIGNMENT (ADDING RULES)
// ============================================================================

/**
 * Helper function to extract the next word after a marker in text
 * Example: "PAYPAL JOHNSSTORE" -> nextWordAfter("paypal", ...) returns "JOHNSSTORE"
 * @param {string} marker - The marker to look for
 * @param {string} desc - The description text
 * @returns {string} The next word after the marker
 */
function nextWordAfter(marker, desc) {
  const lower = (desc || '').toLowerCase();
  const i = lower.indexOf(String(marker).toLowerCase());
  
  if (i === -1) return '';
  
  // Get text after the marker
  let after = (desc || '').slice(i + String(marker).length);
  
  // Strip leading separators (space, dash, colon, slash, asterisk)
  after = after.replace(/^[\s\-:\/*]+/, '');
  
  // Extract first merchant-like token
  const m = after.match(/^([A-Za-z0-9&._]+)/);
  return m ? m[1] : '';
}

/**
 * Derives a sensible keyword from a transaction description
 * Uses heuristics to handle PayPal, VISA-, and generic merchants
 * @param {Object} txn - Transaction object
 * @returns {string} Suggested keyword in uppercase
 */
function deriveKeywordFromTxn(txn) {
  if (!txn) return "";
  
  const desc = String(txn.description || txn.desc || "").trim();
  if (!desc) return "";
  
  const up = desc.toUpperCase();

  // Heuristic 1: PayPal transactions
  // Format is usually "PAYPAL *MERCHANTNAME"
  if (/\bPAYPAL\b/.test(up)) {
    const nxt = nextWordAfter('paypal', desc);
    return ('PAYPAL' + (nxt ? ' ' + nxt : '')).toUpperCase();
  }

  // Heuristic 2: VISA- prefix
  // Some banks show "VISA-MERCHANTNAME"
  const visaPos = up.indexOf("VISA-");
  if (visaPos !== -1) {
    const after = desc.substring(visaPos + 5).trim();
    const token = (after.split(/\s+/)[0] || "");
    if (token) return token.toUpperCase();
  }

  // Heuristic 3: Generic - first merchant-like token
  const m = desc.match(/([A-Za-z0-9&._]{3,})/);
  return m ? m[1].toUpperCase() : "";
}

/**
 * Adds or updates a rule in the rules textarea
 * @param {string} keywordUpper - Keyword in uppercase
 * @param {string} categoryUpper - Category in uppercase
 * @returns {boolean} True if rule was added/updated
 */
function addOrUpdateRuleLine(keywordUpper, categoryUpper) {
  if (!keywordUpper || !categoryUpper) return false;
  
  const box = document.getElementById('rulesBox');
  if (!box) return false;
  
  const lines = String(box.value || '').split(/\r?\n/);

  let updated = false;
  const kwLower = keywordUpper.toLowerCase();
  
  // Check if keyword already exists
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line || line.startsWith('#')) continue;
    
    const parts = line.split(/=>/i);
    if (parts.length >= 2) {
      const existingKw = parts[0].trim().toLowerCase();
      
      if (existingKw === kwLower) {
        // Update existing rule
        lines[i] = `${keywordUpper} => ${categoryUpper}`;
        updated = true;
        break;
      }
    }
  }
  
  // If not found, add new rule
  if (!updated) {
    lines.push(`${keywordUpper} => ${categoryUpper}`);
  }
  
  box.value = lines.join("\n");
  
  // Save to localStorage
  try { localStorage.setItem(LS_KEYS.RULES, box.value); } catch {}
  
  return true;
}

/**
 * Opens category picker to assign a category to a transaction
 * This is called when user clicks the "+" button next to a transaction
 * @param {number} idx - Index of transaction in CURRENT_TXNS array
 */
function assignCategory(idx) {
  // Collect all unique categories from transactions and rules
  const fromTxns = (Array.isArray(CURRENT_TXNS) ? CURRENT_TXNS : [])
    .map(x => (x.category || '').trim());
  
  const fromRules = (Array.isArray(CURRENT_RULES) ? CURRENT_RULES : [])
    .map(r => (r.category || '').trim ? r.category : (r.category || ''));
  
  const merged = Array.from(new Set([...fromTxns, ...fromRules]
    .map(c => (c || '').trim())
    .filter(Boolean)));

  // Build category list
  let base = Array.from(new Set(merged));
  
  // Normalize "UNCATEGORISED" variations
  base = base.map(c => (c.toUpperCase() === 'UNCATEGORISED' ? 'Uncategorised' : c));
  
  if (!base.includes('Uncategorised')) base.unshift('Uncategorised');
  base.unshift('+ Add new category...');

  // Sort alphabetically (except special items)
  const specials = new Set(['+ Add new category...', 'Uncategorised']);
  const rest = base.filter(c => !specials.has(c))
    .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  
  const categories = ['+ Add new category...', 'Uncategorised', ...rest];

  const current = ((CURRENT_TXNS && CURRENT_TXNS[idx] && CURRENT_TXNS[idx].category) || '').trim() || 'Uncategorised';

  // Open the category picker modal (defined in catpicker-modal.js)
  SL_CatPicker.openCategoryPicker({
    categories,
    current,
    onChoose: (chosen) => {
      if (chosen) {
        const ch = String(chosen).trim();
        const lo = ch.toLowerCase();
        const isAdd = ch.startsWith('➕') || ch.startsWith('+') || lo.indexOf('add new category') !== -1;
        
        if (isAdd) {
          // Close modal and use old prompt-based flow
          try { document.getElementById('catpickerBackdrop').classList.remove('show'); } catch {}
          return assignCategory_OLD(idx);
        }
      }
      
      const norm = (chosen === 'Uncategorised') ? '' : String(chosen).trim().toUpperCase();
      
      if (CURRENT_TXNS && CURRENT_TXNS[idx]) {
        CURRENT_TXNS[idx].category = norm;
      }
      
      // Auto-add rule for this merchant
      try {
        if (norm) {
          const kw = deriveKeywordFromTxn(CURRENT_TXNS[idx]);
          
          if (kw) {
            const added = addOrUpdateRuleLine(kw, norm);
            
            if (added && typeof applyRulesAndRender === 'function') {
              applyRulesAndRender({keepPage: true});
            } else {
              renderMonthTotals();
              renderTransactionsTable();
            }
          } else {
            renderMonthTotals();
            renderTransactionsTable();
          }
        } else {
          renderMonthTotals();
          renderTransactionsTable();
        }
      } catch (e) {
        try {
          renderMonthTotals();
          renderTransactionsTable();
        } catch {}
      }
    }
  });
}

/**
 * Old-style category assignment using browser prompts
 * Used when user chooses "+ Add new category..."
 * @param {number} idx - Index of transaction in CURRENT_TXNS array
 */
function assignCategory_OLD(idx) {
  const txn = CURRENT_TXNS[idx];
  if (!txn) return;
  
  const desc = txn.description || "";
  const up = desc.toUpperCase();

  // Build suggested keyword using same heuristics
  let suggestedKeyword = "";
  
  if (/\bPAYPAL\b/.test(up)) {
    const nxt = nextWordAfter('paypal', desc);
    suggestedKeyword = ('PAYPAL' + (nxt ? ' ' + nxt : '')).toUpperCase();
  } else {
    const visaPos = up.indexOf("VISA-");
    if (visaPos !== -1) {
      const after = desc.substring(visaPos + 5).trim();
      suggestedKeyword = (after.split(/\s+/)[0] || "").toUpperCase();
    } else {
      suggestedKeyword = (desc.split(/\s+/)[0] || "").toUpperCase();
    }
  }

  // Ask user for keyword
  const keywordInput = prompt("Enter keyword to match:", suggestedKeyword);
  if (!keywordInput) return;
  const keyword = keywordInput.trim().toUpperCase();

  // Ask user for category
  const defaultCat = (txn.category || "UNCATEGORISED").toUpperCase();
  const catInput = prompt("Enter category name:", defaultCat);
  if (!catInput) return;
  const category = catInput.trim().toUpperCase();

  // Add or update rule in textarea
  const box = document.getElementById('rulesBox');
  const lines = String(box.value || "").split(/\r?\n/);
  let updated = false;
  
  for (let i = 0; i < lines.length; i++) {
    const line = (lines[i] || "").trim();
    if (!line || line.startsWith('#')) continue;
    
    const parts = line.split(/=>/i);
    if (parts.length >= 2) {
      const k = parts[0].trim().toUpperCase();
      if (k === keyword) {
        lines[i] = `${keyword} => ${category}`;
        updated = true;
        break;
      }
    }
  }
  
  if (!updated) lines.push(`${keyword} => ${category}`);
  
  box.value = lines.join("\n");
  
  try { localStorage.setItem(LS_KEYS.RULES, box.value); } catch {}
  
  if (typeof applyRulesAndRender === 'function') {
    applyRulesAndRender({keepPage: true});
  }
}

// ============================================================================
// SECTION 14: LOCAL STORAGE PERSISTENCE
// ============================================================================

/**
 * Saves current transactions to localStorage
 * This ensures data persists even if you close the browser
 */
function saveTxnsToLocalStorage() {
  try {
    const data = JSON.stringify(CURRENT_TXNS || []);
    
    // Save to multiple keys for compatibility with Advanced mode
    localStorage.setItem(LS_KEYS.TXNS_JSON, data);
    localStorage.setItem('spendlite_txns_json_v7', data);
    localStorage.setItem('spendlite_txns_json', data);
  } catch {}
}

// ============================================================================
// SECTION 15: TRANSACTION VISIBILITY TOGGLE
// ============================================================================

/**
 * Checks if transactions section is collapsed (hidden)
 * @returns {boolean} True if collapsed
 */
function isTxnsCollapsed() {
  try {
    return localStorage.getItem(LS_KEYS.TXNS_COLLAPSED) !== 'false';
  } catch {
    return true;  // Default to collapsed
  }
}

/**
 * Sets the collapsed state of transactions section
 * @param {boolean} v - True to collapse, false to expand
 */
function setTxnsCollapsed(v) {
  try {
    localStorage.setItem(LS_KEYS.TXNS_COLLAPSED, v ? 'true' : 'false');
  } catch {}
}

/**
 * Applies the collapsed state to the UI
 */
function applyTxnsCollapsedUI() {
  const body = document.getElementById('transactionsBody');
  const toggle = document.getElementById('txnsToggleBtn');
  const collapsed = isTxnsCollapsed();
  
  if (body) body.style.display = collapsed ? 'none' : '';
  if (toggle) toggle.textContent = collapsed ? 'Show transactions' : 'Hide transactions';
}

/**
 * Toggles the visibility of the transactions section
 * This function is called by the onclick handler in HTML
 */
function toggleTransactions() {
  const collapsed = isTxnsCollapsed();
  setTxnsCollapsed(!collapsed);
  applyTxnsCollapsedUI();
}

// ============================================================================
// SECTION 16: EVENT LISTENERS (UI INTERACTIONS)
// ============================================================================

// CSV file upload
document.getElementById('csvFile').addEventListener('change', (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  
  const reader = new FileReader();
  reader.onload = () => { loadCsvText(reader.result); };
  reader.readAsText(file);
});

// Recalculate button
document.getElementById('recalculateBtn').addEventListener('click', applyRulesAndRender);

// Export buttons
document.getElementById('exportRulesBtn').addEventListener('click', exportRules);
document.getElementById('exportTotalsBtn').addEventListener('click', exportTotals);

// Import rules button
document.getElementById('importRulesBtn').addEventListener('click', () => 
  document.getElementById('importRulesInput').click()
);

document.getElementById('importRulesInput').addEventListener('change', (e) => {
  const f = e.target.files && e.target.files[0];
  if (f) importRulesFromFile(f);
});

// Clear filter button
document.getElementById('clearFilterBtn').addEventListener('click', () => {
  CURRENT_FILTER = null;
  try { localStorage.removeItem(LS_KEYS.FILTER); } catch {}
  
  updateFilterUI();
  CURRENT_PAGE = 1;
  renderTransactionsTable();
  renderMonthTotals(monthFilteredTxns());
});

// Clear month filter button
document.getElementById('clearMonthBtn').addEventListener('click', () => {
  MONTH_FILTER = "";
  try { localStorage.removeItem(LS_KEYS.MONTH); } catch {}
  
  document.getElementById('monthFilter').value = "";
  updateMonthBanner();
  CURRENT_PAGE = 1;
  applyRulesAndRender();
});

// Month filter dropdown
document.getElementById('monthFilter').addEventListener('change', (e) => {
  MONTH_FILTER = e.target.value || "";
  try { localStorage.setItem(LS_KEYS.MONTH, MONTH_FILTER); } catch {}
  
  updateMonthBanner();
  CURRENT_PAGE = 1;
  applyRulesAndRender();
});

// ============================================================================
// SECTION 17: INITIALIZATION (RUNS WHEN PAGE LOADS)
// ============================================================================

/**
 * Main initialization - runs when DOM is ready
 */
window.addEventListener('DOMContentLoaded', async () => {
  // STEP 1: Restore rules from localStorage or load default
  let restored = false;
  
  try {
    const saved = localStorage.getItem(LS_KEYS.RULES);
    if (saved && saved.trim()) {
      document.getElementById('rulesBox').value = saved;
      restored = true;
    }
  } catch {}
  
  // Try loading from rules.txt file if not in localStorage
  if (!restored) {
    try {
      const res = await fetch('rules.txt');
      const text = await res.text();
      document.getElementById('rulesBox').value = text;
      restored = true;
    } catch {}
  }
  
  // Use sample rules if nothing else worked
  if (!restored) {
    document.getElementById('rulesBox').value = SAMPLE_RULES;
  }

  // STEP 2: Restore filters from localStorage
  try {
    const savedFilter = localStorage.getItem(LS_KEYS.FILTER);
    CURRENT_FILTER = savedFilter && savedFilter.trim() ? savedFilter.toUpperCase() : null;
  } catch {}
  
  try {
    const savedMonth = localStorage.getItem(LS_KEYS.MONTH);
    MONTH_FILTER = savedMonth || "";
  } catch {}

  // STEP 3: Update UI
  updateFilterUI();
  CURRENT_PAGE = 1;
  updateMonthBanner();
});

// Apply collapsed state on load
document.addEventListener('DOMContentLoaded', () => {
  applyTxnsCollapsedUI();
  try { updateMonthBanner(); } catch {}
});

// Save transactions before leaving page (safety net)
window.addEventListener('beforeunload', () => {
  try {
    localStorage.setItem(LS_KEYS.TXNS_JSON, JSON.stringify(CURRENT_TXNS || []));
  } catch {}
});

// ============================================================================
// SECTION 23: CLOSE APP WITH AUTO-SAVE FUNCTIONALITY
// ============================================================================

// Track initial rules content and changes
let INITIAL_RULES = '';
let RULES_CHANGED = false;

// Store initial rules content when page loads
window.addEventListener('load', () => {
  const rulesBox = document.getElementById('rulesBox');
  if (rulesBox) {
    INITIAL_RULES = rulesBox.value;
    
    // Track changes to rules textarea
    rulesBox.addEventListener('input', () => {
      RULES_CHANGED = rulesBox.value !== INITIAL_RULES;
    });
  }
});

/**
 * Downloads rules as a text file
 * @param {string} content - The rules content to save
 * @param {string} filename - Name of the file to download
 */
function downloadRulesFile(content, filename = 'rules.txt') {
  const blob = new Blob([content], { type: 'text/plain' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Shows status message to user
 * @param {string} message - Message to display
 * @param {string} type - Type of message ('success' or 'info')
 */
function showSaveStatus(message, type = 'info') {
  const statusEl = document.getElementById('saveStatus');
  if (!statusEl) return;
  
  statusEl.textContent = message;
  statusEl.className = `save-status ${type}`;
  statusEl.style.display = 'block';
  
  // Auto-hide after 3 seconds
  setTimeout(() => {
    statusEl.style.display = 'none';
  }, 3000);
}

/**
 * Handles close app button click
 * Saves rules if changed, shows status message
 */
function handleCloseApp() {
  const rulesBox = document.getElementById('rulesBox');
  if (!rulesBox) return;
  
  const currentRules = rulesBox.value;
  
  // Check if rules have changed
  if (RULES_CHANGED && currentRules !== INITIAL_RULES) {
    // Save rules to file
    downloadRulesFile(currentRules, 'rules.txt');
    
    // Show success message
    showSaveStatus('✓ Rules file updated', 'success');
    
    // Update initial rules to current (so we don't save again)
    INITIAL_RULES = currentRules;
    RULES_CHANGED = false;
  } else {
    // Show info message (no changes)
    showSaveStatus('ℹ No rule changes', 'info');
  }
}

// Attach close button handler
document.addEventListener('DOMContentLoaded', () => {
  const closeBtn = document.getElementById('closeAppBtn');
  if (closeBtn) {
    closeBtn.addEventListener('click', handleCloseApp);
  }
});
