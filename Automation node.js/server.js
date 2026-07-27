/**
 * Excel CRM — Node.js server
 * Multi-account WhatsApp support
 */

const express    = require('express');
const multer     = require('multer');
const path       = require('path');
const fs         = require('fs');
const XLSX       = require('xlsx');
const nunjucks   = require('nunjucks');

const app = express();

// ======================
// Config
// ======================
const UPLOAD_FOLDER  = path.join(__dirname, 'uploads');
const PROFILES_DIR   = path.join(__dirname, 'chrome_profiles');
const ACCOUNTS_FILE  = path.join(__dirname, 'accounts.json');
const PORT           = 5000;

fs.mkdirSync(UPLOAD_FOLDER, { recursive: true });
fs.mkdirSync(PROFILES_DIR,  { recursive: true });

// ======================
// Accounts store
// accounts.json = [ { id, label, profileDir } ]
// ======================
function loadAccounts() {
    try {
        if (fs.existsSync(ACCOUNTS_FILE)) return JSON.parse(fs.readFileSync(ACCOUNTS_FILE, 'utf8'));
    } catch (_) {}
    return [];
}

function saveAccounts(list) {
    fs.writeFileSync(ACCOUNTS_FILE, JSON.stringify(list, null, 2));
}

function makeAccountId() {
    return 'acc_' + Date.now();
}

// ======================
// Nunjucks
// ======================
nunjucks.configure(path.join(__dirname, 'templates'), { autoescape: true, express: app });
app.set('view engine', 'html');

// ======================
// Middleware
// ======================
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use('/static', express.static(path.join(__dirname, 'static')));

const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, UPLOAD_FOLDER),
    filename:    (req, file, cb) => cb(null, file.originalname),
});
const upload = multer({ storage });

// ======================
// Path Safety
// ======================
function safePath(filename) {
    if (!filename) return null;
    const resolved = path.resolve(path.join(UPLOAD_FOLDER, filename));
    if (!resolved.startsWith(path.resolve(UPLOAD_FOLDER))) return null;
    return resolved;
}

// ======================
// Phone Cleaner
// ======================
function cleanPhone(phone) {
    if (phone == null) return [null, 'empty'];
    let p = String(phone).trim().replace(/[{}]/g, '');
    if (['nan', 'none', ''].includes(p.toLowerCase())) return [null, 'empty'];
    const digits = p.replace(/\D/g, '');
    if (!digits) return [null, 'no_digits'];
    let full;
    if (digits.startsWith('601') && digits.length >= 11 && digits.length <= 12) full = digits;
    else if (digits.startsWith('60') && digits.length >= 10) full = digits;
    else if (digits.startsWith('0')) full = '60' + digits.slice(1);
    else if (digits.startsWith('1') && digits.length <= 10) full = '60' + digits;
    else full = digits;
    if (full.length < 10) return [null, 'invalid'];
    const LANDLINE = ['602','603','604','605','606','607','608','609'];
    if (LANDLINE.some(p => full.startsWith(p))) return [null, 'landline'];
    return ['+' + full, 'ok'];
}

// ======================
// Excel helpers
// ======================
function readExcel(filePath) {
    const wb  = XLSX.readFile(filePath);
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: '' });
    return raw.map(row => {
        const out = {};
        for (const [k, v] of Object.entries(row)) out[k.trim().toLowerCase()] = v;
        return out;
    }).filter(row => Object.values(row).some(v => v !== ''));
}

function writeExcel(filePath, rows) {
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, filePath);
}

// ======================
// Mark Sent
// ======================
function markSent(filename, rowIdx) {
    const filePath = safePath(filename);
    if (!filePath || !fs.existsSync(filePath)) return;
    try {
        let rows = readExcel(filePath);
        if (!rows[0].hasOwnProperty('sent')) rows.forEach(r => r.sent = '');
        if (rowIdx >= 0 && rowIdx < rows.length) {
            rows[rowIdx].sent = '✓';
            writeExcel(filePath, rows);
        }
    } catch (_) {}
}

// ======================
// Multi-account browser manager
// Each account has its own browserContext keyed by account id
// ======================
const browsers = {}; // { [accountId]: { context, ready } }

async function ensureBrowser(accountId, profileDir) {
    // Check if existing context is still alive
    if (browsers[accountId]?.ready) {
        try {
            // Try a simple operation to verify the context is still alive
            await browsers[accountId].context.pages();
        } catch (_) {
            // Context is dead — reset and rebuild
            delete browsers[accountId];
        }
    }

    if (browsers[accountId]?.ready) return;

    const { chromium } = require('playwright');
    fs.mkdirSync(profileDir, { recursive: true });
    const context = await chromium.launchPersistentContext(profileDir, {
        headless: false,
        args: ['--window-size=1280,900'],
        viewport: null,
    });

    // Auto-reset if the browser is closed externally
    context.on('close', () => {
        delete browsers[accountId];
    });

    browsers[accountId] = { context, ready: true };
}

function resetBrowser(accountId) {
    if (browsers[accountId]?.context) {
        try { browsers[accountId].context.close(); } catch (_) {}
    }
    delete browsers[accountId];
}

function getBrowserContext(accountId) {
    return browsers[accountId]?.context || null;
}

// ======================
// WhatsApp send logic
// ======================
async function dismissPopups(page) {
    try {
        const selectors = [
            'div[data-animate-modal-backdrop="true"] button[aria-label="Close"]',
            'span[data-icon="x-alt"]',
            'div[role="button"][aria-label="Close"]',
            'button[aria-label="Close"]',
        ];
        for (const sel of selectors) {
            const btn = page.locator(sel).first();
            if (await btn.isVisible({ timeout: 1500 }).catch(() => false)) {
                await btn.click();
                await page.waitForTimeout(500);
                break;
            }
        }
    } catch (_) {}
}

async function sendWhatsappPlaywright(page, phone, message) {
    const phoneNum = phone.replace(/^\+/, '');
    await dismissPopups(page);

    // Robust send button locator — covers old & new WhatsApp Web versions
    function sendBtnLocator() {
        return page.locator([
            'button[aria-label="Send"]',
            'button[data-testid="compose-btn-send"]',
            'span[data-icon="send"]',
            'div[aria-label="Send"]',
            'button[aria-label="Send Message"]',
        ].join(', '));
    }

    // Robust message box locator
    function msgBoxLocator() {
        return page.locator([
            'div[contenteditable="true"][data-tab="10"]',
            'div[contenteditable="true"][aria-placeholder]',
            'div[contenteditable="true"][data-lexical-editor="true"]',
            'footer div[contenteditable="true"]',
        ].join(', ')).first();
    }

    async function typeAndSend() {
        const box = msgBoxLocator();
        await box.waitFor({ state: 'visible', timeout: 20000 });
        await box.click();
        await box.fill('');

        // Split on newlines and type each line, using Shift+Enter for line breaks
        const lines = message.split('\n');
        for (let i = 0; i < lines.length; i++) {
            if (lines[i]) await page.keyboard.type(lines[i], { delay: 20 });
            if (i < lines.length - 1) await page.keyboard.press('Shift+Enter');
        }

        await page.waitForTimeout(800);
        // Try send button first, fallback to Enter
        try {
            const btn = sendBtnLocator();
            await btn.waitFor({ state: 'visible', timeout: 4000 });
            await btn.first().click();
        } catch (_) {
            await box.press('Enter');
        }
        await page.waitForTimeout(1500);
    }

    // Strategy 1: new-chat search
    try {
        const nc = page.locator([
            'span[data-icon="new-chat-outline"]',
            'span[data-icon="new-chat"]',
            'div[aria-label="New chat"]',
            'button[aria-label="New chat"]',
        ].join(', ')).first();
        await nc.waitFor({ state: 'visible', timeout: 10000 });
        await nc.click();
        const sb = page.locator([
            'div[contenteditable="true"][data-tab="3"]',
            'div[contenteditable="true"][aria-autocomplete="list"]',
            'input[type="text"][aria-autocomplete="list"]',
        ].join(', ')).first();
        await sb.waitFor({ state: 'visible', timeout: 8000 });
        await sb.click();
        await sb.fill(phoneNum);
        await page.waitForTimeout(1800);
        const result = page.locator([
            'div[data-testid="cell-frame-container"]',
            'div[role="listitem"]',
            'span[data-testid="cell-frame-title"]',
        ].join(', ')).first();
        await result.waitFor({ state: 'visible', timeout: 6000 });
        await result.click();
        await page.waitForTimeout(800);
        await dismissPopups(page);
        await typeAndSend();
        return [true, ''];
    } catch (_) {}

    // Strategy 2: phone URL
    try {
        const encoded = encodeURIComponent(message);
        await page.goto(`https://web.whatsapp.com/send?phone=${phoneNum}&text=${encoded}`, { timeout: 60000 });
        await page.waitForTimeout(4000);

        // Check for invalid number popup BEFORE dismissing
        // An invalid-number popup has an OK/Close button but NO message input box
        const invalidPopup = await (async () => {
            try {
                const backdrop = page.locator('div[data-animate-modal-backdrop="true"]');
                if (!await backdrop.isVisible({ timeout: 2000 }).catch(() => false)) return false;
                // If message box is also visible, it's a different popup (e.g. "Use here") — not invalid number
                const msgBox = msgBoxLocator();
                const msgVisible = await msgBox.isVisible({ timeout: 2000 }).catch(() => false);
                return !msgVisible; // only treat as invalid if no message box
            } catch (_) { return false; }
        })();

        await dismissPopups(page);
        await page.waitForTimeout(1500);

        if (invalidPopup) {
            return [false, 'Number may be invalid or not on WhatsApp'];
        }

        // Wait for message box (confirms number is valid and chat opened)
        const box = msgBoxLocator();
        await box.waitFor({ state: 'visible', timeout: 30000 });
        await box.click();
        await page.waitForTimeout(500);

        // Clear prefilled text and retype cleanly with proper newline handling
        await page.keyboard.press('Control+A');
        await page.keyboard.press('Backspace');
        const lines2 = message.split('\n');
        for (let i = 0; i < lines2.length; i++) {
            if (lines2[i]) await page.keyboard.type(lines2[i], { delay: 20 });
            if (i < lines2.length - 1) await page.keyboard.press('Shift+Enter');
        }
        await page.waitForTimeout(800);

        // Try Send button, fallback to Enter
        try {
            const btn = sendBtnLocator();
            await btn.waitFor({ state: 'visible', timeout: 5000 });
            await btn.first().click();
        } catch (_) {
            await box.press('Enter');
        }
        await page.waitForTimeout(2000);
        return [true, ''];
    } catch (e) {
        try { await page.screenshot({ path: '/tmp/wa_debug.png' }); } catch (_) {}
        return [false, e.message + ' | Debug screenshot: /tmp/wa_debug.png'];
    }
}

// ======================
// Routes — Home
// ======================
app.get('/', (req, res) => {
    const files = fs.readdirSync(UPLOAD_FOLDER)
        .filter(f => (f.endsWith('.xlsx') || f.endsWith('.xls')) && !f.startsWith('~$'));
    const accounts = loadAccounts();
    res.render('index.html', { files, accounts });
});

// ======================
// Routes — Accounts API
// ======================

// Add account
app.post('/accounts/add', (req, res) => {
    const label = (req.body.label || '').trim();
    if (!label) return res.status(400).json({ error: 'Label required' });
    const accounts = loadAccounts();
    const id = makeAccountId();
    const profileDir = path.join(PROFILES_DIR, id);
    accounts.push({ id, label, profileDir });
    saveAccounts(accounts);
    res.json({ ok: true, id, label });
});

// Delete account
app.post('/accounts/delete', (req, res) => {
    const { id } = req.body;
    let accounts = loadAccounts();
    accounts = accounts.filter(a => a.id !== id);
    saveAccounts(accounts);
    resetBrowser(id);
    // Optionally delete profile folder
    const profileDir = path.join(PROFILES_DIR, id);
    if (fs.existsSync(profileDir)) fs.rmSync(profileDir, { recursive: true, force: true });
    res.json({ ok: true });
});

// Rename account
app.post('/accounts/rename', (req, res) => {
    const { id, label } = req.body;
    const accounts = loadAccounts();
    const acc = accounts.find(a => a.id === id);
    if (!acc) return res.status(404).json({ error: 'Not found' });
    acc.label = (label || '').trim() || acc.label;
    saveAccounts(accounts);
    res.json({ ok: true });
});

// List accounts
app.get('/accounts', (req, res) => {
    res.json(loadAccounts());
});

// ======================
// Routes — Upload
// ======================
app.post('/upload', upload.array('file'), (req, res) => {
    if (!req.files || req.files.length === 0) return res.status(400).send('No file selected');
    res.redirect('/');
});

// ======================
// Routes — View
// ======================
app.get('/view', (req, res) => {
    const filename = req.query.file || '';
    const filePath = safePath(filename);
    if (!filePath) return res.status(400).send('Invalid filename');
    if (!fs.existsSync(filePath)) return res.status(404).send('File not found');

    let rows;
    try { rows = readExcel(filePath); }
    catch (e) { return res.status(400).send('Could not read file: ' + e.message); }

    const accounts = loadAccounts();

    if (rows.length === 0) return res.render('view.html', {
        filename, columns: [], data: [], accounts,
        name_col: null, tel_col: null, tel_col2: null,
        sent_count: 0, show_all: '1', has_sent_col: false, total_count: 0, ok_count: 0,
    });

    const columns = Object.keys(rows[0]);
    const NAME_COLS    = ['name','nama','full_name','customer','contact','clinic name','doctor name','clinic','doctor'];
    const TEL_COLS     = ['tel','phone','mobile','number','telefon','hp','handphone'];
    const TEL_KEYWORDS = ['tel','phone','mobile','hp','handphone','telefon'];

    const name_col = columns.find(c => NAME_COLS.includes(c)) || null;
    const tel_col2 = columns.find(c => !TEL_COLS.includes(c) && TEL_KEYWORDS.some(k => c.includes(k))) || null;
    const tel_col  = columns.find(c => TEL_COLS.includes(c)) || null;

    function bestPhone(row) {
        const candidates = [tel_col2, tel_col].filter(Boolean);
        if (!candidates.length) return [null, 'no_col', null];
        const STATUS_RANK = { landline: 4, invalid: 3, no_digits: 2, empty: 1, no_col: 0 };
        let results = [];
        for (const col of candidates) {
            const [cleaned, status] = cleanPhone(row[col]);
            results.push([cleaned, status, col]);
            if (status === 'ok') return [cleaned, 'ok', col];
        }
        results.sort((a, b) => (STATUS_RANK[b[1]] || 0) - (STATUS_RANK[a[1]] || 0));
        return results[0];
    }

    const show_all     = req.query.show_all || '1';
    const has_sent_col = columns.includes('sent');
    const totalCount   = rows.length;
    const sent_count   = has_sent_col ? rows.filter(r => String(r.sent || '').trim() === '✓').length : 0;

    const data = [];
    rows.forEach((row, i) => {
        if (has_sent_col && show_all !== '1') {
            const sv = String(row.sent || '').trim();
            if (sv && sv !== 'nan') return;
        }
        const [cleaned, status, usedCol] = bestPhone(row);
        data.push({ ...row, index: i, _cleaned_phone: cleaned, _phone_status: status, _phone_col: usedCol, _row_name: name_col ? (row[name_col] || '') : '' });
    });

    const ok_count = data.filter(r => r._phone_status === 'ok').length;
    res.render('view.html', { filename, columns, data, accounts, name_col, tel_col, tel_col2, sent_count, show_all, has_sent_col, total_count: totalCount, ok_count });
});

// ======================
// Routes — Delete File / Rows / Update Cell / Add Row
// ======================
app.post('/delete', (req, res) => {
    const fp = safePath(req.body.filename || '');
    if (!fp) return res.status(400).send('Invalid');
    if (fs.existsSync(fp)) fs.unlinkSync(fp);
    res.send('OK');
});

app.post('/delete_rows', (req, res) => {
    const fp = safePath(req.body.filename || '');
    if (!fp || !fs.existsSync(fp)) return res.status(400).send('Invalid');
    let rows = readExcel(fp);
    const indices = [req.body.indices || []].flat().map(Number).filter(n => !isNaN(n));
    rows = rows.filter((_, i) => !indices.includes(i));
    writeExcel(fp, rows);
    res.send('OK');
});

app.post('/update_cell', (req, res) => {
    const { filename, row: rowIdx, col, value } = req.body;
    const fp = safePath(filename);
    if (!fp || !fs.existsSync(fp)) return res.status(400).json({ error: 'Invalid' });
    let rows = readExcel(fp);
    if (rowIdx < 0 || rowIdx >= rows.length) return res.status(400).json({ error: 'Out of range' });
    const origCol = Object.keys(rows[0]).find(k => k.trim().toLowerCase() === col.trim().toLowerCase());
    if (!origCol) return res.status(400).json({ error: 'Column not found' });
    rows[rowIdx][origCol] = value;
    writeExcel(fp, rows);
    res.json({ ok: true });
});

app.post('/add_row', (req, res) => {
    const { filename, row: rowData } = req.body;
    const fp = safePath(filename);
    if (!fp || !fs.existsSync(fp)) return res.status(400).json({ error: 'Invalid' });
    let rows = readExcel(fp);
    const cols = rows.length ? Object.keys(rows[0]) : Object.keys(rowData || {});
    const newRow = {};
    cols.forEach(c => { newRow[c] = (rowData || {})[c.trim().toLowerCase()] || ''; });
    rows.push(newRow);
    writeExcel(fp, rows);
    res.json({ ok: true, index: rows.length - 1 });
});

// ======================
// Routes — Send (SSE, multi-account)
// ======================
app.post('/send', express.urlencoded({ extended: true }), async (req, res) => {
    const selected        = [req.body.selected || []].flat();
    const messageTemplate = (req.body.message_template || '').trim();
    const filename        = (req.body.filename || '').trim();
    const accountId       = (req.body.account_id || '').trim();

    if (!selected.length) return res.status(400).json({ ok: false, message: 'No contacts selected' });

    // Find account
    const accounts = loadAccounts();
    const account  = accounts.find(a => a.id === accountId);
    if (!account) {
        res.setHeader('Content-Type', 'text/event-stream');
        res.flushHeaders();
        res.write('data: ' + JSON.stringify({ type: 'done', ok: false, message: 'No WhatsApp account selected. Please select an account first.' }) + '\n\n');
        return res.end();
    }

    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('X-Accel-Buffering', 'no');
    res.flushHeaders();

    function push(evt) { res.write('data: ' + JSON.stringify(evt) + '\n\n'); }

    const total   = selected.length;
    let   success = 0;
    const errors  = [];

    const LOGIN_SELECTORS = [
        'div[data-testid="chat-list"]',
        'div[data-testid="default-user"]',
        '#side',
        'div[aria-label="Chat list"]',
    ];

    async function safeVisible(page, sel) {
        try { return await page.locator(sel).isVisible(); } catch { return false; }
    }

    // Open browser for this account
    let waPage;
    try {
        await ensureBrowser(account.id, account.profileDir);
        const ctx   = getBrowserContext(account.id);
        const pages = ctx.pages();
        let pg      = pages.find(p => p.url().includes('web.whatsapp.com')) || null;
        const already = pg && (await Promise.all(LOGIN_SELECTORS.map(s => safeVisible(pg, s)))).some(Boolean);
        if (!pg) pg = await ctx.newPage();
        if (!already) await pg.goto('https://web.whatsapp.com', { timeout: 60000 });
        waPage = pg;
    } catch (e) {
        push({ type: 'done', ok: false, message: 'Could not open browser: ' + e.message });
        return res.end();
    }

    push({ type: 'progress', current: 0, total, status: 'ok', message: `Using account: ${account.label}` });

    // Wait for login
    const alreadyLoggedIn = (await Promise.all(LOGIN_SELECTORS.map(s => safeVisible(waPage, s)))).some(Boolean);
    if (!alreadyLoggedIn) {
        push({ type: 'progress', current: 0, total, status: 'ok', message: 'Opening WhatsApp Web… scan QR if prompted (up to 3 min)' });
        let loggedIn = false;
        for (const sel of LOGIN_SELECTORS) {
            try { await waPage.locator(sel).waitFor({ state: 'visible', timeout: 180000 }); loggedIn = true; break; }
            catch { continue; }
        }
        if (!loggedIn) {
            resetBrowser(account.id);
            push({ type: 'done', ok: false, message: 'Login timed out. Click Send again and scan the QR code.' });
            return res.end();
        }
    } else {
        push({ type: 'progress', current: 0, total, status: 'ok', message: `Session ready ✓ (${account.label})` });
    }

    // Send loop
    for (let idx = 0; idx < selected.length; idx++) {
        const parts  = selected[idx].split('||');
        const phone  = parts[0];
        const name   = parts[1] || '';
        const rowIdx = parseInt(parts[2] ?? '-1');

        if (!phone) {
            const err = 'Empty phone for: ' + name;
            errors.push(err);
            push({ type: 'progress', current: idx + 1, total, status: 'error', message: err });
            continue;
        }

        const msgText = messageTemplate ? messageTemplate.replace(/{name}/g, name) : 'Hello ' + name + '!';
        const [ok, errMsg] = await sendWhatsappPlaywright(waPage, phone, msgText);

        if (ok) {
            success++;
            if (filename && rowIdx >= 0) markSent(filename, rowIdx);
            push({ type: 'progress', current: idx + 1, total, status: 'ok', message: name + ' (' + phone + ')' });
        } else {
            errors.push(phone + ': ' + errMsg);
            push({ type: 'progress', current: idx + 1, total, status: 'error', message: name + ' (' + phone + '): ' + errMsg });
        }

        if (idx < total - 1) {
            const base  = 8000 + Math.random() * 12000;
            const extra = Math.random() < 0.25 ? 5000 + Math.random() * 10000 : 0;
            let   delay = base + extra;
            if ((idx + 1) % 10 === 0) {
                delay = 30000 + Math.random() * 60000;
                push({ type: 'progress', current: idx + 1, total, status: 'ok', message: `⏸ Short break (${Math.round(delay / 1000)}s)…` });
            }
            await new Promise(r => setTimeout(r, delay));
        }
    }

    const msg = errors.length
        ? `Sent ${success}/${total}.\n\nFailed:\n` + errors.join('\n')
        : `Successfully sent to ${success}/${total} contact(s) ✓`;

    push({ type: 'done', ok: success > 0, message: msg });
    res.end();
});

// ======================
// Start
// ======================
app.listen(PORT, () => console.log(`Excel CRM running at http://localhost:${PORT}`));
