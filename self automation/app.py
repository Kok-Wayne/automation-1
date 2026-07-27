from flask import Flask, render_template, request, redirect, url_for, jsonify
import pandas as pd
import os
import re

# ======================
# Config
# ======================
UPLOAD_FOLDER = "uploads"

# Browser profile directory for WhatsApp Web — saves the session so you
# only need to scan the QR code once across multiple runs.
CHROME_PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chrome_profile")

# ======================
# App
# ======================
app = Flask(__name__)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ======================
# Path Safety
# ======================
def safe_path(filename):
    if not filename:
        return None
    path = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.abspath(path).startswith(os.path.abspath(UPLOAD_FOLDER)):
        return None
    return path


# ======================
# Phone Cleaner
# ======================
def clean_phone(phone):
    if phone is None:
        return None, "empty"
    phone = str(phone).strip()
    phone = phone.replace("{", "").replace("}", "")
    if phone.lower() in ("nan", "none", ""):
        return None, "empty"

    digits = re.sub(r"\D", "", phone)
    if not digits:
        return None, "no_digits"

    # ── Step 1: normalise to full number with country code 60 ─
    if digits.startswith("601") and 11 <= len(digits) <= 12:
        full = digits
    elif digits.startswith("60") and len(digits) >= 10:
        full = digits
    elif digits.startswith("0"):
        full = "60" + digits[1:]
    elif digits.startswith("1") and len(digits) <= 10:
        full = "60" + digits
    else:
        full = digits

    if len(full) < 10:
        return None, "invalid"

    # ── Step 2: landline check AFTER normalising ──────────────
    # Malaysian mobiles all start with 601x.
    # Landlines: 603 (KL/Sel), 604 (Penang), 605 (Perak),
    #            606 (NS/Melaka), 607 (Johor), 608 (Pahang/East M'sia),
    #            609 (Kelantan/Terengganu)
    LANDLINE_PREFIXES = ("602", "603", "604", "605", "606", "607", "608", "609")
    if any(full.startswith(p) for p in LANDLINE_PREFIXES):
        return None, "landline"

    return "+" + full, "ok"


# ======================
# Home
# ======================
@app.route("/")
def home():
    files = os.listdir(UPLOAD_FOLDER)
    files = [f for f in files if f.endswith((".xlsx", ".xls")) and not f.startswith("~$")]
    return render_template("index.html", files=files)


# ======================
# Upload
# ======================
@app.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("file")
    if not files or all(f.filename == "" for f in files):
        return "No file selected", 400

    for file in files:
        if not file.filename:
            continue
        if not file.filename.lower().endswith((".xlsx", ".xls")):
            return f"{file.filename} is not .xlsx / .xls", 400
        path = safe_path(file.filename)
        if not path:
            return f"Invalid filename: {file.filename}", 400
        file.save(path)

    return redirect(url_for("home"))


# ======================
# View
# ======================
@app.route("/view")
def view():
    filename = request.args.get("file", "")
    path = safe_path(filename)
    if not path:
        return "Invalid filename", 400
    if not os.path.exists(path):
        return "File not found", 404

    try:
        df = pd.read_excel(path, engine="openpyxl")
    except Exception as e:
        return f"Could not read file: {e}", 400

    df.columns = df.columns.str.strip().str.lower()
    df = df.fillna("").dropna(how="all").reset_index(drop=True)

    columns = df.columns.tolist()

    # Auto-detect name column
    NAME_COLS = ["name", "nama", "full_name", "customer", "contact"]
    name_col = next((c for c in df.columns if c in NAME_COLS), None)

    # Auto-detect phone columns — support two columns (e.g. "tel" and "doctor tel")
    TEL_COLS = ["tel", "phone", "mobile", "number", "telefon", "hp", "handphone"]
    TEL_KEYWORDS = ("tel", "phone", "mobile", "hp", "handphone", "telefon")

    # Prefer compound columns (e.g. "doctor tel") over plain "tel"
    col_lower = {c: c.lower() for c in df.columns}

    # tel_col2: compound match (has keyword but is NOT an exact match)
    tel_col2 = next(
        (c for c in df.columns
         if col_lower[c] not in TEL_COLS and any(k in col_lower[c] for k in TEL_KEYWORDS)),
        None
    )
    # tel_col: exact/plain match
    tel_col = next((c for c in df.columns if col_lower[c] in TEL_COLS), None)

    def best_phone(row):
        """
        Check both phone columns, return the first valid mobile number.
        Priority: compound col (doctor tel) first, plain col (tel) second.
        Returns (cleaned, status, used_col).
        """
        candidates = [c for c in [tel_col2, tel_col] if c]  # compound first
        if not candidates:
            return None, "no_col", None

        results = []
        for col in candidates:
            cleaned, status = clean_phone(row.get(col))
            results.append((cleaned, status, col))
            if status == "ok":
                return cleaned, "ok", col  # found valid mobile, use immediately

        # None valid — return most informative status
        STATUS_RANK = {"landline": 4, "invalid": 3, "no_digits": 2, "empty": 1, "no_col": 0}
        results.sort(key=lambda x: STATUS_RANK.get(x[1], 0), reverse=True)
        return results[0]

    # Check if filtering out already-sent rows
    show_all      = request.args.get("show_all", "1")
    show_all_bool = show_all == "1"
    has_sent_col  = "sent" in df.columns

    data = []
    for i, row in df.iterrows():
        # Skip already-sent rows unless show_all is requested
        if has_sent_col and not show_all_bool:
            sent_val = str(row.get("sent", "")).strip()
            if sent_val and sent_val != "nan":
                continue
        row_dict = row.to_dict()
        row_dict["index"] = i
        if tel_col or tel_col2:
            cleaned, status, used_col = best_phone(row)
            row_dict["_cleaned_phone"] = cleaned
            row_dict["_phone_status"]  = status
            row_dict["_phone_col"]     = used_col
        else:
            row_dict["_cleaned_phone"] = None
            row_dict["_phone_status"]  = "no_col"
            row_dict["_phone_col"]     = None
        data.append(row_dict)

    sent_count = len(df[df["sent"].astype(str).str.strip().isin(["✓"])]) if has_sent_col else 0

    return render_template("view.html", filename=filename, columns=columns, data=data,
                           name_col=name_col, tel_col=tel_col, tel_col2=tel_col2,
                           sent_count=sent_count,
                           show_all=show_all, has_sent_col=has_sent_col,
                           total_count=len(df))


# ======================
# Delete File
# ======================
@app.route("/delete", methods=["POST"])
def delete_file():
    filename = request.form.get("filename", "")
    path = safe_path(filename)
    if not path:
        return "Invalid filename", 400
    if os.path.exists(path):
        os.remove(path)
    return "OK", 200


# ======================
# Delete Rows
# ======================
@app.route("/delete_rows", methods=["POST"])
def delete_rows():
    filename = request.form.get("filename", "")
    indices  = request.form.getlist("indices")
    path = safe_path(filename)
    if not path:
        return "Invalid filename", 400
    if not os.path.exists(path):
        return "File not found", 404

    df = pd.read_excel(path, engine="openpyxl")
    df = df.fillna("").dropna(how="all").reset_index(drop=True)

    def to_int(v):
        try:
            return int(v)
        except (ValueError, TypeError):
            return None

    rows_to_drop = [x for x in (to_int(i) for i in indices) if x is not None]
    df = df.drop(index=rows_to_drop, errors="ignore").reset_index(drop=True)
    df.to_excel(path, index=False, engine="openpyxl")
    return "OK", 200


# ======================
# Update Cell
# ======================
@app.route("/update_cell", methods=["POST"])
def update_cell():
    payload  = request.get_json(force=True) or {}
    filename = payload.get("filename", "")
    row_idx  = payload.get("row")
    col_name = payload.get("col")
    value    = payload.get("value", "")

    if not filename or row_idx is None or not col_name:
        return jsonify({"error": "Missing data"}), 400

    path = safe_path(filename)
    if not path:
        return jsonify({"error": "Invalid filename"}), 400
    if not os.path.exists(path):
        return jsonify({"error": "File not found"}), 404

    try:
        df = pd.read_excel(path, engine="openpyxl")
    except Exception as e:
        return jsonify({"error": str(e)}), 400

    df = df.fillna("").dropna(how="all").reset_index(drop=True)

    col_map      = {c.strip().lower(): c for c in df.columns}
    original_col = col_map.get(col_name.strip().lower())

    if original_col is None:
        return jsonify({"error": "Column not found"}), 400
    if not (0 <= row_idx < len(df)):
        return jsonify({"error": "Row index out of range"}), 400

    col_dtype = df[original_col].dtype
    try:
        if pd.api.types.is_integer_dtype(col_dtype):
            typed_value = int(value) if value != "" else 0
        elif pd.api.types.is_float_dtype(col_dtype):
            typed_value = float(value) if value != "" else 0.0
        else:
            typed_value = value
    except (ValueError, TypeError):
        typed_value = value

    df.at[row_idx, original_col] = typed_value
    df.to_excel(path, index=False, engine="openpyxl")
    return jsonify({"ok": True})


# ======================
# Playwright WhatsApp Web sender
# ======================
# sync_playwright is greenlet-bound: every Playwright call MUST happen on the
# single OS thread that created the browser.  We keep one permanent
# "_browser_thread" alive for the whole Flask process.  The send() route posts
# callables onto _browser_call_queue; the browser thread executes them and
# signals a threading.Event so the caller can collect the result.
# ======================
import queue    as _bq
import threading as _bt

_browser_thread      = None
_browser_call_queue  = _bq.Queue()
_browser_ready_event = _bt.Event()
_browser_context_ref = [None]   # [BrowserContext | None]
_browser_lock        = _bt.Lock()
_BROWSER_STOP        = object()   # sentinel


def _browser_thread_main():
    """Owns the Playwright event loop forever (daemon thread)."""
    from playwright.sync_api import sync_playwright
    os.makedirs(CHROME_PROFILE_DIR, exist_ok=True)
    try:
        with sync_playwright() as pw:
            ctx = pw.chromium.launch_persistent_context(
                user_data_dir = CHROME_PROFILE_DIR,
                headless      = False,
                args          = ["--window-size=1280,900"],
                no_viewport   = True,
            )
            _browser_context_ref[0] = ctx
            _browser_ready_event.set()
            while True:
                task = _browser_call_queue.get()
                if task is _BROWSER_STOP:
                    break
                try:
                    task(ctx)
                except Exception:
                    pass
    except Exception:
        _browser_ready_event.set()


def _ensure_browser_thread():
    global _browser_thread
    with _browser_lock:
        if _browser_thread is None or not _browser_thread.is_alive():
            _browser_ready_event.clear()
            _browser_context_ref[0] = None
            _browser_thread = _bt.Thread(
                target=_browser_thread_main, daemon=True, name="pw-browser"
            )
            _browser_thread.start()


def _run_on_browser(fn):
    """Run fn(ctx) on the browser thread; block until done; return result or raise."""
    _ensure_browser_thread()
    _browser_ready_event.wait(timeout=30)
    ctx = _browser_context_ref[0]
    if ctx is None:
        raise RuntimeError("Playwright browser failed to start.")
    box  = [None, None]
    done = _bt.Event()
    def task(c):
        try:
            box[0] = fn(c)
        except Exception as e:
            box[1] = e
        finally:
            done.set()
    _browser_call_queue.put(task)
    done.wait()
    if box[1] is not None:
        raise box[1]
    return box[0]


def reset_pw_context():
    """Close browser; next send will open a fresh window."""
    global _browser_thread
    _browser_context_ref[0] = None
    _browser_ready_event.clear()
    try:
        _browser_call_queue.put(_BROWSER_STOP)
    except Exception:
        pass
    _browser_thread = None


def _safe_visible(page, selector):
    try:
        return page.locator(selector).is_visible()
    except Exception:
        return False


def _reopen_whatsapp(push_fn, total):
    """
    Browser was closed mid-send.  Tear down the stale context, launch a new one,
    navigate to WA Web and wait for the session to restore (or QR scan).
    Returns the new wa_page, or raises on failure.
    """
    LOGIN_SELECTORS = [
        'div[data-testid="chat-list"]',
        'div[data-testid="default-user"]',
        '#side',
        'div[aria-label="Chat list"]',
    ]

    reset_pw_context()

    push_fn({"type": "progress", "current": 0, "total": total,
             "status": "ok",
             "message": "\u26a0\ufe0f Browser was closed \u2014 reopening automatically\u2026"})

    def open_fresh(ctx):
        pg = ctx.new_page()
        pg.goto("https://web.whatsapp.com", timeout=60_000)
        return pg

    wa_page = _run_on_browser(open_fresh)

    push_fn({"type": "progress", "current": 0, "total": total,
             "status": "ok",
             "message": "Waiting for WhatsApp session to restore (scan QR if needed)\u2026"})

    def wait_login(ctx):
        deadline = 180_000
        for sel in LOGIN_SELECTORS:
            try:
                wa_page.locator(sel).wait_for(state="visible", timeout=deadline)
                return True
            except Exception:
                deadline = 10_000
        return False

    if not _run_on_browser(wait_login):
        raise RuntimeError("Re-login timed out after browser was closed.")

    push_fn({"type": "progress", "current": 0, "total": total,
             "status": "ok", "message": "Session restored \u2713 Resuming sends\u2026"})
    return wa_page


def send_whatsapp_playwright(page, phone, message):
    """Called from the browser thread — all Playwright calls stay on one thread."""
    import urllib.parse

    def type_and_send():
        box = page.locator('div[contenteditable="true"][data-tab="10"]')
        box.wait_for(state="visible", timeout=15_000)
        box.click()
        box.fill(message)
        page.wait_for_timeout(800)
        box.press("Enter")
        page.wait_for_timeout(1500)

    # Strategy 1: search bar (fast, no navigation)
    try:
        nc = page.locator('span[data-icon="new-chat-outline"]')
        nc.wait_for(state="visible", timeout=10_000)
        nc.click()
        sb = page.locator('div[contenteditable="true"][data-tab="3"]')
        sb.wait_for(state="visible", timeout=8_000)
        sb.click()
        sb.fill(phone.lstrip("+"))
        page.wait_for_timeout(1500)
        res = page.locator('div[data-testid="cell-frame-container"]').first
        res.wait_for(state="visible", timeout=5_000)
        res.click()
        type_and_send()
        return True, ""
    except Exception:
        pass

    # Strategy 2: ?phone= URL (works for unsaved numbers)
    try:
        encoded = urllib.parse.quote(message)
        url = "https://web.whatsapp.com/send?phone=" + phone.lstrip("+") + "&text=" + encoded
        page.goto(url, timeout=60_000)
        box = page.locator('div[contenteditable="true"][data-tab="10"]')
        box.wait_for(state="visible", timeout=40_000)
        page.wait_for_timeout(800)
        box.press("Enter")
        page.wait_for_timeout(1500)
        return True, ""
    except Exception as e:
        return False, str(e)


# ======================
# Mark Sent
# ======================
def mark_sent(filename, row_idx):
    path = safe_path(filename)
    if not path or not os.path.exists(path):
        return
    try:
        df = pd.read_excel(path, engine="openpyxl")
        df = df.fillna("").dropna(how="all").reset_index(drop=True)
        if "sent" not in df.columns:
            df["sent"] = ""
        if 0 <= row_idx < len(df):
            df.at[row_idx, "sent"] = "\u2713"
            df.to_excel(path, index=False, engine="openpyxl")
    except Exception:
        pass


# ======================
# Send WhatsApp  (SSE streaming)
# ======================
@app.route("/send", methods=["POST"])
def send():
    import time, random, json as _json
    from flask import Response, stream_with_context
    import queue, threading

    selected         = request.form.getlist("selected")
    message_template = request.form.get("message_template", "").strip()
    filename         = request.form.get("filename", "").strip()

    if not selected:
        return jsonify({"ok": False, "message": "No contacts selected"}), 400

    q = queue.Queue()

    def worker():
        import time, random, traceback
        total   = len(selected)
        success = 0
        errors  = []

        def push(evt):
            q.put(evt)

        def browser_died(exc):
            msg = str(exc).lower()
            return any(k in msg for k in (
                "target page, context or browser has been closed",
                "has been closed",
                "greenlet",
                "browser has been closed",
                "connection closed",
                "page closed",
            ))

        try:
            LOGIN_SELECTORS = [
                'div[data-testid="chat-list"]',
                'div[data-testid="default-user"]',
                '#side',
                'div[aria-label="Chat list"]',
            ]

            # ── Open / reuse WhatsApp tab ─────────────────────────────────
            def open_wa(ctx):
                pages = ctx.pages
                pg = next((p for p in pages if "web.whatsapp.com" in p.url), None)
                already = pg is not None and any(
                    _safe_visible(pg, s) for s in LOGIN_SELECTORS
                )
                if pg is None:
                    pg = ctx.new_page()
                if not already:
                    pg.goto("https://web.whatsapp.com", timeout=60_000)
                return pg, already

            try:
                wa_page, already = _run_on_browser(open_wa)
            except Exception as e:
                err_txt = (
                    "Could not open browser: " + str(e) +
                    "\n\nRun: pip install playwright && playwright install chromium"
                )
                push({"type": "done", "ok": False, "message": err_txt})
                q.put(None)
                return

            if already:
                push({"type": "progress", "current": 0, "total": total,
                      "status": "ok", "message": "Reusing existing WhatsApp session \u2713"})
            else:
                push({"type": "progress", "current": 0, "total": total,
                      "status": "ok",
                      "message": "Opening WhatsApp Web\u2026 scan QR if prompted (up to 3 min)"})

                def wait_login(ctx):
                    deadline = 180_000
                    for sel in LOGIN_SELECTORS:
                        try:
                            wa_page.locator(sel).wait_for(state="visible", timeout=deadline)
                            return True
                        except Exception:
                            deadline = 10_000
                    return False

                try:
                    logged_in = _run_on_browser(wait_login)
                except Exception:
                    logged_in = False

                if not logged_in:
                    reset_pw_context()
                    push({"type": "done", "ok": False,
                          "message": "WhatsApp Web login timed out. Click Send again and scan the QR code."})
                    q.put(None)
                    return

            # ── Send loop ─────────────────────────────────────────────────
            for idx, item in enumerate(selected, start=1):
                parts = item.split("||")
                if len(parts) < 2:
                    err = "Bad format: " + item
                    errors.append(err)
                    push({"type": "progress", "current": idx, "total": total,
                          "status": "error", "message": err})
                    continue

                phone   = parts[0]
                name    = parts[1]
                row_idx = int(parts[2]) if len(parts) > 2 else -1

                if not phone:
                    err = "Empty phone for: " + name
                    errors.append(err)
                    push({"type": "progress", "current": idx, "total": total,
                          "status": "error", "message": err})
                    continue

                if message_template:
                    msg_text = message_template.replace("{name}", name)
                else:
                    msg_text = "Hello " + name + "!"

                # Attempt send; auto-recover once if browser was closed
                ok      = False
                err_msg = ""
                for attempt in range(2):
                    def do_send(ctx, _p=phone, _m=msg_text):
                        return send_whatsapp_playwright(wa_page, _p, _m)
                    try:
                        ok, err_msg = _run_on_browser(do_send)
                        break
                    except Exception as ex:
                        if attempt == 0 and browser_died(ex):
                            try:
                                wa_page = _reopen_whatsapp(push, total)
                            except Exception as reopen_ex:
                                ok      = False
                                err_msg = "Could not reopen browser: " + str(reopen_ex)
                                break
                            # loop continues → retry send with new wa_page
                        else:
                            ok      = False
                            err_msg = str(ex)
                            break

                if ok:
                    success += 1
                    if filename and row_idx >= 0:
                        mark_sent(filename, row_idx)
                    push({"type": "progress", "current": idx, "total": total,
                          "status": "ok", "message": name + " (" + phone + ")"})
                else:
                    errors.append(phone + ": " + err_msg)
                    push({"type": "progress", "current": idx, "total": total,
                          "status": "error",
                          "message": name + " (" + phone + "): " + err_msg})

                if idx < total:
                    if idx % 10 == 0:
                        delay = random.uniform(30, 90)
                        push({"type": "progress", "current": idx, "total": total,
                              "status": "ok",
                              "message": "\u23f8 Taking a short break (" + str(int(delay)) + "s) to avoid detection\u2026"})
                    else:
                        delay = random.uniform(8, 20)
                        if random.random() < 0.25:
                            delay += random.uniform(5, 15)
                    time.sleep(delay)

            if errors:
                msg = "Sent " + str(success) + "/" + str(total) + ".\n\nFailed:\n" + "\n".join(errors)
            else:
                msg = "Successfully sent to " + str(success) + "/" + str(total) + " contact(s) \u2713"
            push({"type": "done", "ok": success > 0, "message": msg})
            q.put(None)

        except Exception:
            tb = traceback.format_exc()
            print("[worker CRASH]\n" + tb)
            push({"type": "done", "ok": False,
                  "message": "Unexpected error in worker:\n" + tb})
            q.put(None)

    t = threading.Thread(target=worker, daemon=True)
    t.start()

    def generate():
        while True:
            try:
                evt = q.get(timeout=10)
            except queue.Empty:
                if not t.is_alive():
                    break
                yield ": heartbeat\n\n"
                continue
            if evt is None:
                break
            yield "data: " + _json.dumps(evt) + "\n\n"

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"}
    )



# ======================
# Add Row
# ======================
@app.route("/add_row", methods=["POST"])
def add_row():
    payload  = request.get_json(force=True) or {}
    filename = payload.get("filename", "")
    row_data = payload.get("row", {})

    path = safe_path(filename)
    if not path:
        return jsonify({"error": "Invalid filename"}), 400
    if not os.path.exists(path):
        return jsonify({"error": "File not found"}), 404

    try:
        df = pd.read_excel(path, engine="openpyxl")
    except Exception as e:
        return jsonify({"error": str(e)}), 400

    # Build new row — fill missing columns with empty string
    new_row = {col: row_data.get(col.strip().lower(), "") for col in df.columns}
    new_df  = pd.DataFrame([new_row])
    df      = pd.concat([df, new_df], ignore_index=True)
    df.to_excel(path, index=False, engine="openpyxl")

    return jsonify({"ok": True, "index": len(df) - 1})


# ======================
# Run
# ======================
if __name__ == "__main__":
    # threaded=True is required — allows SSE streaming while Playwright runs in the background
    app.run(debug=True, threaded=True)  
