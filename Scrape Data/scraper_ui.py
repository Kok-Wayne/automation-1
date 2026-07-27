"""
Google Maps Scraper — Tkinter UI
Wraps scrape.py logic with a simple desktop window.
Run: python scraper_ui.py
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import threading
import queue
import re
import os
import sys

# ── redirect print() into the UI log ──────────────────────────────────────────
_log_queue = queue.Queue()

class _QueueWriter:
    def write(self, text):
        if text.strip():
            _log_queue.put(text)
    def flush(self):
        pass

sys.stdout = _QueueWriter()

# ── import scraper (must come after stdout redirect) ───────────────────────────
try:
    import scrape as sc
except ImportError:
    import tkinter.messagebox as mb
    mb.showerror("Missing file", "scrape.py must be in the same folder as scraper_ui.py")
    sys.exit(1)


# ── main window ───────────────────────────────────────────────────────────────
class ScraperApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Google Maps Scraper")
        self.resizable(True, True)
        self.minsize(620, 560)
        self._running = False
        self._thread  = None
        self._stop_event = threading.Event()
        self._build_ui()
        self._poll_log()

    # ── layout ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        PAD = dict(padx=14, pady=6)

        # ── title bar area ────────────────────────────────────────────────────
        header = tk.Frame(self, bg="#185FA5")
        header.pack(fill="x")
        tk.Label(header, text="  Google Maps Scraper",
                 bg="#185FA5", fg="white",
                 font=("Helvetica", 13, "bold"),
                 pady=10).pack(side="left")

        # ── form ──────────────────────────────────────────────────────────────
        form = tk.Frame(self, padx=14, pady=10)
        form.pack(fill="x")

        # Search queries header
        tk.Label(form, text="Search queries", anchor="w",
                 font=("Helvetica", 10)).grid(row=0, column=0, sticky="w", pady=(0,2))
        tk.Label(form, text="e.g.  GP clinic Johor Bahru  |  pharmacy Penang  |  hotel KL",
                 fg="gray", font=("Helvetica", 9), anchor="w").grid(
                     row=0, column=1, columnspan=2, sticky="w", pady=(0,2))

        # Dynamic query rows container
        self.query_frame = tk.Frame(form)
        self.query_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0,4))
        self.query_vars = []
        self._add_query_row("GP clinic Johor Bahru")

        # Add / Remove buttons
        btn_row = tk.Frame(form)
        btn_row.grid(row=2, column=0, columnspan=3, sticky="w", pady=(0,8))
        tk.Button(btn_row, text="+ Add search",
                  font=("Helvetica", 10), padx=8, pady=2,
                  command=self._add_query_row).pack(side="left")
        tk.Label(btn_row, text="  max 10 queries",
                 fg="gray", font=("Helvetica", 9)).pack(side="left")

        # Output filename
        tk.Label(form, text="Output filename (.xlsx)", anchor="w",
                 font=("Helvetica", 10)).grid(row=3, column=0, sticky="w", pady=(0,2))
        self.output_var = tk.StringVar(value="results.xlsx")
        tk.Entry(form, textvariable=self.output_var,
                 font=("Helvetica", 11), width=34).grid(row=4, column=0, sticky="ew")
        tk.Button(form, text="Browse…", command=self._browse).grid(
            row=4, column=1, padx=(6,0))

        # Scroll times
        tk.Label(form, text="  Scroll times", anchor="w",
                 font=("Helvetica", 10)).grid(row=3, column=2, sticky="w", padx=(16,0))
        self.scroll_var = tk.StringVar(value="100")
        tk.Spinbox(form, from_=10, to=500, increment=10,
                   textvariable=self.scroll_var,
                   font=("Helvetica", 11), width=6).grid(
                       row=4, column=2, sticky="w", padx=(16,0))

        # Exclude keywords
        tk.Label(form, text="Exclude keywords (comma separated)", anchor="w",
                 font=("Helvetica", 10)).grid(row=5, column=0, sticky="w", pady=(8,2))
        tk.Label(form, text="e.g.  haiwan, veterinar, pharmacy, dental",
                 fg="gray", font=("Helvetica", 9), anchor="w").grid(
                     row=5, column=1, columnspan=2, sticky="w", pady=(8,2))
        self.exclude_var = tk.StringVar(value="")
        tk.Entry(form, textvariable=self.exclude_var,
                 font=("Helvetica", 11)).grid(row=6, column=0, columnspan=3, sticky="ew", pady=(0,4))

        form.columnconfigure(0, weight=1)

        # ── progress bar ──────────────────────────────────────────────────────
        prog_frame = tk.Frame(self, padx=14)
        prog_frame.pack(fill="x")
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(prog_frame, variable=self.progress_var,
                                            maximum=100, mode="determinate")
        self.progress_bar.pack(fill="x", pady=(0,4))
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(prog_frame, textvariable=self.status_var,
                 fg="gray", font=("Helvetica", 9), anchor="w").pack(fill="x")

        # ── buttons ───────────────────────────────────────────────────────────
        btn_frame = tk.Frame(self, padx=14, pady=6)
        btn_frame.pack(fill="x")
        self.start_btn = tk.Button(btn_frame, text="▶  Start scraping",
                                   bg="#185FA5", fg="white",
                                   font=("Helvetica", 11, "bold"),
                                   padx=14, pady=6,
                                   relief="flat", cursor="hand2",
                                   command=self._start)
        self.start_btn.pack(side="left")
        self.stop_btn = tk.Button(btn_frame, text="■  Stop",
                                  bg="#A32D2D", fg="white",
                                  font=("Helvetica", 11),
                                  padx=14, pady=6,
                                  relief="flat", cursor="hand2",
                                  state="disabled",
                                  command=self._stop)
        self.stop_btn.pack(side="left", padx=(8,0))
        tk.Button(btn_frame, text="Clear log",
                  font=("Helvetica", 10),
                  padx=10, pady=6,
                  command=self._clear_log).pack(side="right")

        # ── log window ────────────────────────────────────────────────────────
        log_frame = tk.Frame(self, padx=14, pady=4)
        log_frame.pack(fill="both", expand=True)
        tk.Label(log_frame, text="Log", anchor="w",
                 font=("Helvetica", 10)).pack(fill="x")
        self.log_box = scrolledtext.ScrolledText(
            log_frame,
            font=("Courier", 10),
            bg="#1e1e1e", fg="#a8d8a8",
            insertbackground="white",
            state="disabled",
            wrap="word")
        self.log_box.pack(fill="both", expand=True)

    # ── helpers ───────────────────────────────────────────────────────────────
    def _add_query_row(self, default=""):
        if len(self.query_vars) >= 10:
            return
        idx = len(self.query_vars)
        row_frame = tk.Frame(self.query_frame)
        row_frame.pack(fill="x", pady=1)

        tk.Label(row_frame, text=f"{idx+1}.", width=2,
                 font=("Helvetica", 10)).pack(side="left")

        var = tk.StringVar(value=default)
        entry = tk.Entry(row_frame, textvariable=var,
                         font=("Helvetica", 11))
        entry.pack(side="left", fill="x", expand=True, padx=(4,4))

        def remove(rf=row_frame, v=var):
            if len(self.query_vars) <= 1:
                return
            self.query_vars.remove(v)
            rf.destroy()
            self._renumber_rows()

        tk.Button(row_frame, text="✕", font=("Helvetica", 9),
                  fg="gray", relief="flat", cursor="hand2",
                  command=remove).pack(side="left")

        self.query_vars.append(var)

    def _renumber_rows(self):
        for i, frame in enumerate(self.query_frame.winfo_children()):
            labels = [w for w in frame.winfo_children()
                      if isinstance(w, tk.Label)]
            if labels:
                labels[0].config(text=f"{i+1}.")

    def _autofill_filename(self, *_):
        pass

    def _browse(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=self.output_var.get())
        if path:
            self.output_var.set(os.path.basename(path))

    def _clear_log(self):
        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")

    def _append_log(self, text):
        self.log_box.config(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")
        # Update progress bar based on keywords in log
        low = text.lower()
        if "scrolling..." in low:
            try:
                parts = text.split()
                frac = parts[-1]          # e.g. "47/100"
                done, total = frac.split("/")
                pct = int(done) / int(total) * 30   # scrolling = first 30%
                self.progress_var.set(pct)
            except Exception:
                pass
        elif "area " in low and "/" in low:
            try:
                # e.g. "Area 3/7"
                match = re.search(r'area\s+(\d+)/(\d+)', low)
                if match:
                    pct = 30 + int(match.group(1)) / int(match.group(2)) * 70
                    self.progress_var.set(pct)
                    self.status_var.set(text.strip())
            except Exception:
                pass
        elif "all areas done" in low or "no areas discovered" in low or "no new data" in low:
            self.progress_var.set(100)
            self.status_var.set("Done!")

    def _poll_log(self):
        """Check log queue every 100 ms and flush to the log box."""
        while not _log_queue.empty():
            msg = _log_queue.get_nowait()
            self._append_log(msg.rstrip())
        self.after(100, self._poll_log)

    # ── scrape thread ─────────────────────────────────────────────────────────
    def _start(self):
        queries = [v.get().strip() for v in self.query_vars if v.get().strip()]
        output = self.output_var.get().strip()
        if not queries:
            messagebox.showwarning("Missing input", "Please enter at least one search query.")
            return
        if not output.endswith(".xlsx"):
            output += ".xlsx"
            self.output_var.set(output)

        try:
            scroll_times = int(self.scroll_var.get())
        except ValueError:
            scroll_times = 100

        self._running = True
        self._stop_event.clear()
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.progress_var.set(0)
        self.status_var.set("Starting…")
        self._clear_log()

        exclude_kws = [k.strip().lower() for k in self.exclude_var.get().split(",") if k.strip()]

        self._thread = threading.Thread(
            target=self._run_scraper,
            args=(queries, output, scroll_times, exclude_kws),
            daemon=True)
        self._thread.start()

    def _stop(self):
        self._running = False
        self._stop_event.set()
        self.status_var.set("Stopping after current listing…")
        self.stop_btn.config(state="disabled")

    def _run_scraper(self, queries, output, scroll_times, exclude_kws=None):
        try:
            total_queries = len(queries)
            for qi, query in enumerate(queries):
                if not self._running:
                    _log_queue.put("Stopped by user.")
                    break

                _log_queue.put(f"\n{'=' * 40}")
                _log_queue.put(f"Query {qi+1}/{total_queries}: {query}")
                _log_queue.put(f"{'=' * 40}")

                words = query.split()
                city = " ".join(words[-2:]) if len(words) >= 3 else query

                # Step 1: discover areas
                _log_queue.put("Step 1: Discovering areas...")
                areas = sc.discover_areas(query, city, quick_scrolls=20, stop_event=self._stop_event)

                if not areas:
                    _log_queue.put("  No areas discovered. Scraping directly...")
                    sc.scrape(query, output, scroll_times, headless=False, stop_event=self._stop_event, exclude_kws=exclude_kws)
                else:
                    _log_queue.put(f"  Discovered {len(areas)} areas:")
                    for i, area in enumerate(areas, 1):
                        _log_queue.put(f"    {i:>2}. {area}")

                    _log_queue.put(f"\nStep 2: Scraping {len(areas)} areas...")
                    for i, area in enumerate(areas):
                        if not self._running:
                            _log_queue.put("Stopped by user.")
                            break
                        sub_query = f"{query.rsplit(city, 1)[0].strip()} {area} {city}"
                        _log_queue.put(f"\nArea {i+1}/{len(areas)}: {sub_query}")
                        sc.scrape(sub_query, output, scroll_times, headless=False, stop_event=self._stop_event, exclude_kws=exclude_kws)

                # Update overall progress
                pct = 30 + (qi + 1) / total_queries * 70
                self.progress_var.set(pct)

            _log_queue.put("\n" + "=" * 40)
            _log_queue.put("All areas done!")
            _log_queue.put(f"Results saved to {output}")
            _log_queue.put("=" * 40)

        except Exception as e:
            _log_queue.put(f"Error: {e}")
        finally:
            self.after(0, self._on_finish)

    def _on_finish(self):
        self._running = False
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.status_var.set("Finished. Ready for next run.")


# ── entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = ScraperApp()
    app.mainloop()
