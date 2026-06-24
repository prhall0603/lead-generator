#!/usr/bin/env python3
"""Lead Generator — OpenWeb Ninja API + ScrapeGraphAI"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import requests
import threading
from datetime import datetime
import os
import json

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import openpyxl
    from openpyxl.styles import (Font, PatternFill, Alignment,
                                  Border, Side, GradientFill)
    from openpyxl.utils import get_column_letter
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    from scrapegraphai.graphs import SmartScraperGraph
    SCRAPE_OK = True
except ImportError:
    SCRAPE_OK = False

# ── API credentials ──────────────────────────────────────────────────────────
# OpenWeb Ninja direct key (ak_...) — used with X-Api-Key header
OWN_API_KEY    = ""

# RapidAPI key
RAPIDAPI_KEY   = "644279a34bmshbda39876cdd9abcp17180bjsnfaa7d94a6b3d"

# Anthropic API key for ScrapeGraphAI
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

# (display label, internal key, pixel width)
COLUMNS = [
    ("Company Name",    "name",    220),
    ("Address",         "address", 260),
    ("Contact Number",  "phone",   140),
    ("Business Owner",  "owner",   160),
    ("Email",           "email",   210),
]

# ─────────────────────────────────────────────────────────── normalisation ──

def _first(lst, default=""):
    return lst[0] if lst else default

def _str(v, default=""):
    return str(v).strip() if v not in (None, "", []) else default

def normalize(record: dict) -> dict:
    # ── address ──────────────────────────────────────────────────────────────
    addr_obj = record.get("address") or {}
    if isinstance(addr_obj, str):
        full_address = addr_obj
    else:
        parts = [
            addr_obj.get("street") or addr_obj.get("streetAddress", ""),
            addr_obj.get("city", ""),
            addr_obj.get("state", ""),
            addr_obj.get("zipCode") or addr_obj.get("zip", ""),
        ]
        full_address = ", ".join(p for p in parts if p)

    if not full_address:
        full_address = (record.get("full_address")
                        or record.get("formatted_address")
                        or record.get("location", ""))

    # ── phone ─────────────────────────────────────────────────────────────────
    phones = record.get("phone_numbers") or []
    phone  = _first(phones) if isinstance(phones, list) else ""
    phone  = phone or _str(record.get("phone") or record.get("phone_number"))

    # ── email ─────────────────────────────────────────────────────────────────
    emails = record.get("emails") or []
    email  = _first(emails) if isinstance(emails, list) else ""
    email  = email or _str(record.get("email") or record.get("email_address"))

    # ── owner ─────────────────────────────────────────────────────────────────
    owner = (record.get("owner_name")
             or record.get("owner")
             or record.get("proprietor")
             or record.get("contact_name")
             or record.get("manager")
             or "")

    return {
        "name":    _str(record.get("name") or record.get("business_name")),
        "address": full_address,
        "phone":   phone,
        "owner":   _str(owner),
        "email":   email,
    }


# ──────────────────────────────────────────────────────────────── API call ──

def _unwrap(data) -> list[dict]:
    """Extract a list of business records from any common API envelope shape."""
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("data", "results", "businesses", "items", "records"):
            if key in data and isinstance(data[key], list):
                return data[key]
        if "name" in data:
            return [data]
    return []


def _try_request(url: str, headers: dict, params: dict) -> list[dict] | None:
    """
    Make one GET request. Returns a list of records on success,
    None on 404 (try next endpoint), raises RuntimeError on auth failure,
    raises requests.HTTPError on other HTTP errors.
    """
    resp = requests.get(url, headers=headers, params=params, timeout=30)
    if resp.status_code == 404:
        return None
    if resp.status_code in (401, 403):
        raise RuntimeError(
            f"HTTP {resp.status_code} — API key rejected.\n\n"
            f"Response: {resp.text[:300]}\n\n"
            "Check that you are using the correct key for this endpoint\n"
            "(OpenWeb Ninja direct key vs RapidAPI key)."
        )
    resp.raise_for_status()
    return _unwrap(resp.json())


def search_businesses(business_type: str, location: str, limit: int) -> list[dict]:
    query  = f"{business_type} in {location}"
    params = {"query": query, "location": location, "limit": limit, "country": "us"}

    # ── attempt list: (url, headers) ─────────────────────────────────────────
    attempts = []

    # 1. OpenWeb Ninja direct endpoints (ak_ key)
    if OWN_API_KEY:
        own_headers = {"X-Api-Key": OWN_API_KEY, "Accept": "application/json"}
        for path in (
            "https://local-business-data.openwebninja.com/search",
            "https://api.openwebninja.com/local-business-data/search",
            "https://datastore.openwebninja.com/local-business-data/search",
        ):
            attempts.append((path, own_headers))

    # 2. RapidAPI endpoint (if the user has a RapidAPI key)
    if RAPIDAPI_KEY:
        rapid_headers = {
            "X-RapidAPI-Key":  RAPIDAPI_KEY,
            "X-RapidAPI-Host": "local-business-data.p.rapidapi.com",
            "Accept":          "application/json",
        }
        attempts.append((
            "https://local-business-data.p.rapidapi.com/search",
            rapid_headers,
        ))

    if not attempts:
        raise RuntimeError("No API key configured. Set OWN_API_KEY or RAPIDAPI_KEY.")

    errors = []
    for url, headers in attempts:
        try:
            records = _try_request(url, headers, params)
            if records is None:          # 404 — try next
                errors.append(f"404 Not Found — {url}")
                continue
            return [normalize(r) for r in records if isinstance(r, dict)]

        except RuntimeError:
            raise                        # auth errors bubble up immediately
        except requests.HTTPError as e:
            errors.append(
                f"HTTP {e.response.status_code} — {url}\n"
                f"  {e.response.text[:300]}"
            )
        except requests.RequestException as e:
            errors.append(f"Connection error — {url}: {e}")

    raise RuntimeError("All endpoints failed.\n\n" + "\n".join(errors))


# ───────────────────────────────────────────────────── ScrapeGraphAI ───────

def scrape_leads(url: str) -> list[dict]:
    if not SCRAPE_OK:
        raise RuntimeError(
            "scrapegraphai is not installed.\n"
            "Run:  pip install scrapegraphai && playwright install"
        )
    if not ANTHROPIC_API_KEY:
        raise RuntimeError(
            "ANTHROPIC_API_KEY is not set.\n"
            "Set it in a .env file or as an environment variable."
        )

    graph_config = {
        "llm": {
            "api_key": ANTHROPIC_API_KEY,
            "model": "anthropic/claude-sonnet-4-6",
        },
        "verbose": False,
        "headless": True,
    }

    scraper = SmartScraperGraph(
        prompt=(
            "Extract all business leads from this page. For each business, "
            "extract: company name, full address, phone number, email address, "
            "and business owner or contact person name. Return a JSON list of "
            "objects with keys: name, address, phone, email, owner."
        ),
        source=url,
        config=graph_config,
    )

    result = scraper.run()

    if isinstance(result, dict):
        for key in ("businesses", "leads", "results", "data", "items"):
            if key in result and isinstance(result[key], list):
                result = result[key]
                break
        else:
            if "name" in result:
                result = [result]
            else:
                result = list(result.values()) if result else []
                if result and not isinstance(result[0], dict):
                    result = []

    if not isinstance(result, list):
        result = []

    normalized = []
    for item in result:
        if not isinstance(item, dict):
            continue
        normalized.append({
            "name":    str(item.get("name") or item.get("company_name") or "").strip(),
            "address": str(item.get("address") or item.get("full_address") or "").strip(),
            "phone":   str(item.get("phone") or item.get("phone_number") or "").strip(),
            "owner":   str(item.get("owner") or item.get("contact") or item.get("contact_name") or "").strip(),
            "email":   str(item.get("email") or item.get("email_address") or "").strip(),
        })

    return normalized


# ───────────────────────────────────────────────────────── Excel export ─────

def export_excel(records: list[dict], path: str,
                 business_type: str, location: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Leads"

    # ── colour palette ────────────────────────────────────────────────────────
    PURPLE_DARK  = "4C1D95"   # header bg
    PURPLE_MED   = "7C3AED"   # title bar
    WHITE        = "FFFFFF"
    ROW_EVEN     = "F5F3FF"   # light lavender
    ROW_ODD      = "FFFFFF"
    BORDER_CLR   = "D1D5DB"

    thin = Side(style="thin", color=BORDER_CLR)
    box  = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── title row ─────────────────────────────────────────────────────────────
    ws.merge_cells("A1:E1")
    title_cell = ws["A1"]
    title_cell.value = (f"Business Leads — {business_type.title()} "
                        f"in {location}   |   Generated {datetime.now():%Y-%m-%d %H:%M}")
    title_cell.font      = Font(name="Calibri", bold=True, size=13, color=WHITE)
    title_cell.fill      = PatternFill("solid", fgColor=PURPLE_MED)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # ── header row ────────────────────────────────────────────────────────────
    headers = [c[0] for c in COLUMNS]
    ws.append(headers)
    for col_idx, _ in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col_idx)
        cell.font      = Font(name="Calibri", bold=True, size=11, color=WHITE)
        cell.fill      = PatternFill("solid", fgColor=PURPLE_DARK)
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=False)
        cell.border    = box
    ws.row_dimensions[2].height = 22

    # ── data rows ─────────────────────────────────────────────────────────────
    keys = [c[1] for c in COLUMNS]
    for row_num, rec in enumerate(records, start=3):
        fill_color = ROW_EVEN if row_num % 2 == 0 else ROW_ODD
        row_fill   = PatternFill("solid", fgColor=fill_color)
        for col_idx, key in enumerate(keys, start=1):
            cell = ws.cell(row=row_num, column=col_idx,
                           value=rec.get(key, ""))
            cell.font      = Font(name="Calibri", size=10)
            cell.fill      = row_fill
            cell.alignment = Alignment(horizontal="left",
                                       vertical="center", wrap_text=False)
            cell.border    = box
        ws.row_dimensions[row_num].height = 18

    # ── column widths ─────────────────────────────────────────────────────────
    col_widths = {"A": 30, "B": 40, "C": 20, "D": 24, "E": 32}
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # ── freeze panes & auto-filter ────────────────────────────────────────────
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:E{len(records) + 2}"

    # ── summary sheet ─────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    ws2.column_dimensions["A"].width = 26
    ws2.column_dimensions["B"].width = 40
    summary_rows = [
        ("Search Query",    business_type),
        ("Location",        location),
        ("Total Leads",     len(records)),
        ("With Phone",      sum(1 for r in records if r.get("phone"))),
        ("With Email",      sum(1 for r in records if r.get("email"))),
        ("With Owner",      sum(1 for r in records if r.get("owner"))),
        ("Generated On",    datetime.now().strftime("%Y-%m-%d %H:%M")),
    ]
    for i, (label, value) in enumerate(summary_rows, start=1):
        a = ws2.cell(row=i, column=1, value=label)
        b = ws2.cell(row=i, column=2, value=value)
        a.font = Font(name="Calibri", bold=True, size=11)
        b.font = Font(name="Calibri", size=11)
        if i % 2 == 0:
            for cell in (a, b):
                cell.fill = PatternFill("solid", fgColor=ROW_EVEN)

    wb.save(path)


# ─────────────────────────────────────────────────────────────────── GUI ────

class LeadGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Lead Generator — OpenWeb Ninja")
        self.geometry("1080x680")
        self.minsize(860, 520)
        self.configure(bg="#1e1e2e")
        self._results: list[dict] = []
        self._build_ui()

    # ──────────────────────────────────────────────────── build UI ───────────

    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        BG   = "#1e1e2e"
        CARD = "#2a2a3e"
        ACC  = "#7c3aed"
        ACC2 = "#6d28d9"
        FG   = "#e2e8f0"
        MUTED= "#94a3b8"
        GREEN= "#059669"
        GRN2 = "#047857"

        style.configure(".", background=BG, foreground=FG, font=("Segoe UI", 10))
        style.configure("Card.TFrame", background=CARD)
        style.configure("TLabel",      background=BG, foreground=FG)
        style.configure("Card.TLabel", background=CARD, foreground=FG)
        style.configure("Sub.TLabel",  background=CARD, foreground=MUTED,
                        font=("Segoe UI", 9))

        style.configure("Accent.TButton", background=ACC, foreground=WHITE_STR,
                        font=("Segoe UI", 10, "bold"), padding=(14, 8), relief="flat")
        style.map("Accent.TButton",
                  background=[("active", ACC2), ("disabled", "#4b4b6a")])

        style.configure("Excel.TButton", background=GREEN, foreground=WHITE_STR,
                        font=("Segoe UI", 10, "bold"), padding=(13, 8), relief="flat")
        style.map("Excel.TButton",
                  background=[("active", GRN2), ("disabled", "#4b4b6a")])

        style.configure("Treeview",
                        background="#12121e", foreground=FG,
                        fieldbackground="#12121e", rowheight=24,
                        font=("Segoe UI", 9))
        style.configure("Treeview.Heading",
                        background=CARD, foreground=FG,
                        font=("Segoe UI", 9, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", ACC)])
        style.configure("Horizontal.TProgressbar",
                        troughcolor=CARD, background=ACC, thickness=4)

        # ── header bar ───────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=CARD, pady=13)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  Lead Generator", bg=CARD, fg=FG,
                 font=("Segoe UI", 17, "bold")).pack(side="left", padx=10)
        tk.Label(hdr, text="powered by OpenWeb Ninja", bg=CARD, fg=MUTED,
                 font=("Segoe UI", 9)).pack(side="left")

        # ── search card ──────────────────────────────────────────────────────
        sc = ttk.Frame(self, style="Card.TFrame", padding=(22, 14))
        sc.pack(fill="x", padx=18, pady=(12, 0))

        # labels
        for col, text in enumerate(("Business Type", "Location", "Max Results")):
            ttk.Label(sc, text=text, style="Sub.TLabel").grid(
                row=0, column=col, sticky="w", padx=(0 if col == 0 else 14, 0))

        def entry(parent, var, w=26):
            e = tk.Entry(parent, textvariable=var,
                         bg="#12121e", fg=FG, insertbackground=FG,
                         relief="flat", font=("Segoe UI", 11), width=w)
            return e

        self.biz_var = tk.StringVar(value="restaurant")
        self.loc_var = tk.StringVar(value="New York, NY")
        self.lim_var = tk.IntVar(value=25)

        biz_e = entry(sc, self.biz_var)
        biz_e.grid(row=1, column=0, sticky="ew", ipady=7)

        loc_e = entry(sc, self.loc_var)
        loc_e.grid(row=1, column=1, sticky="ew", padx=(14, 0), ipady=7)

        spin = tk.Spinbox(sc, from_=5, to=200, increment=5,
                          textvariable=self.lim_var, width=6,
                          bg="#12121e", fg=FG, buttonbackground=CARD,
                          relief="flat", font=("Segoe UI", 11))
        spin.grid(row=1, column=2, sticky="w", padx=(14, 0), ipady=5)

        self.search_btn = ttk.Button(sc, text="  Search Leads",
                                     style="Accent.TButton",
                                     command=self._on_search)
        self.search_btn.grid(row=1, column=3, padx=(18, 0), sticky="w")

        sc.columnconfigure(0, weight=1)
        sc.columnconfigure(1, weight=1)

        for e in (biz_e, loc_e):
            e.bind("<Return>", lambda _: self._on_search())

        # ── scrape row ───────────────────────────────────────────────────
        ttk.Label(sc, text="— or scrape a website directly —",
                  style="Sub.TLabel").grid(row=2, column=0, columnspan=2,
                                           sticky="w", pady=(10, 2))

        self.url_var = tk.StringVar()
        url_e = entry(sc, self.url_var, w=50)
        url_e.configure(font=("Segoe UI", 10))
        url_e.grid(row=3, column=0, columnspan=2, sticky="ew", ipady=6)
        url_e.bind("<Return>", lambda _: self._on_scrape())

        self.scrape_btn = ttk.Button(sc, text="  Scrape Website",
                                     style="Accent.TButton",
                                     command=self._on_scrape)
        self.scrape_btn.grid(row=3, column=3, padx=(18, 0), sticky="w")

        url_e.configure(insertontime=600)
        url_e.delete(0, tk.END)
        url_e.insert(0, "")
        self.url_var.set("")

        # ── status row ───────────────────────────────────────────────────────
        sr = ttk.Frame(self, style="Card.TFrame", padding=(22, 5))
        sr.pack(fill="x", padx=18, pady=(3, 0))
        self.status_var = tk.StringVar(
            value="Enter a business type and location, then click Search.")
        ttk.Label(sr, textvariable=self.status_var, style="Sub.TLabel").pack(side="left")
        self.progress = ttk.Progressbar(sr, mode="indeterminate",
                                        style="Horizontal.TProgressbar", length=100)
        self.progress.pack(side="right")

        # ── results table ────────────────────────────────────────────────────
        tbl = tk.Frame(self, bg=BG)
        tbl.pack(fill="both", expand=True, padx=18, pady=(10, 0))

        cols = [c[0] for c in COLUMNS]
        self.tree = ttk.Treeview(tbl, columns=cols, show="headings",
                                 selectmode="extended")
        for label, _, width in COLUMNS:
            self.tree.heading(label, text=label,
                              command=lambda l=label: self._sort(l))
            self.tree.column(label, width=width, minwidth=60, anchor="w")

        vsb = ttk.Scrollbar(tbl, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(tbl, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tbl.rowconfigure(0, weight=1)
        tbl.columnconfigure(0, weight=1)

        self.tree.tag_configure("odd",  background="#12121e")
        self.tree.tag_configure("even", background="#16162a")

        # ── bottom bar ───────────────────────────────────────────────────────
        bot = ttk.Frame(self, style="Card.TFrame", padding=(18, 9))
        bot.pack(fill="x", padx=18, pady=(5, 13))

        self.count_var = tk.StringVar(value="0 leads")
        ttk.Label(bot, textvariable=self.count_var, style="Card.TLabel",
                  font=("Segoe UI", 10, "bold")).pack(side="left")

        self.excel_btn = ttk.Button(bot, text="  Download Excel",
                                    style="Excel.TButton",
                                    command=self._export_excel)
        self.excel_btn.pack(side="right")

        ttk.Button(bot, text="Clear",
                   command=self._clear, padding=(10, 6)).pack(side="right", padx=(0, 10))

    # ─────────────────────────────────────────────────── search logic ────────

    def _on_search(self):
        biz = self.biz_var.get().strip()
        loc = self.loc_var.get().strip()
        if not biz or not loc:
            messagebox.showwarning("Missing Input",
                                   "Please enter both a business type and a location.")
            return
        self.search_btn.state(["disabled"])
        self.excel_btn.state(["disabled"])
        self.status_var.set(f"Searching for '{biz}' in {loc} …")
        self.progress.start(12)
        self._clear_tree()
        threading.Thread(target=self._run,
                         args=(biz, loc, self.lim_var.get()),
                         daemon=True).start()

    def _run(self, biz, loc, limit):
        try:
            results = search_businesses(biz, loc, limit)
            self.after(0, self._populate, results)
        except Exception as exc:
            self.after(0, self._err, str(exc))

    def _populate(self, results: list[dict]):
        self.progress.stop()
        self.search_btn.state(["!disabled"])
        self.excel_btn.state(["!disabled"])
        self._results = results

        if not results:
            self.status_var.set("No results returned — try a broader query or different location.")
            self.count_var.set("0 leads")
            return

        for i, rec in enumerate(results):
            tag = "even" if i % 2 == 0 else "odd"
            vals = tuple(rec.get(k, "") for _, k, _ in COLUMNS)
            self.tree.insert("", "end", values=vals, tags=(tag,))

        n = len(results)
        self.count_var.set(f"{n} lead{'s' if n != 1 else ''} found")
        self.status_var.set(
            f"{n} result{'s' if n != 1 else ''} for "
            f"'{self.biz_var.get()}' in {self.loc_var.get()}")

    def _err(self, msg: str):
        self.progress.stop()
        self.search_btn.state(["!disabled"])
        self.excel_btn.state(["!disabled"])
        self.status_var.set("Search failed — see error dialog.")
        messagebox.showerror("API Error",
                             f"Could not retrieve results:\n\n{msg}\n\n"
                             "Check that your API key is valid and the "
                             "endpoint path matches your plan.")

    # ─────────────────────────────────────────────────── scrape logic ──────────

    def _on_scrape(self):
        url = self.url_var.get().strip()
        if not url:
            messagebox.showwarning("Missing URL",
                                   "Please enter a URL to scrape.")
            return
        self.scrape_btn.state(["disabled"])
        self.search_btn.state(["disabled"])
        self.excel_btn.state(["disabled"])
        self.status_var.set(f"Scraping {url} …")
        self.progress.start(12)
        self._clear_tree()
        threading.Thread(target=self._run_scrape, args=(url,),
                         daemon=True).start()

    def _run_scrape(self, url):
        try:
            results = scrape_leads(url)
            self.after(0, self._populate_scrape, results, url)
        except Exception as exc:
            self.after(0, self._err_scrape, str(exc))

    def _populate_scrape(self, results, url):
        self.progress.stop()
        self.scrape_btn.state(["!disabled"])
        self.search_btn.state(["!disabled"])
        self.excel_btn.state(["!disabled"])
        self._results = results

        if not results:
            self.status_var.set("No leads extracted — try a different URL.")
            self.count_var.set("0 leads")
            return

        for i, rec in enumerate(results):
            tag = "even" if i % 2 == 0 else "odd"
            vals = tuple(rec.get(k, "") for _, k, _ in COLUMNS)
            self.tree.insert("", "end", values=vals, tags=(tag,))

        n = len(results)
        self.count_var.set(f"{n} lead{'s' if n != 1 else ''} scraped")
        self.status_var.set(f"{n} lead{'s' if n != 1 else ''} scraped from {url}")

    def _err_scrape(self, msg):
        self.progress.stop()
        self.scrape_btn.state(["!disabled"])
        self.search_btn.state(["!disabled"])
        self.excel_btn.state(["!disabled"])
        self.status_var.set("Scrape failed — see error dialog.")
        messagebox.showerror("Scrape Error",
                             f"Could not scrape the URL:\n\n{msg}")

    # ─────────────────────────────────────────────────── table helpers ────────

    def _clear_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _clear(self):
        self._clear_tree()
        self._results = []
        self.count_var.set("0 leads")
        self.status_var.set("Cleared.")

    def _sort(self, label: str):
        key_map = {c[0]: c[1] for c in COLUMNS}
        key = key_map.get(label, "")
        self._results.sort(key=lambda r: (r.get(key) or "").lower())
        self._clear_tree()
        for i, rec in enumerate(self._results):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end",
                             values=tuple(rec.get(k, "") for _, k, _ in COLUMNS),
                             tags=(tag,))

    # ──────────────────────────────────────────────── Excel export ────────────

    def _export_excel(self):
        if not self._results:
            messagebox.showinfo("Nothing to Export", "Run a search first.")
            return
        if not EXCEL_OK:
            messagebox.showerror("Missing Library",
                                 "openpyxl is not installed.\n"
                                 "Run:  pip install openpyxl")
            return

        biz = self.biz_var.get()
        loc = self.loc_var.get()
        ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
        default = (f"leads_{biz.replace(' ','_')}_"
                   f"{loc.replace(', ','_').replace(' ','_')}_{ts}.xlsx")

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All files", "*.*")],
            initialfile=default,
        )
        if not path:
            return

        try:
            export_excel(self._results, path, biz, loc)
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc))
            return

        self.status_var.set(
            f"Exported {len(self._results)} leads → {os.path.basename(path)}")
        if messagebox.askyesno("Export Complete",
                               f"Saved {len(self._results)} leads to:\n{path}\n\n"
                               "Open the file now?"):
            import subprocess, sys
            if sys.platform.startswith("linux"):
                subprocess.Popen(["xdg-open", path])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                os.startfile(path)


WHITE_STR = "#ffffff"   # module-level constant used by _build_ui before class body ends

if __name__ == "__main__":
    app = LeadGeneratorApp()
    app.mainloop()
