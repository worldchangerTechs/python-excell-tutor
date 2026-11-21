"""
EXCEL TUTOR AI  -  run with:   python app.py      (or: python app.py "C:\\path\\file.xlsx")

 * Attaches to the Excel file you already have open (or opens one / a sample).
 * A small HAND pointer moves over the Excel sheet, taps cells and TYPES into them live.
 * Tables, histograms, charts, filters, heat-maps are built INSIDE the Excel sheet.
 * The assistant panel sits beside Excel. It closes by itself when Excel closes.
   Closing the panel NEVER closes Excel.
 * Voice in (SpeechRecognition) and voice out (pyttsx3). Brain: NVIDIA NIM (OpenAI-compatible).

Install once:
    pip install openai xlwings pyautogui SpeechRecognition pyttsx3 pyaudio

Emergency stop: red STOP button, say "stop", or throw the mouse into a screen corner.
"""
import csv, ctypes, json, math, os, queue, random, re, sys, threading, time, traceback
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, scrolledtext

try:
    import xlwings as xw
except Exception:
    print("xlwings is missing. Run:  pip install xlwings")
    sys.exit(1)
try:
    from openai import OpenAI
except Exception:
    print("openai is missing. Run:  pip install openai")
    sys.exit(1)
try:
    import pyautogui
    pyautogui.FAILSAFE = True
except Exception:
    pyautogui = None
try:
    import pythoncom
except Exception:
    pythoncom = None
try:
    import speech_recognition as sr
except Exception:
    sr = None
try:
    import pyttsx3
except Exception:
    pyttsx3 = None

try:
    from dotenv import load_dotenv
    load_dotenv(override=True)
except Exception:
    pass

# ============================== SETTINGS ==============================
NVIDIA_KEY = os.environ.get("NVIDIA_API_KEY", "")
NVIDIA_BASE = "https://integrate.api.nvidia.com/v1"
NVIDIA_MODELS = [
    os.environ.get("NVIDIA_MODEL", "mistralai/mistral-large-4-1t-a52b"),
    "mistralai/mistral-large-4-1t-a52b",
    "nvidia/nemotron-3.5-lightning-30b-a3b",
    "minimaxai/minimax-m3.1-flash",
    "moonshotai/kimi-k3.1-pro",
]
NVIDIA_MODELS = list(dict.fromkeys(NVIDIA_MODELS))[:2]     # only 2 tries, so it never waits for minutes
SPEECH_LANG = "en-US"                  # try "en-KE" for a Kenyan accent
GUI_FRACTION = 0.30                    # share of screen width for the assistant panel
OWN_BY_EXCEL = False                   # True = link windows to Excel at Windows level (risky on some PCs)
EXPORT_DIR = os.path.join(os.path.expanduser("~"), "Documents", "ExcelTutor_Exports")
# ======================================================================

ai = OpenAI(api_key=NVIDIA_KEY, base_url=NVIDIA_BASE, timeout=40, max_retries=0)

try:
    _HERE = os.path.dirname(os.path.abspath(__file__))
except NameError:
    _HERE = os.getcwd()
LOGFILE = os.path.join(_HERE, "excel_tutor.log")


def log_line(msg):
    try:
        with open(LOGFILE, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%H:%M:%S}  {msg}\n")
    except Exception:
        pass


threading.excepthook = lambda a: log_line("THREAD ERROR: " + "".join(
    traceback.format_exception(a.exc_type, a.exc_value, a.exc_traceback)))

T0 = time.time()


def stage(msg):
    line = f"[{time.time() - T0:5.1f}s] {msg}"
    print(line, flush=True)
    log_line(line)


UIQ = queue.Queue()                    # everything the GUI must show or animate goes through here
STOP = threading.Event()
LAYOUT = {"gui_left": None}
W = None                               # ExcelWorker, created in main()


# ---------------------------------------------------------------- keep PC awake
def prevent_sleep():
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001 | 0x00000002)
    except Exception:
        pass


def allow_sleep():
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
    except Exception:
        pass


SYSTEM = """You are an expert Excel tutor and data-mining assistant. You CONTROL the user's open Excel sheet
live, like a human with a pointer: you type into cells, build tables, histograms and charts, filter rows.
The user watches the SHEET change. The chat only gets a short summary.

You receive: workbook, sheet, used_range, rows, first/last data row, free_area_start, a list of columns
(letter, header, type, examples) and a data sample.
Reply with ONLY one JSON object, no markdown, no code fences, no thinking text:
{"say": "friendly summary under 50 words of what you built and what it shows",
 "actions": [ ... ]}

Actions:
  {"type":"write","cell":"Y1","value":"Data Mining Dashboard"}
  {"type":"write_data","start_cell":"Y4","data":[["Column","Average"],["Age","=AVERAGE(C2:C10128)"]]}
  {"type":"histogram","column":"C","bins":10,"title":"Age distribution",
        "where":[{"column":"B","op":"=","value":"F"}],"anchor":"Y4"}
        (live histogram: bin table with COUNTIFS formulas + chart. "where" and "anchor" optional;
         op is one of = <> > < >= <=)
  {"type":"chart","kind":"column","source":"Y4:Z12","title":"...","anchor":"AB4"}
        kinds: column bar line line_markers pie scatter area
  {"type":"table","range":"Y4:Z12","name":"Summary","style":"TableStyleMedium9"}
  {"type":"filter","column":"B","criteria":"F"}                 (show only that group; criteria like "F" or ">50")
  {"type":"clear_filter"}
  {"type":"highlight","range":"Y1:AD1","color":"blue"}          yellow green blue orange red purple none
  {"type":"colorscale","range":"Z5:AD17"}                        (3-colour heat-map)
  {"type":"bold","range":"Y1:AD1"}   {"type":"font_size","range":"Y1","size":16}   {"type":"autofit"}
  {"type":"select","range":"A1:D1"}  {"type":"goto","cell":"Y1"}  {"type":"click","cell":"B4"}
  {"type":"scroll","rows":10}        {"type":"autoscroll","rows_per_step":4,"steps":8,"delay":0.7}
  {"type":"export","format":"csv"}   formats: csv pdf xlsx png
RULES:
- Do the real work IN EXCEL with actions. Never put tables or long text in "say".
- Use the exact column letters, header names and row numbers from the state. Never invent ranges.
- Put new tables, summaries and charts in the free area (free_area_start) and below earlier ones. Never
  overwrite the user's data. Use real formulas (English names, commas).
- For "mine / analyse / profile": title cell, a numeric profile table (Average, Median, Min, Max, StDev),
  category counts with COUNTIF, a colorscale, then a histogram or chart of the key finding.
- For "histogram of X for group Y": use the histogram action with "where".
- Keep the whole reply compact so it never gets cut off.
- NEVER write your reasoning, plan or analysis. The first character of your reply must be { and the last }.
- For "explain/mine/histogram of <column>" use ONE histogram action with the column header: it works for
  numbers (bins) AND for text columns like Marital_Status (category counts, percent and chart)."""

COLORS = {"yellow": (255, 255, 0), "green": (146, 208, 80), "blue": (155, 194, 230),
          "orange": (255, 192, 0), "red": (255, 120, 120), "purple": (178, 100, 220)}
CHART_TYPES = {"column": "column_clustered", "bar": "bar_clustered", "line": "line",
               "line_markers": "line_markers", "pie": "pie", "scatter": "xy_scatter", "area": "area"}


# ====================================================================== pure helpers
def col_letter(n):
    s, n = "", int(n)
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_number(letters):
    n = 0
    for ch in str(letters).upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def safe_name(name, default="Table"):
    n = re.sub(r"[^A-Za-z0-9_]", "_", str(name or default))
    if not re.match(r"[A-Za-z_]", n):
        n = "T_" + n
    return n[:60]


def speakable(text):
    t = re.sub(r"[*`#_>|]", " ", str(text))
    return re.sub(r"\s+", " ", t).strip()[:600]


def salvage(raw):
    """Rescue complete action objects from a reply that was cut off or is slightly broken."""
    i = raw.find('"actions"')
    j = raw.find("[", i) if i != -1 else -1
    if j == -1:
        return None
    dec, pos, acts = json.JSONDecoder(), j + 1, []
    while True:
        while pos < len(raw) and raw[pos] in " \n\r\t,":
            pos += 1
        if pos >= len(raw) or raw[pos] != "{":
            break
        try:
            obj, pos = dec.raw_decode(raw, pos)
        except Exception:
            break
        if isinstance(obj, dict):
            acts.append(obj)
    if not acts:
        return None
    say = "Working on it."
    m = re.search(r'"say"\s*:\s*"((?:[^"\\]|\\.)*)"', raw)
    if m:
        try:
            say = json.loads('"' + m.group(1) + '"')
        except Exception:
            pass
    return {"say": say, "actions": acts}


def try_parse(raw):
    """Return {'say','actions'} if the text contains a real JSON reply, else None."""
    raw = re.sub(r"<think>.*?</think>", "", raw or "", flags=re.S).strip()
    s0, e0 = raw.find("{"), raw.rfind("}")
    if s0 != -1 and e0 > s0:
        try:
            data = json.loads(raw[s0:e0 + 1])
            if isinstance(data, dict) and ("actions" in data or "say" in data):
                if not isinstance(data.get("actions"), list):
                    data["actions"] = []
                return data
        except Exception:
            pass
    return salvage(raw)


def parse_reply(raw):
    return try_parse(raw) or {"say": (raw or "").strip()[:400] or "No reply. Try again.", "actions": []}


def build_messages(history, request, state_json):
    msgs = []
    for req, ans in history[-3:]:
        msgs.append({"role": "user", "content": req})
        msgs.append({"role": "assistant", "content": ans})
    msgs.append({"role": "user", "content": f"[EXCEL STATE]\n{state_json}\n\n[REQUEST]\n{request}"})
    return msgs


def nice_edges(lo, hi, bins):
    if hi <= lo:
        hi = lo + 1
    raw = (hi - lo) / bins
    mag = 10 ** math.floor(math.log10(raw))
    step = raw
    for m in (1, 2, 2.5, 5, 10):
        step = m * mag
        if step >= raw:
            break
    start = math.floor(lo / step) * step
    n = int(math.floor((hi - start) / step)) + 1
    return start, step, n


def match_value(v, op, target):
    if v in (None, ""):
        return False
    try:
        a, b = float(v), float(target)
        return {"=": a == b, "<>": a != b, ">": a > b, "<": a < b, ">=": a >= b, "<=": a <= b}.get(op, a == b)
    except (TypeError, ValueError):
        a, b = str(v).strip().lower(), str(target).strip().lower()
        return (a != b) if op == "<>" else (a == b)


# ====================================================================== AI brain (NVIDIA NIM)
def ask_brain(messages):
    """NVIDIA NIM. Uses ONLY the final answer (never the model's thinking text). If a model answers with
    thoughts instead of JSON, it is asked again once, then the next model is tried."""
    base = [{"role": "system", "content": SYSTEM}] + messages
    nudge = {"role": "user", "content": "Reply NOW with ONLY the JSON object. The first character must be {. "
                                        "Do not explain your plan and do not think out loud."}
    last = None
    for m in NVIDIA_MODELS:
        for attempt in range(2):
            kw = dict(model=m, messages=base if attempt == 0 else base + [nudge],
                      max_tokens=4000, temperature=0.2)
            try:
                try:
                    resp = ai.chat.completions.create(response_format={"type": "json_object"}, **kw)
                except Exception as e:
                    low = str(e).lower()
                    if "response_format" in low or "json" in low or "unsupported" in low or "400" in low:
                        resp = ai.chat.completions.create(**kw)
                    else:
                        raise
                text = resp.choices[0].message.content or ""
                if try_parse(text):
                    return text, f"NVIDIA {m}"
                last = RuntimeError(f"{m} answered with thoughts instead of actions")
                print("[WARN]", last)
            except Exception as e:
                last = e
                print(f"[WARN] model {m} failed: {e}")
                break                                   # network/404 problem: go to the next model
    raise RuntimeError(f"The AI did not return usable actions. Last problem: {last}")


# ====================================================================== hand (thread-safe proxies)
def hand_move(x, y, tap=False):
    evt = threading.Event()
    UIQ.put(("hand", (x, y, tap, evt)))
    evt.wait(3)


def hand_hide():
    UIQ.put(("hand_hide", None))


def own_by_excel(tk_window):
    """Make a Tk window OWNED by the Excel window: it then stays above Excel only, goes behind other
    apps, and hides when Excel is minimized. Returns False if Windows refuses (fallback is used)."""
    try:
        if not OWN_BY_EXCEL or not W or not W.hwnd:
            return False
        tk_window.update_idletasks()
        u = ctypes.WinDLL("user32", use_last_error=True)
        u.GetParent.argtypes = [ctypes.c_void_p]
        u.GetParent.restype = ctypes.c_void_p
        u.SetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p]
        u.SetWindowLongPtrW.restype = ctypes.c_void_p
        wid = tk_window.winfo_id()
        hwnd = u.GetParent(wid) or wid
        u.SetWindowLongPtrW(hwnd, -8, W.hwnd)          # GWLP_HWNDPARENT = set the owner window
        return True
    except Exception:
        return False


# ====================================================================== Excel worker thread
class ExcelWorker(threading.Thread):
    """Every Excel (COM) call runs on this one thread, one at a time. It also watches Excel's health."""

    def __init__(self):
        super().__init__(daemon=True)
        self.jobs = queue.Queue()
        self.wb = None
        self.app = None
        self.pid = None
        self.hwnd = None
        self.on_dead = None
        self.fails = 0
        self.dead_sent = False
        self.start()

    def run(self):
        if pythoncom:
            pythoncom.CoInitialize()
        while True:
            try:
                fn, args, out = self.jobs.get(timeout=3)
            except queue.Empty:
                self.check_alive()
                continue
            try:
                out.put((True, self._retry(fn, args)))
            except BaseException as e:      # noqa: BLE001
                out.put((False, e))

    @staticmethod
    def _retry(fn, args):
        for attempt in range(6):
            try:
                return fn(*args)
            except Exception as e:
                msg = str(e)
                busy = "-2147418111" in msg or "rejected" in msg.lower() or "-2147417846" in msg
                if busy and attempt < 5:
                    time.sleep(0.8)         # Excel busy (e.g. you are typing in a cell): wait, retry
                    continue
                raise

    def call(self, fn, *args, timeout=600):
        out = queue.Queue()
        self.jobs.put((fn, args, out))
        ok, val = out.get(timeout=timeout)
        if ok:
            return val
        raise val

    def check_alive(self):
        if self.pid is None or self.dead_sent:
            return
        try:
            alive = self.pid in list(xw.apps.keys())
        except Exception:
            return
        self.fails = 0 if alive else self.fails + 1
        if self.fails >= 2 and self.on_dead:
            self.dead_sent = True
            self.on_dead()


# ====================================================================== Excel helpers (run on worker)
def current():
    try:
        wb = W.app.books.active
    except Exception:
        wb = None
    wb = wb or W.wb
    try:
        wb.activate()
    except Exception:
        pass
    return wb, wb.sheets.active


def has_data(rng):
    vals = rng.options(ndim=2).value
    return any(v not in (None, "") for row in vals for v in row)


def goto(rng_com):
    try:
        W.app.api.Goto(rng_com, True)
    except Exception:
        try:
            rng_com.Select()
        except Exception:
            pass


def ensure_visible(rng_com):
    try:
        if W.app.api.Intersect(rng_com, W.app.api.ActiveWindow.VisibleRange) is None:
            W.app.api.Goto(rng_com, True)
    except Exception:
        goto(rng_com)


def screen_pos(rng_com):
    win = W.app.api.ActiveWindow
    c = rng_com.Cells(1, 1)
    return (int(win.PointsToScreenPixelsX(c.Left + c.Width / 2)),
            int(win.PointsToScreenPixelsY(c.Top + c.Height / 2)))


def hand_to(rng_com, tap=False):
    """Slide the visible hand over a cell (and tap it)."""
    try:
        ensure_visible(rng_com)
        x, y = screen_pos(rng_com)
        gl = LAYOUT["gui_left"]
        if x > 0 and y > 0 and (gl is None or x < gl - 10):
            hand_move(x, y, tap)
    except Exception:
        pass


def resolve_column(sh, ref):
    """Accepts a header name ('Age') or a column letter ('C'). Header names win, so 'Age' is not read as a letter."""
    ref = str(ref or "").strip()
    ur = sh.used_range
    headers = ur.resize(1).options(ndim=1).value
    headers = headers if isinstance(headers, list) else [headers]
    for j, h in enumerate(headers):
        if h is not None and str(h).strip().lower() == ref.lower():
            return col_letter(ur.column + j)
    if re.fullmatch(r"[A-Za-z]{1,3}", ref):
        return ref.upper()
    raise ValueError(f"Column '{ref}' not found. Use a column letter or an exact header name.")


def first_numeric_column(sh):
    ur = sh.used_range
    nrows, ncols = ur.shape
    vals = ur.resize(min(nrows, 31), ncols).options(ndim=2).value
    for j in range(ncols):
        col = [vals[i][j] for i in range(1, len(vals))]
        if col and sum(1 for v in col if is_number(v)) >= max(1, len(col) // 2):
            return col_letter(ur.column + j)
    raise ValueError("I could not find a numeric column.")


def write_data(sh, cell, data, overwrite=False, human=True):
    """Type data into the sheet cell by cell, with the hand following, like a person typing."""
    if not isinstance(data, list) or not data:
        raise ValueError("data must be a non-empty list of rows")
    rows = [list(r) if isinstance(r, (list, tuple)) else [r] for r in data]
    width = max(len(r) for r in rows)
    rows = [r + [None] * (width - len(r)) for r in rows]
    target = sh.range(cell).resize(len(rows), width)
    if not overwrite and has_data(target):
        raise ValueError(f"{target.address.replace('$', '')} already has data. Pick an empty area "
                         f"(or ask me to overwrite).")
    total = len(rows) * width
    if not human or total > 600:                      # big block: write at once
        target.value = [[None if is_formula(v) else v for v in r] for r in rows]
        for i, r in enumerate(rows):
            for j, v in enumerate(r):
                if is_formula(v):
                    sh.range((target.row + i, target.column + j)).formula = v
        goto(target.api)
        return target
    delay = 0.14 if total <= 60 else (0.05 if total <= 200 else 0.01)
    for i, r in enumerate(rows):
        if STOP.is_set():
            raise RuntimeError("Stopped by user")
        for j, v in enumerate(r):
            if v is None:
                continue
            c = sh.range((target.row + i, target.column + j))
            if total <= 60 or (j == 0 and i % (1 if total <= 200 else 4) == 0):
                hand_to(c.api, tap=total <= 60)
            try:
                c.select()
            except Exception:
                pass
            if is_formula(v):
                c.formula = v
            else:
                c.value = v
            time.sleep(delay)
    return target


def snapshot():
    wb, sh = current()
    ur = sh.used_range
    nrows, ncols = ur.shape
    sample = ur.resize(min(nrows, 25), min(ncols, 40)).options(ndim=2).value
    while sample and all(v in (None, "") for v in sample[-1]):
        sample.pop()
    cols = []
    for j in range(len(sample[0]) if sample else 0):
        vals = [row[j] for row in sample[1:] if j < len(row) and row[j] not in (None, "")]
        nums = [v for v in vals if is_number(v)]
        info = {"col": col_letter(ur.column + j), "header": sample[0][j],
                "type": "number" if vals and len(nums) >= 0.7 * len(vals) else "text"}
        if info["type"] == "text":
            info["examples"] = sorted({str(v) for v in vals})[:6]
        elif nums:
            info["sample_min_max"] = [min(nums), max(nums)]
        cols.append(info)
    try:
        selection = wb.app.selection.address.replace("$", "")
    except Exception:
        selection = "unknown"
    return {
        "workbook": wb.name, "sheet": sh.name, "all_sheets": [s.name for s in wb.sheets],
        "selection": selection, "used_range": ur.address.replace("$", ""),
        "rows": nrows, "columns": ncols,
        "header_row": ur.row, "first_data_row": ur.row + 1, "last_data_row": ur.row + nrows - 1,
        "free_area_start": f"{col_letter(ur.column + ncols + 1)}{ur.row}",
        "columns_info": cols, "data_sample": sample[:12],
    }


def add_chart(sh, src, anchor, kind, title, width=440, height=280, gap=None, xtitle=None):
    ch = sh.charts.add(left=anchor.left, top=anchor.top, width=float(width), height=float(height))
    api = ch.api
    xl = api[1] if isinstance(api, tuple) else api
    xl.SetSourceData(sh.api.Range(src))
    ch.chart_type = CHART_TYPES.get(kind, "column_clustered")
    try:
        xl.HasTitle = True
        xl.ChartTitle.Text = str(title or "Chart")
    except Exception:
        pass
    if gap is not None:
        try:
            xl.ChartGroups(1).GapWidth = gap
            xl.HasLegend = False
        except Exception:
            pass
    if xtitle:
        try:
            xl.Axes(1).HasTitle = True
            xl.Axes(1).AxisTitle.Text = str(xtitle)
        except Exception:
            pass
    hand_to(anchor.api, tap=True)
    return ch


def auto_source(sh, max_rows=30):
    ur = sh.used_range
    nrows, ncols = ur.shape
    if nrows < 2 or ncols < 2:
        raise ValueError("I need at least 2 rows and 2 columns of data to chart.")
    last = min(nrows, max_rows + 1)
    vals = ur.resize(last, ncols).options(ndim=2).value
    num = None
    for j in range(1, ncols):
        col = [vals[i][j] for i in range(1, last)]
        if sum(1 for v in col if is_number(v)) >= max(1, len(col) // 2):
            num = j
            break
    if num is None:
        raise ValueError("I could not find a numeric column to chart.")
    a, b = col_letter(ur.column), col_letter(ur.column + num)
    r1, r2 = ur.row, ur.row + last - 1
    return f"{a}{r1}:{b}{r2}" if num == 1 else f"{a}{r1}:{a}{r2},{b}{r1}:{b}{r2}"


def make_chart(sh, a):
    kind = str(a.get("kind") or a.get("chart_type") or "column").lower()
    src = a.get("source") or a.get("data_range") or auto_source(sh)
    ur = sh.used_range
    if a.get("anchor"):
        anchor = sh.range(a["anchor"])
    else:
        anchor = sh.cells(ur.row, ur.column + ur.shape[1] + 1)
    add_chart(sh, src, anchor, kind, a.get("title"), a.get("width", 440), a.get("height", 280))
    return f"Chart ({kind}) from {src} at {anchor.address.replace('$', '')}"


def find_free_row(sh, col_idx, need_rows, width, start_row):
    r = start_row
    for _ in range(60):
        if not has_data(sh.range((r, col_idx)).resize(need_rows, width)):
            return r
        r += need_rows + 3
    return r


def place_block(sh, ur, a, need_rows, width):
    if a.get("anchor"):
        c = sh.range(a["anchor"])
        return c.row, c.column
    col0 = ur.column + ur.shape[1] + 1
    return find_free_row(sh, col0, need_rows, width, ur.row), col0


def style_table(sh, tgt, name):
    try:
        t = sh.api.ListObjects.Add(1, tgt.api, None, 1)
        t.Name = safe_name(name)
        t.TableStyle = "TableStyleMedium9"
    except Exception:
        tgt.resize(1).api.Font.Bold = True


def make_histogram(sh, a):
    """Histogram of a numeric column (bins) OR frequency chart of a text column (categories)."""
    col = resolve_column(sh, a.get("column") or first_numeric_column(sh))
    ur = sh.used_range
    r1, r2 = ur.row + 1, ur.row + ur.shape[0] - 1
    if r2 < r1:
        raise ValueError("There is no data under the header row.")
    hdrs = ur.resize(1).options(ndim=1).value
    hdrs = hdrs if isinstance(hdrs, list) else [hdrs]
    header = hdrs[col_number(col) - ur.column]
    values = sh.range(f"{col}{r1}:{col}{r2}").options(ndim=1).value
    values = values if isinstance(values, list) else [values]
    conds = []
    for w in (a.get("where") or []):
        wc = resolve_column(sh, w.get("column"))
        wv = sh.range(f"{wc}{r1}:{wc}{r2}").options(ndim=1).value
        conds.append((wc, str(w.get("op", "=")), w.get("value"), wv if isinstance(wv, list) else [wv]))
    ok = [i for i, v in enumerate(values)
          if v not in (None, "") and all(match_value(c[3][i], c[1], c[2]) for c in conds)]
    if not ok:
        raise ValueError("No rows match that group, so there is nothing to plot.")
    nums = [values[i] for i in ok if is_number(values[i])]
    full = f"${col}${r1}:${col}${r2}"
    extra = ""
    for wc, op, val, _ in conds:
        extra += f',${wc}${r1}:${wc}${r2},"{op}{str(val).replace(chr(34), chr(34) * 2)}"'
    suffix = " (filtered group)" if conds else ""

    if len(nums) < 0.6 * len(ok):                      # ---- text column: category frequencies
        counts = {}
        for i in ok:
            k = str(values[i]).strip()
            counts[k] = counts.get(k, 0) + 1
        cats = [k for k, _ in sorted(counts.items(), key=lambda kv: -kv[1])][:20]
        n = len(cats)
        row0, col0 = place_block(sh, ur, a, n + 5, 3)
        write_data(sh, (row0, col0), [[f"{header} - distribution{suffix}"]], overwrite=True)
        sh.range((row0, col0)).api.Font.Bold = True
        trow = row0 + 1
        table = [[header, "Count", "Percent"]]
        for k in range(n):
            crit = cats[k].replace('"', '""')
            r = trow + 1 + k
            cl, pl = col_letter(col0 + 1), col_letter(col0 + 2)
            table.append([cats[k], f'=COUNTIFS({full},"={crit}"{extra})',
                          f"={cl}{r}/${cl}${trow + n + 1}"])
        cl = col_letter(col0 + 1)
        table.append(["Total", f"=SUM({cl}{trow + 1}:{cl}{trow + n})",
                      f"=SUM({col_letter(col0 + 2)}{trow + 1}:{col_letter(col0 + 2)}{trow + n})"])
        tgt = write_data(sh, (trow, col0), table, overwrite=False)
        try:
            sh.range((trow + 1, col0 + 2)).resize(n + 1, 1).api.NumberFormat = "0.0%"
        except Exception:
            pass
        style_table(sh, tgt, f"Cat_{header}_{row0}")
        src = sh.range((trow, col0)).resize(n + 1, 2).address.replace("$", "")
        add_chart(sh, src, sh.cells(row0, col0 + 4), "column", a.get("title") or f"{header} distribution",
                  width=460, height=290, gap=60, xtitle=header)
        return f"Frequency chart of '{header}' ({len(ok)} rows, {n} categories) at {tgt.address.replace('$', '')}"

    bins = max(3, min(25, int(a.get("bins", 10))))      # ---- numeric column: bins
    start, step, n = nice_edges(min(nums), max(nums), bins)
    table = [[f"{header} range", "Count"]]
    for k in range(n):
        lo, hi = round(start + k * step, 10), round(start + (k + 1) * step, 10)
        table.append([f"{lo:g}-{hi:g}",
                      f'=COUNTIFS({full},">={lo:.10g}",{full},"<{hi:.10g}"{extra})'])
    row0, col0 = place_block(sh, ur, a, len(table) + 2, 2)
    write_data(sh, (row0, col0), [[f"Histogram: {header}{suffix}"]], overwrite=True)
    sh.range((row0, col0)).api.Font.Bold = True
    tgt = write_data(sh, (row0 + 1, col0), table, overwrite=False)
    style_table(sh, tgt, f"Hist_{header}_{row0}")
    addr = tgt.address.replace("$", "")
    add_chart(sh, addr, sh.cells(row0, col0 + 3), "column", a.get("title") or f"Histogram of {header}",
              width=460, height=290, gap=5, xtitle=header)
    return f"Histogram of '{header}' ({len(nums)} rows, {n} bins) at {addr}"


def make_table(sh, a):
    if a.get("data"):
        cell = a.get("cell") or str(a.get("range", "A1")).split(":")[0]
        rng = write_data(sh, cell, a["data"], overwrite=bool(a.get("overwrite")))
    elif a.get("range"):
        rng = sh.range(a["range"])
    else:
        rng = sh.used_range
    hand_to(rng.api, tap=True)
    names = []
    try:
        lo = sh.api.ListObjects
        names = [lo.Item(i).Name for i in range(1, lo.Count + 1)]
    except Exception:
        pass
    base = safe_name(a.get("name"), "Table")
    name, n = base, 1
    while name in names:
        n += 1
        name = f"{base}_{n}"
    try:
        t = sh.api.ListObjects.Add(1, rng.api, None, 1)
        t.Name = name
        t.TableStyle = a.get("style") or "TableStyleMedium9"
        note = f"Excel table '{name}'"
    except Exception:
        h = rng.resize(1)
        h.api.Font.Bold = True
        h.api.Interior.Color = 0x7F4F1F
        h.api.Font.Color = 0xFFFFFF
        rng.api.Borders.LineStyle = 1
        note = "formatted table (range was already a table)"
    try:
        rng.columns.autofit()
    except Exception:
        pass
    return f"{note} at {rng.address.replace('$', '')}"


def do_scroll(n):
    win = W.app.api.ActiveWindow
    n = int(n)
    step = 1 if abs(n) <= 30 else 3
    done = 0
    try:
        while done < abs(n):
            if STOP.is_set():
                return
            k = min(step, abs(n) - done)
            win.ScrollRow = max(1, min(1048576, int(win.ScrollRow) + (k if n > 0 else -k)))
            done += k
            time.sleep(0.035)                         # smooth, visible scrolling
    except Exception:
        if pyautogui:
            pyautogui.scroll(-n * 40)


def park_hand_center():
    try:
        win = W.app.api.ActiveWindow
        vr = win.VisibleRange
        x = int(win.PointsToScreenPixelsX(vr.Left + vr.Width / 2))
        y = int(win.PointsToScreenPixelsY(vr.Top + vr.Height / 3))
        gl = LAYOUT["gui_left"]
        if gl is not None:
            x = min(x, gl - 120)
        hand_move(x, y, False)
    except Exception:
        pass


def do_autoscroll(rows_per_step, steps, delay):
    park_hand_center()
    for _ in range(max(1, min(int(steps), 60))):
        if STOP.is_set():
            return "Auto-scroll stopped."
        do_scroll(rows_per_step)
        time.sleep(max(0.15, float(delay)))
    return "Auto-scroll finished."


def do_export(sh, a):
    fmt = str(a.get("format", "csv")).lower().lstrip(".")
    os.makedirs(EXPORT_DIR, exist_ok=True)
    base = safe_name(os.path.splitext(current()[0].name)[0], "workbook")
    path = os.path.join(EXPORT_DIR, f"{base}_{safe_name(sh.name)}_{datetime.now().strftime('%H%M%S')}.{fmt}")
    if fmt == "csv":
        vals = sh.used_range.options(ndim=2).value
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows([["" if v is None else v for v in row] for row in vals])
    elif fmt == "pdf":
        sh.api.ExportAsFixedFormat(0, path)
    elif fmt in ("xlsx", "excel"):
        path = path[:-len(fmt)] + "xlsx"
        wb = current()[0]
        sh.api.Copy()                                  # copy of the sheet (with charts) in a new workbook
        nwb = W.app.api.ActiveWorkbook
        try:
            nwb.SaveAs(path, 51)                       # 51 = real .xlsx, even if the source is a CSV
        finally:
            nwb.Close(False)
            try:
                wb.activate()
            except Exception:
                pass
    elif fmt == "png":
        if sh.charts.count == 0:
            raise ValueError("There is no chart on this sheet yet.")
        sh.charts[sh.charts.count - 1].to_png(path)
    else:
        raise ValueError(f"Unknown export format '{fmt}'. Use csv, pdf, xlsx or png.")
    return f"Exported to {path}"


def run_action(a):
    t = str(a.get("type", "")).lower()
    wb, sh = current()
    if t in ("write", "edit"):
        cell = a.get("cell") or a.get("range")
        data = a.get("data") if a.get("data") is not None else [[a.get("value", "")]]
        r = write_data(sh, cell, data, overwrite=bool(a.get("overwrite")))
        return f"Typed into {r.address.replace('$', '')}"
    if t in ("write_data", "write_matrix", "table_data"):
        start = a.get("start_cell") or a.get("cell") or "A1"
        r = write_data(sh, start, a.get("data") or a.get("values"), overwrite=bool(a.get("overwrite")))
        return f"Typed table at {r.address.replace('$', '')}"
    if t == "select":
        rng = sh.api.Range(a.get("range") or a.get("cell"))
        hand_to(rng, tap=True)
        rng.Select()
        return f"Selected {a.get('range') or a.get('cell')}"
    if t == "goto":
        rng = sh.api.Range(a["cell"])
        goto(rng)
        hand_to(rng, tap=True)
        return f"Went to {a['cell']}"
    if t == "click":
        rng = sh.api.Range(a["cell"])
        hand_to(rng, tap=True)
        rng.Select()
        return f"Clicked {a['cell']}"
    if t == "highlight":
        c = str(a.get("color", "yellow")).lower()
        rng = sh.api.Range(a["range"])
        hand_to(rng, tap=True)
        if c == "none":
            rng.Interior.ColorIndex = -4142
        else:
            r, g, b = COLORS.get(c, COLORS["yellow"])
            rng.Interior.Color = r + g * 256 + b * 65536
        return f"Highlighted {a['range']} ({c})"
    if t == "colorscale":
        rng = sh.api.Range(a["range"])
        hand_to(rng, tap=True)
        rng.FormatConditions.Delete()
        rng.FormatConditions.AddColorScale(3)
        return f"Heat-map on {a['range']}"
    if t == "bold":
        sh.api.Range(a["range"]).Font.Bold = True
        return f"Bold {a['range']}"
    if t == "font_size":
        sh.api.Range(a["range"]).Font.Size = int(a.get("size", 12))
        return f"Font size on {a['range']}"
    if t == "autofit":
        sh.used_range.columns.autofit()
        return "Columns auto-fitted"
    if t == "scroll":
        park_hand_center()
        do_scroll(a.get("rows", 5))
        return f"Scrolled {a.get('rows', 5)} rows"
    if t == "autoscroll":
        return do_autoscroll(a.get("rows_per_step", 4), a.get("steps", 8), a.get("delay", 0.7))
    if t == "table":
        return make_table(sh, a)
    if t == "chart":
        return make_chart(sh, a)
    if t == "histogram":
        return make_histogram(sh, a)
    if t == "filter":
        ur = sh.used_range
        col = resolve_column(sh, a.get("column"))
        field = col_number(col) - ur.column + 1
        hand_to(sh.cells(ur.row, col_number(col)).api, tap=True)
        sh.api.Range(a.get("range") or ur.address).AutoFilter(field, str(a.get("criteria", "")))
        return f"Filtered column {col} for {a.get('criteria')}"
    if t == "clear_filter":
        try:
            if sh.api.AutoFilterMode:
                sh.api.AutoFilterMode = False
        except Exception:
            pass
        return "Filter cleared"
    if t == "export":
        return do_export(sh, a)
    if t == "wait":
        time.sleep(min(float(a.get("seconds", 1)), 10))
        return "Waited"
    return f"Unknown action '{t}' (skipped)"


# ---- offline helpers for the buttons (work even if the AI is down)
def local_table():
    return make_table(current()[1], {})


def local_chart():
    return make_chart(current()[1], {"kind": "column", "title": "Chart of your data"})


def local_histogram():
    return make_histogram(current()[1], {"bins": 10})


def local_autoscroll():
    sh = current()[1]
    nrows = sh.used_range.shape[0]
    goto(sh.api.Range("A1"))
    return do_autoscroll(6, max(1, min(30, nrows // 6 + 1)), 0.6)


def local_export(fmt):
    return do_export(current()[1], {"format": fmt})


def guess_column(sh, text):
    ur = sh.used_range
    hdrs = ur.resize(1).options(ndim=1).value
    hdrs = hdrs if isinstance(hdrs, list) else [hdrs]
    low = text.lower().replace("_", " ")
    best = None
    for h in hdrs:
        if h is None:
            continue
        hl = str(h).lower().replace("_", " ")
        if hl and hl in low and (best is None or len(hl) > len(best)):
            best = str(h)
    return best


def local_mine():
    """Instant data-mining dashboard, no AI: stats table, heat-map, histogram, category chart."""
    sh = current()[1]
    ur = sh.used_range
    nrows, ncols = ur.shape
    if nrows < 3:
        raise ValueError("There is not enough data to mine.")
    sample = ur.resize(min(nrows, 60), ncols).options(ndim=2).value
    r1, r2 = ur.row + 1, ur.row + nrows - 1
    nums, texts = [], []
    for j in range(ncols):
        vals = [row[j] for row in sample[1:] if row[j] not in (None, "")]
        if not vals:
            continue
        name = str(sample[0][j])
        if sum(1 for v in vals if is_number(v)) >= 0.8 * len(vals):
            if not re.search(r"(^id$|_id$|number|clientnum|index)", name.lower()):
                nums.append((j, name))
        elif len({str(v) for v in vals}) <= 12:
            texts.append((j, name))
    if not nums:
        raise ValueError("I found no numeric columns to profile.")
    nums = nums[:8]
    rows = [["Column", "Average", "Median", "Min", "Max", "StDev"]]
    for j, name in nums:
        c = col_letter(ur.column + j)
        rg = f"${c}${r1}:${c}${r2}"
        rows.append([name, f"=AVERAGE({rg})", f"=MEDIAN({rg})", f"=MIN({rg})", f"=MAX({rg})",
                     f"=STDEV.S({rg})"])
    row0, col0 = place_block(sh, ur, {}, len(rows) + 3, 6)
    write_data(sh, (row0, col0), [["DATA MINING DASHBOARD"]], overwrite=True)
    title = sh.range((row0, col0)).resize(1, 6)
    title.api.Font.Bold = True
    title.api.Font.Size = 14
    title.api.Interior.Color = 155 + 194 * 256 + 230 * 65536
    tgt = write_data(sh, (row0 + 1, col0), rows, overwrite=False)
    style_table(sh, tgt, f"Profile_{row0}")
    try:
        body = sh.range((row0 + 2, col0 + 1)).resize(len(nums), 5)
        body.api.NumberFormat = "#,##0.00"
        sh.range((row0 + 2, col0 + 1)).resize(len(nums), 1).api.FormatConditions.AddColorScale(3)
    except Exception:
        pass
    done = [f"profile of {len(nums)} numeric columns"]
    done.append(make_histogram(sh, {"column": nums[0][1]}))
    if texts:
        done.append(make_histogram(sh, {"column": texts[0][1]}))
    return "Data-mining dashboard built: " + "; ".join(done)


BUILTIN_WORDS = re.compile(r"\b(histograms?|distribution|frequency|charts?|graphs?|tables?|scroll(ing)?|"
                           r"auto-?scroll|export|mine|mining|profile|dashboard|analy[sz]e)\b", re.I)


def wants_builtin(text):
    return bool(BUILTIN_WORDS.search(text or ""))


def offline_command(text):
    """Understands simple commands without any AI, so something always happens in Excel."""
    sh = current()[1]
    low = (text or "").lower()
    if re.search(r"\b(mine|mining|profile|dashboard|analy[sz]e)\b", low):
        return local_mine()
    if re.search(r"\b(histograms?|distribution|frequency)\b", low):
        return make_histogram(sh, {"column": guess_column(sh, text), "bins": 10})
    if re.search(r"\b(charts?|graphs?)\b", low):
        return make_chart(sh, {"kind": "column", "title": "Chart of your data"})
    if re.search(r"\btables?\b", low):
        return make_table(sh, {})
    if "scroll" in low:
        return local_autoscroll()
    if "export" in low:
        for fmt in ("pdf", "xlsx", "csv"):
            if fmt in low:
                return do_export(sh, {"format": fmt})
        return do_export(sh, {"format": "csv"})
    return None


def startup_check():
    wb, sh = current()
    goto(sh.api.Range("A1"))
    hand_to(sh.api.Range("A1"), tap=True)
    return f"Hand check: you should see the hand tap cell A1 of '{wb.name}'."


def handle_request(text, history, say, speak, ask_fn):
    """The whole command pipeline. Built-in tools act FIRST (instantly), then the AI explains."""
    STOP.clear()
    log_line(f"REQUEST: {text}")
    built = None
    if wants_builtin(text):
        say("⚡ Doing it in Excel now...", "sys")
        try:
            built = W.call(offline_command, text)
            if built:
                say("▶ " + str(built), "act")
                speak("Done. Look at the sheet.")
        except Exception as e:
            log_line("BUILTIN ERROR: " + traceback.format_exc())
            say(f"⚠️ {e}", "err")
    try:
        state = W.call(snapshot)
    except Exception as e:
        state = {"note": f"Could not read Excel: {e}"}
        say(f"⚠️ Could not read Excel ({e}).", "err")
    req = text
    if built:
        req = (f"{text}\n[NOTE] I already built this in Excel: {built}. Reply with a short spoken explanation "
               f"of what the user sees and the main insight. Use \"actions\": [] unless truly needed.")
    say("🧠 Asking the AI...", "sys")
    try:
        raw, brain = ask_fn(build_messages(history, req, json.dumps(state, default=str)))
    except Exception as e:
        log_line(f"AI ERROR: {e}")
        if built:
            say(f"(The AI explanation is unavailable: {e}. The sheet is already done.)", "sys")
            return
        say(f"AI problem: {e}", "err")
        try:
            msg = W.call(offline_command, text)
        except Exception as e2:
            log_line("OFFLINE ERROR: " + traceback.format_exc())
            return say(f"⚠️ {e2}", "err")
        if msg:
            say("▶ (built-in) " + str(msg), "act")
            speak("I did that with my built in tools.")
        else:
            say("I could not reach the AI and this request needs it. Try: histogram of <column>, "
                "mine data, chart, table, auto scroll, export pdf.", "err")
        return
    data = parse_reply(raw)
    history.append((text, raw))
    msg = str(data.get("say", ""))
    say(f"Tutor ({brain}): {msg}", "tutor")
    speak(msg)
    if built:
        return                                   # sheet already built; do not duplicate
    for a in data["actions"]:
        if STOP.is_set():
            break
        if not isinstance(a, dict):
            continue
        try:
            say("▶ " + str(W.call(run_action, a)), "act")
        except Exception as e:
            log_line("ACTION ERROR: " + traceback.format_exc())
            say(f"⚠️ {a.get('type', 'action')} failed: {e}", "err")
    if data["actions"]:
        say("✅ Done.", "sys")


def test_excel():
    wb, sh = current()
    info = snapshot()
    park_hand_center()
    goto(sh.api.Range("A1"))
    hand_to(sh.api.Range("A1"), tap=True)
    do_scroll(12)
    time.sleep(0.5)
    do_scroll(-12)
    sh.api.Range("A1").Select()
    return (f"Excel control works. Workbook '{info['workbook']}', sheet '{info['sheet']}', "
            f"data in {info['used_range']}.")



def layout_excel(px_left, px_top, px_w, px_h, dpi):
    k = 72.0 / float(dpi)
    api = W.app.api
    api.WindowState = -4143
    api.Left, api.Top = px_left * k, px_top * k
    api.Width, api.Height = px_w * k, px_h * k
    return "ok"


def create_sample():
    wb = xw.Book()
    sh = wb.sheets[0]
    sh.name = "Sales"
    rnd = random.Random(7)
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    prices = {"Laptop": 900, "Phone": 600, "Tablet": 350}
    rows = [["Month", "Region", "Product", "Units", "Unit Price", "Revenue"]]
    for i in range(60):
        prod = rnd.choice(list(prices))
        rows.append([months[i % 12], rnd.choice(["North", "South", "East", "West"]), prod,
                     rnd.randint(5, 60), prices[prod], f"=D{i + 2}*E{i + 2}"])
    write_data(sh, "A1", rows, overwrite=True, human=False)
    sh.api.Range("A1:F1").Font.Bold = True
    sh.used_range.columns.autofit()
    return wb


def wb_has_data(b):
    try:
        s = b.sheets.active
        return s.used_range.address != "$A$1" or s.range("A1").value is not None
    except Exception:
        return False


def find_open_workbook():
    try:
        stage("looking for running Excel...")
        n = xw.apps.count
        stage(f"Excel instances visible to Python: {n}")
        if n == 0:
            return None
        cands = [b for app in xw.apps for b in app.books]
        stage(f"open workbooks found: {[b.name for b in cands]}")
    except Exception as e:
        stage(f"could not list Excel: {e}")
        return None
    if not cands:
        return None
    try:
        act = xw.apps.active.books.active
    except Exception:
        act = None
    if act is not None and wb_has_data(act):
        return act
    for b in cands:
        if wb_has_data(b):
            return b
    return act or cands[0]


def open_workbook(path, allow_sample):
    sample = False
    if path:
        stage(f"opening file: {path}")
        wb = xw.Book(path)
    else:
        wb = find_open_workbook()
        if wb is None:
            if not allow_sample:
                return None
            wb = create_sample()
            sample = True
    stage(f"connected to workbook: {wb.name}")
    W.wb, W.app, W.pid = wb, wb.app, wb.app.pid
    try:
        W.hwnd = int(wb.app.api.Hwnd)
    except Exception:
        W.hwnd = None
    try:
        wb.app.visible = True
        wb.app.activate(steal_focus=True)
    except Exception:
        pass
    return {"name": wb.name, "sample": sample}


# ====================================================================== speech output
class Speaker(threading.Thread):
    def __init__(self):
        super().__init__(daemon=True)
        self.q = queue.Queue()
        self.busy = threading.Event()
        self.enabled = True
        self.start()

    def run(self):
        if pythoncom:
            pythoncom.CoInitialize()
        while True:
            text = self.q.get()
            if text is None:
                return
            self.busy.set()
            try:
                if pyttsx3:
                    eng = pyttsx3.init()
                    eng.setProperty("rate", 170)
                    eng.say(text)
                    eng.runAndWait()
                    eng.stop()
            except Exception:
                pass
            finally:
                if self.q.empty():
                    self.busy.clear()

    def say(self, text):
        if self.enabled and pyttsx3 and text:
            self.q.put(speakable(text))

    def clear(self):
        try:
            while True:
                self.q.get_nowait()
        except queue.Empty:
            pass


# ====================================================================== the pointing hand
class Hand:
    """A click-through hand pointer that glides over Excel and taps cells."""

    def __init__(self, root):
        self.root = root
        self.key = "#010101"
        self.win = tk.Toplevel(root)
        self.win.withdraw()
        self.win.overrideredirect(True)
        self.win.configure(bg=self.key)
        try:
            self.win.attributes("-transparentcolor", self.key)
        except Exception:
            pass
        self.lbl = tk.Label(self.win, text="👆", font=("Segoe UI Emoji", 34), bg=self.key,
                            fg="#ffd24a", bd=0)
        self.lbl.pack()
        self.x = self.y = None
        self.visible = False
        try:
            self.win.update_idletasks()
            u = ctypes.windll.user32
            hwnd = u.GetParent(self.win.winfo_id()) or self.win.winfo_id()
            u.SetWindowLongW(hwnd, -20, u.GetWindowLongW(hwnd, -20) | 0x80000 | 0x20 | 0x80)
        except Exception:
            pass
        self.owned = own_by_excel(self.win)

    def go(self, tx, ty, tap, evt):
        try:
            if not self.visible:
                self.win.deiconify()
                self.visible = True
            self.win.lift()
            self.win.update_idletasks()
            tx, ty = tx - self.lbl.winfo_reqwidth() // 2, ty - 6
            if self.x is None:
                self.x, self.y = tx - 90, ty + 90
            sx, sy = self.x, self.y
            steps = 12

            def step(i):
                e = 1 - (1 - i / steps) ** 3
                self.x, self.y = sx + (tx - sx) * e, sy + (ty - sy) * e
                self.win.geometry(f"+{int(self.x)}+{int(self.y)}")
                if i < steps:
                    self.root.after(14, lambda: step(i + 1))
                elif tap:
                    self.lbl.config(text="👇")
                    self.root.after(110, lambda: (self.lbl.config(text="👆"), evt.set()))
                else:
                    evt.set()
            step(0)
        except Exception:
            evt.set()

    def hide(self):
        try:
            self.win.withdraw()
        except Exception:
            pass
        self.visible = False


# ====================================================================== GUI
class App:
    BG, FG, PANEL = "#1e1e1e", "#d4d4d4", "#252526"

    def __init__(self, root, info):
        self.root = root
        self.history = []
        self.busy = False
        self.closing = False
        self.hf_on = False
        self.speaker = Speaker()
        prevent_sleep()

        sw, shh = root.winfo_screenwidth(), root.winfo_screenheight()
        gw = max(400, int(sw * GUI_FRACTION))
        gx, gh = sw - gw, shh - 70
        LAYOUT["gui_left"] = gx
        root.title("Excel Tutor AI")
        root.geometry(f"{gw}x{gh}+{gx}+0")
        root.configure(bg=self.BG)
        root.protocol("WM_DELETE_WINDOW", self.close)
        self.hand = Hand(root)
        self.owned = own_by_excel(root)
        if not self.owned:                    # fallback: float only while Excel is the active window
            self.follow_excel()

        hdr = tk.Frame(root, bg="#0078d4")
        hdr.pack(fill="x")
        tk.Label(hdr, text="📊 Excel Tutor AI", bg="#0078d4", fg="white",
                 font=("Segoe UI", 12, "bold")).pack(side="left", padx=10, pady=6)
        tk.Label(hdr, text=f"🟢 {info['name']}", bg="#0078d4", fg="#c8ffd4",
                 font=("Segoe UI", 9, "bold")).pack(side="right", padx=10)

        self.log = scrolledtext.ScrolledText(root, wrap="word", state="disabled", font=("Consolas", 10),
                                             bg=self.BG, fg=self.FG, relief="flat", bd=0)
        self.log.pack(fill="both", expand=True, padx=8, pady=6)
        for tag, color in (("you", "#9cdcfe"), ("tutor", "#4ec9b0"), ("sys", "#808080"),
                           ("act", "#ce9178"), ("err", "#f44747")):
            self.log.tag_config(tag, foreground=color)

        self.entry = tk.Entry(root, font=("Segoe UI", 11), bg=self.PANEL, fg=self.FG,
                              insertbackground="white", relief="flat", bd=4)
        self.entry.pack(fill="x", padx=8, ipady=4)
        self.entry.bind("<Return>", lambda e: self.on_send())
        self.entry.focus_set()

        def row(buttons):
            f = tk.Frame(root, bg=self.BG)
            f.pack(fill="x", padx=8, pady=(4, 0))
            for i, (txt, cmd, color) in enumerate(buttons):
                f.columnconfigure(i, weight=1)
                tk.Button(f, text=txt, command=cmd, relief="flat", bg=color, fg="white",
                          activebackground="#555", activeforeground="white", cursor="hand2",
                          font=("Segoe UI", 9, "bold")).grid(row=0, column=i, sticky="ew", padx=2, ipady=3)
        D = "#2d2d2d"
        row([("▶ Send", self.on_send, "#0078d4"), ("🎤 Speak", self.listen_once, "#0b6e4f"),
             ("■ STOP", self.stop, "#c01c28")])
        row([("⛏️ Mine Data", lambda: self.submit("mine this data"), "#7030a0"),
             ("🧠 Explain sheet", lambda: self.submit(
                 "Explain what this sheet contains, column by column. Highlight the headers and key columns "
                 "and scroll through the data while you explain."), "#7030a0")])
        row([("📈 Histogram", lambda: self.local(local_histogram), D), ("📋 Table", lambda: self.local(local_table), D),
             ("📊 Chart", lambda: self.local(local_chart), D), ("↕ Auto-scroll", lambda: self.local(local_autoscroll), D)])
        row([("CSV", lambda: self.local(local_export, "csv"), D), ("PDF", lambda: self.local(local_export, "pdf"), D),
             ("XLSX", lambda: self.local(local_export, "xlsx"), D), ("📁", self.open_folder, D),
             ("Test", lambda: self.local(test_excel), D), ("Clear", self.clear_log, D)])

        opts = tk.Frame(root, bg=self.BG)
        opts.pack(fill="x", padx=8, pady=(4, 8))
        self.voice_var = tk.BooleanVar(value=True)
        self.hf_var = tk.BooleanVar(value=False)
        for text, var, cmd in (("Voice replies", self.voice_var, self.toggle_voice),
                               ("Hands-free (always listening)", self.hf_var, self.toggle_hf)):
            tk.Checkbutton(opts, text=text, variable=var, command=cmd, bg=self.BG, fg=self.FG,
                           selectcolor="#3e3e3e", activebackground=self.BG, activeforeground=self.FG,
                           font=("Segoe UI", 9)).pack(side="left", padx=4)

        self.say(f"🧠 Brain: NVIDIA NIM ({NVIDIA_MODELS[0]})", "sys")
        if info.get("sample"):
            self.say("Opened a sample sales workbook. Use your own file: python app.py yourfile.xlsx", "sys")
        if str(info["name"]).lower().endswith(".csv"):
            self.say("Note: a CSV cannot store charts or colours. Use the XLSX button to save a full copy.", "sys")
        greet = "Hello! I am ready. Say: mine this data, or make a histogram of age for females."
        self.say("Tutor: " + greet, "tutor")
        self.speaker.say(greet)
        self.say(f"Log file: {LOGFILE}", "sys")
        root.after(60, self.poll)
        root.after(1200, lambda: self.local(startup_check))

    # ---- thread-safe UI plumbing
    def say(self, text, tag="sys"):
        log_line(text)
        UIQ.put(("log", (text, tag)))

    def poll(self):
        try:
            while True:
                kind, p = UIQ.get_nowait()
                if kind == "log":
                    self.log.config(state="normal")
                    self.log.insert("end", p[0] + "\n\n", p[1])
                    self.log.see("end")
                    self.log.config(state="disabled")
                    print(p[0])
                elif kind == "hand":
                    self.hand.go(*p)
                elif kind == "hand_hide":
                    self.hand.hide()
                elif kind == "quit":
                    self.hand.hide()
                    self.log.config(state="normal")
                    self.log.insert("end", "Excel was closed. Closing the assistant...\n", "err")
                    self.log.config(state="disabled")
                    self.root.after(1500, self.close)
        except queue.Empty:
            pass
        if not self.closing:
            self.root.after(40, self.poll)

    def follow_excel(self):
        if self.closing:
            return
        try:
            u = ctypes.windll.user32
            fg = u.GetForegroundWindow()
            ids = set()
            for w in (self.root, self.hand.win):
                wid = w.winfo_id()
                ids.add(u.GetParent(wid) or wid)
            top = bool(fg == W.hwnd or fg in ids)
            self.root.attributes("-topmost", top)
            self.hand.win.attributes("-topmost", top)
        except Exception:
            pass
        self.root.after(300, self.follow_excel)

    def excel_closed(self):
        UIQ.put(("quit", None))

    def close(self):
        """Closes only this assistant. Excel stays open."""
        if self.closing:
            return
        self.closing = True
        self.hf_on = False
        self.speaker.clear()
        self.speaker.q.put(None)
        allow_sleep()
        try:
            self.root.destroy()
        except Exception:
            pass

    def clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    # ---- controls
    def toggle_voice(self):
        self.speaker.enabled = self.voice_var.get()
        if not self.speaker.enabled:
            self.speaker.clear()

    def toggle_hf(self):
        self.hf_on = self.hf_var.get()
        if self.hf_on:
            if not sr:
                self.hf_var.set(False)
                self.hf_on = False
                return self.say("Voice input missing. Run: pip install SpeechRecognition pyaudio", "err")
            threading.Thread(target=self.hands_free_loop, daemon=True).start()
            self.say("🎤 Hands-free ON. Just speak. Say 'stop' to interrupt.", "sys")
        else:
            self.say("Hands-free OFF.", "sys")

    def stop(self):
        STOP.set()
        self.speaker.clear()
        self.say("■ Stopped.", "err")

    def open_folder(self):
        try:
            os.makedirs(EXPORT_DIR, exist_ok=True)
            os.startfile(EXPORT_DIR)
        except Exception:
            self.say(f"Exports are saved in: {EXPORT_DIR}", "sys")

    def on_send(self):
        text = self.entry.get()
        self.entry.delete(0, "end")
        self.submit(text)

    # ---- voice input
    def listen_once(self):
        if not sr:
            return self.say("Voice input missing. Run: pip install SpeechRecognition pyaudio", "err")

        def work():
            try:
                r = sr.Recognizer()
                with sr.Microphone() as src:
                    self.say("🎤 Listening... speak now", "sys")
                    r.adjust_for_ambient_noise(src, 0.4)
                    audio = r.listen(src, timeout=8, phrase_time_limit=15)
                text = r.recognize_google(audio, language=SPEECH_LANG)
            except Exception as e:
                return self.say(f"I didn't catch that ({type(e).__name__}). Try again or type.", "err")
            self.submit(text)
        threading.Thread(target=work, daemon=True).start()

    def hands_free_loop(self):
        try:
            r = sr.Recognizer()
            with sr.Microphone() as src:
                r.adjust_for_ambient_noise(src, 0.8)
                while self.hf_on and not self.closing:
                    if self.speaker.busy.is_set() or self.busy:
                        time.sleep(0.3)
                        continue
                    try:
                        audio = r.listen(src, timeout=3, phrase_time_limit=15)
                    except sr.WaitTimeoutError:
                        continue
                    try:
                        text = r.recognize_google(audio, language=SPEECH_LANG)
                    except sr.UnknownValueError:
                        continue
                    except Exception as e:
                        self.say(f"Speech service problem: {e}", "err")
                        time.sleep(2)
                        continue
                    self.submit(text)
        except Exception as e:
            self.say(f"Microphone problem: {e}", "err")
            self.hf_on = False

    # ---- commands
    def submit(self, text):
        text = (text or "").strip()
        if not text:
            return
        if text.lower().strip(" .!") in ("stop", "cancel", "stop scrolling", "be quiet", "quiet"):
            return self.stop()
        if self.busy:
            return self.say("⏳ Still working on the last request...", "sys")
        self.busy = True
        self.say("You: " + text, "you")
        threading.Thread(target=self._command, args=(text,), daemon=True).start()

    def ask_with_progress(self, msgs):
        box = {}

        def run():
            try:
                box["r"] = ask_brain(msgs)
            except Exception as e:
                box["e"] = e
        t = threading.Thread(target=run, daemon=True)
        t.start()
        waited = 0
        while t.is_alive():
            t.join(6)
            if t.is_alive():
                waited += 6
                if STOP.is_set():
                    raise RuntimeError("Stopped")
                if waited >= 60:
                    raise TimeoutError("the AI took longer than 60 seconds")
                self.say(f"⏳ AI still thinking... {waited}s", "sys")
        if "e" in box:
            raise box["e"]
        return box["r"]

    def _command(self, text):
        try:
            handle_request(text, self.history, self.say, self.speaker.say, self.ask_with_progress)
        except Exception as e:
            log_line("COMMAND ERROR: " + traceback.format_exc())
            self.say(f"Error: {e}", "err")
        finally:
            hand_hide()
            self.busy = False

    def local(self, fn, *args):
        if self.busy:
            return self.say("⏳ Still working on the last request...", "sys")
        self.busy = True

        def work():
            STOP.clear()
            try:
                self.say("▶ " + str(W.call(fn, *args)), "act")
            except Exception as e:
                traceback.print_exc()
                self.say(f"⚠️ {e}", "err")
            finally:
                hand_hide()
                self.busy = False
        threading.Thread(target=work, daemon=True).start()


# ====================================================================== start-up
def make_dpi_aware():
    try:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def call_status(fn, *args, label="", limit=40):
    """Run an Excel job but never wait silently: print progress and give up after `limit` seconds."""
    out = queue.Queue()
    W.jobs.put((fn, args, out))
    waited = 0
    while True:
        try:
            ok, val = out.get(timeout=3)
            break
        except queue.Empty:
            waited += 3
            stage(f"still waiting ({waited}s): {label}")
            if waited >= limit:
                raise TimeoutError(f"Timed out after {limit}s while: {label}")
    if ok:
        return val
    raise val


def choice_window(root, title, text, buttons):
    """A window that is ALWAYS visible and in front (a normal dialog can hide behind Excel)."""
    win = tk.Toplevel(root)
    win.title(title)
    win.configure(bg="#1e1e1e")
    win.attributes("-topmost", True)
    result = {"v": None}

    def pick(v):
        result["v"] = v
        win.destroy()
    tk.Label(win, text=text, justify="left", wraplength=520, bg="#1e1e1e", fg="#d4d4d4",
             font=("Segoe UI", 10)).pack(padx=20, pady=(16, 10))
    bar = tk.Frame(win, bg="#1e1e1e")
    bar.pack(pady=(4, 16))
    for label, value in buttons:
        tk.Button(bar, text=label, command=lambda v=value: pick(v), relief="flat", bg="#0078d4", fg="white",
                  font=("Segoe UI", 10, "bold"), padx=12, pady=5).pack(side="left", padx=6)
    win.protocol("WM_DELETE_WINDOW", lambda: pick(None))
    win.update_idletasks()
    x = (win.winfo_screenwidth() - win.winfo_reqwidth()) // 2
    y = (win.winfo_screenheight() - win.winfo_reqheight()) // 3
    win.geometry(f"+{max(0, x)}+{max(0, y)}")
    win.lift()
    win.focus_force()
    root.wait_window(win)
    return result["v"]


def connect_to_excel(root, path):
    """Returns the workbook info, or None if the user gave up. Never hangs silently."""
    reason = ""
    while True:
        try:
            if path:
                return call_status(open_workbook, path, True, label=f"opening {path}", limit=60)
            info = call_status(open_workbook, None, False, label="looking for an open Excel workbook", limit=40)
            if info is not None:
                return info
            reason = "I could not see an open Excel workbook."
        except Exception as e:
            stage(f"connect problem: {e}")
            reason = f"Problem while connecting to Excel: {e}"
        pick = choice_window(
            root, "Excel Tutor - connect to Excel",
            reason + "\n\nIf your file IS open in Excel, then Excel and Python are probably running with "
            "different permissions (for example Excel 'as administrator'). Easiest fix: close that Excel "
            "window and click 'Choose a file' so I open it for you.",
            [("Choose a file...", "file"), ("Use sample data", "sample"), ("Try again", "retry"),
             ("Quit", None)])
        if pick is None:
            return None
        if pick == "retry":
            continue
        if pick == "sample":
            path = ""
            try:
                return call_status(open_workbook, "", True, label="creating the sample workbook", limit=60)
            except Exception as e:
                reason = f"Could not create the sample workbook: {e}"
                continue
        chosen = filedialog.askopenfilename(
            parent=root, title="Choose your Excel or CSV file",
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.xls *.csv"), ("All files", "*.*")])
        if chosen:
            path = chosen


def main():
    global W
    make_dpi_aware()
    root = tk.Tk()
    root.withdraw()                              # the panel only appears once Excel is open
    stage("starting Excel Tutor")
    path = sys.argv[1] if len(sys.argv) > 1 else ""
    if path and not os.path.exists(path):
        print(f"File not found: {path}")
        path = ""
    W = ExcelWorker()
    info = connect_to_excel(root, path)
    if info is None:
        stage("no workbook, quitting")
        return
    stage("opening the assistant window")
    try:
        app = App(root, info)
    except Exception:
        err = traceback.format_exc()
        log_line("APP START ERROR: " + err)
        stage("assistant window failed to start:\n" + err)
        choice_window(root, "Excel Tutor - error",
                      "The assistant window could not start:\n\n" + err[-700:], [("OK", "ok")])
        return
    W.on_dead = app.excel_closed
    try:                                          # Excel on the left, panel on the right
        call_status(layout_excel, 0, 0, LAYOUT["gui_left"], root.winfo_screenheight() - 70,
                    root.winfo_fpixels("1i"), label="placing the Excel window", limit=20)
    except Exception as e:
        stage(f"could not resize the Excel window (not fatal): {e}")
    stage("ready: showing the panel")
    root.deiconify()
    root.mainloop()


if __name__ == "__main__":
    main()