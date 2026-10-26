"""
IREMBOPAY CHECKOUT CHECK (step 5, Exception Resolution - every bank)
======================================================================
For an invoice not found in the IremboPay payment file, asks IremboPay's checkout what it knows about it, the way
the page https://checkout.irembopay.com/<invoice> does. The answer is only SHOWN to the reviewer ("IremboPay check"
column): it never fills a decision, and it is not saved with the decisions or in the report.

    Paid                       IremboPay says the invoice has already been paid (the channel is not given)
    Expired                    the invoice has expired
    Open, not paid yet         the invoice can still be paid
    Not an IremboPay invoice   IremboPay does not know this number
    Could not check            no answer (no internet, the site is down, an unexpected reply)

The lookups run in a background thread, up to 50 invoices at the same time and about 50 per second (26/10/2026: for
Ecobank and BPR every invoice is checked before the run), so the page stays usable. When IremboPay asks to slow down
(HTTP 429) or is busy (5xx), the lookup waits and tries again (1, 2, then 4 seconds). Each invoice is looked up once
per browser session.
"""
from __future__ import annotations

import json
import re
import threading
import time
import urllib.error
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
import urllib.request

URL = "https://checkout.irembopay.com/api-proxy/invoices/checkout/invoice/{}"
WORKERS = 50                      # invoices looked up at the same time
PAUSE = 1.0                       # at most WORKERS lookups start in this many seconds (about 50 per second)
RETRY_WAIT = (1, 2, 4)            # seconds to wait before trying again when IremboPay says "too many" or is busy
TIMEOUT = 15                      # seconds to wait for one answer
CHECKING = "Checking…"
NO_INVOICE = "No invoice number"          # a bank row without an invoice number: nothing to look up
COULD_NOT = "Could not check"
ANSWERS = {"BAD_INVOICES_PAID": "Paid", "BAD_INVOICES_PAYMENT_EXPIRED": "Expired",
           "INVOICE_NOT_FOUND": "Not an IremboPay invoice"}


def check(invoice: str) -> str:
    """One invoice -> one of the answers above (never raises)."""
    invoice = str(invoice).strip()
    if not re.fullmatch(r"[0-9A-Za-z-]{1,40}", invoice):
        return "Not an IremboPay invoice"
    req = urllib.request.Request(URL.format(invoice), headers={"Accept": "application/json",
                                                               "User-Agent": "Mozilla/5.0 (Bank Reconciliation)"})
    for wait_s in (*RETRY_WAIT, None):
        try:
            with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
                status, body = r.status, r.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as err:
            status, body = err.code, err.read().decode("utf-8", "replace")
        except Exception:                                  # no internet, timeout, DNS, ...
            status, body = None, ""
        if (status == 429 or status is None or status >= 500) and wait_s is not None:
            time.sleep(wait_s)                             # IremboPay asks to slow down, or is busy: try again
            continue
        break
    if status is None:
        return COULD_NOT
    if status == 200:
        return "Open, not paid yet"
    try:
        errors = json.loads(body).get("errors") or []
    except ValueError:
        return COULD_NOT
    for e in errors:
        if e.get("code") in ANSWERS:
            return ANSWERS[e["code"]]
    if errors and errors[0].get("detail"):
        return f"Other: {errors[0]['detail']}"
    return COULD_NOT


class Checker:
    """The lookups of one browser session: add() queues invoices, a background thread answers them one by one."""

    def __init__(self, lookup=check, pause: float = PAUSE, workers: int = WORKERS):
        self._lookup, self._pause, self._workers = lookup, pause, workers
        self._results: dict[str, str] = {}
        self._queue: list[str] = []
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None

    def add(self, invoices) -> None:
        """Queue the invoices not looked up yet (each one once), and start the thread if it is not running."""
        with self._lock:
            for inv in invoices:
                inv = str(inv).strip()
                if inv not in ("", "nan", "None", "<NA>") and inv not in self._results and inv not in self._queue:
                    self._queue.append(inv)
            if self._queue and (self._thread is None or not self._thread.is_alive()):
                self._thread = threading.Thread(target=self._run, daemon=True, name="irembopay-checkout")
                self._thread.start()

    def _run(self) -> None:
        """Up to WORKERS lookups in flight at all times (a slow answer does not hold the others back); a new one
        starts at most every PAUSE / WORKERS seconds (about 50 per second)."""
        gap = self._pause / self._workers
        running, last_start = {}, 0.0
        with ThreadPoolExecutor(max_workers=self._workers, thread_name_prefix="irembopay-check") as pool:
            while True:
                with self._lock:
                    waiting = [i for i in self._queue if i not in running.values()]
                while waiting and len(running) < self._workers:
                    delay = last_start + gap - time.monotonic()
                    if delay > 0:
                        time.sleep(delay)
                    inv = waiting.pop(0)
                    running[pool.submit(self._lookup, inv)] = inv
                    last_start = time.monotonic()
                if not running:
                    return
                done, _ = wait(list(running), timeout=0.5, return_when=FIRST_COMPLETED)
                with self._lock:
                    for future in done:
                        inv = running.pop(future)
                        try:
                            self._results[inv] = future.result()
                        except Exception:
                            self._results[inv] = COULD_NOT
                        if inv in self._queue:
                            self._queue.remove(inv)

    def answer(self, invoice) -> str:
        invoice = str(invoice).strip()
        if invoice in ("", "nan", "None", "<NA>"):
            return NO_INVOICE
        with self._lock:
            return self._results.get(invoice, CHECKING)

    def progress(self, invoices) -> tuple[int, int, int]:
        """(checked, total, could not check) for these invoices (rows without an invoice number are left out)."""
        invoices = [i for i in (str(i).strip() for i in invoices) if i not in ("", "nan", "None", "<NA>")]
        with self._lock:
            done = [self._results[i] for i in invoices if i in self._results]
        return len(done), len(invoices), sum(a == COULD_NOT for a in done)

    def retry_failed(self, invoices) -> None:
        """Look up again the invoices whose answer was 'Could not check'."""
        with self._lock:
            for i in invoices:
                i = str(i).strip()
                if self._results.get(i) == COULD_NOT:
                    del self._results[i]
        self.add(invoices)
