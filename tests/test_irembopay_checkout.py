"""The IremboPay checkout check of step 5 (Ecobank). No test reaches IremboPay: the web answer is faked."""
import io
import json
import time
import urllib.error

import pytest

from integrations import irembopay_checkout as C


def fake_http(status, payload):
    """A urlopen replacement answering `status` with `payload` (JSON)."""
    body = json.dumps(payload).encode()

    def urlopen(req, timeout=None):
        assert req.full_url.startswith("https://checkout.irembopay.com/api-proxy/invoices/checkout/invoice/")
        if status == 200:
            class R(io.BytesIO):
                status = 200
                def __enter__(self): return self
                def __exit__(self, *a): return False
            return R(body)
        raise urllib.error.HTTPError(req.full_url, status, "error", {}, io.BytesIO(body))
    return urlopen


@pytest.mark.parametrize("status, code, expected", [
    (400, "BAD_INVOICES_PAID", "Paid"),
    (400, "BAD_INVOICES_PAYMENT_EXPIRED", "Expired"),
    (404, "INVOICE_NOT_FOUND", "Not an IremboPay invoice"),
])
def test_answers_are_read_from_irembopay_codes(monkeypatch, status, code, expected):
    monkeypatch.setattr(C.urllib.request, "urlopen", fake_http(status, {"success": False,
                                                                         "errors": [{"code": code, "detail": "x"}]}))
    assert C.check("886213126169") == expected


def test_an_invoice_that_can_still_be_paid_is_open(monkeypatch):
    monkeypatch.setattr(C.urllib.request, "urlopen", fake_http(200, {"success": True, "data": {}}))
    assert C.check("886213126169") == "Open, not paid yet"


def test_no_internet_or_odd_answer_is_could_not_check(monkeypatch):
    def down(req, timeout=None):
        raise OSError("no internet")
    monkeypatch.setattr(C.urllib.request, "urlopen", down)
    assert C.check("886213126169") == C.COULD_NOT
    monkeypatch.setattr(C.urllib.request, "urlopen", fake_http(500, {"oops": True}))
    assert C.check("886213126169") == C.COULD_NOT


def test_a_value_that_is_not_an_invoice_number_is_not_sent(monkeypatch):
    def never(req, timeout=None):
        raise AssertionError("must not be called")
    monkeypatch.setattr(C.urllib.request, "urlopen", never)
    assert C.check("../../etc") == "Not an IremboPay invoice"


def wait(checker, invoices, seconds=5):
    end = time.time() + seconds
    while checker.progress(invoices)[0] < len(invoices) and time.time() < end:
        time.sleep(0.02)


def test_checker_runs_in_the_background_once_per_invoice():
    calls = []

    def lookup(inv):
        calls.append(inv)
        return "Paid"
    checker = C.Checker(lookup=lookup, pause=0)
    invoices = ["1", "2", "3"]
    checker.add(invoices)
    assert checker.answer("9") == C.CHECKING                       # not asked: still "Checking…"
    wait(checker, invoices)
    assert checker.progress(invoices) == (3, 3, 0)
    assert [checker.answer(i) for i in invoices] == ["Paid"] * 3
    checker.add(invoices + ["2"])                                   # asked again: not looked up twice
    time.sleep(0.05)
    assert sorted(calls) == ["1", "2", "3"]


def test_check_again_only_retries_the_failed_ones():
    answers = {"1": ["Paid"], "2": [C.COULD_NOT, "Expired"]}
    calls = []

    def lookup(inv):
        calls.append(inv)
        return answers[inv].pop(0)
    checker = C.Checker(lookup=lookup, pause=0)
    checker.add(["1", "2"])
    wait(checker, ["1", "2"])
    assert checker.progress(["1", "2"]) == (2, 2, 1)
    checker.retry_failed(["1", "2"])
    wait(checker, ["1", "2"])
    assert checker.answer("2") == "Expired" and calls.count("1") == 1


def test_ten_invoices_are_looked_up_at_the_same_time():
    """25 invoices taking 0.2 s each: in batches of 10 they take ~0.6 s, one by one they would take 5 s."""
    def slow(inv):
        time.sleep(0.2)
        return "Paid"
    checker = C.Checker(lookup=slow, pause=0)
    invoices = [str(i) for i in range(25)]
    started = time.time()
    checker.add(invoices)
    wait(checker, invoices, seconds=10)
    assert checker.progress(invoices) == (25, 25, 0)
    assert time.time() - started < 2


def test_rows_without_an_invoice_number_are_not_looked_up_and_do_not_block_the_progress():
    calls = []
    checker = C.Checker(lookup=lambda inv: calls.append(inv) or "Paid", pause=0)
    invoices = ["886213126169", "", "nan"]
    checker.add(invoices)
    wait(checker, invoices)
    assert checker.progress(invoices) == (1, 1, 0)                    # done: the blank rows are not counted
    assert checker.answer("") == C.NO_INVOICE and checker.answer("nan") == C.NO_INVOICE
    assert calls == ["886213126169"]
