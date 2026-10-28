"""services/dashboard_view.py: merchant_bars() says the right reason when there is nothing to show - an old report
of this app's own (made before merchants were tracked) vs. the reconciliation notebooks, which never keep one."""
import pandas as pd

from services import dashboard_view as dv


def test_no_merchant_data_explains_an_old_report_by_default():
    html = dv.merchant_bars(pd.DataFrame(), "RWF")
    assert "made before 12/10/2026" in html


def test_no_merchant_data_explains_the_notebooks_when_every_report_is_one():
    html = dv.merchant_bars(pd.DataFrame(), "RWF", all_notebook=True)
    assert "reconciliation notebooks" in html and "made before" not in html
