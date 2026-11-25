"""Dashboard ⚙ Figures (24/11/2026): "As reported" = the saved reports; "Latest" = plus the pending items marked
"Received later", as of the last "Update the dashboard" on the Follow-up page (here a fake Drive)."""
import pandas as pd

from services import dashboard as d
from services import dashboard_view, followup
from test_followup import FakeDrive, item


def fact(kind="month", start="2026-10-01", end="2026-10-31", bank="IM"):
    counts = {"processed": {"RWF": 100}, "received": {"RWF": 97}, "sla": {"RWF": 0}, "pending": {"RWF": 3},
              "duplicate": {"RWF": 0}}
    amounts = {"processed": {"RWF": 1_567_395_147.0}, "received": {"RWF": 1_483_016_451.0}, "sla": {"RWF": 0.0},
               "pending": {"RWF": 1_500_000.0}, "duplicate": {"RWF": 0.0}}
    return d.Fact(bank, kind, pd.Timestamp(start), pd.Timestamp(end), counts, amounts, None)


def test_latest_adds_the_items_received_later_to_the_notified_money():
    month, week = fact(), fact("week", "2026-10-05", "2026-10-11")
    items = [item("a", amount=1_200_000, paid="2026-10-08"), item("b", amount=300_000, paid="2026-10-20"),
             item("c", amount=50, paid="2026-10-08", bank="ECO")]
    dec = {"a": {"decision": "Received later"}, "b": {"decision": "Refunded"}, "c": {"decision": "Received later"}}
    new_month, new_week = d.with_followup([month, week], items, dec)
    for f in (new_month, new_week):                       # a: I&M, paid on 08/10, in the month and in its week
        assert f.amounts["received"]["RWF"] == 1_484_216_451.0
        assert f.amounts["pending"]["RWF"] == 300_000.0 and f.counts["pending"]["RWF"] == 2
        assert f.counts["received"]["RWF"] == 97 and f.counts["later"]["RWF"] == 1   # the rate stays the report's own
        assert f.amounts["processed"]["RWF"] == 1_567_395_147.0           # and the money collected
    assert month.amounts["received"]["RWF"] == 1_483_016_451.0            # the facts read are never changed
    assert d.later_amounts(items, dec, "2026-10-01", "2026-10-31", ["IM"]) == {"RWF": 1_200_000}
    assert d.with_followup([month], items, {}) == [month]


def test_the_tile_shows_3_decimals_and_the_full_amounts_on_hover():
    t = d.totals(d.with_followup([fact()], [item("a", amount=1_200_000, paid="2026-10-08")],
                                 {"a": {"decision": "Received later"}}))
    html = dashboard_view.kpi_cards(t, None, later={"RWF": 1_200_000.0})
    assert "<b>1.567B</b>" in html and "Notified 1.484B" in html
    assert t["received"] == 97 and t["later"] == 1 and d.rates(t)["receipt"] == 0.97        # the rate: the report's
    assert "<b>98</b>" in html                                                               # the Notified tile: +1
    assert "<td>Notified</td><td>1,484,216,451</td>" in html
    assert "<td>in the reports</td><td>1,483,016,451</td>" in html and "<td>received later</td><td>+1,200,000</td>" in html
    plain = dashboard_view.kpi_cards(d.totals([fact()]), None)            # As reported: no follow-up rows
    assert "Notified 1.483B" in plain and "received later" not in plain


def test_update_the_dashboard_keeps_the_received_later_items_of_that_moment():
    drive = FakeDrive()
    assert followup.load_published(drive, "root") == {}
    followup.save_decisions(drive, "root", {"a": {"decision": "Received later", "comment": ""},
                                            "b": {"decision": "Refunded", "comment": ""}}, "ann@irembo.com")
    pub = followup.publish(drive, "root", "ann@irembo.com")
    assert pub["keys"] == ["a"] and pub["by"] == "ann@irembo.com"
    # a decision saved afterwards waits for the next update; saving keeps what the dashboard shows
    followup.save_decisions(drive, "root", {"c": {"decision": "Received later", "comment": ""}}, "bob@irembo.com")
    assert followup.load_published(drive, "root")["keys"] == ["a"]
    assert set(followup.load_decisions(drive, "root")[0]) == {"a", "b", "c"}
    assert followup.publish(drive, "root", "bob@irembo.com")["keys"] == ["a", "c"]


def test_update_the_dashboard_for_the_period_and_bank_chosen_only():
    """The Follow-up page's period and bank: only their items are updated; the others keep what the dashboard had."""
    drive = FakeDrive()
    followup.save_decisions(drive, "root", {"oct_im": {"decision": "Received later", "comment": ""}}, "ann@irembo.com")
    followup.publish(drive, "root", "ann@irembo.com", scope=["oct_im"])
    followup.save_decisions(drive, "root", {"oct_im": {"decision": "", "comment": ""},           # undone in October
                                            "oct_eco": {"decision": "Received later", "comment": ""},
                                            "nov_im": {"decision": "Received later", "comment": ""}}, "bob@irembo.com")
    pub = followup.publish(drive, "root", "bob@irembo.com", scope=["nov_im"])     # November · I&M only
    assert pub["keys"] == ["nov_im", "oct_im"]                  # October stays as it was until it is updated
    assert followup.publish(drive, "root", "bob@irembo.com", scope=["oct_im", "oct_eco"])["keys"] == ["nov_im", "oct_eco"]
