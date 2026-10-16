"""Step 2: the IremboPay files of the chosen bank (services/pay_files.py)."""
from types import SimpleNamespace as F

import engine as e
from services import pay_files as P

DRIVE = [F(id="1", name="payment details.csv", folder="Eco Bank / 09/2026"),
         F(id="2", name="payments_2026-08.csv", folder="I&M / 08/2026"),
         F(id="3", name="payment_details.csv", folder="I&M / 09/2026"),
         F(id="4", name="loose.csv", folder=""),
         F(id="5", name="old.csv", folder="Archive / 2025")]


def test_folders_are_recognised_whatever_the_spelling():
    for folder, code in [("Eco Bank / 09/2026", "ECO"), ("ECOBANK", "ECO"), ("ecobank / x", "ECO"),
                         ("I&M / 08/2026", "IM"), ("I&M Bank", "IM"), ("Bank of Africa / 09/2026", "BOA"),
                         ("BOA", "BOA"), ("Bank of Africa Rwanda", "BOA"), ("Archive", None), ("", None)]:
        assert P.bank_of_folder(folder, e.BANKS) == code, folder


def test_only_the_chosen_banks_files_and_the_loose_ones():
    mine, loose, has = P.files_for_bank(DRIVE, e.BANKS["ECO"], e.BANKS)
    assert [f.id for f in mine] == ["1"] and [f.id for f in loose] == ["4"] and has
    mine, loose, has = P.files_for_bank(DRIVE, e.BANKS["IM"], e.BANKS)
    assert [f.id for f in mine] == ["2", "3"]
    mine, loose, has = P.files_for_bank(DRIVE, e.BANKS["BOA"], e.BANKS)
    assert mine == [] and not has and [f.id for f in loose] == ["4"]          # no BOA folder yet
