"""The custom report (Results) after the reconciliation: the rows the rules decided automatically are in the table
their decision chose, so the custom report has no 'Exceptions' section, option, figure or line."""
import engine


def test_no_exceptions_section_option_or_report_type():
    assert "exceptions" not in engine.CUSTOM_SECTIONS
    assert "exceptions" not in engine.CUSTOM_INCLUDE
    assert not any("xception" in name or "xception" in desc for name, (desc, _) in engine.REPORT_TYPES.items())
    for _, (_, sections) in engine.REPORT_TYPES.items():
        assert set(sections) <= set(engine.CUSTOM_SECTIONS)        # every type only uses existing sections
    assert engine.REPORT_TYPES["Pending & Duplicates Report"][1] == {"pending": engine.TABLE,
                                                                     "duplicates": engine.TABLE,
                                                                     "other_channel": engine.TABLE}
