"""Temporary files of one session (02/11/2026): two steps using a file of the same name at the same time (the
Compare report and the Dashboard's copy of the same report) never delete each other's file."""
from services.session_files import SessionFolder


def test_same_name_twice_gives_two_files_and_deleting_one_keeps_the_other(tmp_path):
    session = SessionFolder(root=tmp_path)
    name = ("month", "IM_recon_summary_01-09-2026_to_30-09-2026.xlsx")
    compare, dashboard = session.file(*name), session.file(*name)
    assert compare != dashboard and compare.name == dashboard.name        # same name, kept for the engine
    compare.write_bytes(b"compare"), dashboard.write_bytes(b"dashboard")
    assert session.remove(dashboard)
    assert compare.read_bytes() == b"compare"                              # still there for the run
    assert session.remove(compare) and session.files() == []


def test_a_report_written_in_a_private_folder_goes_with_it(tmp_path):
    session = SessionFolder(root=tmp_path)
    report = session.private() / "month" / "IM_recon_summary_01-10-2026_to_31-10-2026.xlsx"
    report.parent.mkdir(parents=True)
    report.write_bytes(b"x")
    assert session.remove(report) and session.files() == [] and not report.parent.parent.exists()


def test_nothing_outside_the_session_is_deleted(tmp_path):
    session = SessionFolder(root=tmp_path / "sessions")
    outside = tmp_path / "keep.xlsx"
    outside.write_bytes(b"x")
    session.file("a.xlsx")
    assert not session.remove(outside) and outside.exists()
    session.clear()
    assert outside.exists() and session.files() == []
