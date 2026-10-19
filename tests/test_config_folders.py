"""Each bank's statements/exports folder (dianah branch: folders instead of uploads), set in folders.toml."""
import config


def test_local_and_drive_folders_are_read_per_bank(tmp_path):
    toml = tmp_path / "folders.toml"
    toml.write_text('[banks.IM]\nstatements = "C:/data/I&M"\nexports = "drive:1AbCdEf"\n\n'
                    '[banks.eco]\nstatements = ""\nexports = "C:/data/Ecobank"\n')
    folders = config.load_folders(toml)
    assert folders["IM"]["statements"] == config.FolderSpec("local", "C:/data/I&M")
    assert folders["IM"]["exports"] == config.FolderSpec("drive", "1AbCdEf")
    assert folders["ECO"]["statements"] is None                  # blank = not set
    assert folders["ECO"]["exports"] == config.FolderSpec("local", "C:/data/Ecobank")


def test_no_file_means_no_folders(tmp_path):
    assert config.load_folders(tmp_path / "missing.toml") == {}
