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


def test_save_folders_sets_one_bank_and_keeps_the_others(tmp_path):
    toml = tmp_path / "folders.toml"
    config.save_folders("IM", "C:/data/I&M", "drive:1AbCdEf", toml)
    config.save_folders("ECO", "C:/data/irembo", "C:/data/Ecobank", toml)
    folders = config.load_folders(toml)
    assert folders["IM"]["statements"] == config.FolderSpec("local", "C:/data/I&M")
    assert folders["IM"]["exports"] == config.FolderSpec("drive", "1AbCdEf")
    assert folders["ECO"]["statements"] == config.FolderSpec("local", "C:/data/irembo")
    assert folders["ECO"]["exports"] == config.FolderSpec("local", "C:/data/Ecobank")


def test_save_folders_can_clear_a_bank(tmp_path):
    toml = tmp_path / "folders.toml"
    config.save_folders("IM", "C:/data/I&M", "C:/data/I&M-exports", toml)
    config.save_folders("IM", "", "", toml)
    assert config.load_folders(toml).get("IM") is None
