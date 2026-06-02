from pathlib import Path
import zipfile

from build_chm import configured_root_files, create_zip, should_include_root_file


def touch(tmp_path: Path, name: str) -> Path:
    path = tmp_path / name
    path.write_text("x", encoding="utf-8")
    return path


def test_patch_version_includes_minor_version_root_notes(tmp_path: Path):
    editor_note = touch(tmp_path, "6.5序列库编者注.txt")
    changelog = touch(tmp_path, "V6.5序列库更新日志.txt")

    assert should_include_root_file(editor_note, "v6.5.2")
    assert should_include_root_file(changelog, "v6.5.2")


def test_patch_version_excludes_other_minor_version_root_notes(tmp_path: Path):
    old_editor_note = touch(tmp_path, "6.4序列库编者注.txt")
    old_changelog = touch(tmp_path, "V6.4序列库更新日志.txt")

    assert not should_include_root_file(old_editor_note, "v6.5.2")
    assert not should_include_root_file(old_changelog, "v6.5.2")


def test_configured_root_files_apply_to_patch_version(tmp_path: Path):
    editor_note = touch(tmp_path, "6.5序列库编者注.txt")
    changelog = touch(tmp_path, "V6.5序列库更新日志.txt")
    old_editor_note = touch(tmp_path, "6.4序列库编者注.txt")
    config = {
        "include_root_files": {
            "6.5": [
                "6.5序列库编者注.txt",
                "V6.5序列库更新日志.txt",
            ]
        }
    }
    root_files_config = configured_root_files(config, "v6.5.2")

    assert should_include_root_file(editor_note, "v6.5.2", root_files_config)
    assert should_include_root_file(changelog, "v6.5.2", root_files_config)
    assert not should_include_root_file(old_editor_note, "v6.5.2", root_files_config)


def test_create_zip_uses_configured_root_files(tmp_path: Path):
    touch(tmp_path, "6.5序列库编者注.txt")
    touch(tmp_path, "V6.5序列库更新日志.txt")
    touch(tmp_path, "6.4序列库编者注.txt")
    (tmp_path / "序列库").mkdir()
    touch(tmp_path / "序列库", "001】示例.txt")
    output_path = tmp_path / "out.zip"
    root_files_config = {
        "6.5序列库编者注.txt",
        "V6.5序列库更新日志.txt",
    }

    create_zip(
        tmp_path,
        output_path,
        "v6.5.2",
        zip_content_dirs=["序列库"],
        root_files_config=root_files_config,
    )

    with zipfile.ZipFile(output_path) as zf:
        names = set(zf.namelist())

    assert "6.5序列库编者注.txt" in names
    assert "V6.5序列库更新日志.txt" in names
    assert "6.4序列库编者注.txt" not in names
    assert "序列库/001】示例.txt" in names
