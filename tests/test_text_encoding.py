from pathlib import Path

from text_encoding import decode_text_bytes, read_text


def test_utf16_bom_is_decoded_before_legacy_fallbacks():
    text, encoding = decode_text_bytes("050】第一天灾-虫群\r\n正文".encode("utf-16"))

    assert text == "050】第一天灾-虫群\r\n正文"
    assert encoding == "utf-16"


def test_all_repository_txt_resources_are_stored_as_utf8():
    repo_root = Path(__file__).resolve().parents[1]
    resources = [*repo_root.joinpath("序列库").rglob("*.txt"), *repo_root.joinpath("荣誉室").rglob("*.txt")]

    assert resources
    for path in resources:
        path.read_bytes().decode("utf-8")


def test_first_disaster_title_and_content_are_readable():
    path = Path(__file__).resolve().parents[1] / "序列库/特质改造/生化改造类/050】第一天灾-虫群.txt"
    text, encoding = read_text(path)

    assert encoding == "utf-8"
    assert text.splitlines()[0] == "050】第一天灾-虫群"
    assert "【生化改造】" in text
