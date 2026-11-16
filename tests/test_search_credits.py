from web.backend.app.search import SearchIndex


def test_creator_facets_and_filters_exclude_review_only_credits(tmp_path):
    resources = tmp_path / "序列库" / "职业"
    resources.mkdir(parents=True)
    manuscripts = {
        "001】甲.txt": "001】甲\n\n（制作人：甲）（审核人：乙）\n（复查人：丙）\n正文",
        "002】乙.txt": "002】乙\n\n（审核人：甲）（制作人：乙）\n（修改人：乙）\n正文",
        "003】丙.txt": "003】丙\n\n（审核人：丙）\n正文",
    }
    for filename, content in manuscripts.items():
        (resources / filename).write_text(content, encoding="utf-8")

    index = SearchIndex(tmp_path, ("序列库",))
    result = index.search("")
    assert result["facets"]["authors"] == [
        {"name": "乙", "count": 1}, {"name": "甲", "count": 1},
    ]
    for creator, filename in (("甲", "001】甲.txt"), ("乙", "002】乙.txt")):
        filtered = index.search("", authors=[creator], include_content=True)
        assert filtered["count"] == 1
        assert filtered["items"][0]["filename"] == filename
        assert filtered["items"][0]["content"] == manuscripts[filename]
    assert index.search("", authors=["丙"])["count"] == 0
